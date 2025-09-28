# hwpx_redaction.py
import os
import re
import zipfile
import shutil
import logging
import xml.etree.ElementTree as ET

# RULES 불러오기 (방법 B: 패키지/스크립트 모두 지원)
try:
    from .redac_rules import RULES
except ImportError:
    from redac_rules import RULES

logger = logging.getLogger("hwpx_redaction")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s"))
    logger.addHandler(ch)

# 알려진 HWPX 네임스페이스 (폭넓게 대응)
NS_LIST = [
    "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "http://www.hancom.co.kr/hwpml/2011/wordprocessor",
    "http://www.hancom.co.kr/hwpml/2011/shared",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",  # 방어적
]

def _local(tag: str) -> str:
    """{ns}local 형태에서 local만 추출"""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag

def _iter_hwpx_xml_files(tmp_dir: str):
    """
    HWPX는 ZIP 안에 Contents/section*.xml 등에 문단/런 텍스트가 들어있음.
    기본적으로 Contents 아래 모든 .xml을 순회.
    """
    contents = os.path.join(tmp_dir, "Contents")
    if not os.path.isdir(contents):
        # 일부 도구는 root 바로 아래 둘 수도 있으므로 폴백
        contents = tmp_dir
    for root, _dirs, files in os.walk(contents):
        for f in files:
            if f.lower().endswith(".xml"):
                yield os.path.join(root, f)

def _collect_paragraph_nodes(tree: ET.ElementTree):
    """
    문서에서 문단(<*p>) 단위로 텍스트 런(<*t> 또는 <*text>) 노드를 수집.
    반환: [ [ [node, text], ... ] , ... ]  # 문단들
    """
    root = tree.getroot()
    paragraphs = []

    # 모든 요소를 순회하며 localname이 'p'인 문단을 찾음
    for p in root.iter():
        if _local(p.tag) != "p":
            continue

        # 문단 내부에서 텍스트 런 수집 (순서대로)
        nodes = []
        for el in p.iter():
            lname = _local(el.tag)
            if lname in ("t", "text") and el.text:
                nodes.append([el, el.text])
        if nodes:
            paragraphs.append(nodes)

    # 문단을 못 찾았으면, 파일별로 단순히 모든 t/text를 하나의 문단으로 취급 (방어)
    if not paragraphs:
        nodes = []
        for el in root.iter():
            lname = _local(el.tag)
            if lname in ("t", "text") and el.text:
                nodes.append([el, el.text])
        if nodes:
            paragraphs.append(nodes)

    return paragraphs

def _merge_overlaps(spans):
    """겹치는 (s,e) 병합"""
    if not spans:
        return []
    spans = sorted(spans)
    merged = [spans[0]]
    for s, e in spans[1:]:
        ps, pe = merged[-1]
        if s <= pe:
            merged[-1] = (ps, max(pe, e))
        else:
            merged.append((s, e))
    return merged

def _apply_replacements_to_nodes(nodes, spans, mask="*"):
    """
    nodes: [[node, text], ...]
    spans: [(start, end)]   # 결합 문자열 기준
    정책: 매칭된 길이만큼 마스킹하되, 하이픈('-')은 그대로 보존.
    """
    mask_char = (mask or "*")[0]

    # 전역 오프셋
    offsets = []
    acc = 0
    for _node, txt in nodes:
        offsets.append((acc, acc + len(txt)))
        acc += len(txt)

    # 치환
    for s, e in spans:
        i = 0
        while i < len(nodes) and e > s:
            node, txt = nodes[i]
            ns, ne = offsets[i]
            if ne <= s:
                i += 1
                continue
            if ns >= e:
                break
            ls = max(ns, s) - ns
            le = max(0, min(ne, e) - ns)
            if ls < le:
                piece = txt[ls:le]
                masked = "".join(ch if ch == "-" else mask_char for ch in piece)
                nodes[i][1] = txt[:ls] + masked + txt[le:]
            i += 1

    # 반영
    for node, new_text in nodes:
        node.text = new_text

def _find_matches(text: str):
    """RULES로 search + validator → [(pname, (s,e), matched), ...]"""
    matches = []
    for pname, rule in RULES.items():
        comp = rule["regex"]
        validator = rule["validator"]
        for m in comp.finditer(text):
            val = m.group(0)
            ok = False
            try:
                ok = bool(validator(val))
            except Exception as e:
                logger.debug("[VALIDATOR ERROR] %s value='%s' err=%s", pname, val, e)
                ok = False
            if ok:
                matches.append((pname, (m.start(), m.end()), val))
                logger.debug("[MATCH] %s '%s' span=%s", pname, val, (m.start(), m.end()))
    return matches

def redact_hwpx(input_hwpx: str, output_hwpx: str, mask="*"):
    """
    HWPX 레닥션:
      - Contents 폴더의 모든 XML에서 문단(<*p>)을 찾아 텍스트 런(<*t>, <*text>) 결합
      - RULES 기반 탐지 → 숫자/문자 마스킹(하이픈 '-' 보존)
    """
    tmp_dir = "hwpx_tmp"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # 1) 해제
    with zipfile.ZipFile(input_hwpx, "r") as z:
        z.extractall(tmp_dir)

    total_spans = 0
    changed_files = 0

    # 2) XML 처리
    for xml_path in _iter_hwpx_xml_files(tmp_dir):
        try:
            tree = ET.parse(xml_path)
        except ET.ParseError:
            continue  # 이미지/미디어 등 무시

        paragraphs = _collect_paragraph_nodes(tree)
        if not paragraphs:
            continue

        file_changed = False

        for nodes in paragraphs:
            joined = "".join(txt for _, txt in nodes)
            found = _find_matches(joined)
            if not found:
                continue
            spans = _merge_overlaps([span for _pn, span, _v in found])
            _apply_replacements_to_nodes(nodes, spans, mask=mask)
            total_spans += len(spans)
            file_changed = True

        if file_changed:
            tree.write(xml_path, encoding="utf-8", xml_declaration=True)
            changed_files += 1

    logger.info("[HWPX] files changed=%d, total redacted groups=%d", changed_files, total_spans)

    # 3) 재압축
    if os.path.exists("redacted.zip"):
        os.remove("redacted.zip")
    shutil.make_archive("redacted", "zip", tmp_dir)

    if os.path.exists(output_hwpx):
        try:
            os.remove(output_hwpx)
        except PermissionError:
            logger.error("Output file is open. Close '%s' and run again.", output_hwpx)
            raise

    shutil.move("redacted.zip", output_hwpx)
    logger.info("[DONE] Saved: %s", output_hwpx)

if __name__ == "__main__":
    # 사용 예시: 하이픈('-') 보존, 매칭 길이만큼 '*'
    redact_hwpx("demo_sensitive.hwpx", "demo_redacted.hwpx", mask="*")
