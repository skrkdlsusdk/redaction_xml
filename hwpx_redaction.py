# hwpx_redaction.py
import os
import re
import zipfile
import shutil
import logging
import xml.etree.ElementTree as ET

# RULES 불러오기 (패키지/스크립트 둘 다 지원)
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

# 알려진 HWPX 네임스페이스 (방어적 대응)
NS_LIST = [
    "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "http://www.hancom.co.kr/hwpml/2011/wordprocessor",
    "http://www.hancom.co.kr/hwpml/2011/shared",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
]

# 하이픈/대시(보존 대상): - ‐ - ‒ – — ― −
KEEP = set("-\u2010\u2011\u2012\u2013\u2014\u2015\u2212")

def _local(tag: str) -> str:
    """{ns}local 형태에서 local만 추출"""
    return tag.split("}", 1)[1] if "}" in tag else tag

def _iter_hwpx_xml_files(tmp_dir: str):
    """
    HWPX는 ZIP 내부 Contents/section*.xml 등에 텍스트가 존재.
    안전하게 Contents/ 하위 모든 .xml 순회 (없으면 루트 폴더 폴백)
    """
    contents = os.path.join(tmp_dir, "Contents")
    if not os.path.isdir(contents):
        contents = tmp_dir
    for root, _dirs, files in os.walk(contents):
        for f in files:
            if f.lower().endswith(".xml"):
                yield os.path.join(root, f)

def _collect_paragraph_nodes(tree: ET.ElementTree):
    """
    문단(<*p>) 단위로 텍스트 런(<*t>, <*text>) 수집.
    반환: 문단 리스트. 각 문단은 [ [node, text], ... ]
    """
    root = tree.getroot()
    paragraphs = []

    # 문단 기준 수집
    for p in root.iter():
        if _local(p.tag) != "p":
            continue
        nodes = []
        for el in p.iter():
            lname = _local(el.tag)
            if lname in ("t", "text") and el.text is not None:
                nodes.append([el, el.text])
        if nodes:
            paragraphs.append(nodes)

    # 문단이 전혀 없으면 파일 전체에서 t/text를 하나의 문단으로 간주(방어)
    if not paragraphs:
        nodes = []
        for el in root.iter():
            lname = _local(el.tag)
            if lname in ("t", "text") and el.text is not None:
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
    정책: 매칭된 길이만큼 마스킹하되, 하이픈/대시는 보존.
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
                # 하이픈/대시/공백 보존, 나머지는 마스킹
                masked = "".join(ch if (ch in KEEP or ch.isspace()) else mask_char for ch in piece)
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

def _rezip_dir(src_dir: str, out_path: str):
    """src_dir의 '내용'을 ZIP 루트에 그대로 담아 HWPX로 재압축"""
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(src_dir):
            for fname in files:
                abs_path = os.path.join(root, fname)
                arc = os.path.relpath(abs_path, src_dir).replace(os.sep, "/")
                z.write(abs_path, arc)

def redact_hwpx(input_hwpx: str, output_hwpx: str, mask="*"):
    """
    HWPX 레닥션:
      - Contents 폴더의 모든 XML에서 문단(<*p>) 탐색
      - 문단 내 텍스트 런(<*t>, <*text>)을 결합 → RULES로 후보 식별 → 길이 유지 마스킹(하이픈/대시 보존)
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

    # 3) 안전 재압축(루트 구조 보존)
    _rezip_dir(tmp_dir, output_hwpx)
    shutil.rmtree(tmp_dir, ignore_errors=True)
    logger.info("[DONE] Saved: %s", output_hwpx)

def _is_candidate(fname: str) -> bool:
    """배치 후보 필터"""
    if not fname.lower().endswith(".hwpx"):
        return False
    base, _ = os.path.splitext(fname)
    if base.lower().endswith("_redacted"):
        return False
    return os.path.isfile(fname)

if __name__ == "__main__":
    import sys
    args = sys.argv[1:]
    if len(args) >= 1:
        # 단일 파일
        src = args[0]
        dst = args[1] if len(args) >= 2 else "output_redacted.hwpx"
        redact_hwpx(src, dst, mask="*")
    else:
        # 배치: 현재 폴더의 모든 .hwpx 처리
        files = [f for f in os.listdir(".") if _is_candidate(f)]
        if not files:
            print("현재 폴더에 처리할 HWPX가 없습니다.")
            raise SystemExit(0)
        for f in files:
            base, ext = os.path.splitext(f)
            out = f"{base}_redacted{ext}"
            try:
                print(f"[HWPX] {f} → {out}")
                redact_hwpx(f, out, mask="*")
            except Exception as e:
                print(f"[ERROR] {f}: {e}")
