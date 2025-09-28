# xlsx_redaction.py
import os
import re
import zipfile
import shutil
import logging
import xml.etree.ElementTree as ET

# RULES는 redac_rules.py의 것 사용 (방법 B와 호환)
try:
    from .redac_rules import RULES
except ImportError:
    from redac_rules import RULES

logger = logging.getLogger("xlsx_redaction")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(ch)

# Excel main namespace
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# -----------------------------
# 공통 유틸 (PPTX 버전과 유사)
# -----------------------------
def _merge_overlaps(spans):
    """겹치는 (s,e) 구간 병합"""
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
    nodes: [[node, text], ...]   # node.text 를 가진 텍스트 노드들의 리스트
    spans: [(start, end)]        # 결합 문자열 기준 (start 포함, end 제외)
    mask: 마스킹 문자 (첫 글자만 사용)

    정책:
      - 매칭된 길이만큼 마스킹하되, 하이픈('-')은 그대로 보존
      - 전체 길이는 유지 → 오프셋 보정 불필요
    """
    mask_char = (mask or "*")[0]

    # 전역 오프셋
    offsets = []
    acc = 0
    for _node, txt in nodes:
        offsets.append((acc, acc + len(txt)))
        acc += len(txt)

    # 각 span을 노드 조각 단위로 치환
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
                logger.debug("[MATCH] pattern=%s text='%s' span=%s", pname, val, (m.start(), m.end()))
    return matches


# -----------------------------
# 엑셀(XML) 텍스트 수집 유틸
# -----------------------------
def _collect_nodes_shared_string(si):
    """
    sharedStrings.xml의 <si> 하나에서 텍스트 노드 모으기.
    - 단순 <t>
    - 리치 텍스트 <r>/<t> 시퀀스
    반환: [[node, text], ...]
    """
    nodes = []
    # 리치 텍스트(run) 우선
    runs = si.findall("./s:r", NS)
    if runs:
        for r in runs:
            t = r.find("./s:t", NS)
            if t is not None:
                nodes.append([t, t.text or ""])
        return nodes
    # 단일 t
    t = si.find("./s:t", NS)
    if t is not None:
        nodes.append([t, t.text or ""])
    return nodes


def _collect_nodes_inline_str(cell):
    """
    시트 XML에서 inlineStr 형태의 텍스트 수집
    - <c t="inlineStr"><is><t>...</t> 또는 <is><r><t>...</t>...
    반환: [[node, text], ...]
    """
    nodes = []
    is_el = cell.find("./s:is", NS)
    if is_el is None:
        return nodes

    runs = is_el.findall("./s:r", NS)
    if runs:
        for r in runs:
            t = r.find("./s:t", NS)
            if t is not None:
                nodes.append([t, t.text or ""])
        return nodes

    t = is_el.find("./s:t", NS)
    if t is not None:
        nodes.append([t, t.text or ""])
    return nodes


# -----------------------------
# 처리기: sharedStrings.xml
# -----------------------------
def _process_shared_strings(tmp_dir: str, mask="*") -> int:
    """
    sharedStrings.xml 내부 텍스트 마스킹
    반환: 처리된 span(매치 그룹) 개수 총합
    """
    sst_path = os.path.join(tmp_dir, "xl", "sharedStrings.xml")
    if not os.path.exists(sst_path):
        return 0

    tree = ET.parse(sst_path)
    root = tree.getroot()
    changed = False
    total_spans = 0

    for si in root.findall("./s:si", NS):
        nodes = _collect_nodes_shared_string(si)
        if not nodes:
            continue
        joined = "".join(txt for _, txt in nodes)
        found = _find_matches(joined)
        if not found:
            continue

        spans = _merge_overlaps([span for _pn, span, _v in found])
        _apply_replacements_to_nodes(nodes, spans, mask=mask)
        changed = True
        total_spans += len(spans)

    if changed:
        tree.write(sst_path, encoding="utf-8", xml_declaration=True)
        logger.info("[sharedStrings] redacted groups: %d", total_spans)
    return total_spans


# -----------------------------
# 처리기: 각 워크시트 (inlineStr)
# -----------------------------
def _process_sheets_inline(tmp_dir: str, mask="*") -> int:
    """
    xl/worksheets/sheet*.xml 내부의 inlineStr 텍스트 마스킹
    (sharedStrings 인덱스를 참조하는 셀(t="s")은 sharedStrings 처리가 담당)
    """
    ws_dir = os.path.join(tmp_dir, "xl", "worksheets")
    if not os.path.isdir(ws_dir):
        return 0

    total_spans = 0
    for fname in sorted(os.listdir(ws_dir)):
        if not fname.startswith("sheet") or not fname.endswith(".xml"):
            continue
        fpath = os.path.join(ws_dir, fname)
        tree = ET.parse(fpath)
        root = tree.getroot()
        changed = False

        # 모든 셀 순회
        for c in root.findall(".//s:c", NS):
            t_attr = c.get("t")
            if t_attr != "inlineStr":
                continue
            nodes = _collect_nodes_inline_str(c)
            if not nodes:
                continue

            joined = "".join(txt for _, txt in nodes)
            found = _find_matches(joined)
            if not found:
                continue

            spans = _merge_overlaps([span for _pn, span, _v in found])
            _apply_replacements_to_nodes(nodes, spans, mask=mask)
            changed = True
            total_spans += len(spans)

        if changed:
            tree.write(fpath, encoding="utf-8", xml_declaration=True)
            logger.info("[sheet] %s redacted groups: %d", fname, total_spans)

    return total_spans


# -----------------------------
# 공개 함수
# -----------------------------
def redact_xlsx(input_xlsx: str, output_xlsx: str, mask="*"):
    """
    .xlsx 레닥션:
      - sharedStrings.xml의 모든 문자열
      - 각 시트의 inlineStr 문자열
    에 대해 RULES 기반 탐지 → 숫자/문자 마스킹(하이픈 '-' 보존)
    """
    tmp_dir = "xlsx_tmp"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # 해제
    with zipfile.ZipFile(input_xlsx, "r") as z:
        z.extractall(tmp_dir)

    # 처리
    total = 0
    total += _process_shared_strings(tmp_dir, mask=mask)
    total += _process_sheets_inline(tmp_dir, mask=mask)
    logger.info("Total redacted groups: %d", total)

    # 재압축
    if os.path.exists("redacted.zip"):
        os.remove("redacted.zip")
    shutil.make_archive("redacted", "zip", tmp_dir)

    # 대상 파일 치환
    if os.path.exists(output_xlsx):
        try:
            os.remove(output_xlsx)
        except PermissionError:
            logger.error("Output file is open. Close '%s' and run again.", output_xlsx)
            raise

    shutil.move("redacted.zip", output_xlsx)
    logger.info("[DONE] Saved: %s", output_xlsx)


if __name__ == "__main__":
    # 사용 예시
    # 하이픈('-')은 보존하고, 매칭된 길이만큼 mask('*')로 치환
    redact_xlsx("demo_sensitive.xlsx", "demo_redacted.xlsx", mask="*")
