# pptx_redaction.py
import os
import re
import zipfile
import shutil
import logging
import xml.etree.ElementTree as ET

from redac_rules import RULES

logger = logging.getLogger("pptx_redaction")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(ch)

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _iter_slides(tmp_dir: str):
    slides_path = os.path.join(tmp_dir, "ppt", "slides")
    if not os.path.isdir(slides_path):
        return
    for fname in sorted(os.listdir(slides_path)):
        if fname.startswith("slide") and fname.endswith(".xml"):
            yield os.path.join(slides_path, fname)


def _collect_text_nodes_in_paragraph(p):
    """<a:p> 안의 <a:t>들을 순서대로 수집 -> [[node, text], ...]"""
    nodes = []
    for t in p.findall(".//a:r/a:t", NS):
        nodes.append([t, t.text or ""])
    return nodes


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
    nodes: [[node, text], ...]
    spans: [(start, end)]  # 결합 문자열 기준
    mask: 마스킹 문자(여러 글자 들어와도 첫 글자만 사용)

    동작:
      - 매칭된 길이만큼 마스킹하되, 하이픈('-')은 그대로 보존.
      - 예) '5107-3759-8931-6325' -> '****-****-****-****'
    """
    mask_char = (mask or "*")[0]

    # 각 노드의 전역 오프셋 계산
    offsets = []
    acc = 0
    for _node, txt in nodes:
        offsets.append((acc, acc + len(txt)))
        acc += len(txt)

    # 각 span을 노드 조각 단위로 치환 (하이픈 '-'은 그대로 유지)
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
                # 하이픈('-')은 보존, 그 외는 mask_char로 치환 → 길이 동일 유지
                masked = "".join(ch if ch == "-" else mask_char for ch in piece)
                nodes[i][1] = txt[:ls] + masked + txt[le:]
                # 길이 동일 → offsets 보정 불필요
            i += 1

    # XML 반영
    for node, new_text in nodes:
        node.text = new_text


def _find_matches(text: str):
    """RULES로 search + validator. 반환: [(pname, (s,e), matched), ...]"""
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


def redact_pptx(input_pptx: str, output_pptx: str, mask="*"):
    # 1) temp 초기화
    tmp_dir = "pptx_tmp"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # 2) 압축 해제
    with zipfile.ZipFile(input_pptx, "r") as z:
        z.extractall(tmp_dir)

    # 3) 슬라이드 처리
    total_spans = 0
    for slide_path in _iter_slides(tmp_dir):
        tree = ET.parse(slide_path)
        root = tree.getroot()
        changed = False

        for p in root.findall(".//a:p", NS):
            nodes = _collect_text_nodes_in_paragraph(p)
            if not nodes:
                continue
            joined = "".join(txt for _, txt in nodes)
            found = _find_matches(joined)
            if not found:
                continue

            # 겹치는 범위 병합 후 동일 길이 마스킹(하이픈 보존)
            spans = _merge_overlaps([span for _pn, span, _v in found])
            _apply_replacements_to_nodes(nodes, spans, mask=mask)
            changed = True
            total_spans += len(spans)

        if changed:
            tree.write(slide_path, encoding="utf-8", xml_declaration=True)

    logger.info("Total redacted ranges: %d", total_spans)

    # 4) ZIP 재생성 (기존 파일 정리)
    if os.path.exists("redacted.zip"):
        os.remove("redacted.zip")
    shutil.make_archive("redacted", "zip", tmp_dir)

    if os.path.exists(output_pptx):
        try:
            os.remove(output_pptx)
        except PermissionError:
            # 파워포인트에서 열려 있으면 실패하므로 사용자에게 알림
            logger.error("Output file is open. Close '%s' and run again.", output_pptx)
            raise

    shutil.move("redacted.zip", output_pptx)
    logger.info("[DONE] Saved: %s", output_pptx)


if __name__ == "__main__":
    # 사용 예시: 매칭 길이만큼 마스킹, 모든 패턴에서 '-'는 그대로 보존
    redact_pptx("demo_sensitive.pptx", "demo_redacted.pptx", mask="*")
