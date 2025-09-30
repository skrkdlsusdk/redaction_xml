# pptx_redaction.py
import os
import re
import zipfile
import shutil
import logging
import xml.etree.ElementTree as ET

# RULES 임포트 (패키지/스크립트 양쪽 지원)
try:
    from .redac_rules import RULES
except ImportError:
    from redac_rules import RULES

logger = logging.getLogger("pptx_redaction")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s"))
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
    하이픈('-')은 보존, 나머지는 mask로 동일 길이 치환
    """
    mask_char = (mask or "*")[0]

    # 전역 오프셋
    offsets = []
    acc = 0
    for _node, txt in nodes:
        offsets.append((acc, acc + len(txt)))
        acc += len(txt)

    # 각 span 치환
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

def _rezip_dir(src_dir: str, out_path: str):
    """src_dir의 '내용'을 ZIP 루트에 그대로 담아 PPTX로 재압축"""
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(src_dir):
            for fname in files:
                abs_path = os.path.join(root, fname)
                arc = os.path.relpath(abs_path, src_dir).replace(os.sep, "/")
                z.write(abs_path, arc)

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

            spans = _merge_overlaps([span for _pn, span, _v in found])
            _apply_replacements_to_nodes(nodes, spans, mask=mask)
            changed = True
            total_spans += len(spans)

        if changed:
            tree.write(slide_path, encoding="utf-8", xml_declaration=True)

    logger.info("Total redacted ranges: %d", total_spans)

    # 4) 안전 재압축
    _rezip_dir(tmp_dir, output_pptx)
    shutil.rmtree(tmp_dir, ignore_errors=True)
    logger.info("[DONE] Saved: %s", output_pptx)

def _is_candidate(fname: str) -> bool:
    """배치 후보 필터"""
    if not fname.lower().endswith(".pptx"):
        return False
    if fname.startswith("~$"):  # PowerPoint 임시파일 제외
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
        dst = args[1] if len(args) >= 2 else "output_redacted.pptx"
        redact_pptx(src, dst, mask="*")
    else:
        # 배치: 현재 폴더의 모든 .pptx 처리
        files = [f for f in os.listdir(".") if _is_candidate(f)]
        if not files:
            print("현재 폴더에 처리할 PPTX가 없습니다.")
            raise SystemExit(0)
        for f in files:
            base, ext = os.path.splitext(f)
            out = f"{base}_redacted{ext}"
            try:
                print(f"[PPTX] {f} → {out}")
                redact_pptx(f, out, mask="*")
            except Exception as e:
                print(f"[ERROR] {f}: {e}")
