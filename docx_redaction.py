# docx_redaction_lxml.py
import os, re, zipfile, shutil, logging
from lxml import etree as LET
from redac_rules import RULES

logger = logging.getLogger("docx_redaction")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s"))
    logger.addHandler(ch)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

KEEP = set("-\u2010\u2011\u2012\u2013\u2014\u2015\u2212")  # 모든 하이픈/대시

def _iter_doc_parts(tmp_dir: str):
    word = os.path.join(tmp_dir, "word")
    if not os.path.isdir(word): return
    def _y(name):
        p = os.path.join(word, name)
        if os.path.exists(p): yield p
    yield from _y("document.xml")
    for fn in sorted(os.listdir(word)):
        if fn.startswith("header") and fn.endswith(".xml"): yield os.path.join(word, fn)
        if fn.startswith("footer") and fn.endswith(".xml"): yield os.path.join(word, fn)
    yield from _y("comments.xml")
    yield from _y("footnotes.xml")
    yield from _y("endnotes.xml")

def _collect_text_nodes_in_paragraph(p):
    nodes = []
    for t in p.xpath(".//w:r/w:t", namespaces=NS):
        txt = t.text if t.text is not None else ""
        nodes.append([t, txt])
    return nodes

def _merge_overlaps(spans):
    if not spans: return []
    spans.sort()
    merged = [spans[0]]
    for s,e in spans[1:]:
        ps,pe = merged[-1]
        if s <= pe: merged[-1] = (ps, max(pe, e))
        else: merged.append((s,e))
    return merged

def _apply_replacements_to_nodes(nodes, spans, mask="*"):
    mask_char = (mask or "*")[0]
    # 전역 오프셋
    offs, acc = [], 0
    for _n, txt in nodes:
        offs.append((acc, acc+len(txt))); acc += len(txt)
    # 치환
    for s,e in spans:
        i = 0
        while i < len(nodes) and e > s:
            node, txt = nodes[i]
            ns, ne = offs[i]
            if ne <= s: i += 1; continue
            if ns >= e: break
            ls = max(ns, s)-ns
            le = max(0, min(ne, e)-ns)
            if ls < le:
                piece = txt[ls:le]
                # 하이픈/공백 보존, 나머지는 마스킹
                masked = "".join(ch if (ch in KEEP or ch.isspace()) else mask_char for ch in piece)
                new_text = txt[:ls] + masked + txt[le:]
                nodes[i][1] = new_text
            i += 1
    # XML 반영: 공백 보존, 빈 텍스트 방지(셀프클로징 방지)
    for node, new_text in nodes:
        if new_text == "":
            new_text = " "
        if new_text.startswith(" ") or new_text.endswith(" ") or ("\u00A0" in new_text):
            node.set(XML_SPACE, "preserve")
        node.text = new_text

def _find_matches(text: str):
    out = []
    for pname, rule in RULES.items():
        comp, validator = rule["regex"], rule["validator"]
        for m in comp.finditer(text):
            val = m.group(0)
            try:
                if validator(val):
                    out.append((pname, (m.start(), m.end()), val))
                    logger.debug("[MATCH] %s '%s' %s", pname, val, (m.start(), m.end()))
            except Exception as e:
                logger.debug("[VALIDATOR ERROR] %s value='%s' err=%s", pname, val, e)
    return out

def redact_docx(input_docx: str, output_docx: str, mask="*"):
    tmp_dir = "docx_tmp"
    if os.path.exists(tmp_dir): shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    with zipfile.ZipFile(input_docx, "r") as z:
        z.extractall(tmp_dir)

    parser = LET.XMLParser(remove_blank_text=False, resolve_entities=False, strip_cdata=False)

    total = 0
    for xml_path in _iter_doc_parts(tmp_dir):
        try:
            with open(xml_path, "rb") as f:
                data = f.read()
            root = LET.fromstring(data, parser=parser)
            changed = False
            for p in root.xpath(".//w:p", namespaces=NS):
                nodes = _collect_text_nodes_in_paragraph(p)
                if not nodes: continue
                joined = "".join(txt for _, txt in nodes)
                found = _find_matches(joined)
                if not found: continue
                spans = _merge_overlaps([span for _pn, span, _v in found])
                _apply_replacements_to_nodes(nodes, spans, mask=mask)
                changed = True
                total += len(spans)
            if changed:
                xml_bytes = LET.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=None,   # 원본 선언 유지 성향
                    pretty_print=False # 포맷 보존
                )
                with open(xml_path, "wb") as f:
                    f.write(xml_bytes)
        except Exception as e:
            logger.exception("Processing error in %s: %s", xml_path, e)

    logger.info("Total redacted ranges: %d", total)

    if os.path.exists(output_docx):
        try: os.remove(output_docx)
        except PermissionError:
            logger.error("Output is open. Close '%s' and run again.", output_docx)
            raise

    with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for folder, _dirs, files in os.walk(tmp_dir):
            for f in files:
                fp = os.path.join(folder, f)
                arc = os.path.relpath(fp, tmp_dir)
                z.write(fp, arc)
    logger.info("[DONE] Saved: %s", output_docx)

if __name__ == "__main__":
    redact_docx("demo_sensitive.docx", "demo_redacted.docx", mask="*")
