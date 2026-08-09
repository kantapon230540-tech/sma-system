"""Export an assessment back into the original check-sheet workbook, byte-for-byte.

Rebuilding the workbook with openpyxl loses 29 of its 50 parts — both radar charts'
styling, printer settings, shared strings, custom XML — so the result never matches
what the site actually uses. Instead the real workbook IS the template: we rewrite
only the Judgement/Reason <c> cells inside the one sheet XML that needs them and copy
every other part through untouched. The (Ref.)Dashboard formulas and both radars keep
working because nothing they depend on was touched.
"""
import io, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XMLNS_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("", NS)

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "export_templates"

# The sheet's dropdown allows exactly Yes / No / N/A. 'not_rolled_out' has no slot,
# so it is left blank and said in the Reason column rather than forced into N/A.
JUDGEMENT = {"yes": "Yes", "no": "No", "na": "N/A", "not_rolled_out": None}

# per type: template file, and per pillar the sheet name + judgement/reason columns
LAYOUTS = {
    "warehouse": {
        "template": "warehouse_checksheet.xlsx",
        "sheets": {
            "leadership":    ("Leadership",     "K", "L"),
            "tm_engagement": ("TM Engagement",  "K", "L"),
            "organization":  ("Organization",   "K", "L"),
            "system":        ("System",         "L", "M"),
        },
        # column holding the question No, per sheet
        "no_col": {"Leadership": "D", "TM Engagement": "D", "Organization": "D", "System": "E"},
    },
}

def _col_idx(col):
    n = 0
    for ch in col: n = n * 26 + (ord(ch) - 64)
    return n

def _split(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))

def _sheet_paths(z):
    wb   = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid  = {r.get("Id"): r.get("Target") for r in rels}
    out  = {}
    for sh in wb.find(f"{{{NS}}}sheets"):
        t = rid[sh.get(f"{{{R}}}id")]
        out[sh.get("name")] = "xl/" + t.lstrip("/").replace("xl/", "", 1)
    return out

def _cell_text(row, col):
    """Plain text of a cell, resolving nothing — the No column is numeric/inline."""
    for c in row.findall(f"{{{NS}}}c"):
        if _split(c.get("r"))[0] != col: continue
        v = c.find(f"{{{NS}}}v")
        if v is not None and v.text: return v.text.strip()
        t = c.find(f"{{{NS}}}is/{{{NS}}}t")
        if t is not None and t.text: return t.text.strip()
    return None

def _patch_sheet(xml_bytes, no_col, edits_by_no):
    """edits_by_no: {question_no: {col: value}} — value None clears the cell."""
    root = ET.fromstring(xml_bytes)
    data = root.find(f"{{{NS}}}sheetData")
    written = 0
    for row in data.findall(f"{{{NS}}}row"):
        raw = _cell_text(row, no_col)
        if not raw or not raw.isdigit(): continue
        edits = edits_by_no.get(int(raw))
        if not edits: continue
        for col, val in edits.items():
            ref = f"{col}{row.get('r')}"
            cell = next((c for c in row.findall(f"{{{NS}}}c") if c.get("r") == ref), None)
            if cell is None:
                cell = ET.Element(f"{{{NS}}}c", {"r": ref})
                pos = 0
                for i, c in enumerate(row.findall(f"{{{NS}}}c")):
                    if _col_idx(_split(c.get("r"))[0]) < _col_idx(col): pos = i + 1
                row.insert(pos, cell)
            for ch in list(cell): cell.remove(ch)      # drop old value, keep @s (style)
            if val in (None, ""):
                cell.attrib.pop("t", None); continue
            cell.set("t", "inlineStr")
            t = ET.SubElement(ET.SubElement(cell, f"{{{NS}}}is"), f"{{{NS}}}t")
            t.text = str(val)
            if str(val) != str(val).strip(): t.set(XMLNS_SPACE, "preserve")
            written += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), written

def supported(assessment_type):
    return assessment_type in LAYOUTS

def build(assessment_type, pillars, responses):
    """Return (bytes, cells_written). `responses` is {question_id: {answer, comment}}."""
    layout = LAYOUTS.get(assessment_type)
    if not layout: raise ValueError(f"no check-sheet template for {assessment_type!r}")
    path = TEMPLATE_DIR / layout["template"]
    if not path.exists(): raise FileNotFoundError(f"missing export template {path}")

    # {sheet name: {question_no: {col: value}}}
    plan, total = {}, 0
    for pillar in pillars:
        conf = layout["sheets"].get(pillar["id"])
        if not conf: continue
        sheet, jcol, rcol = conf
        per = plan.setdefault(sheet, {})
        for element in pillar["elements"]:
            for q in element["questions"]:
                r = responses.get(q["id"]) or {}
                ans = (r.get("answer") or "").strip()
                if not ans and not (r.get("comment") or "").strip(): continue
                note = (r.get("comment") or "").strip()
                if ans == "not_rolled_out":
                    note = ("Not rolled out" + (" — " + note if note else ""))
                per[int(q["no"])] = {jcol: JUDGEMENT.get(ans), rcol: note or None}
                total += 1

    with zipfile.ZipFile(path) as z:
        paths = _sheet_paths(z)
        patched = {}
        for sheet, edits in plan.items():
            target = paths.get(sheet)
            if not target: continue
            patched[target], _ = _patch_sheet(z.read(target), layout["no_col"][sheet], edits)
        # The template ships with cached formula results from an EMPTY sheet. Without this
        # Excel can show those stale blanks — dashboard at 0, radars flat — until something
        # forces a recalc. fullCalcOnLoad makes it recompute the moment the file opens.
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wbxml:
            wbxml = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml, count=1)
            patched["xl/workbook.xml"] = wbxml.encode("utf-8")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, patched.get(item.filename) or z.read(item.filename))
    return buf.getvalue(), total
