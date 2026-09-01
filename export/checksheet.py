"""Export an assessment back into the original check-sheet workbook, byte-for-byte.

Rebuilding the workbook with openpyxl loses 29 of its 50 parts — both radar charts'
styling, printer settings, shared strings, custom XML — so the result never matches
what the site actually uses. Instead the real workbook IS the template: we rewrite
only the Judgement/Reason <c> cells inside the one sheet XML that needs them and copy
every other part through untouched. The (Ref.)Dashboard formulas and both radars keep
working because nothing they depend on was touched.
"""
import io, json, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XMLNS_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("", NS)

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "export_templates"

# Judgement vocabulary, per layout. 'not_rolled_out' is the one that differs, and it is
# NOT cosmetic: every element on these sheets scores only once every question in it carries
# a value ("=IF(SUM(C17:Q17)=R17,"done","to be done")" on the warehouse Calculation sheet,
# the same gate the MFG '2nd Cal' sheets use). Leaving a cell blank therefore drops the
# whole element to "-" and blanks the dashboard, however many other questions were answered.
JUDGEMENT = {"yes": "Yes", "no": "No", "na": "N/A", "not_rolled_out": None}

# The warehouse sheet's own dropdown offers "Yes, No, N/A, Not rolled out yet", and its
# Calculation sheet counts that literal into the same "Other" bucket as N/A
# (=COUNTIFS(...,"N/A")+COUNTIFS(...,"Not rolled out yet")). So writing it is the sheet's
# own vocabulary, keeps the element complete, and still leaves the answer out of the
# Yes/No ratio — it does not count against the site.
WH_JUDGEMENT = {**JUDGEMENT, "not_rolled_out": "Not rolled out yet"}

# Verbatim from the warehouse workbook's own dropdown source, 'Calculation sheet'!D2:D3 —
# the cell is validated against that list, so the wording cannot be paraphrased.
ASSESSMENT_TYPE_LABEL = {"self": "Self-assessment セルフアセスメント",
                         "validation": "Validation assessment 妥当性評価"}


def _assessor_names(assessment):
    """Names actually recorded on this assessment, newest field first. Never falls back to
    the default roster — an exported sheet must not name people who were not entered."""
    raw = assessment.get("assessors_json")
    if raw:
        try:
            names = [str(n).strip() for n in json.loads(raw) if str(n).strip()]
            if names: return ", ".join(names)
        except (ValueError, TypeError):
            pass
    return ", ".join(n.strip() for n in (assessment.get("assessor_a"),
                                         assessment.get("assessor_b")) if (n or "").strip())

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
        "judgement": WH_JUDGEMENT,
        # No Cover Sheet in this workbook — the single (Ref.)Dashboard carries the identity
        # block, and it is the page the site actually reads. Labels C4:E4 etc. are merged,
        # so the values belong in F. F7's wording must match the sheet's own dropdown
        # ('Calculation sheet'!D2:D3) or the cell falls outside its validation list.
        "identity": [{"sheet": "(Ref.)Dashboard", "site": "F4", "date": "F5",
                      "assessors": "F6", "kind": "F7"}],
    },
    # Att-1 "Result of safety maturity validation assessment". Judgement/Reason are Y/Z
    # on the first three sheets and Z/AA on System — that sheet is shifted one column
    # right, the same offset its U/V/W criteria columns carry. Read from the header row,
    # not assumed. Serves plain `manufacturing` assessments and the SMA track of SSDPMA.
    "manufacturing": {
        "template": "mfg_checksheet.xlsx",
        "sheets": {
            "leadership":    ("Leadership Checklist",          "Y", "Z"),
            "tm_engagement": ("Teammate Engagement Checklist",  "Y", "Z"),
            "organization":  ("Organization Checklist",         "Y", "Z"),
            "system":        ("System Checklist",               "Z", "AA"),
        },
        "no_col": {"Leadership Checklist": "F", "Teammate Engagement Checklist": "F",
                   "Organization Checklist": "F", "System Checklist": "G"},
        # Cover Sheet cells filled from the assessment record; the site's own injury
        # figures and previous-validation scores are left blank for them to complete.
        "cover": {"sheet": "Cover Sheet", "site": "D6", "year": "H6", "date": "J6",
                  "assessors": ["B9", "B10", "B11"]},
        # The two dashboard sheets carry their OWN Site/Date inputs — the labels in C4/C5
        # are merged C:E, so the value belongs in F. They are not formulas off the Cover
        # Sheet, so without this the summary the site actually reads is anonymous.
        # F6 is "Assessment type", and the master ships it hardcoded to
        # "Validation assessment 妥当性評価" — so a SELF-assessment exported without this
        # would carry two summary pages labelled Validation. It is written from the record.
        "identity": [{"sheet": "(Ref.)MA Dashboard",             "site": "F4", "date": "F5",
                      "kind": "F6"},
                     {"sheet": "(Ref.)MA Dashboard (DP_divide)", "site": "F4", "date": "F5",
                      "kind": "F6"}],
    },
}

# SSDPMA exports as three separate workbooks, one per score track, because the three
# come from three different source workbooks and the site receives them separately.
SSDPMA_TRACKS = {
    "sma":          {"label": "SMA-MFG",                 "kind": "questions",
                     "layout": "manufacturing"},
    "safety_solid": {"label": "Solidification-Safety",   "kind": "solid",
                     "template": "safety_solid_checksheet.xlsx"},
    "dp_solid":     {"label": "Solidification-DP",       "kind": "solid",
                     "template": "dp_solid_checksheet.xlsx"},
}

# Site/Date on the solidification workbooks. These have no Cover Sheet block like the MFG
# master — the Report and its (ref.) cover carry `Site` in B4 (merged B:C, value in D4,
# merged D:F) and `Date` in G4 (value in J4, merged J:O). Without this they go out
# anonymous, which is how BMT's name survived on them for so long.
#
# Only the half this workbook actually fills is named. Att-2's DP half is a superseded
# revision and is deliberately never written, so putting a site name on `DP Cover` there
# would assert a DP assessment that did not happen.
SOLID_IDENTITY = {
    "safety_solid": [{"sheet": "Safety Report",       "site": "D4", "date": "J4"},
                     {"sheet": "(ref.) Safety Cover", "site": "D4", "date": "J4"}],
    "dp_solid":     [{"sheet": "DP Report",           "site": "D4", "date": "J4"},
                     {"sheet": "DP Cover",            "site": "D4", "date": "J4"},
                     # the Safety twin half IS filled in this workbook, so it is named too
                     {"sheet": "Safety Cover",        "site": "D4", "date": "J4"}],
}

_MAPS_CACHE = {}
def _solid_maps():
    """Row/cell routing for the two solidification workbooks, derived from the source
    files themselves (see export_templates/ssdpma_solid_maps.json) rather than guessed.

    The two workbooks are wired in OPPOSITE directions, which decides where a judgement
    has to be written:
      * Safety — 'Safety overall result' Judge cells are blank inputs, and the role
        sheets read their criteria from it. The overall sheet is authoritative; role
        sheets get the same value mirrored in so the interview views aren't left blank.
      * DP — 'DP overall' Judge cells are FORMULAS pulling from the role sheets
        (=OP!H4, =SV!H5 ...). Writing there would destroy that wiring, so the role
        sheet is the real input. Only the 25 BSAPIC cells, which no role sheet feeds,
        are written directly onto 'DP overall'.
    """
    if "maps" not in _MAPS_CACHE:
        _MAPS_CACHE["maps"] = json.loads(
            (TEMPLATE_DIR / "ssdpma_solid_maps.json").read_text())
    return _MAPS_CACHE["maps"]

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

# ── byte-level cell surgery ───────────────────────────────────────────────────
# The sheet XML is spliced as BYTES, never parsed and re-serialised. ElementTree
# rewrites the namespace prefixes on the way out: mc: becomes ns1:, and the
# x14ac/xr/xr2/xr3 declarations are dropped because no element it kept still uses
# them — while mc:Ignorable="x14ac xr xr2 xr3" goes on naming them. Referencing
# undeclared prefixes is invalid Markup Compatibility, so Excel rejects the file
# ("unreadable content"), even though openpyxl and a plain XML parse are both happy.
# Splicing bytes leaves every byte we did not deliberately change exactly as it was.

_ROW_START = re.compile(rb"<row\b")
_CELL_START = re.compile(rb"<c\b")
_ATTR_R = re.compile(rb'\sr="([^"]+)"')
_ATTR_S = re.compile(rb'\ss="(\d+)"')

def _tag_end(xml, start):
    """Index just past the '>' of the tag starting at `start`, and whether it self-closed."""
    i = xml.index(b">", start)
    return i + 1, xml[i - 1:i] == b"/"

def _elements(xml, pattern, close_tag, lo, hi):
    """Yield (attrs_bytes, elem_start, elem_end) for each matching element in [lo,hi)."""
    pos = lo
    while True:
        m = pattern.search(xml, pos)
        if not m or m.start() >= hi: return
        after, selfclose = _tag_end(xml, m.start())
        start_tag = xml[m.start():after]
        end = after if selfclose else xml.index(close_tag, after) + len(close_tag)
        yield start_tag, m.start(), end
        pos = end

def _esc(s):
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(s))
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _cell_bytes(ref, style, val):
    s = b' s="%s"' % style if style else b""
    if val in (None, ""):
        return b'<c r="%s"%s/>' % (ref.encode(), s)
    space = b' xml:space="preserve"' if str(val) != str(val).strip() else b""
    return (b'<c r="%s"%s t="inlineStr"><is><t%s>%s</t></is></c>'
            % (ref.encode(), s, space, _esc(val).encode("utf-8")))

def _patch_row_bytes(xml, row_start, row_end, rnum, edits):
    """Rewrite the target cells inside one <row> element. Returns (new_row_bytes, written)."""
    after, selfclose = _tag_end(xml, row_start)
    if selfclose:                      # <row .../> — no cells to replace, build them all
        inner_lo = inner_hi = after
        head, tail = xml[row_start:after - 2] + b">", b"</row>"
    else:
        inner_lo, inner_hi = after, xml.index(b"</row>", after)
        head, tail = xml[row_start:after], b"</row>"

    cells = []
    for start_tag, cs, ce in _elements(xml, _CELL_START, b"</c>", inner_lo, inner_hi):
        m = _ATTR_R.search(start_tag)
        cells.append([_split(m.group(1).decode())[0] if m else "", cs, ce, start_tag])

    out, written, used = [], 0, set()
    for col, cs, ce, start_tag in cells:
        if col in edits:
            used.add(col)
            sm = _ATTR_S.search(start_tag)
            out.append(_cell_bytes(f"{col}{rnum}", sm.group(1) if sm else None, edits[col]))
            if edits[col] not in (None, ""): written += 1
        else:
            out.append(xml[cs:ce])
    # cells the row did not have yet, inserted in column order
    for col, val in edits.items():
        if col in used: continue
        new = _cell_bytes(f"{col}{rnum}", None, val)
        if val not in (None, ""): written += 1
        at = len(out)
        for i, (c, *_rest) in enumerate(cells):
            if c and _col_idx(c) > _col_idx(col):
                at = min(at, i + sum(1 for k in edits if k not in used and
                                     _col_idx(k) < _col_idx(col)))
                break
        out.insert(at, new)
    return head + b"".join(out) + tail, written

def _patch_rows(xml_bytes, edits_by_row):
    """edits_by_row: {row_number: {col: value}} — value None clears the cell.
    Returns (xml, written, rows_not_found) so a bad row map is caught, not silently
    skipped: writing an audit judgement against the wrong row must never pass quietly."""
    xml = xml_bytes
    pieces, last, written, seen = [], 0, 0, set()
    for start_tag, rs, re_ in _elements(xml, _ROW_START, b"</row>", 0, len(xml)):
        m = _ATTR_R.search(start_tag)
        if not m: continue
        rnum = int(m.group(1))
        edits = edits_by_row.get(rnum)
        if not edits: continue
        seen.add(rnum)
        new, n = _patch_row_bytes(xml, rs, re_, rnum, edits)
        pieces.append(xml[last:rs]); pieces.append(new)
        last, written = re_, written + n
    pieces.append(xml[last:])
    return b"".join(pieces), written, sorted(set(edits_by_row) - seen)

def _cell_text_bytes(xml, row_start, row_end, col):
    """Plain text of one cell in a row — the No column is numeric/inline."""
    after, selfclose = _tag_end(xml, row_start)
    if selfclose: return None
    inner_hi = xml.index(b"</row>", after)
    for start_tag, cs, ce in _elements(xml, _CELL_START, b"</c>", after, inner_hi):
        m = _ATTR_R.search(start_tag)
        if not m or _split(m.group(1).decode())[0] != col: continue
        body = xml[cs:ce]
        v = re.search(rb"<v>([^<]*)</v>", body) or re.search(rb"<t[^>]*>([^<]*)</t>", body)
        return v.group(1).decode().strip() if v else None
    return None

def _patch_sheet(xml_bytes, no_col, edits_by_no):
    """edits_by_no: {question_no: {col: value}} — resolved to rows via the No column."""
    by_row = {}
    for start_tag, rs, re_ in _elements(xml_bytes, _ROW_START, b"</row>", 0, len(xml_bytes)):
        m = _ATTR_R.search(start_tag)
        if not m: continue
        raw = _cell_text_bytes(xml_bytes, rs, re_, no_col)
        if not raw or not raw.isdigit(): continue
        edits = edits_by_no.get(int(raw))
        if edits: by_row[int(m.group(1))] = edits
    xml, written, _ = _patch_rows(xml_bytes, by_row)
    return xml, written

def _repack(template_path, patched_sheets, force_recalc=True):
    """Copy the workbook through, swapping in the rewritten sheet XML. Everything else —
    charts, styles, printer settings, custom XML — is passed byte-for-byte."""
    with zipfile.ZipFile(template_path) as z:
        paths = _sheet_paths(z)
        patched = {}
        for sheet, xml in patched_sheets.items():
            target = paths.get(sheet)
            if target: patched[target] = xml
        # The template ships with cached formula results from an EMPTY sheet. Without this
        # Excel can show those stale blanks — dashboard at 0, radars flat — until something
        # forces a recalc. fullCalcOnLoad makes it recompute the moment the file opens.
        if force_recalc:
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            if "fullCalcOnLoad" not in wbxml:
                wbxml = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>',
                               wbxml, count=1)
                patched["xl/workbook.xml"] = wbxml.encode("utf-8")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, patched.get(item.filename) or z.read(item.filename))
    return buf.getvalue()

def _sheet_xml(template_path, sheet_name):
    with zipfile.ZipFile(template_path) as z:
        path = _sheet_paths(z).get(sheet_name)
        if path is None: raise KeyError(f"no sheet named {sheet_name!r} in the template")
        return z.read(path)

def supported(assessment_type):
    return assessment_type in LAYOUTS or assessment_type == "ssdpma"

def tracks(assessment_type):
    """The separate workbooks this type exports as. One for the plain types; three for
    SSDPMA — its SMA, Safety and DP scores come from three different source workbooks."""
    if assessment_type == "ssdpma":
        return [(k, v["label"]) for k, v in SSDPMA_TRACKS.items()]
    return [(None, "Check sheet")] if assessment_type in LAYOUTS else []

def build(assessment_type, pillars, responses, assessment=None):
    """Return (bytes, cells_written). `responses` is {question_id: {answer, comment}}."""
    layout = LAYOUTS.get(assessment_type)
    if not layout: raise ValueError(f"no check-sheet template for {assessment_type!r}")
    path = TEMPLATE_DIR / layout["template"]
    if not path.exists(): raise FileNotFoundError(f"missing export template {path}")

    judgement = layout.get("judgement", JUDGEMENT)

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
                # only say it in the Reason column when the Judgement column cannot
                if ans == "not_rolled_out" and not judgement.get(ans):
                    note = ("Not rolled out" + (" — " + note if note else ""))
                per[int(q["no"])] = {jcol: judgement.get(ans), rcol: note or None}
                total += 1

    patched = {}
    for sheet, edits in plan.items():
        patched[sheet], _ = _patch_sheet(_sheet_xml(path, sheet), layout["no_col"][sheet], edits)

    cover = layout.get("cover")
    if cover and assessment:
        edits = {}
        def _put(ref, val):
            if not val: return
            col, row = _split(ref)
            edits.setdefault(row, {})[col] = val
        date = (assessment.get("assessment_date") or "").strip()
        _put(cover["site"], assessment.get("site_name"))
        _put(cover["year"], date[:4] if len(date) >= 4 and date[:4].isdigit() else None)
        _put(cover["date"], date)
        names = [n for n in (assessment.get("assessor_a"), assessment.get("assessor_b")) if n]
        for ref, name in zip(cover["assessors"], names):
            _put(ref, name)
        if edits:
            patched[cover["sheet"]], _, _ = _patch_rows(_sheet_xml(path, cover["sheet"]), edits)

    _write_identity(path, patched, layout.get("identity"), assessment)

    return _repack(path, patched), total


def _write_identity(path, patched, idents, assessment):
    """Fill the Site/Date/Assessor/Type block on the sheets the site actually reads.
    Skipped silently when the record has nothing to say — an exported sheet must never
    invent identity, and must never leave the previous site's showing."""
    if not assessment or not idents: return
    date = (assessment.get("assessment_date") or "").strip()
    for ident in idents:
        edits = {}
        for key, val in (("site", assessment.get("site_name")),
                         ("date", date),
                         ("assessors", _assessor_names(assessment)),
                         ("kind", ASSESSMENT_TYPE_LABEL.get(
                             (assessment.get("kind") or "").strip()))):
            ref = ident.get(key)
            if not ref or not val: continue
            col, row = _split(ref)
            edits.setdefault(row, {})[col] = val
        if edits:
            patched[ident["sheet"]], _, _ = _patch_rows(
                _sheet_xml(path, ident["sheet"]), edits)

def build_solid(track, sections, responses, twin_sections=None, assessment=None):
    """One solidification workbook (Safety or DP) with every judged rubric level written
    into its Judge cell. `sections` is the bank's list for that track; `responses` is
    keyed <item id>__L<level>, the same key the assess page saves under.
    `twin_sections` is the OTHER track's list, used to fill this workbook's twin half
    where that half is the same revision (see the note below).

    Raises if the row map and the bank disagree — a silent skip here would mean a
    judgement landing on the wrong item, which is worse than a failed download.
    """
    conf = SSDPMA_TRACKS[track]
    path = TEMPLATE_DIR / conf["template"]
    if not path.exists(): raise FileNotFoundError(f"missing export template {path}")
    maps = _solid_maps()[track]

    plan, total = {}, 0
    _fill(plan, track, maps, sections, responses, strict=True)
    total = _count(maps, sections, responses)

    # Each source workbook carries the OTHER track's sheets too. Where that twin half is
    # the same revision it gets filled, so the site receives one complete record — Att-3's
    # 'Safety cal.' matches on 31 of 36 items (5 absent from that revision, left blank).
    # Att-2 has no usable DP twin: its 'DP cal.' is a superseded 28-item checklist whose
    # criteria were rewritten, so writing current judgements there would pin them to
    # criteria the site was never assessed against. Deliberately not filled.
    twin = maps.get("twin")
    if twin and twin_sections:
        _fill(plan, track + ":twin", twin, twin_sections, responses, strict=False)

    patched, missing = {}, {}
    for sheet, edits in plan.items():
        patched[sheet], _, miss = _patch_rows(_sheet_xml(path, sheet), edits)
        if miss: missing[sheet] = miss
    if missing:
        raise ValueError(f"{track}: rows not found in {missing} — row map is stale")
    _write_identity(path, patched, SOLID_IDENTITY.get(track), assessment)
    return _repack(path, patched), total

def _count(maps, sections, responses):
    n = 0
    for s in sections:
        if s["id"] not in maps["rows"]: continue
        for lv in s.get("solid_rubric") or []:
            if _judgement(responses, f"{s['id']}__L{lv['level']}"): n += 1
    return n

def _judgement(responses, key):
    r = responses.get(key)
    ans = r.get("answer") if isinstance(r, dict) else r
    return JUDGEMENT.get((ans or "").strip())

def _fill(plan, label, maps, sections, responses, strict):
    """Add one half's judgements to `plan`. strict=True means every bank item MUST have a
    row — that half is the workbook's own track, so a gap there is a stale map. strict=False
    is the twin half, where items genuinely absent from that revision are skipped."""
    rows, judge_cols = maps["rows"], maps["judge_cols"]
    if strict:
        unknown = [s["id"] for s in sections if s["id"] not in rows]
        if unknown:
            raise ValueError(f"{label}: no row mapped for {unknown[:5]} — row map is stale")
    for s in sections:
        row = rows.get(s["id"])
        if row is None: continue                 # not in this revision of the twin half
        for lv in s.get("solid_rubric") or []:
            key = f"{s['id']}__L{lv['level']}"
            val = _judgement(responses, key)
            if not val: continue
            # Safety-shaped half: the overall sheet IS the input, so every level is written
            # there and mirrored onto the role sheet when one shows it.
            # DP-shaped half: the overall sheet's Judge cells are formulas reading the role
            # sheets, so the role sheet is the input; only levels no role sheet feeds are
            # written onto the overall sheet directly.
            target, direct = maps["role_targets"].get(key), maps["direct"].get(key)
            if target:
                sheet, col, r = target
                plan.setdefault(sheet, {}).setdefault(int(r), {})[col] = val
            if maps["overall_is_input"]:
                col, r = judge_cols[int(lv["level"]) - 1], int(row)
            elif direct:
                col, r = _split(direct)
            elif target:
                continue        # that cell is a formula reading the role sheet
            else:
                raise ValueError(f"{label}: {key} has nowhere to write — map is stale")
            plan.setdefault(maps["overall_sheet"], {}).setdefault(r, {})[col] = val
