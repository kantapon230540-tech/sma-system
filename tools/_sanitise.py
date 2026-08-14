"""Deep-sanitise the export templates.

Clearing cell values is NOT enough. A completed workbook keeps the previous site's data in
four more places, none of which show up in a structural check:

  1. xl/sharedStrings.xml — clearing a cell removes the reference, not the string. TBSCN's
     site name, assessor names and OFI findings text all survive as orphaned entries.
  2. xl/charts/chartN.xml — every chart caches the values it last plotted, so the radars
     render the previous site's scores until Excel redraws them.
  3. docProps/core.xml — dc:creator and cp:lastModifiedBy carry real people's names.
  4. xl/comments1.xml — the comment author's name.

Everything is edited as BYTES (see ref_xlsx_namespace_trap): re-serialising any of these
with ElementTree rewrites namespace prefixes and Excel then rejects the file.

An <si> is only blanked when NO remaining cell references it, so live criteria/question
text is untouched. Chart numCache values are zeroed but the cache structure is kept, and
fullCalcOnLoad makes Excel recompute and redraw from the live cells on open.
"""
import re, sys, zipfile
from pathlib import Path

TPL_DIR = Path(__file__).resolve().parent.parent / "export_templates"
SHEET_RE = re.compile(rb"xl/worksheets/sheet\d+\.xml$")
CHART_RE = re.compile(r"xl/charts/chart\d+\.xml$")


def referenced_shared_strings(z):
    """Indices still pointed at by a cell (t="s" means <v> is a sharedStrings index)."""
    used = set()
    cell = re.compile(rb'<c\b[^>]*\bt="s"[^>]*>\s*(?:<f[^>]*>.*?</f>\s*)?<v>(\d+)</v>', re.S)
    for n in z.namelist():
        if not SHEET_RE.search(n.encode()):
            continue
        for m in cell.finditer(z.read(n)):
            used.add(int(m.group(1)))
    return used


def blank_orphan_strings(xml, used):
    """Replace the text of every unreferenced <si> with nothing, keeping its index."""
    out, pos, idx, blanked = [], 0, 0, 0
    for m in re.finditer(rb"<si>.*?</si>", xml, re.S):
        out.append(xml[pos:m.start()])
        if idx in used:
            out.append(m.group(0))
        else:
            out.append(b"<si><t/></si>")
            blanked += 1
        pos, idx = m.end(), idx + 1
    out.append(xml[pos:])
    return b"".join(out), blanked, idx


def zero_chart_cache(xml):
    """Zero the cached numeric points; leave categories and cache structure intact."""
    n = 0
    def repl(m):
        nonlocal n
        inner, cnt = re.subn(rb"(<c:v>)[^<]*(</c:v>)", rb"\g<1>0\g<2>", m.group(2))
        n += cnt
        return m.group(1) + inner + m.group(3)
    xml = re.sub(rb"(<c:numCache>)(.*?)(</c:numCache>)", repl, xml, flags=re.S)
    return xml, n


def scrub(path, report):
    z = zipfile.ZipFile(path)
    used = referenced_shared_strings(z)
    patched, stats = {}, {}

    if "xl/sharedStrings.xml" in z.namelist():
        xml, blanked, total = blank_orphan_strings(z.read("xl/sharedStrings.xml"), used)
        patched["xl/sharedStrings.xml"] = xml
        stats["sharedStrings blanked"] = f"{blanked} of {total} (kept {total-blanked} in use)"

    charts = 0
    for n in z.namelist():
        if not CHART_RE.match(n):
            continue
        xml, c = zero_chart_cache(z.read(n))
        if c:
            patched[n] = xml
            charts += c
    stats["chart cached points zeroed"] = charts

    if "docProps/core.xml" in z.namelist():
        x = z.read("docProps/core.xml")
        names = re.findall(rb"<dc:creator>([^<]*)</dc:creator>", x) + \
                re.findall(rb"<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>", x)
        x = re.sub(rb"<dc:creator>[^<]*</dc:creator>", b"<dc:creator></dc:creator>", x)
        x = re.sub(rb"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>",
                   b"<cp:lastModifiedBy></cp:lastModifiedBy>", x)
        patched["docProps/core.xml"] = x
        stats["docProps names removed"] = [n.decode() for n in names if n]

    for n in z.namelist():
        if not re.match(r"xl/comments\d+\.xml$", n):
            continue
        x = z.read(n)
        authors = re.findall(rb"<author>([^<]*)</author>", x)
        x = re.sub(rb"<author>[^<]*</author>", b"<author></author>", x)
        patched[n] = x
        stats["comment authors removed"] = [a.decode() for a in authors if a]

    out = path.with_suffix(".tmp")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
        for item in z.infolist():
            w.writestr(item, patched.get(item.filename) or z.read(item.filename))
    z.close()
    out.replace(path)
    report[path.name] = stats


if __name__ == "__main__":
    report = {}
    for f in ["mfg_checksheet.xlsx", "safety_solid_checksheet.xlsx",
              "dp_solid_checksheet.xlsx", "warehouse_checksheet.xlsx"]:
        p = TPL_DIR / f
        if p.exists():
            scrub(p, report)
    for name, st in report.items():
        print(f"\n{name}")
        for k, v in st.items():
            print(f"   {k}: {v}")
