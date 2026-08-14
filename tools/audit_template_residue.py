"""Exhaustive residue audit: scan EVERY part of a template for previous-site data.

Written after four rounds of "it's clean now" that were each wrong. Instead of checking the
places I happen to think of, this walks every byte of every part and flags anything matching
site names, people, finding language or known leak vectors.
"""
import re, sys, zipfile
from pathlib import Path

TPL = Path(__file__).resolve().parent.parent / "export_templates"

# site codes, the people who appear in these workbooks, and finding vocabulary
MARKERS = [
    r"TBSCN", r"BSFC", r"BTMT", r"JSDC", r"ZBDC",
    r"Jarinya", r"Thanida", r"Supattra", r"Justin\s*Thomson", r"Kantapon",
    r"鈴木", r"山下",
    r"OFI\s*:", r"No evidence", r"Oil tank", r"Same issue from",
    r"Confidential", r"salesbridge\.sharepoint", r"absPath",
]
PAT = re.compile("|".join(MARKERS), re.I)

# text that is legitimately part of the blank checklist and must NOT be flagged
ALLOW = re.compile(r"Confidential(ity)?\s+(of|information|is)"
                   r"|not enough to check/confirm it in GENBA"   # B+ rubric criterion,
                   r"|No evidence to check/confirm in GENBA",    # not a finding
                   re.I)

BINARY_OK = (".png", ".emf", ".wmf", ".bin", ".jpeg", ".jpg", ".gif")


def audit(path):
    z = zipfile.ZipFile(path)
    findings = []
    for n in z.namelist():
        if n.endswith(BINARY_OK):
            continue
        try:
            blob = z.read(n).decode("utf-8", "ignore")
        except Exception:
            continue
        for m in PAT.finditer(blob):
            ctx = blob[max(0, m.start() - 70):m.end() + 70].replace("\n", " ")
            if ALLOW.search(ctx):
                continue
            findings.append((n, m.group(0), ctx))
    # structural leak vectors that should now be empty
    vectors = {}
    vectors["embedded pictures"] = [n for n in z.namelist() if n.startswith("xl/media/")]
    vectors["autofilter criteria"] = sum(
        len(re.findall(rb"<filter\s", z.read(n)))
        for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    nz = 0
    for n in z.namelist():
        if re.match(r"xl/charts/chart\d+\.xml$", n):
            for cache in re.findall(rb"<c:numCache>.*?</c:numCache>", z.read(n), re.S):
                nz += len([v for v in re.findall(rb"<c:v>([^<]*)</c:v>", cache)
                           if v not in (b"0", b"")])
    vectors["non-zero chart cache points"] = nz
    core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in z.namelist() else ""
    vectors["docProps creator/lastModifiedBy"] = re.findall(
        r"<dc:creator>([^<]*)</dc:creator>|<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>", core)
    z.close()
    return findings, vectors


ok = True
for f in ["mfg_checksheet.xlsx", "safety_solid_checksheet.xlsx",
          "dp_solid_checksheet.xlsx", "warehouse_checksheet.xlsx"]:
    p = TPL / f
    if not p.exists():
        continue
    findings, vectors = audit(p)
    print("=" * 78)
    print(f"{f}   ({p.stat().st_size/1024:.0f} KB)")
    print(f"   embedded pictures            : {[x.split('/')[-1] for x in vectors['embedded pictures']]}")
    print(f"   autofilter criteria remaining: {vectors['autofilter criteria']}")
    print(f"   non-zero chart cache points  : {vectors['non-zero chart cache points']}")
    names = [a or b for a, b in vectors["docProps creator/lastModifiedBy"] if (a or b)]
    print(f"   docProps names               : {names or 'none'}")
    if findings:
        ok = False
        print(f"   *** {len(findings)} TEXT MATCHES ***")
        for n, tok, ctx in findings[:12]:
            print(f"      {n}  [{tok}]  ...{ctx[:110]}...")
    else:
        print("   text markers                 : none")
    if names or vectors["autofilter criteria"] or vectors["non-zero chart cache points"]:
        ok = False

print("\nRESULT:", "NO RESIDUE FOUND" if ok else "RESIDUE REMAINS — see above")
sys.exit(0 if ok else 1)
