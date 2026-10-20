# QA: IFC PDF vs approved IFR PDF — page count, FOR CONSTRUCTION on every page, no leftover IFR stamp, rev text
import sys, re, glob, fitz
from pathlib import Path
E = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo\Design\Engineering")
ifc_dir = E / "1. Drawings/4. IFC(Client)"
sp = list(E.glob("13. Client Sharepoint/1.IFR/*/Approved to IFC/*.pdf"))
for p in sorted(ifc_dir.glob("*_Rev0_IFC.pdf")):
    doc = p.name.split("_")[0]
    ifr = [q for q in sp if q.name.startswith(doc + "_") or q.name.startswith(doc + " ")]
    d = fitz.open(p); n = d.page_count
    nfr = fitz.open(ifr[0]).page_count if ifr else None
    bad = []
    for i, pg in enumerate(d):
        t = pg.get_text().upper()
        if "FOR CONSTRUCTION" not in t: bad.append(f"p{i+1}:noFC")
        if "ISSUED FOR REVIEW" in t or "FOR REVIEW" in t: bad.append(f"p{i+1}:IFRleft")
    verdict = "PASS" if not bad and nfr == n else "CHECK"
    print(f"{verdict} {doc}: pages {n} vs IFR {nfr} ({ifr[0].name if ifr else 'NO IFR'}) {bad}")
