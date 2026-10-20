# Keep only the IFC pages that match a page the client approved in the IFR PDF (extra layouts in the DWG were never issued).
# Original IFC PDF backed up to 4. IFC(Client)\SS\<name>.untrimmed.pdf first.
import sys, shutil, fitz
from pathlib import Path
E = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo\Design\Engineering")
ifc_dir = E / "1. Drawings/4. IFC(Client)"; sp = list(E.glob("13. Client Sharepoint/1.IFR/*/Approved to IFC/*.pdf"))
APPLY = "--apply" in sys.argv
def vec(pg):
    pm = pg.get_pixmap(matrix=fitz.Matrix(256/pg.rect.width), colorspace=fitz.csGRAY)
    return pm.samples, (pm.width, pm.height)
def sim(a, b):
    return 0 if a[1] != b[1] else 1 - sum(abs(x-y) for x,y in zip(a[0], b[0]))/255/len(a[0])
for p in sorted(ifc_dir.glob("*_Rev0_IFC.pdf")):
    doc = p.name.split("_")[0]; ifr = [q for q in sp if q.name.startswith(doc + "_")]
    d = fitz.open(p); r = fitz.open(ifr[0])
    if d.page_count == r.page_count: continue
    keep = []
    for rp in r:
        rv = vec(rp); s = [sim(vec(x), rv) for x in d]; b = max(range(len(s)), key=s.__getitem__)
        keep.append((b, round(s[b], 3)))
    ok = all(sc >= 0.95 for _, sc in keep) and len({b for b, _ in keep}) == len(keep)
    print(doc, f"{d.page_count}->{len(keep)} keep={keep}", "OK" if ok else "AMBIGUOUS-skip")
    if APPLY and ok:
        d.close(); ss = ifc_dir / "SS"; ss.mkdir(exist_ok=True)
        bk = ss / (p.stem + ".untrimmed.pdf"); shutil.copy2(p, bk)
        src = fitz.open(bk); out = fitz.open()
        for b, _ in keep: out.insert_pdf(src, from_page=b, to_page=b)
        out.save(p); print("   trimmed; backup", bk.name)
