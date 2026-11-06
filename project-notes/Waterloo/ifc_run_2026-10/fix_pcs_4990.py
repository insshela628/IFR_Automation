# 2026-11-06 PCS label 9567MV-5000 -> 9567MV-4990 (01_Commercial V1.1 quotation / NR primary B1.01) on the
# two Rev0 IFC sheets that still carry it, then republish with v10's own _publish_single_pdf. No rev bump
# (PM call: label correction on a Rev0 IFC, overwrite in place). Usage: python fix_pcs_4990.py [--apply]
import sys, time, re, shutil
from pathlib import Path
import pythoncom, win32com.client
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
E = P / "Design/Engineering"; N = E / "1. Drawings/1. Native"; OUT = E / "1. Drawings/4. IFC(Client)"
BK = Path(r"C:\Users\jilin\AppData\Local\Temp\claude\0fa2f3b7e932577a\35a5c24e-0deb-4a42-adef-2bacece47b32\scratchpad\pcs4990")
JOBS = [N/"GG38-E-SLD-001_Single Line Diagram_Rev/GG38-E-SLD-001_Single_Line_Diagram_Rev0_IFC.dwg",
        N/"GG38-E-BLD-001_Auxiliary Cable Block Diagram_Rev/GG38-E-BLD-001_Auxiliary_Cable_Block_Diagram_Rev0_IFC.dwg"]
PAT = re.compile(r"(9567MV-)(?:5000|4990)(?:-V[\d.]+)?", re.I)   # -V2 = 5000 datasheet version, not model; NR names it PCS-9567MV-4990 (idempotent)
APPLY = "--apply" in sys.argv
ifc = v.IFCManager(P, dry_run=True, preserve_ifr=True); a = ifc._get_acad()


def hits(doc):
    """Every TEXT/MTEXT/ATTRIB mentioning 9567MV — model, every layout, every block definition."""
    out = []
    for b in doc.Blocks:                       # includes *Model_Space / *Paper_Space* / user blocks
        n = b.Count
        for i in range(n):
            e = b.Item(i)
            on = e.ObjectName
            if on in ("AcDbText", "AcDbMText", "AcDbAttributeDefinition"):
                s = e.TextString
                if "9567" in s.upper(): out.append((b.Name, on, e.Handle, e, "TextString", s))
            elif on == "AcDbBlockReference" and e.HasAttributes:
                for at in e.GetAttributes():
                    if "9567" in at.TextString.upper():
                        out.append((b.Name, "ATTRIB:" + at.TagString, at.Handle, at, "TextString", at.TextString))
    return out


for dwg in JOBS:
    print("==", dwg.name, flush=True)
    assert dwg.exists(), dwg
    if APPLY:
        pdf0 = next(OUT.glob(dwg.stem.split("_")[0] + "_*_Rev0_IFC.pdf"))
        for f in (dwg, pdf0):                  # first run only: the backup is the pre-fix original
            if not (BK / f.name).exists(): shutil.copy2(f, BK / f.name)
    while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
    doc = ifc._com_retry(lambda: a.Documents.Open(str(dwg), not APPLY)); time.sleep(3)
    changed = 0
    for blk, on, h, ent, attr, s in hits(doc):
        new = PAT.sub(r"\g<1>4990", s)
        flag = "  -> " + repr(new) if new != s else "  (keep)"
        print(f"  {blk:22} {on:28} {h:>6} {s!r}{flag if APPLY or new != s else ''}", flush=True)
        if APPLY and new != s:
            ent.TextString = new; changed += 1
    if APPLY:
        doc.Regen(1); doc.Save(); time.sleep(1)
        print("  changed", changed, flush=True)
    doc.Close(False); time.sleep(1)
    if APPLY and changed:
        pdf = next(OUT.glob(dwg.stem.split("_")[0] + "_*_Rev0_IFC.pdf"))
        print("  republish", ifc._publish_single_pdf(a, dwg, pdf), flush=True)
