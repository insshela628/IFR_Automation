import sys, time
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
N = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo\Design\Engineering\1. Drawings\1. Native")
D = {"BLD-001": N/"GG38-E-BLD-001_Auxiliary Cable Block Diagram_Rev/GG38-E-BLD-001_Auxiliary_Cable_Block_Diagram_Rev0_IFC.dwg",
     "BLD-003": N/"GG38-E-BLD-003_Communication Block Diagram_Rev/GG38-E-BLD-003_Communications_Block_Diagram_Rev0_IFC.dwg",
     "SLD-001": N/"GG38-E-SLD-001_Single Line Diagram_Rev/GG38-E-SLD-001_Single_Line_Diagram_Rev0_IFC.dwg",
     "PLN-017": N/"GG38-C-PLN-017_Kerb Layout Plan_Rev/GG38-C-PLN-017_Kerb_Layout_Plan_Rev0_IFC.dwg"}
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
ifc = v.IFCManager(P, dry_run=True, preserve_ifr=True); a = ifc._get_acad()
for k, dwg in D.items():
    doc = ifc._com_retry(lambda: a.Documents.Open(str(dwg), True)); time.sleep(3)
    print("==", k, [l.Name for l in doc.Layouts])
    for lay in doc.Layouts:
        if lay.Name.lower() == "model": continue
        blk = lay.Block
        for i in range(blk.Count):
            e = blk.Item(i); on = e.ObjectName
            try: mn, mx = e.GetBoundingBox(); bb = f"({mn[0]:.0f},{mn[1]:.0f})-({mx[0]:.0f},{mx[1]:.0f})"
            except Exception: bb = "?"
            if on in ("AcDbText", "AcDbMText"):
                t = e.TextString.upper()
                if "FOR " in t or "INFORMATION" in t or "CONSTRUCTION" in t: print(" ", lay.Name, on, e.Layer, repr(e.TextString[:50]), bb, e.Handle)
            elif on == "AcDbBlockReference":
                tag = {x.TagString: x.TextString for x in e.GetAttributes()} if e.HasAttributes else {}
                hit = {t: s for t, s in tag.items() if "00X" in s or "PLN-0" in s or "DRG" in t.upper() or "DWG" in t.upper()}
                print(" ", lay.Name, "BLK", e.EffectiveName, e.Layer, bb, e.Handle, hit)
            elif on == "AcDbPolyline" and e.Layer == "IFC_STAMP" or (on == "AcDbPolyline" and bb != "?" and 15 < (mx[1]-mn[1]) < 40 and 80 < (mx[0]-mn[0]) < 140):
                print(" ", lay.Name, on, e.Layer, bb, e.Handle)
            elif on == "AcDbViewport": print(" ", lay.Name, "VP", bb)
    doc.Close(False); time.sleep(1)
