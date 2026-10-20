# Read-only: for every IFC DWG, list title blocks and whether a '0 / FOR CONSTRUCTION' revision row exists
import sys, time, re, glob
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo"); E = P/"Design/Engineering"
ifc = v.IFCManager(P, dry_run=True, preserve_ifr=True); a = ifc._get_acad()
dw = sorted(set(glob.glob(str(E/"1. Drawings/1. Native/*/*_Rev0_IFC.dwg")) + glob.glob(str(E/"2. Calcs & Reports/Reports/Civil & Structure/*/*_Rev0_IFC.dwg"))))
only = sys.argv[1:]
for f in dw:
    if only and not any(o in f for o in only): continue
    import os, subprocess, hashlib
    fp = Path(f)
    if len(f) > 250:
        j = Path(os.environ["TEMP"]) / ("ifcj_" + hashlib.md5(str(fp.parent).encode()).hexdigest()[:8])
        if not j.exists(): subprocess.run(["cmd","/c","mklink","/J",str(j),str(fp.parent)],check=True,capture_output=True)
        fp = j/fp.name
    doc = ifc._com_retry(lambda: a.Documents.Open(str(fp), True)); time.sleep(2)
    out = []
    for t in ifc._find_all_title_blocks(doc):
        at = t[1]; g = lambda k: ifc._safe_get_text(at[k]).strip() if k in at else None
        rows = [(n, g(f"{n}REV"), g(f"{n}DESCRIPTION")) for n in range(1, 9) if g(f"{n}REV") or g(f"{n}DESCRIPTION")]
        out.append((t[3] if len(t) > 3 else "?", rows if rows else sorted(at)[:12]))
    print(Path(f).name.split("_")[0], out, flush=True)
    doc.Close(False); time.sleep(1)
