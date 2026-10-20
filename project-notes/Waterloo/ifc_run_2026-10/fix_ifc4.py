# Fix 5 IFC DWGs found by visual QA, then republish their PDFs with the tool's own _publish_single_pdf.
# Reversible: each DWG copied to <folder>\SS\<name>.pre-fix.dwg BEFORE any save. Script (v10) untouched.
import sys, time, shutil, re, os, subprocess, hashlib
from datetime import datetime
from pathlib import Path
import pythoncom, win32com.client
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
E = P / "Design/Engineering"; N = E / "1. Drawings/1. Native"; OUT = E / "1. Drawings/4. IFC(Client)"
R = E / "2. Calcs & Reports/Reports/Civil & Structure"
JOBS = [
 ("GG38-E-BLD-001", "del_info", N/"GG38-E-BLD-001_Auxiliary Cable Block Diagram_Rev/GG38-E-BLD-001_Auxiliary_Cable_Block_Diagram_Rev0_IFC.dwg"),
 ("GG38-E-SLD-001", "del_info", N/"GG38-E-SLD-001_Single Line Diagram_Rev/GG38-E-SLD-001_Single_Line_Diagram_Rev0_IFC.dwg"),
 ("GG38-E-BLD-003", "move_stamp", N/"GG38-E-BLD-003_Communication Block Diagram_Rev/GG38-E-BLD-003_Communications_Block_Diagram_Rev0_IFC.dwg"),
 ("GG38-C-PLN-017", "drgno", N/"GG38-C-PLN-017_Kerb Layout Plan_Rev/GG38-C-PLN-017_Kerb_Layout_Plan_Rev0_IFC.dwg"),
 ("GG38-C-PLN-009", "roll", R/"GG38-C-RPT-006_BESS Foundation Design Report_Rev/GG38-C-PLN-009_BESS_Foundation_Design_Rev0_IFC.dwg"),
]
only = set(sys.argv[1:])
ifc = v.IFCManager(P, dry_run=True, preserve_ifr=True); a = ifc._get_acad()
pt = lambda x, y: win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, 0.0))


def short(dwg):
    if len(str(dwg)) <= 250: return dwg
    j = Path(os.environ["TEMP"]) / ("ifcj_" + hashlib.md5(str(dwg.parent).encode()).hexdigest()[:8])
    if not j.exists(): subprocess.run(["cmd", "/c", "mklink", "/J", str(j), str(dwg.parent)], check=True, capture_output=True)
    return j / dwg.name


def paper(doc):
    for lay in doc.Layouts:
        if lay.Name.lower() != "model": yield lay.Name, lay.Block


def del_info(doc):
    bad = set()
    for b in doc.Blocks:
        if not b.Name.startswith("A$C"): continue
        for i in range(b.Count):
            e = b.Item(i)
            if e.ObjectName in ("AcDbText", "AcDbMText") and "FOR INFORMATION" in e.TextString.upper().replace("\\P", " "):
                bad.add(b.Name); break
    done = []
    for ln, blk in paper(doc):
        for e in [blk.Item(i) for i in range(blk.Count)]:
            if e.ObjectName == "AcDbBlockReference" and e.Name in bad:
                done.append(f"{ln}:{e.Name}@{e.Handle}"); e.Delete()
    return done


def move_stamp(doc, dy=139.0):
    done = []
    for ln, blk in paper(doc):
        for e in [blk.Item(i) for i in range(blk.Count)]:
            if e.Layer == "IFC_STAMP":
                e.Move(pt(0, 0), pt(0, dy)); done.append(f"{ln}:{e.ObjectName}@{e.Handle}")
    return done


def drgno(doc):
    done = []
    spaces = [("Model", doc.ModelSpace)] + list(paper(doc))
    for ln, blk in spaces:
        for i in range(blk.Count):
            e = blk.Item(i)
            if e.ObjectName == "AcDbBlockReference" and e.HasAttributes:
                for at in e.GetAttributes():
                    if "PLN-00X" in at.TextString.upper():
                        old = at.TextString; at.TextString = re.sub("(?i)PLN-00X", "PLN-017", old)
                        done.append(f"{ln}:{at.TagString} {old!r}->{at.TextString!r}")
    return done


def _d(s):
    try: return datetime.strptime(s.strip(), "%d/%m/%y")
    except Exception: return datetime.min


def roll(doc):
    done = []
    for t in ifc._find_all_title_blocks(doc):
        at = t[1]; g = lambda k: ifc._safe_get_text(at[k]).strip() if k in at else ""
        n_rows = max(int(m.group(1)) for k in at if (m := re.match(r"^(\d)REV$", k)))
        if any("CONSTRUCTION" in g(f"{n}DESCRIPTION").upper() for n in range(1, n_rows + 1)):
            done.append(f"{t[3]}: already has IFC row"); continue
        sfx = sorted({k[1:] for k in at if re.match(r"^\dREV$|^\d[A-Z]", k) and k[0].isdigit() and int(k[0]) <= n_rows})
        rows = [{s: g(f"{n}{s}") for s in sfx} for n in range(1, n_rows + 1)]
        rows = [r for r in rows if r.get("REV")]
        rows.sort(key=lambda r: _d(r.get("DATE", "")))
        newest = dict(rows[-1])
        for s in ifc.PERSONNEL_TAGS:  # per-field backfill, newest-first (same rule as tool)
            if s in sfx and not newest.get(s):
                newest[s] = next((r[s] for r in reversed(rows) if r.get(s)), "")
        new = {s: (newest.get(s, "") if s in ifc.PERSONNEL_TAGS else "") for s in sfx}
        new.update(REV="0", DATE=datetime.now().strftime("%d/%m/%y"), DESCRIPTION="FOR CONSTRUCTION")
        rows = (rows + [new])[-n_rows:]
        for n, r in enumerate(rows, 1):
            for s in sfx:
                if f"{n}{s}" in at: ifc._safe_set_text(at[f"{n}{s}"], r.get(s, ""))
        done.append(f"{t[3]}: " + " | ".join(f"{r['REV']} {r.get('DATE','')} {r.get('DESCRIPTION','')}" for r in rows))
    return done


FIX = {"del_info": del_info, "move_stamp": move_stamp, "drgno": drgno, "roll": roll}
for doc_id, kind, dwg in JOBS:
    if only and doc_id not in only: continue
    try:
        sp = short(dwg); assert sp.exists(), sp
        pdf = next(OUT.glob(doc_id + "_*_Rev0_IFC.pdf"))
        bk = sp.parent / "SS"; bk.mkdir(exist_ok=True)
        dst = bk / (sp.stem + ".pre-fix" + sp.suffix)
        if not dst.exists():
            shutil.copy2(sp, dst); assert dst.stat().st_size == sp.stat().st_size
        while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
        doc = ifc._com_retry(lambda: a.Documents.Open(str(sp), False)); time.sleep(3)
        done = FIX[kind](doc)
        if not done: raise RuntimeError("nothing to fix")
        doc.Save(); time.sleep(1); doc.Close(False); time.sleep(1)
        print(doc_id, kind, "FIXED", done, "backup:", dst.name, flush=True)
        ok = ifc._publish_single_pdf(a, sp, pdf)
        print(doc_id, "republish", ok, pdf.name, flush=True)
    except Exception as e:
        print(doc_id, "ERR", repr(e), flush=True)
        try:
            while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
        except Exception: pass
print("DONE", flush=True)
