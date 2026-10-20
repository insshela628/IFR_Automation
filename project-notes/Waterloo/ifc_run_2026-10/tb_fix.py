# B0: source fix — revision row with DATE/DESCRIPTION but blank REV gets REV = submitted IFR rev.
# Reversible: original DWG copied to <folder>\SS\<name>.pre-REVfix.dwg BEFORE any save.
import sys, time, shutil
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
src = open(Path(__file__).with_name("ifc_phaseB.py"), encoding="utf-8").read()
ns = {}; exec(src.split("only = set")[0].split("import ifr_automation_v10 as v")[1], {"Path": Path, "v": v}, ns)
P, JOBS = ns["P"], ns["JOBS"]
ifc = v.IFCManager(P, dry_run=True, preserve_ifr=True); a = ifc._get_acad()


def needs_fix(tbs):
    plan = []
    if any(ifc._read_latest_ifr_row(t[1]) for t in tbs):
        return plan
    for j, t in enumerate(tbs):
        at = t[1]
        g = lambda k: ifc._safe_get_text(at[k]).strip() if k in at else ''
        filled = [n for n in range(1, 7) if g(f"{n}DATE") or g(f"{n}DESCRIPTION")]
        if filled and f"{filled[-1]}REV" in at and not g(f"{filled[-1]}REV"):
            n = filled[-1]
            plan.append((at, n, f"tb{j+1} {n}REV ({g(f'{n}DATE')} {g(f'{n}DESCRIPTION')})"))
    return plan


for doc_id, desc, dwg, rev, _ in JOBS:
    try:
        doc = ifc._com_retry(lambda: a.Documents.Open(str(dwg), True)); time.sleep(3)
        tbs = ifc._find_all_title_blocks(doc)
        plan = needs_fix(tbs)
        ok = any(ifc._read_latest_ifr_row(t[1]) for t in tbs)
        doc.Close(False); time.sleep(1)
        if not plan:
            print(doc_id, "ok" if ok else "STILL-NO-ROW", f"tbs={len(tbs)}", flush=True)
            continue
        bk = dwg.parent / "SS"; bk.mkdir(exist_ok=True)
        dst = bk / (dwg.stem + ".pre-REVfix" + dwg.suffix)
        shutil.copy2(dwg, dst); assert dst.stat().st_size == dwg.stat().st_size
        doc = ifc._com_retry(lambda: a.Documents.Open(str(dwg), False)); time.sleep(3)
        tbs = ifc._find_all_title_blocks(doc)
        done = []
        for at, n, label in needs_fix(tbs):
            ifc._safe_set_text(at[f"{n}REV"], rev); done.append(f"{label} -> '{rev}'")
        doc.Save(); doc.Close(False); time.sleep(1)
        print(doc_id, "FIXED", done, "backup:", dst.name, flush=True)
    except Exception as e:
        print(doc_id, "ERR", e, flush=True)
        try:
            while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
        except Exception: pass
