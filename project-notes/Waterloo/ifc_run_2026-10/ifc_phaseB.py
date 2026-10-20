# Phase B: IFC-convert the 16 approved drawings with the tool's IFCManager (no script change).
# Explicit source map: foundation plans live in C-RPT report folders; 002-1/2/3 share one regex id.
import sys, json
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
E = P / "Design/Engineering"
N = E / "1. Drawings/1. Native"
R = E / "2. Calcs & Reports/Reports/Civil & Structure"
JOBS = [  # doc_id, IFR description (as submitted), source DWG, native-folder cleanup?
 ("GG38-C-PLN-001", "Civil Site Layout", N/"GG38-C-PLN-001_Civil Site Layout Plan_Rev/GG38-C-PLN-001_rA.dwg", "A", True),
 ("GG38-C-PLN-002-1", "Road&Hardstand Coordinate", N/"GG38-C-PLN-002_Site Coordinates design (Civil Scope)_Rev/GG38-C-PLN-002-CS_rB.dwg", "B", True),
 ("GG38-C-PLN-002-2", "Equipment foundation Coordinate", N/"GG38-C-PLN-002_Site Coordinates design (Electrical Scope)_Rev/GG38-C-PLN-002-ES_rA.dwg", "A", True),
 ("GG38-C-PLN-002-3", "Fence & Gate Coordinate", N/"GG38-C-PLN-002_Site Coordinates design (Fencing Scope)_Rev/GG38-C-PLN-002-FNC_rA.dwg", "A", True),
 ("GG38-C-PLN-003", "Fence & Gate Design", N/"GG38-C-PLN-003_Fence & Gate Layout Plan_Rev/GG38-C-PLN-003_Fence & Gate Layout Plan_RevA.dwg", "A", True),
 ("GG38-C-PLN-005", "Trench Alignment Layout Plan", N/"GG38-C-PLN-005_Trench Alignment Layout Plan_Rev/GG38-C-PLN-005_Trench Alignment Layout Plan_RevB.dwg", "B", True),
 ("GG38-C-PLN-006", "MVPS Foundation Design", R/"GG38-C-RPT-001_MV Power Station (MVPS) Foundation Report_Rev/GG38-C-PLN-006_MV Power Station (MVPS) Foundation Plan_RevA.dwg", "A", False),
 ("GG38-C-PLN-008", "HVSB & Platform Foundation Design", R/"GG38-C-RPT-003_HVSB Platform & Footing Report_Rev/GG38-C-PLN-008_HVSB Foundation Plan_RevC.dwg", "C", False),
 ("GG38-C-PLN-009", "BESS Foundation Design", R/"GG38-C-RPT-006_BESS Foundation Design Report_Rev/GG38-C-PLN-009_BESS Foundation_RevE.dwg", "E", False),
 ("GG38-C-PLN-013", "Light & CCTV Pole Foundation Design", R/"GG38-C-RPT-010_Light & CCTV Pole Foundation Design Report_Rev/GG38-C-PLN-013_Light Pole Foundation_RevA.dwg", "A", False),
 ("GG38-C-PLN-016", "Aux Transformer Foundation Design", R/"GG38-C-RPT-014_Aux Transformer Foundation Design Report_Rev/GG38-C-PLN-016_AUX Transformer Foundation Plan_RevB.dwg", "B", False),
 ("GG38-C-PLN-017", "Kerb Layout Plan", N/"GG38-C-PLN-017_Kerb Layout Plan_Rev/GG38-C-PLN-017_Kerb Layout Plan_RevA.dwg", "A", True),
 ("GG38-E-BLD-001", "Auxiliary Cable Block Diagram", N/"GG38-E-BLD-001_Auxiliary Cable Block Diagram_Rev/GG38-E-BLD-001_Rev.A - 344 Waterloo Rd - Auxiliary Block Diagram.dwg", "A", True),
 ("GG38-E-BLD-003", "Communications Block Diagram", N/"GG38-E-BLD-003_Communication Block Diagram_Rev/GG38-E-BLD-003_Rev.A - 344 Waterloo Rd - Comms Block Diagram - Rev.A.dwg", "A", True),
 ("GG38-E-PLN-001", "DC Cable Route", N/"GG38-E-PLN-001_DC Cable Route Layout Plan_Rev/GG38-E-PLN-001-rB.dwg", "B", True),
 ("GG38-E-SLD-001", "Single Line Diagram", N/"GG38-E-SLD-001_Single Line Diagram_Rev/GG38-E-SLD-001_Single Line Diagram_RevD.dwg", "D", True),
]
only = set(sys.argv[1:])
import subprocess, os
def short(dwg):  # >250-char source: open via short junction to its folder (same trick as AsBuiltManager)
    if len(str(dwg)) <= 250: return dwg
    j = Path(os.environ.get("TEMP", r"C:\Temp")) / ("ifcj_" + __import__("hashlib").md5(str(dwg.parent).encode()).hexdigest()[:8])
    if not j.exists(): subprocess.run(["cmd", "/c", "mklink", "/J", str(j), str(dwg.parent)], check=True, capture_output=True)
    return j / dwg.name
am = v.ApprovedIFCManager(P, dry_run=False)
ifc = v.IFCManager(P, dry_run=False, native_root=am.native_root, ifc_output=am.ifc_output, preserve_ifr=True)
rp = Path(__file__).with_name("phaseB_results.json")
out = json.load(open(rp)) if (only and rp.exists()) else []
for i, (doc, desc, dwg, rev, clean) in enumerate(JOBS, 1):
    if only and doc not in only: continue
    dwg = short(dwg)
    assert dwg.exists(), dwg
    info = {'doc_id': doc, 'folder': dwg.parent, 'latest_ifr_dwg': dwg, 'latest_ifr_rev': rev,
            'existing_ifc_rev': None, 'needs_ifc': True, 'description': desc,
            'all_ifr_dwgs': [], 'all_ifc_dwgs': [], 'other_dwgs': []}
    if clean:  # tool's own scan -> move stale revs/_recover copies of THIS folder to SS
        s = am._scan_native_dwgs(dwg.parent, doc) or {}
        info['all_ifr_dwgs'] = s.get('all_ifr_dwgs', []); info['other_dwgs'] = s.get('other_dwgs', [])
        print(f"[{i}/16] {doc} cleanup:", am._cleanup_native_folder(info), flush=True)
    print(f"[{i}/16] {doc}: {dwg.name} Rev{rev} -> IFC Rev0", flush=True)
    r = ifc.convert_to_ifc(info)
    print(f"   success={r['success']} pdf={r['pdf_path']} err={r['errors']}", flush=True)
    out.append({k: (str(x) if isinstance(x, Path) else x) for k, x in r.items()})
    if not r['success']:
        try:
            a = ifc._get_acad()
            while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
        except Exception: ifc._acad = None
json.dump(out, open(Path(__file__).with_name("phaseB_results.json"), "w"), indent=1)
print("DONE ok=%d fail=%d" % (sum(r['success'] for r in out), sum(not r['success'] for r in out)), flush=True)
