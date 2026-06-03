"""
Robust AS BUILT batch runner — hang-proof.

Why this exists: driving AutoCAD via COM can wedge indefinitely (modal dialog,
PUBLISH stall, RPC_E_CALL_REJECTED cascade). A single hung file must NOT take
down the whole batch. This runner converts ONE source per child subprocess with
a hard per-file timeout; if a file hangs, the child is killed, AutoCAD is force-
restarted, and the batch continues with the next file.

Usage:
    python run_ab_batch_safe.py cole2            # all sources
    python run_ab_batch_safe.py cole2 PLN-005    # filter doc-id substring
    python run_ab_batch_safe.py cole2 --timeout 300

Each child writes a one-line JSON result; the parent aggregates and prints a
summary, then runs read-only PDF QA on the output folder.
"""
import sys, os, json, time, subprocess
from pathlib import Path

# Children emit UTF-8 (Chinese log lines). Force UTF-8 everywhere so the parent
# can decode child stdout and print summaries without cp1252 crashes.
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

_CHILD_ENV = dict(os.environ, PYTHONIOENCODING='utf-8')

PY = r"C:\Users\jilin\AppData\Local\Programs\Python\Python312\python.exe"
SCRIPT_DIR = r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√"

PROJECTS = {
    "cole2": r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\1.NSW\Coleambally #2",
    "warnertown": r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS",
    "lms": r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\LMS",
}

DEFAULT_TIMEOUT = 240   # seconds per single-source conversion
ACAD_SETTLE = 4         # seconds to wait after killing AutoCAD


def kill_acad(reason=""):
    """Force-kill all AutoCAD instances, clear the autosave/recovery files the
    kill leaves behind, and wait for a clean state.

    Each forced kill drops a .sv$ autosave; they accumulate and make the NEXT
    launch pop the Drawing Recovery dialog → COM hangs (even trivial files time
    out). Clearing them here keeps repeated kill/relaunch cycles healthy."""
    subprocess.run(["taskkill", "/F", "/IM", "acad.exe"],
                   capture_output=True, text=True)
    time.sleep(ACAD_SETTLE)
    # Remove leftover autosave/recovery files so the next launch starts clean.
    tmp = os.environ.get("TEMP") or os.environ.get("TMP") or r"C:\Windows\Temp"
    try:
        for f in Path(tmp).glob("*.sv$"):
            try: f.unlink()
            except Exception: pass
        for f in Path(tmp).glob("*.dwk"):
            try: f.unlink()
            except Exception: pass
    except Exception:
        pass
    if reason:
        print(f"    [recover] killed AutoCAD + cleared recovery files ({reason})",
              flush=True)


# Child worker: convert exactly ONE source (doc_id + source_dir substring).
# Run via `python run_ab_batch_safe.py --worker <project_path> <doc_id> <srcsub>`.
def _worker(project_path, doc_id, src_sub):
    import pythoncom, logging
    logging.basicConfig(level=logging.ERROR)
    sys.path.insert(0, SCRIPT_DIR)
    pythoncom.CoInitialize()
    try:
        from ifr_automation_v10 import AsBuiltManager
        mgr = AsBuiltManager(project_path, dry_run=False)
        scan = mgr.scan_native_for_ab()
        target = None
        for s in scan:
            if s['doc_id'].upper() == doc_id.upper() and \
               (not src_sub or src_sub.lower() in s['ifc_source']['source_dir'].name.lower()):
                target = s
                break
        if target is None:
            print(json.dumps({"doc_id": doc_id, "src": src_sub,
                              "ok": False, "err": "source not found"}))
            return
        target['existing_ab_rev'] = None  # force REV 1
        r = mgr._convert_with_qa_retry(target)
        print(json.dumps({
            "doc_id": doc_id, "src": src_sub,
            "ok": bool(r.get('success')),
            "pdf": Path(r['pdf_path']).name if r.get('pdf_path') else None,
            "errs": r.get('errors', [])[:2],
            "qa": r.get('post_qa', [])[:3],
        }))
    finally:
        pythoncom.CoUninitialize()


def list_sources(project_path):
    """Scan sources in a short-lived subprocess (no AutoCAD)."""
    code = (
        "import sys; sys.path.insert(0, r'%s');"
        "from ifr_automation_v10 import AsBuiltManager; import json;"
        "m=AsBuiltManager(r'%s', dry_run=True);"
        "print(json.dumps([[s['doc_id'], s['ifc_source']['source_dir'].name] "
        "for s in m.scan_native_for_ab()]))"
    ) % (SCRIPT_DIR, project_path)
    r = subprocess.run([PY, "-c", code], capture_output=True, text=True,
                       encoding='utf-8', errors='replace', env=_CHILD_ENV, timeout=120)
    line = [l for l in r.stdout.splitlines() if l.startswith('[')]
    return json.loads(line[-1]) if line else []


def main():
    args = sys.argv[1:]
    if args and args[0] == '--worker':
        _worker(args[1], args[2], args[3] if len(args) > 3 else "")
        return

    project = args[0] if args else "cole2"
    timeout = DEFAULT_TIMEOUT
    filt = ""
    i = 1
    while i < len(args):
        if args[i] == '--timeout':
            timeout = int(args[i + 1]); i += 2
        else:
            filt = args[i]; i += 1

    project_path = PROJECTS.get(project, project)
    print(f"=== SAFE AS BUILT BATCH: {project} (per-file timeout={timeout}s) ===", flush=True)

    sources = list_sources(project_path)
    if filt:
        # filt may be a comma-separated list of substrings; match if ANY term hits
        terms = [t.strip().lower() for t in filt.split(',') if t.strip()]
        sources = [s for s in sources
                   if any(t in s[0].lower() or t in s[1].lower() for t in terms)]
    print(f"Sources to convert: {len(sources)}", flush=True)

    results = []
    for idx, (doc_id, src_name) in enumerate(sources, 1):
        # Pass the FULL source-dir name so multi-scope doc-ids (PLN-002 Civil /
        # Equipment / Fencing — identical first 40 chars) are pinned exactly.
        src_sub = src_name
        print(f"[{idx}/{len(sources)}] {doc_id} | {src_name[:50]} ...", flush=True)
        # Fresh AutoCAD per file: chaining workers against one warm instance left
        # COM in a dirty state and most conversions failed. Killing first makes
        # every file run in the same clean condition as a verified single run.
        kill_acad()
        cmd = [PY, os.path.join(SCRIPT_DIR, "run_ab_batch_safe.py"),
               "--worker", project_path, doc_id, src_sub]
        t0 = time.time()
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  encoding='utf-8', errors='replace',
                                  env=_CHILD_ENV, timeout=timeout)
            jline = [l for l in proc.stdout.splitlines() if l.startswith('{')]
            res = json.loads(jline[-1]) if jline else {
                "doc_id": doc_id, "src": src_sub, "ok": False,
                "err": "no result line", "stdout_tail": proc.stdout[-200:]}
            dt = time.time() - t0
            tag = "PASS" if res.get('ok') and not res.get('qa') else \
                  ("WARN" if res.get('ok') else "FAIL")
            print(f"    [{tag}] {dt:.0f}s "
                  f"{res.get('pdf') or res.get('err') or ''}"
                  f"{' QA:' + '; '.join(res.get('qa', [])) if res.get('qa') else ''}",
                  flush=True)
            results.append(res)
        except subprocess.TimeoutExpired:
            print(f"    [TIMEOUT >{timeout}s] killing AutoCAD, recovering...", flush=True)
            kill_acad(f"{doc_id} hung")
            results.append({"doc_id": doc_id, "src": src_sub,
                            "ok": False, "err": f"timeout >{timeout}s"})

    ok = sum(1 for r in results if r.get('ok'))
    warn = sum(1 for r in results if r.get('ok') and r.get('qa'))
    fail = len(results) - ok
    print(f"\n=== DONE: ok={ok} (warn={warn}) fail={fail} of {len(results)} ===", flush=True)
    for r in results:
        if not r.get('ok'):
            print(f"  FAIL {r['doc_id']}: {r.get('err') or r.get('errs')}", flush=True)
        elif r.get('qa'):
            print(f"  WARN {r['doc_id']}: {'; '.join(r['qa'])}", flush=True)


if __name__ == "__main__":
    main()
