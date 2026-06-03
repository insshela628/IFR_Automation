"""
Robust AS BUILT batch engine — hang-proof. CLI + importable by the bot.

Why: driving AutoCAD via COM can wedge indefinitely (modal dialog, PUBLISH stall,
RPC_E_CALL_REJECTED, a heavy file that crashes AutoCAD). A single hung/crashed file
must NOT take down the whole batch. This converts ONE source per child subprocess
with a hard per-file timeout; on hang/crash the child is killed, AutoCAD is force-
restarted (and its .sv$ recovery files cleared), and the batch continues.

Two entry points share one engine (`run_safe_batch`):
  - CLI:  python run_ab_batch_safe.py cole2 [filter] [--timeout N] [--rev1]
  - Bot:  from run_ab_batch_safe import run_safe_batch
          run_safe_batch(project_path, force_rev_1=True, on_event=cb)

`run_safe_batch(project_path, terms=None, timeout=240, force_rev_1=False,
on_event=None)` returns {total, ok, fail, warn, results:[{doc_id, ok, pdf, errs,
qa, seconds, ...}]} and streams per-file dicts to `on_event` (for Telegram
progress). NOT project-specific — works for any project_path.
"""
import sys, os, json, time, subprocess
from pathlib import Path

# Children emit UTF-8 (Chinese log lines). Force UTF-8 everywhere so the parent
# can decode child stdout and print/aggregate without cp1252 crashes.
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

_CHILD_ENV = dict(os.environ, PYTHONIOENCODING='utf-8')

PY = r"C:\Users\jilin\AppData\Local\Programs\Python\Python312\python.exe"
SCRIPT_DIR = r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√"

# CLI convenience aliases only — run_safe_batch takes a full project_path.
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
    tmp = os.environ.get("TEMP") or os.environ.get("TMP") or r"C:\Windows\Temp"
    try:
        for pat in ("*.sv$", "*.dwk"):
            for f in Path(tmp).glob(pat):
                try: f.unlink()
                except Exception: pass
    except Exception:
        pass
    if reason:
        print(f"    [recover] killed AutoCAD + cleared recovery files ({reason})",
              flush=True)


def _worker(project_path, doc_id, src_sub, force_rev_1=True, force_redo=True):
    """Child: convert exactly ONE source. Prints one JSON result line.

    force_rev_1: lock output to REV 1. force_redo: convert even if an AS BUILT PDF
    already exists (bypass the incremental skip). force_rev_1 implies force_redo.
    """
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
        # Incremental skip (preserve batch_convert behaviour) unless forced.
        if not (force_rev_1 or force_redo) and mgr._check_ab_incremental(target):
            print(json.dumps({"doc_id": doc_id, "src": src_sub,
                              "ok": True, "skipped": True}))
            return
        if force_rev_1:
            target['existing_ab_rev'] = None   # lock output to REV 1
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


def run_safe_batch(project_path, terms=None, timeout=DEFAULT_TIMEOUT,
                   force_rev_1=False, force_redo=False, on_event=None):
    """Hang-proof batch over a project. Returns a summary dict; streams per-file
    dicts to on_event(ev). Cross-project: only needs project_path.

    terms: substring filter (str CSV or list) — convert only matching doc-ids.
    force_rev_1: lock output to REV 1. force_redo: ignore the incremental skip.
    A `terms` filter implies force_redo (you asked for those specifically).
    on_event: {"type":"start","total"} | {"type":"file",...,"ok","pdf","qa","err",
    "skipped","seconds"} | {"type":"done","total","ok","fail","warn","skipped"}.
    """
    def emit(ev):
        if on_event:
            try: on_event(ev)
            except Exception: pass

    sources = list_sources(project_path)
    if terms:
        tl = terms if isinstance(terms, list) else [t for t in terms.split(',')]
        tl = [t.strip().lower() for t in tl if t.strip()]
        if tl:
            sources = [s for s in sources
                       if any(t in s[0].lower() or t in s[1].lower() for t in tl)]
            force_redo = True   # an explicit filter means "do these now"
    emit({"type": "start", "total": len(sources)})

    results = []
    for idx, (doc_id, src_name) in enumerate(sources, 1):
        # Fresh AutoCAD per file (chaining workers against one warm instance left
        # COM dirty and most conversions failed). kill_acad also clears .sv$.
        kill_acad()
        cmd = [PY, os.path.join(SCRIPT_DIR, "run_ab_batch_safe.py"),
               "--worker", project_path, doc_id, src_name,
               "1" if force_rev_1 else "0", "1" if force_redo else "0"]
        t0 = time.time()
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  encoding='utf-8', errors='replace',
                                  env=_CHILD_ENV, timeout=timeout)
            jline = [l for l in proc.stdout.splitlines() if l.startswith('{')]
            res = json.loads(jline[-1]) if jline else {
                "doc_id": doc_id, "src": src_name, "ok": False,
                "err": "no result line", "stdout_tail": proc.stdout[-200:]}
        except subprocess.TimeoutExpired:
            kill_acad(f"{doc_id} hung")
            res = {"doc_id": doc_id, "src": src_name, "ok": False,
                   "err": f"timeout >{timeout}s"}
        res["seconds"] = round(time.time() - t0)
        res["idx"] = idx
        results.append(res)
        emit({"type": "file", "total": len(sources), **res})

    skipped = sum(1 for r in results if r.get('skipped'))
    ok = sum(1 for r in results if r.get('ok') and not r.get('skipped'))
    warn = sum(1 for r in results if r.get('ok') and r.get('qa'))
    fail = sum(1 for r in results if not r.get('ok'))
    summary = {"total": len(results), "ok": ok, "fail": fail, "warn": warn,
               "skipped": skipped, "results": results}
    emit({"type": "done", "total": len(results), "ok": ok, "fail": fail,
          "warn": warn, "skipped": skipped})
    return summary


def main():
    args = sys.argv[1:]
    if args and args[0] == '--worker':
        _worker(args[1], args[2], args[3] if len(args) > 3 else "",
                (args[4] == "1") if len(args) > 4 else True,
                (args[5] == "1") if len(args) > 5 else True)
        return

    project = args[0] if args else "cole2"
    timeout = DEFAULT_TIMEOUT
    filt = ""
    force_rev_1 = False
    force_redo = False
    i = 1
    while i < len(args):
        if args[i] == '--timeout':
            timeout = int(args[i + 1]); i += 2
        elif args[i] == '--rev1':
            force_rev_1 = True; i += 1
        elif args[i] == '--redo':
            force_redo = True; i += 1
        else:
            filt = args[i]; i += 1

    project_path = PROJECTS.get(project, project)
    print(f"=== SAFE AS BUILT BATCH: {project} "
          f"(per-file timeout={timeout}s, rev1={force_rev_1}) ===", flush=True)

    def on_event(ev):
        if ev["type"] == "start":
            print(f"Sources to convert: {ev['total']}", flush=True)
        elif ev["type"] == "file":
            if ev.get('skipped'):
                tag = "SKIP"
            elif str(ev.get('err', '')).startswith('timeout'):
                tag = "TIMEOUT"
            elif ev.get('ok') and not ev.get('qa'):
                tag = "PASS"
            elif ev.get('ok'):
                tag = "WARN"
            else:
                tag = "FAIL"
            detail = ev.get('pdf') or ev.get('err') or ''
            if ev.get('qa'):
                detail += " QA:" + "; ".join(ev['qa'])
            print(f"[{ev['idx']}/{ev['total']}] {ev['doc_id']}\n"
                  f"    [{tag}] {ev['seconds']}s {detail}", flush=True)
        elif ev["type"] == "done":
            print(f"\n=== DONE: ok={ev['ok']} (warn={ev['warn']}) "
                  f"fail={ev['fail']} skip={ev['skipped']} of {ev['total']} ===",
                  flush=True)

    summary = run_safe_batch(project_path, filt or None, timeout,
                             force_rev_1, force_redo, on_event)
    for r in summary["results"]:
        if not r.get('ok'):
            print(f"  FAIL {r['doc_id']}: {r.get('err') or r.get('errs')}", flush=True)
        elif r.get('qa'):
            print(f"  WARN {r['doc_id']}: {'; '.join(r['qa'])}", flush=True)


if __name__ == "__main__":
    main()
