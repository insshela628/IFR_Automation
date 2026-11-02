"""Native-folder tidy for ONE project — the daily patrol's real-write slice.

    python native_tidy_cli.py "<project path>"            # plan only
    python native_tidy_cli.py "<project path>" --apply    # move (reversible)

Scope = asbuilt-native-folder-standard, nothing more:
  * dead-versions folders → ONE `Superseded` (SS / Superseeded / OLD REVISION …)
  * loose junk (.bak .dwl .log .err .sv$ ~$) → Superseded/
  * older revs below the loose master, byte-identical "- Copy" DWGs → Superseded/
  * drawings OPEN right now (fresh .dwl) are skipped entirely
  * DWG/BAK sitting loose at `1. Drawings` root (outside Native) → FLAG only:
    it has no doc number yet, a human assigns one.
Everything moves, nothing is deleted. No AutoCAD.
Last stdout line = `RESULT_JSON: {...}` (machine contract, same as run_pipeline_cli).
"""
import json
import sys
from pathlib import Path

from ifr_automation_v10 import AsBuiltManager


def main() -> int:
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 2
    proj = Path(args[0])
    apply = "--apply" in args
    out = {"status": "ok", "project": proj.name, "apply": apply,
           "moved": [], "flagged": []}
    try:
        mgr = AsBuiltManager(str(proj), dry_run=not apply)
        acts = mgr.canonicalize_dead_dirs(report_only=not apply)
        acts += mgr.cleanup_native_loose(report_only=not apply)
        for a in acts:
            rec = {"doc_id": a.get("doc_id"), "kind": a.get("kind"),
                   "name": Path(str(a.get("path") or a.get("dir") or "")).name,
                   "why": a.get("why", "")}
            (out["flagged"] if a.get("flag") else out["moved"]).append(rec)
        drawings = mgr.native_root.parent
        for f in sorted(drawings.iterdir()) if drawings.is_dir() else []:
            if f.is_file() and f.suffix.lower() == ".dwg":
                out["flagged"].append({"doc_id": None, "kind": "unfiled_dwg",
                                       "name": f.name,
                                       "why": "loose at 1. Drawings root - needs a doc number"})
    except Exception as e:  # never crash the patrol
        out = {"status": "error", "message": f"{type(e).__name__}: {e}"}
    print("RESULT_JSON: " + json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
