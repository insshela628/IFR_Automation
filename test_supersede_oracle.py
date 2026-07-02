# -*- coding: utf-8 -*-
"""Offline test for the DLV-authority AS BUILT collision gate
(AsBuiltManager.supersede_ab_deliverables + _ab_title_reconciles).

No AutoCAD, no client folders — builds fake deliverable dirs in a temp folder,
injects a controlled register via the _ab_reg_titles cache, and asserts the
membership-oracle outcomes: SLD synonym drift auto-merges, 50023-CA-001
Trench≠Cable-Route stay split, a title matching no register name is flagged
(anti-data-loss), no register → conservative wholesale defer, single-title
doc-ids keep legacy behavior. report_only=True → nothing is actually moved.

Run: python test_supersede_oracle.py  (exit 0 = all pass).
"""
import os
import sys
import time
import shutil
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ifr_automation_v10 as M


def _mk(d: Path, name: str, mtime: float):
    p = d / name
    p.write_text("x", encoding="utf-8")
    os.utime(p, (mtime, mtime))
    return p


def run_case(name, files, register, expect_moved, expect_flag_docs,
             expect_stray_docs):
    tmp = Path(tempfile.mkdtemp(prefix="abtest_"))
    try:
        ab = tmp / "5. As Built" / "3. As Built Client"
        ab.mkdir(parents=True)
        t0 = time.time() - 10000
        for fn, dt in files:
            _mk(ab, fn, t0 + dt)

        mgr = M.AsBuiltManager.__new__(M.AsBuiltManager)  # bypass COM __init__
        mgr._ab_output_override = ab                      # drives ab_output property
        mgr._ab_reg_titles = {k.upper(): v for k, v in register.items()}  # inject

        out = mgr.supersede_ab_deliverables(report_only=True)
        moved = {Path(r["path"]).name for r in out["moved"]}
        flag_docs = {f["doc_id"] for f in out["flagged"] if "variants" in f}
        # stray vs deferred both carry 'variants'; distinguish by the reason text
        stray_docs = {f["doc_id"] for f in out["flagged"]
                      if "variants" in f and "对不上客户清单" in f.get("reason", "")}

        ok = (moved == set(expect_moved)
              and flag_docs == set(expect_flag_docs)
              and stray_docs == set(expect_stray_docs))
        print(f"[{'PASS' if ok else 'FAIL'}] {name}")
        if not ok:
            print(f"    moved      got={sorted(moved)} want={sorted(expect_moved)}")
            print(f"    flag_docs  got={sorted(flag_docs)} "
                  f"want={sorted(expect_flag_docs)}")
            print(f"    stray_docs got={sorted(stray_docs)} "
                  f"want={sorted(expect_stray_docs)}")
        return ok
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    results = []

    # 1. SLD synonym drift → auto-merge + version (Rev0 SLD older → superseded).
    results.append(run_case(
        "SLD == Single Line Diagram auto-merges (no human)",
        files=[
            ("50023-EL-001 REV 0 SLD_AS BUILT.pdf", 0),
            ("50023-EL-001 REV 1 Single Line Diagram_AS BUILT.pdf", 100),
        ],
        register={"50023-EL-001": ["Single Line Diagram"]},
        expect_moved=["50023-EL-001 REV 0 SLD_AS BUILT.pdf"],
        expect_flag_docs=[],
        expect_stray_docs=[],
    ))

    # 2. CA-001 same number / different drawings → split, versioned independently.
    #    Trench has 2 revs (Rev0 old → SS); Cable Route single rev untouched.
    results.append(run_case(
        "CA-001 Trench != Cable Route stay split, each versioned",
        files=[
            ("50023-CA-001 REV 0 Trench Alignment_AS BUILT.pdf", 0),
            ("50023-CA-001 REV 1 Trench Alignment_AS BUILT.pdf", 100),
            ("50023-CA-001 REV 1 Cable Route GA_AS BUILT.pdf", 50),
        ],
        register={"50023-CA-001": ["Trench Alignment Layout Plan",
                                   "Cable Route GA"]},
        expect_moved=["50023-CA-001 REV 0 Trench Alignment_AS BUILT.pdf"],
        expect_flag_docs=[],
        expect_stray_docs=[],
    ))

    # 3. No register for a multi-title doc-id → conservative wholesale defer.
    results.append(run_case(
        "No register -> wholesale defer (nothing moved, flagged)",
        files=[
            ("50023-EA-999 REV 1 Foo Layout_AS BUILT.pdf", 0),
            ("50023-EA-999 REV 1 Bar Diagram_AS BUILT.pdf", 100),
        ],
        register={},
        expect_moved=[],
        expect_flag_docs=["50023-EA-999"],
        expect_stray_docs=[],
    ))

    # 4. Stray: one title reconciles to register, a foreign title matches nothing.
    results.append(run_case(
        "Stray title (matches no register name) -> flagged, not moved",
        files=[
            ("50023-XX-001 REV 0 Alpha Plan_AS BUILT.pdf", 0),
            ("50023-XX-001 REV 1 Alpha Plan_AS BUILT.pdf", 100),
            ("50023-XX-001 REV 1 Zeta Section_AS BUILT.pdf", 50),
        ],
        register={"50023-XX-001": ["Alpha Plan"]},
        expect_moved=["50023-XX-001 REV 0 Alpha Plan_AS BUILT.pdf"],
        expect_flag_docs=["50023-XX-001"],
        expect_stray_docs=["50023-XX-001"],
    ))

    # 5. Single-title doc-id (no collision) → legacy behavior, register irrelevant.
    results.append(run_case(
        "Single-title doc-id versions normally (register not required)",
        files=[
            ("50023-GA-001 REV 0 Site Plan_AS BUILT.pdf", 0),
            ("50023-GA-001 REV 1 Site Plan_AS BUILT.pdf", 100),
        ],
        register={},
        expect_moved=["50023-GA-001 REV 0 Site Plan_AS BUILT.pdf"],
        expect_flag_docs=[],
        expect_stray_docs=[],
    ))

    print()
    ok = all(results)
    print("ALL PASS" if ok else "SOME FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
