# -*- coding: utf-8 -*-
"""Offline test: register-canonical grouping in VersionManager (PDF track), in BOTH
ifr_automation_v10.py and version_manager_v5.py. No AutoCAD, no client folders —
builds temp PDFs, injects a register via the _reg_titles_cache, and asserts a
re-worded revision of one drawing (SLD -> Single Line Diagram) now SUPERSEDES its
predecessor (the under-merge fix), while WITHOUT a register nothing merges (zero
regression) and a genuinely-different same-doc-id drawing stays split."""
import os, sys, time, tempfile, shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
results = []


def _mk(d, name, mtime):
    p = d / name
    p.write_text("x", encoding="utf-8")
    os.utime(p, (mtime, mtime))
    return p


def check(name, got, want):
    ok = got == want
    print(f"[{'PASS' if ok else 'FAIL'}] {name}")
    if not ok:
        print(f"    got={got!r} want={want!r}")
    results.append(ok)


def run(make_mgr, label):
    # 1. Drifted-title revs of ONE drawing + register present → old rev moves.
    tmp = Path(tempfile.mkdtemp(prefix="vmreg_"))
    try:
        d = tmp / "4. IFC(Client)"
        d.mkdir(parents=True)
        t0 = time.time() - 10000
        _mk(d, "50023-EL-001 Rev 0 SLD.pdf", t0)
        _mk(d, "50023-EL-001 Rev 1 Single Line Diagram.pdf", t0 + 100)
        mgr = make_mgr(tmp)
        mgr._reg_titles_cache = {str(d): {"50023-EL-001": ["Single Line Diagram"]}}
        mgr._version_strays = []
        ss = mgr._find_ss_folder(d)
        groups = mgr.scan_files_pdf(d) if hasattr(mgr, "scan_files_pdf") else mgr.scan_files(d)
        moves = mgr.identify_old_versions(groups, ss)
        moved = {Path(m[0]).name for m in moves}
        check(f"[{label}] drifted revs merge → Rev0 SLD superseded",
              moved, {"50023-EL-001 Rev 0 SLD.pdf"})
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 2. Same files, NO register → base_name differs → NO merge (zero regression).
    tmp = Path(tempfile.mkdtemp(prefix="vmreg_"))
    try:
        d = tmp / "4. IFC(Client)"
        d.mkdir(parents=True)
        t0 = time.time() - 10000
        _mk(d, "50023-EL-001 Rev 0 SLD.pdf", t0)
        _mk(d, "50023-EL-001 Rev 1 Single Line Diagram.pdf", t0 + 100)
        mgr = make_mgr(tmp)
        mgr._reg_titles_cache = {str(d): {}}       # register absent
        mgr._version_strays = []
        ss = mgr._find_ss_folder(d)
        groups = mgr.scan_files_pdf(d) if hasattr(mgr, "scan_files_pdf") else mgr.scan_files(d)
        moves = mgr.identify_old_versions(groups, ss)
        check(f"[{label}] no register → drifted titles NOT merged", len(moves), 0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 3. CA-001 Trench vs Cable Route (different drawings) → stay split, no move.
    tmp = Path(tempfile.mkdtemp(prefix="vmreg_"))
    try:
        d = tmp / "4. IFC(Client)"
        d.mkdir(parents=True)
        t0 = time.time() - 10000
        _mk(d, "50023-CA-001 Rev 1 Trench Alignment.pdf", t0)
        _mk(d, "50023-CA-001 Rev 1 Cable Route GA.pdf", t0 + 100)
        mgr = make_mgr(tmp)
        mgr._reg_titles_cache = {str(d): {"50023-CA-001":
                                          ["Trench Alignment Layout Plan", "Cable Route GA"]}}
        mgr._version_strays = []
        ss = mgr._find_ss_folder(d)
        groups = mgr.scan_files_pdf(d) if hasattr(mgr, "scan_files_pdf") else mgr.scan_files(d)
        moves = mgr.identify_old_versions(groups, ss)
        check(f"[{label}] CA-001 different drawings stay split (no move)", len(moves), 0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def run_native(make_native, label):
    """Native DWG track: a doc-id folder holding TWO real drawings (CA-001 =
    Trench + Cable) must keep BOTH newest revs; WITHOUT a register it collapses to
    one 'newest' = data loss (today's behavior, demonstrated as the control)."""
    def build(d):
        t0 = time.time() - 10000
        _mk(d, "50023-CA-001_Trench Alignment_RevA.dwg", t0)
        _mk(d, "50023-CA-001_Trench Alignment_RevB.dwg", t0 + 300)
        _mk(d, "50023-CA-001_Cable Route GA_RevA.dwg", t0 + 100)
        _mk(d, "50023-CA-001_Cable Route GA_RevB.dwg", t0 + 200)

    def moves_of(result):
        return {Path(a.source).name for a in result.actions if 'move' in a.action}

    reg = {"50023-CA-001": ["Trench Alignment Layout Plan", "Cable Route GA"]}

    # WITH register → split: each drawing keeps its RevB, only the two RevA move.
    tmp = Path(tempfile.mkdtemp(prefix="nvreg_"))
    try:
        folder = tmp / "50023-CA-001"
        folder.mkdir(parents=True)
        build(folder)
        mgr = make_native(tmp)
        mgr._reg_titles_cache = {str(folder): reg}
        res = mgr.process_folder(folder)
        renamed = {Path(a.source).name for a in res.actions if a.action == 'rename'}
        check(f"[{label}] CA-001 two drawings: only the two RevA superseded",
              moves_of(res), {"50023-CA-001_Trench Alignment_RevA.dwg",
                              "50023-CA-001_Cable Route GA_RevA.dwg"})
        check(f"[{label}] multi-drawing folder → standardize-rename skipped",
              renamed, set())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # WITHOUT register → collapses to one newest: a real drawing's RevB is lost.
    tmp = Path(tempfile.mkdtemp(prefix="nvreg_"))
    try:
        folder = tmp / "50023-CA-001"
        folder.mkdir(parents=True)
        build(folder)
        mgr = make_native(tmp)
        mgr._reg_titles_cache = {str(folder): {}}
        res = mgr.process_folder(folder)
        mv = moves_of(res)
        check(f"[{label}] no register → 3 files moved (control: data-loss shape)",
              len(mv), 3)
        check(f"[{label}] no register → a real drawing's newest (Cable RevB) is superseded",
              "50023-CA-001_Cable Route GA_RevB.dwg" in mv, True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


import ifr_automation_v10 as ENG
run(lambda tmp: ENG.VersionManager(str(tmp), dry_run=True), "engine")
run_native(lambda tmp: ENG.NativeVersionManager(str(tmp), dry_run=True), "engine-native")

import version_manager_v5 as V5
run(lambda tmp: V5.VersionManager(str(tmp), dry_run=True), "v5")
run_native(lambda tmp: V5.NativeVersionManager(str(tmp), dry_run=True), "v5-native")

print()
print("ALL PASS" if all(results) else "SOME FAILED")
sys.exit(0 if all(results) else 1)
