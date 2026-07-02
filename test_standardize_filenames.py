# -*- coding: utf-8 -*-
"""Offline test for the standardize_filenames pipeline step: dry-run never touches
disk; --apply renames only the tidy-able files; collision + existing-target are
skipped (never merged/overwritten); case-only rename works on a case-insensitive
FS. No AutoCAD, no client folders — temp dirs only."""
import sys, tempfile, shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import standardize_filenames as S

results = []


def check(name, ok):
    print(f"[{'PASS' if ok else 'FAIL'}] {name}")
    results.append(ok)


def mk(d, name):
    (d / name).write_text("x", encoding="utf-8")


def names(d):
    return sorted(p.name for p in d.iterdir() if p.is_file())


# 1. DRY-RUN leaves disk untouched but reports the planned rename.
tmp = Path(tempfile.mkdtemp(prefix="stdfn_"))
try:
    mk(tmp, "50023-EL-001 REV 1 as built.pdf")
    before = names(tmp)
    renames, skips = S.plan_renames(tmp)
    S.run(tmp, apply=False)
    check("dry-run: disk unchanged", names(tmp) == before)
    check("dry-run: one rename planned", len(renames) == 1 and len(skips) == 0)
    check("dry-run: target is standardized form",
          renames[0][1].name == "50023-EL-001 Rev 1_AS BUILT.pdf")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# 2. --apply performs the rename; an already-tidy file is left alone.
tmp = Path(tempfile.mkdtemp(prefix="stdfn_"))
try:
    mk(tmp, "50023-EL-001 REV 1.pdf")            # tidy-able
    mk(tmp, "50023-EL-002 Rev 2.pdf")            # already standard
    S.run(tmp, apply=True)
    check("apply: messy file renamed", "50023-EL-001 Rev 1.pdf" in names(tmp))
    check("apply: tidy file untouched", "50023-EL-002 Rev 2.pdf" in names(tmp))
    check("apply: no leftover of old name",
          "50023-EL-001 REV 1.pdf" not in names(tmp))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# 3. COLLISION: two files normalize to the same name → BOTH skipped, none renamed.
tmp = Path(tempfile.mkdtemp(prefix="stdfn_"))
try:
    mk(tmp, "50023-EL-001 REV 1.pdf")
    mk(tmp, "50023-EL-001  rev.1.pdf")           # both -> '...Rev 1.pdf'
    before = names(tmp)
    renames, skips = S.plan_renames(tmp)
    S.run(tmp, apply=True)
    check("collision: nothing renamed", names(tmp) == before)
    check("collision: both reported as skips", len(renames) == 0 and len(skips) == 2)
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# 4. EXISTING TARGET: normalized name already taken by a DIFFERENT file → skip.
# (Uses a double-space vs single-space pair — on a case-insensitive FS a pair
# differing only in case can't coexist, but these two genuinely can.)
tmp = Path(tempfile.mkdtemp(prefix="stdfn_"))
try:
    mk(tmp, "50023-EL-001  Rev 1.pdf")           # 2 spaces -> '...Rev 1.pdf'
    mk(tmp, "50023-EL-001 Rev 1.pdf")            # already occupies the target
    before = names(tmp)
    renames, skips = S.plan_renames(tmp)
    S.run(tmp, apply=True)
    check("existing target: nothing renamed/overwritten", names(tmp) == before)
    check("existing target: reported as skip",
          len(renames) == 0 and len(skips) == 1)
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# 5. CASE-ONLY rename (e.g. '_as built' -> '_AS BUILT') succeeds on any FS.
tmp = Path(tempfile.mkdtemp(prefix="stdfn_"))
try:
    mk(tmp, "50023-EL-001 Rev 1_as built.pdf")
    S.run(tmp, apply=True)
    check("case-only: renamed to uppercase suffix",
          "50023-EL-001 Rev 1_AS BUILT.pdf" in names(tmp))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print()
print("ALL PASS" if all(results) else "SOME FAILED")
sys.exit(0 if all(results) else 1)
