# -*- coding: utf-8 -*-
"""Pipeline step: standardize deliverable FILENAME FORMAT (case/space/underscore/
Rev) across a folder — the execute side of the format PREVIEW that every collision
gate already prints.

WHY a separate step (not folded into the version gates): renaming real client
files is a deliberate, per-project action, so it lives behind an explicit --apply
flag and DEFAULTS TO DRY-RUN. The gates only ever PREVIEW; this is the one place
that can actually rename, and only when a human runs it with --apply.

SAFETY:
  • Uses register_membership.normalize_filename_format — IDENTITY-PRESERVING
    (case/space/Rev only, description words untouched), so it can never collapse
    two genuinely-different drawings into one. No register lookup needed.
  • COLLISION GUARD: if two source files normalize to the same name, or the target
    name already exists as a different file, BOTH are skipped and reported — never
    overwrite, never merge.
  • Case-only renames (Windows case-insensitive FS) go through a temp name.
  • Dry-run by default; --apply required to touch the disk.

MODES (`--mode`):
  format  (default) register_membership.normalize_filename_format -- house style
          "Rev 1" with SPACE separators. Matches the GG-31 / AS BUILT corpus.
  rev     DeliverableManager.normalize_filename with NO deliverable_desc, and
          only the changes that touch the REVISION TOKEN are kept. Target is
          "_RevA": measured 2026-09-23 across Waterloo + Forbes that is the
          de-facto plurality (9 of 27 versioned deliverables; `_rA` 6, `_Rev.A`
          6, `-rA` 1). Use this on a project whose files are underscore-style --
          running `format` there would rewrite the ALREADY-CORRECT `_RevA` files
          into ` Rev A` and churn the whole folder.
  ⚠ The separator rule (`-`/` ` -> `_` right after the FILE NO) is deliberately
    DROPPED in `rev` mode. Measured, it does real damage: it rewrites the
    sub-sheet number in `GG38-C-PLN-002-1-...` (002-1/-2/-3 are three DIFFERENT
    drawings, not revisions of one) and it turns
    `NSW113 - AuxTx ... - Rev.A.pdf` into `..._-_RevA.pdf`. Changing a document
    numbering convention is not the same job as fixing a revision token.
  ⚠ Neither mode ever touches DESCRIPTION words. cross_check's own `suggested`
    name does -- it pulls the title from the DLV, which on Forbes would rename a
    real drawing `NSW113-C-PLN-012_Steel Platform_RevB.pdf` to `..._Reserved_...`.

Usage:
    python standardize_filenames.py <dir>                 # preview only
    python standardize_filenames.py <dir> --mode rev      # revision token only
    python standardize_filenames.py <dir> --apply         # actually rename
    python standardize_filenames.py <dir> -r --apply      # recurse into subfolders
    python standardize_filenames.py <dir> --ext .pdf .dwg # restrict to extensions
"""
import argparse
import os
import sys
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).resolve().parent))
import register_membership as _rm


def _iter_files(root: Path, recurse: bool, exts):
    it = root.rglob('*') if recurse else root.glob('*')
    for p in it:
        if not p.is_file():
            continue
        if exts and p.suffix.lower() not in exts:
            continue
        yield p


_dm_cache = {}


def _rev_only_namer(path: Path) -> str:
    """`rev` mode namer: DeliverableManager.normalize_filename(name) with NO
    deliverable_desc (Rule 3 -- description rewrite -- is then skipped entirely),
    keeping ONLY the changes whose reason mentions the revision. Returns the
    original name when nothing revision-shaped needs fixing."""
    import ifr_automation_v10 as _ifr
    key = str(path.parent)
    dm = _dm_cache.get(key)
    if dm is None:
        dm = _dm_cache[key] = _ifr.DeliverableManager.__new__(_ifr.DeliverableManager)
        dm.project_path = path.parent
        dm.dry_run = True
        dm.logger = __import__("logging").getLogger("standardize_filenames")
    try:
        new, changes = dm.normalize_filename(path.name)
    except Exception:
        return path.name
    if not changes or any("revision" not in c for c in changes):
        return path.name          # separator-only / mixed -> leave it to a human
    return new


def plan_renames(root: Path, recurse=False, exts=None, mode="format"):
    """Return (renames, skips): renames = [(src, dst)], skips = [(src, reason)].
    Files already tidy are silently omitted. Collisions land in `skips`."""
    exts = {e.lower() for e in exts} if exts else None
    namer = _rev_only_namer if mode == "rev" else (
        lambda p: _rm.normalize_filename_format(p.name))
    # group planned targets per parent dir to catch two-into-one collisions.
    planned = defaultdict(list)          # parent -> [(src, new_name)]
    for src in _iter_files(root, recurse, exts):
        new_name = namer(src)
        if new_name == src.name:
            continue                     # already standard
        planned[src.parent].append((src, new_name))

    renames, skips = [], []
    for parent, items in planned.items():
        # existing on-disk names in this dir (lower-cased for case-insensitive FS).
        existing = {p.name.lower() for p in parent.iterdir() if p.is_file()}
        # count how many sources want each target name (two-into-one guard).
        want = defaultdict(list)
        for src, new_name in items:
            want[new_name.lower()].append(src)
        for src, new_name in items:
            dst = parent / new_name
            key = new_name.lower()
            if len(want[key]) > 1:
                skips.append((src, f"collision: {len(want[key])} files normalize "
                                   f"to '{new_name}'"))
                continue
            # target already exists as a DIFFERENT file (not just src's own name in
            # a different case — that IS the rename we want to perform).
            if key in existing and key != src.name.lower():
                skips.append((src, f"target '{new_name}' already exists"))
                continue
            renames.append((src, dst))
    return renames, skips


def _do_rename(src: Path, dst: Path):
    """Rename src→dst, routing a case-only change through a temp name so a
    case-insensitive filesystem doesn't treat it as a no-op / same-file clash.
    NB: a plain `src == dst` guard is WRONG here — WindowsPath compares case-
    insensitively, so it would swallow the very case-only renames we must perform;
    plan_renames already guarantees dst.name differs from src.name."""
    if src.name.lower() == dst.name.lower() and src.name != dst.name:
        tmp = src.with_name(src.name + ".rncase.tmp")
        os.replace(src, tmp)
        os.replace(tmp, dst)
    else:
        os.replace(src, dst)


def run(root: Path, apply=False, recurse=False, exts=None, mode="format"):
    renames, skips = plan_renames(root, recurse, exts, mode)
    label = "APPLY" if apply else "DRY-RUN (preview only — pass --apply to rename)"
    print(f"=== standardize_filenames [{label}] mode={mode} : {root} ===")
    if not renames and not skips:
        print("Nothing to do — all filenames already standard.")
        return 0
    done, failed = 0, 0
    for src, dst in renames:
        if apply:
            try:
                _do_rename(src, dst)
                done += 1
                print(f"  [renamed] {src.name}  ->  {dst.name}")
            except OSError as e:
                failed += 1
                print(f"  [FAILED ] {src.name}  ->  {dst.name}  ({e})")
        else:
            print(f"  [preview] {src.name}  ->  {dst.name}")
    for src, reason in skips:
        print(f"  [SKIP   ] {src.name}  ({reason})")
    print(f"--- {len(renames)} to rename"
          + (f" ({done} done, {failed} failed)" if apply else "")
          + f", {len(skips)} skipped ---")
    return 1 if failed else 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Standardize deliverable filename "
                                             "format (case/space/Rev). Dry-run "
                                             "unless --apply.")
    ap.add_argument("directory", help="folder to process")
    ap.add_argument("--apply", action="store_true",
                    help="actually rename (default: preview only)")
    ap.add_argument("-r", "--recurse", action="store_true",
                    help="recurse into subfolders")
    ap.add_argument("--ext", nargs="+", metavar="EXT",
                    help="restrict to these extensions, e.g. --ext .pdf .dwg")
    ap.add_argument("--mode", choices=("format", "rev"), default="format",
                    help="format = house 'Rev 1' style (default); "
                         "rev = revision token only, target '_RevA' (see header)")
    args = ap.parse_args(argv)
    root = Path(args.directory)
    if not root.is_dir():
        print(f"error: not a directory: {root}", file=sys.stderr)
        return 2
    return run(root, apply=args.apply, recurse=args.recurse, exts=args.ext,
               mode=args.mode)


if __name__ == "__main__":
    sys.exit(main())
