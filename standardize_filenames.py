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

Usage:
    python standardize_filenames.py <dir>                 # preview only
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


def plan_renames(root: Path, recurse=False, exts=None):
    """Return (renames, skips): renames = [(src, dst)], skips = [(src, reason)].
    Files already tidy are silently omitted. Collisions land in `skips`."""
    exts = {e.lower() for e in exts} if exts else None
    # group planned targets per parent dir to catch two-into-one collisions.
    planned = defaultdict(list)          # parent -> [(src, new_name)]
    for src in _iter_files(root, recurse, exts):
        new_name = _rm.normalize_filename_format(src.name)
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


def run(root: Path, apply=False, recurse=False, exts=None):
    renames, skips = plan_renames(root, recurse, exts)
    label = "APPLY" if apply else "DRY-RUN (preview only — pass --apply to rename)"
    print(f"=== standardize_filenames [{label}] : {root} ===")
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
    args = ap.parse_args(argv)
    root = Path(args.directory)
    if not root.is_dir():
        print(f"error: not a directory: {root}", file=sys.stderr)
        return 2
    return run(root, apply=args.apply, recurse=args.recurse, exts=args.ext)


if __name__ == "__main__":
    sys.exit(main())
