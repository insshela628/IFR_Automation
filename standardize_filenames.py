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
import re
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


# ── mode `dlv`: the Deliverables List is the naming authority ────────────────
# 用户 2026-09-23 的第一性: 「我们的文件命名应该按 deliverable 来, 由于 engineer 会有自己的
# 命名, 那你应该把他们自己命名的改成跟 deliverable 统一」+「Deliverable 里没有的, 一概不需要
# 交付」。所以这一模式**不问人**, 直接按 DLV 把工程师自己起的名改成统一的。
#
# 目标格式来自 `naming_schema_v3.json`(SSOT, 10 个消费者在读):
#     {doc_id}_{description}_Rev{revision}[{note}].{ext}
#     separator `_` · revision_prefix `Rev` · revision_no_space true
# 它的 incorrect 示例里就有 `..._Rev A.pdf` ("Space between Rev and revision letter"),
# 所以 `--mode format` 产出的 ` Rev A` 与这份 SSOT 冲突 —— 别在 GG/NSW 项目上用那个模式。
#
# 四条闸 (每条都是实测出来的, 不是假想):
#  ① 无 doc-ID → **跳过, 且不算问题**。DLV 里没有 = 不是交付物 = 不需要 FILE NO。
#     实测 Waterloo 52 个文件落这一类: 土壤模型/计算附页/厂家 datasheet/底图。
#     以前把它们报成"文件健康异常"是把第一性搞反了。
#  ② DLV 标题是占位 (`Reserved`/空/`N/A`/`TBC`) → **保留文件自己的描述**。实测 Waterloo 有
#     12 行 `Reserved`+9 行空, Forbes 14+9; 不设这道闸, 真图 `NSW113-C-PLN-012_Steel
#     Platform_RevB.pdf` 会被改名成 `..._Reserved_RevB.pdf`。
#  ③ 子图号 (`GG38-C-PLN-002-1/-2/-3` 三张不同的图, DLV 只有 `-002` 一行) → doc-ID 保留子图号,
#     描述**用文件自己的** (DLV 没有这一行, 拿父行标题会把三张图压成同一个名)。
#  ④ 版本号之后的一切 = 备注, **原样保留** (用户: 「保留备注, 你只改前面的文件名和版本号」)。
#     贴着版本号写的备注 (`_RevBWait nr datasheet`) 补一个 `_` 分开, 否则版本号解不出来。
_PH_DESC = {"reserved", "n/a", "na", "tbc", "-", ""}     # DLV 占位值, 绝不当标题用

# group(1) = DLV 查得到的 doc-ID; group(2) = 子图号(可选)。
# ⚠ 两段**分开捕获**是必须的: 合成一段再用 `-\d+$` 剥子图, 会把序号本身 (`-001`) 当子图剥掉,
#   于是 `GG38-E-PLN-001` 的 base 变成 `GG38-E-PLN`, DLV 永远查不到 (实测踩过)。
_DLV_DOCID = re.compile(
    r'^((?:GG\d{2}|NSW\d{2,3}|[A-Z]{2,5}\d{2,5})(?:-[CE])?'
    r'-(?:PLN|SEC|SLD|BLD|GAD|RPT|SCH|DLV|SPC)-?\d{3})(-\d+)?', re.IGNORECASE)

# 版本号后面**紧贴**着备注时, register_membership._RE_REV 的后瞻 `(?=[_.\s]|$)` 会失配
# (`_RevBWait…` 的 B 后面是 W)。这条是它的无后瞻版本, 只在前者落空时才用。
_DLV_REV_GLUED = re.compile(r'[_\s-](?:[Rr]ev|[Rr])\.?\s*([A-Za-z0-9])')

_dlv_map_cache = {}


def _dlv_descriptions(project_root: Path) -> dict:
    """{doc_id: DLV 里的标题}。读不到 → {} (于是每个文件都落闸②/③, 只修版本号写法, 不乱改描述)。"""
    key = str(project_root)
    if key in _dlv_map_cache:
        return _dlv_map_cache[key]
    out = {}
    try:
        import io as _io
        import contextlib as _cl
        import openpyxl as _ox
        import ifr_automation_v10 as _ifr
        with _cl.redirect_stdout(_io.StringIO()):
            dm = _ifr.DeliverableManager(project_root, dry_run=True)
            xl = dm.find_deliverable_excel()
            if xl:
                wb = _ox.load_workbook(xl, data_only=True)
                ws = wb[wb.sheetnames[0]]
                items, _ = dm.read_excel_items(ws, dm.detect_layout(ws))
                out = {k.upper(): (v.get("description") or "") for k, v in items.items()}
                wb.close()
    except Exception:
        out = {}
    _dlv_map_cache[key] = out
    return out


def _multi_desc_docids(root: Path, recurse: bool, exts) -> set:
    """盘上同一个 doc-ID 下出现了 2 种以上"自带描述"的号。返回这些号。"""
    seen = {}
    for f in _iter_files(root, recurse, exts):
        if any(x.lower() in ("ss", "superseded", "superceded") for x in f.parts):
            continue
        m = _DLV_DOCID.match(Path(f.name).stem)
        if not m:
            continue
        key = m.group(1).upper() + (m.group(2) or "")
        stem = Path(f.name).stem
        # ⚠ 别用 re.split 找版本号: 模式 `[_\s-][Rr]…` 会先命中 doc-ID 里的 `-RPT`/`-RevA`
        #   的 `-R`, 于是描述恒为空、这道闸恒不触发 (实测踩过)。版本号只认一处取法 ——
        #   和 `dlv_target` 同一把尺, 取**最后**一个匹配。
        revs = list(_rm._RE_REV.finditer(stem)) or list(_DLV_REV_GLUED.finditer(stem))
        if not revs:
            continue
        desc = stem[m.end():revs[-1].start()].strip(" _-")
        seen.setdefault(key, set()).add(desc.lower())
    return {k for k, v in seen.items() if len(v) > 1}


def dlv_target(name: str, dlv_map: dict, own_desc_only: bool = False):
    """文件名 → (规范名, 依据)。不该动就返回 (原名, 原因)。纯函数, 好测。"""
    stem, ext = Path(name).stem, Path(name).suffix
    m = _DLV_DOCID.match(stem)
    if not m:
        return name, "skip:不在命名体系内 (DLV 里没有 → 不是交付物)"
    base, sub = m.group(1).upper(), (m.group(2) or "")
    revs = list(_rm._RE_REV.finditer(stem)) or list(_DLV_REV_GLUED.finditer(stem))
    if not revs:
        return name, "skip:没有版本号"
    r = revs[-1]
    rev = r.group(2 if r.re is _rm._RE_REV else 1).upper()
    raw = stem[r.end():]                                  # 闸④: 版本号之后的一切 = 备注
    note = ("_" + raw.strip()) if (raw and raw[0] not in " _-") else raw
    own = stem[m.end():r.start()].strip(" _-")
    title = "" if own_desc_only else (dlv_map.get(base) or "").strip()
    if own_desc_only:
        desc, why = own, "同号多文档 → 只规范版本号, 描述保持原样"
    elif sub:
        desc, why = own, "子图 (DLV 无此行, 用文件自己的描述)"
    elif title.lower() in _PH_DESC:
        desc, why = own, f"DLV 是占位 '{title or '空'}' → 保留原描述"
    else:
        desc, why = title, "按 DLV 标题"
    return f"{base}{sub}_{desc}_Rev{rev}{note}{ext}", why


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


def plan_renames(root: Path, recurse=False, exts=None, mode="format",
                 project_root=None):
    """Return (renames, skips): renames = [(src, dst)], skips = [(src, reason)].
    Files already tidy are silently omitted. Collisions land in `skips`."""
    exts = {e.lower() for e in exts} if exts else None
    if mode == "dlv":
        if project_root is None:
            raise ValueError("--mode dlv 需要 --project <项目根>")
        dlv_map = _dlv_descriptions(Path(project_root))

        # 盘上同一个 doc-ID 带着**两种不同描述** = 两份不同的文档共用一个号 (实测
        # GG38-E-RPT-003 同时有 `DC Cable Calculation Report` 和 `…Spreadsheet`)。
        # DLV 只有一行、只有一个标题, 裁不了这种 —— 照着改会把 Spreadsheet 改名成
        # Report, **两份文档合并成同一个身份**, 而且是不可逆的语义损失 (碰撞守卫只挡得住
        # 目标名已存在的那一半, 另一半会静静改成功)。这种号一律不改描述。
        multi = _multi_desc_docids(root, recurse, exts)

        def namer(p):
            # 面闸 (mainv3 §「文件命名 (注意与 IFR/IFC 不同)」): AS BUILT 是**另一张脸**
            # (`{DocID} REV {N} {Description}_AS BUILT.pdf`, 空格分隔), 它的权威产出器是
            # `AsBuiltManager._build_ab_pdf_filename`。拿 IFR/IFC 的尺去量那张脸 = 把整个
            # AS BUILT 文件夹改成错的形状。比命名前必先按面分组, 组间绝不比。
            segs = [seg.lower() for seg in p.parts]
            # ① 面闸 (mainv3 §「文件命名 (注意与 IFR/IFC 不同)」): AS BUILT 是**另一张脸**
            #    (`{DocID} REV {N} {Description}_AS BUILT.pdf`, 空格分隔), 权威产出器是
            #    `AsBuiltManager._build_ab_pdf_filename`。拿 IFR/IFC 的尺去量那张脸 =
            #    把整个 AS BUILT 夹改成错的形状。比命名前必先按面分组, 组间绝不比。
            if any("as built" in x or "asbuilt" in x for x in segs):
                return p.name
            # ② 废版夹: 里面躺的是被取代的旧件, 改它纯属搅动 + 徒增碰撞。
            if any(x in ("ss", "superseded", "superceded", "old") for x in segs):
                return p.name
            # ③ 出站闸 (DENY by default): 文件名一旦随 transmittal 出门就是**对外标识符**,
            #    机器永不改 —— 而且 Dropbox 文件级分享链 (`scl/fi`) 改名即永久失效, 改回
            #    也不复活, 且**不报错**。已发出的错名正解是下一版按正名重发, 不是改名。
            #    ⚠ 现在没有逐文档的发出台账 (`transmittals` 只有单头没有行), 所以这里
            #    **按路径一刀挡**, 而不是"查不到就放行" —— 没有覆盖 = HOLD, 不是 ALLOW。
            if any(("ifr(client)" in x or "ifc(client)" in x or "sharepoint" in x)
                   for x in segs):
                return p.name
            tgt, _why = dlv_target(p.name, dlv_map)
            m = _DLV_DOCID.match(Path(p.name).stem)
            if m and (m.group(1).upper() + (m.group(2) or "")) in multi:
                # 多文档共号 → 挡的是**描述**那一半 (DLV 一行裁不了两份文档), 版本号写法
                # 照修。走同一个 dlv_target 而不是退回 _rev_only_namer —— 后者解不出
                # 贴着备注写的版本号 (`_RevBWait nr datasheet`), 于是这类文件会整个漏掉。
                return dlv_target(p.name, dlv_map, own_desc_only=True)[0]
            return tgt
    elif mode == "rev":
        namer = _rev_only_namer
    else:
        namer = lambda p: _rm.normalize_filename_format(p.name)
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
    # ⚠ 长路径 (2026-09-23 实测): 规范名往往**比原名长** (描述按 DLV 补全), 于是原本
    #   刚好在 260 以内的路径一改就越界, `os.replace` 抛 WinError 3 "找不到路径" ——
    #   那句报错会被读成"文件不见了", 其实是 MAX_PATH。所有 I/O 一律走 `\?\` 前缀,
    #   存的仍是干净路径 (记忆 windows-long-path)。
    def _lp(x: Path) -> str:
        t = str(Path(x).resolve())
        return t if t.startswith("\\\\?\\") else "\\\\?\\" + t
    if src.name.lower() == dst.name.lower() and src.name != dst.name:
        tmp = src.with_name(src.name + ".rncase.tmp")
        os.replace(_lp(src), _lp(tmp))
        os.replace(_lp(tmp), _lp(dst))
    else:
        os.replace(_lp(src), _lp(dst))


def run(root: Path, apply=False, recurse=False, exts=None, mode="format",
        project_root=None):
    renames, skips = plan_renames(root, recurse, exts, mode, project_root)
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
    ap.add_argument("--mode", choices=("format", "rev", "dlv"), default="format",
                    help="format = generic tidier (PREVIEW/compare key only); "
                         "rev = revision token only; "
                         "dlv = Deliverables List is the authority (see header)")
    ap.add_argument("--project", help="project root; required by --mode dlv")
    args = ap.parse_args(argv)
    root = Path(args.directory)
    if not root.is_dir():
        print(f"error: not a directory: {root}", file=sys.stderr)
        return 2
    # ⚠ `format` 是**通用整理器/比较键**, 不是任何一张面的权威产出器 —— 它产出的
    #   ` Rev A` 同时违反 mainv3 §2 的 IFR/IFC 形状和 §AS BUILT 的空格形状。
    #   它当 `--apply` 的默认档会把**已经正确**的 `_RevA` 改坏, 所以真改盘时必须显式选面。
    if args.apply and "--mode" not in (argv if argv is not None else sys.argv):
        print("error: --apply 时必须显式 --mode (format 只是预览用的通用整理器, "
              "拿它真改名会把 _RevA 改成 ' Rev A')", file=sys.stderr)
        return 2
    return run(root, apply=args.apply, recurse=args.recurse, exts=args.ext,
               mode=args.mode, project_root=args.project)


if __name__ == "__main__":
    sys.exit(main())
