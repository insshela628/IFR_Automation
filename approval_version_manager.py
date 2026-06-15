#!/usr/bin/env python3
"""
Approval Version Manager — SAPN 审批文档版本管理 (跨所有 SA 项目)
================================================================
对 `9. Approval/SAPN/` 下**每个类别文件夹**做版本管理: 同一文档的多版本里挑出
当前版留在类别根目录, 旧版移进该类别自己的 `SS/`。免去人工一个个 folder 点版本。

与图纸 VersionManager (IFR 字母 / IFC 数字, 按 doc-ID) **互不干扰** —— 不同路径、
不同版本语义。本模块用 SAPN 审批文档的判定 (SSOT: mainv3 §4.3):

  当前版三级判定 (高→低):
    1. 状态   executed > countersigned > signed > 未签
    2. 版本字母  Rev/Issue  A < B < C < D
    3. mtime  修改时间最新 (仅前两者平手时)
  (光看文件名/纯 mtime 都不可靠 —— Dropbox 改 mtime, 见 V6 CLAUDE.md「Incremental Check」)

类别结构 (mainv3 §4.3, kickoff 已建好骨架, 本模块不建/不归类, 只在既有类别内排版本):
  1. Application / 2. Connection Offer (CO) / 3. Engineering Report (ER) /
  4. Protection / 5. XR Impedance / 6. Correspondence  (+ 主目录 SS)

范围: 扫 `Project (EPC)/2.SA/*/Design/Engineering/9. Approval/SAPN` —— 自动覆盖
**未来所有 SA 项目**, 无需逐项目配置。NSW(Essential Energy/Ausgrid) 阶段不同, 暂不纳入。

用法:
    python approval_version_manager.py                      # 全部 SA 项目, dry-run
    python approval_version_manager.py --apply              # 落盘
    python approval_version_manager.py --project Waterloo   # 仅某项目 (名称含子串)
    python approval_version_manager.py --root "D:\\...\\2.SA"  # 自定义 SA 根
"""

import os
import re
import sys
import shutil
import argparse
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# SAPN 类别文件夹 (mainv3 §4.3)。键=文件夹名前缀匹配, 值=该类别的子类型识别器。
_SAPN_CATEGORIES = [
    "1. Application", "2. Connection Offer (CO)", "3. Engineering Report (ER)",
    "4. Protection", "5. XR Impedance", "6. Correspondence",
]
_SS_NAMES = {"ss", "superseded", "superceded"}


def _flat(s: str) -> str:
    """分隔符归一: _ - & → 空格, 收敛多空格 (关键词匹配不受下划线/连字符影响)。"""
    return re.sub(r"\s+", " ", re.sub(r"[_\-&]", " ", str(s).lower())).strip()


def status_rank(name: str) -> int:
    """签署状态: executed(4) > countersigned(3) > signed(2) > 其他(1)。"""
    n = _flat(name)
    if "executed" in n:
        return 4
    if "countersigned" in n:
        return 3
    if "signed" in n:
        return 2
    return 1


def rev_rank(name: str) -> int:
    """版本字母 Rev/Issue A<B<C → 1,2,3; 无则 0。'revised' 不误命中 (词边界)。"""
    m = re.search(r"\b(?:rev|issue)\s*([a-z])\b", _flat(name))
    return (ord(m.group(1)) - ord("a") + 1) if m else 0


def version_key(path: Path):
    """三级排序键 (越大越新): (状态, 版本字母, mtime)。"""
    name = path.name
    try:
        mt = path.stat().st_mtime
    except OSError:
        mt = 0.0
    return (status_rank(name), rev_rank(name), mt)


def _subtype(category: str, name: str) -> str:
    """同一类别内把"同一文档的不同版本"归到一组 (不同文档→不同子类型, 不互相 supersede)。
    SAPN 文档语义跨所有 SA 项目一致 (同一网络商), 故规则内置。SSOT 类别 doctrine: mainv3 §4.3。"""
    n = _flat(name)
    if category.startswith("2."):                       # Connection Offer (CO)
        if "3610" in n:
            return "tc"                                 # 条款文件 (名里也含 ongoing, 须先判)
        if "ld 500" in n or n.startswith("11 ") or " cn" in n:
            return "ld-ref"
        if ("ongoing" in n or "c s contract" in n or "connection and supply" in n
                or "supply contract" in n):
            return "contract"
        if "agreement" in n:
            return "agreement"
        return "offer"
    if category.startswith("3."):                       # Engineering Report (ER)
        if "letter" in n:
            return "er-letter"
        return "er-report"
    if category.startswith("1."):                       # Application
        if "hlf" in n:
            return "hlf"
        if "enquiry" in n:
            return "enquiry"
        if "indicative" in n:
            return "indicative"
        return "prelim"
    # Protection / XR Impedance / Correspondence: 默认每个文件名"基名"自成一组,
    # 用归一化基名分组 (去 rev/issue/状态/日期/序号)。
    return "base:" + _norm_base(name)


def _norm_base(name: str) -> str:
    s = os.path.splitext(name)[0].lower()
    s = re.sub(r"\b(?:rev|issue)[\s._-]*[a-z0-9.]+", " ", s)
    s = re.sub(r"\b(signed|executed|countersigned|revised|final|draft)\b", " ", s)
    s = re.sub(r"\b20\d{2}[-_.]?\d{0,2}[-_.]?\d{0,2}\b", " ", s)
    s = re.sub(r"^\s*\d+[a-z]?\s+", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _to_long(p: Path) -> str:
    r"""Windows 长路径 (>260): \\?\ 前缀 (仅 I/O)。"""
    s = os.path.abspath(str(p))
    return ("\\\\?\\" + s) if (os.name == "nt" and not s.startswith("\\\\?\\")) else s


def _detect_sa_root() -> Path | None:
    """探测 Project (EPC)/2.SA 根 (兼容 ACE / jilin / home)。"""
    cands = [
        Path(r"C:\Users\jilin\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA"),
        Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA"),
        Path.home() / "GREEN GOLD ENERGY Dropbox" / "Projects" / "Project (EPC)" / "2.SA",
    ]
    for c in cands:
        if c.is_dir():
            return c
    return None


class ApprovalVersionManager:
    """SAPN 审批文档版本管理 (跨所有 SA 项目)。"""

    def __init__(self, sa_root: Path, dry_run: bool = True):
        self.sa_root = Path(sa_root)
        self.dry_run = dry_run

    # ── 发现各 SA 项目的 SAPN 路径 ──────────────────────────────
    def find_sapn_dirs(self, project_filter: str | None = None):
        out = []
        if not self.sa_root.is_dir():
            return out
        for proj in sorted(self.sa_root.iterdir()):
            if not proj.is_dir():
                continue
            if project_filter and project_filter.lower() not in proj.name.lower():
                continue
            # 标准 Design/Engineering 或 flat Engineering
            for mid in ("Design/Engineering", "Engineering"):
                sapn = proj / mid / "9. Approval" / "SAPN"
                if sapn.is_dir():
                    out.append((proj.name, sapn))
                    break
        return out

    # ── 单个类别文件夹内排版本 ──────────────────────────────────
    def _process_category(self, cat_dir: Path):
        """返回 [(src, dest, group, decision)]。当前版留根, 旧版 → 该类别 SS/。"""
        ss_dir = self._ss_folder(cat_dir)
        try:
            files = [f for f in cat_dir.iterdir()
                     if f.is_file() and not f.name.startswith("_")
                     and not f.name.endswith(".meta.json")]
        except OSError:
            return []
        groups = defaultdict(list)
        for f in files:
            groups[_subtype(cat_dir.name, f.name)].append(f)
        moves = []
        for grp, items in groups.items():
            if len(items) < 2:
                continue                               # 单文件 = 已是当前版
            current = max(items, key=version_key)
            for f in items:
                if f is current:
                    continue
                moves.append((f, ss_dir / f.name, grp, "→SS"))
        return moves

    @staticmethod
    def _ss_folder(cat_dir: Path) -> Path:
        try:
            for it in cat_dir.iterdir():
                if it.is_dir() and it.name.lower() in _SS_NAMES:
                    return it
        except OSError:
            pass
        return cat_dir / "SS"

    # ── 处理一个项目的整个 SAPN ─────────────────────────────────
    def process_sapn(self, sapn: Path):
        plan = []   # (category, moves)
        for cat in _SAPN_CATEGORIES:
            cat_dir = sapn / cat
            if not cat_dir.is_dir():
                continue
            mv = self._process_category(cat_dir)
            if mv:
                plan.append((cat, mv))
        return plan

    def _do_move(self, src: Path, dest: Path) -> bool:
        for i in range(5):
            try:
                os.makedirs(_to_long(dest.parent), exist_ok=True)
                tgt = dest
                if os.path.exists(_to_long(tgt)):       # 同名 → 时间戳避让, 不覆盖
                    ts = datetime.now().strftime("%H%M%S")
                    tgt = dest.with_name(f"{dest.stem}_{ts}{dest.suffix}")
                shutil.move(_to_long(src), _to_long(tgt))
                return True
            except OSError:
                import time
                time.sleep(2)
        return False

    # ── 跑全部 ──────────────────────────────────────────────────
    def run(self, project_filter: str | None = None) -> dict:
        sapns = self.find_sapn_dirs(project_filter)
        tag = "DRY-RUN" if self.dry_run else "APPLY"
        print(f"\n=== Approval Version Manager [{tag}] — {len(sapns)} 个 SA 项目有 SAPN ===")
        total = {"projects": 0, "moved": 0, "scanned_cats": 0}
        for pname, sapn in sapns:
            plan = self.process_sapn(sapn)
            n = sum(len(mv) for _, mv in plan)
            if not plan:
                print(f"\n[{pname}] ✓ 已最新, 无需移动")
                total["projects"] += 1
                continue
            print(f"\n[{pname}]  待降 SS: {n}")
            for cat, moves in plan:
                print(f"  {cat}:")
                for src, dest, grp, _ in moves:
                    print(f"    [{grp}] {src.name}")
                    if not self.dry_run:
                        ok = self._do_move(src, dest)
                        if ok:
                            total["moved"] += 1
                        else:
                            print(f"        ⚠ 移动失败(占用?): {src.name}")
                    total["scanned_cats"] += 1
            if self.dry_run:
                total["moved"] += n
            total["projects"] += 1
        print(f"\n=== 汇总: {total['projects']} 项目, "
              f"{'将降' if self.dry_run else '已降'} SS {total['moved']} 个 ===")
        if self.dry_run:
            print("(预览, 未移动任何文件; 加 --apply 落盘)")
        return total


def main():
    ap = argparse.ArgumentParser(description="SAPN 审批文档版本管理 (跨所有 SA 项目)")
    ap.add_argument("--apply", action="store_true", help="落盘 (默认 dry-run 预览)")
    ap.add_argument("--project", help="仅处理名称含此子串的项目 (如 Waterloo)")
    ap.add_argument("--root", help="自定义 2.SA 根路径")
    args = ap.parse_args()

    root = Path(args.root) if args.root else _detect_sa_root()
    if not root or not root.is_dir():
        print(f"ERROR: 找不到 2.SA 根 ({root}). 用 --root 指定。")
        sys.exit(1)
    mgr = ApprovalVersionManager(root, dry_run=not args.apply)
    mgr.run(args.project)


if __name__ == "__main__":
    main()
