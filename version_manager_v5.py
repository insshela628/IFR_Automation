#!/usr/bin/env python3
"""
Version Manager - 旧版本文件管理工具 v5.0 (STANDALONE FALLBACK)

NOTE: This functionality has been merged into ifr_automation_v7.py as part of the
unified Engineering Pipeline. This file is kept as a standalone fallback.
Use ifr_automation_v7.py for the integrated pipeline experience.

Automatically identifies and moves older versions of files to SS folder.

Manages file versions across:
  Mode 1 (PDF): IFR_internal / IFR(Client) / IFC(Client) — PDF files
    - NEW: Detects IFC files in IFR(Client) directories and moves them to SS
  Mode 2 (Native): Native + Reports + Schedule folders — all working files
    - Scans multiple root directories (Native, Reports, Schedule)
    - Classifies files as IFR (letter rev) or IFC (numeric rev)
    - Keeps newest IFR + newest IFC per extension per folder (v5.0: 每种扩展名各保留最新)
    - Renames non-standard filenames to match naming_schema_v3 rules
    - .bak files follow their corresponding .dwg
    - Auto-detects SS/SUPERSEDED/Superceded folders

Usage:
    python version_manager_v4.py                          # 交互式模式
    python version_manager_v4.py --native                 # Native 模式 (dry-run, 全部范围)
    python version_manager_v4.py --native --scope native  # 仅扫描 Native
    python version_manager_v4.py --native --scope reports # 仅扫描 Reports
    python version_manager_v4.py --native --execute
    python version_manager_v4.py --native --folder PLN-001

Author: Generated with Claude Code
Version: 4.0
"""

import os
import re
import sys
import json
import shutil
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from collections import defaultdict
from dataclasses import dataclass, field
from itertools import groupby


class VersionManager:
    """Manage file versions by moving older versions to SS folder."""

    # Version patterns to identify version suffixes in filenames
    VERSION_PATTERNS = [
        r'[_\s]Rev\.?\s*([A-Z0-9]+)',      # _RevA, _Rev A, _Rev.A, Rev01
        r'[_\s]R([A-Z])(?=[_\.\s]|$)',      # _RA, _RB (single letter)
        r'[_\s]R(\d+)(?=[_\.\s]|$)',        # _R1, _R2 (number)
        r'[_\s]V(\d+)',                      # _V1, _V2
        r'[_\s]v(\d+)',                      # _v1, _v2
        r'-([A-Z])(?=[_\.\s]|$)',           # -A, -B at end
    ]

    # Subdirectories to process within a project
    TARGET_SUBDIRS = [
        r"Design\Engineering\1. Drawings\2. IFR_internal",
        r"Design\Engineering\1. Drawings\3. IFR(Client)\1.Drawing",
        r"Design\Engineering\1. Drawings\3. IFR(Client)\2.Reports",
        r"Design\Engineering\1. Drawings\3. IFR(Client)\3.Deliverables",
        r"Design\Engineering\1. Drawings\4. IFC(Client)",
        r"Design\Engineering\13. Client Sharepoint\1.IFR\1.Report",
        r"Design\Engineering\13. Client Sharepoint\1.IFR\2.Drawing",
    ]

    # Pattern to identify project folders
    PROJECT_PATTERNS = [
        r'^GG-?\d+',       # GG-31, GG31
        r'^NSW\d+',        # NSW152
        r'^[A-Z]{2,4}-\d+', # XX-123 format
    ]

    def __init__(self, root_path: str, dry_run: bool = False):
        """
        Initialize Version Manager.

        Args:
            root_path: Root directory containing all projects
            dry_run: If True, only preview changes without moving files
        """
        self.root_path = Path(root_path)
        self.dry_run = dry_run
        self.total_stats = {"scanned": 0, "groups": 0, "moved": 0}

    def extract_base_name_and_version(self, filename: str) -> Tuple[str, Optional[str]]:
        """
        Extract base filename and version from a filename.

        Returns:
            Tuple of (base_name, version) where version may be None
        """
        name_without_ext = Path(filename).stem
        extension = Path(filename).suffix

        # Strip browser download-duplicate suffix like " (1)", " (2)" etc.
        name_without_ext = re.sub(r'\s*\(\d+\)\s*$', '', name_without_ext)

        # Try each version pattern
        for pattern in self.VERSION_PATTERNS:
            match = re.search(pattern, name_without_ext, re.IGNORECASE)
            if match:
                version = match.group(1).upper()
                # Remove the version part to get base name
                base_name = re.sub(pattern, '', name_without_ext, flags=re.IGNORECASE)
                # Clean up any trailing underscores or spaces
                base_name = re.sub(r'[_\s]+$', '', base_name)
                return (base_name + extension, version)

        # No version found
        return (Path(name_without_ext).stem + extension if name_without_ext != Path(filename).stem else filename, None)

    def scan_for_projects(self) -> List[Path]:
        """
        Scan root directory for project folders.

        Returns:
            List of project paths
        """
        projects = []

        if not self.root_path.exists():
            return projects

        # Search for projects - use iterdir with limited depth instead of rglob
        # to avoid long path issues
        def scan_level(path: Path, depth: int = 0, max_depth: int = 3):
            if depth > max_depth:
                return
            try:
                for item in path.iterdir():
                    if not item.is_dir():
                        continue
                    try:
                        # Check if this folder has the IFR structure
                        ifr_path = item / "Design" / "Engineering" / "1. Drawings"
                        if ifr_path.exists():
                            # Check if folder name matches project patterns
                            for pattern in self.PROJECT_PATTERNS:
                                if re.match(pattern, item.name, re.IGNORECASE):
                                    projects.append(item)
                                    break
                            else:
                                # Also include if it has IFR_internal or IFR(Client)
                                if (ifr_path / "2. IFR_internal").exists() or (ifr_path / "3. IFR(Client)").exists():
                                    projects.append(item)
                        else:
                            # Continue searching in subdirectories
                            scan_level(item, depth + 1, max_depth)
                    except (OSError, PermissionError):
                        # Skip paths that cause errors (too long, no permission, etc.)
                        continue
            except (OSError, PermissionError):
                pass

        scan_level(self.root_path)

        # Remove duplicates and sort
        projects = sorted(set(projects), key=lambda p: p.name)
        return projects

    @staticmethod
    def _find_ss_folder(target_dir: Path) -> Path:
        """Find existing SS/Superseded/Superceded folder, or return default SS/ path."""
        try:
            for item in target_dir.iterdir():
                if item.is_dir() and item.name.lower() in ('ss', 'superseded', 'superceded'):
                    return item
        except (OSError, PermissionError):
            pass
        return target_dir / 'SS'

    def _reg_titles_for(self, target_dir: Path) -> Dict[str, List[str]]:
        """Client *Deliverables List* titles for register-canonical grouping — the
        shared cross-track oracle (register_membership). SOFT dep: {} → grouping
        stays byte-identical to base_name. Cached per target_dir; strays surfaced
        via self._version_strays. Mirror of VersionManager._reg_titles_for in
        ifr_automation_v10.py (same SSOT)."""
        cache = getattr(self, '_reg_titles_cache', None)
        if cache is None:
            cache = self._reg_titles_cache = {}
            self._version_strays = []
        k = str(target_dir)
        if k not in cache:
            try:
                import register_membership as _rm
                cache[k] = _rm.load_register_titles(target_dir)
            except Exception:
                cache[k] = {}
        return cache[k]

    def scan_files(self, target_dir: Path) -> Dict[str, List[Tuple[Path, str, datetime]]]:
        """
        Scan directory and group files by base name.

        Returns:
            Dictionary mapping base_name to list of (file_path, version, modified_time)
        """
        if not target_dir.exists():
            return {}

        file_groups: Dict[str, List[Tuple[Path, str, datetime]]] = defaultdict(list)

        # Only scan PDF files in the target directory (not subdirectories, exclude SS folder)
        try:
            pdf_files = list(target_dir.glob("*.pdf"))
        except (OSError, PermissionError):
            return file_groups

        reg = self._reg_titles_for(target_dir)
        for file_path in pdf_files:
            try:
                if file_path.is_file():
                    base_name, version = self.extract_base_name_and_version(file_path.name)
                    key = base_name
                    if reg:   # register-canonical key: collapse re-worded revs of
                        try:  # one drawing; keep different-title clashes split
                            import register_membership as _rm
                            did = _rm.extract_doc_id(file_path.name)
                            titles = reg.get((did or '').upper())
                            if titles:
                                key, _m, stray = _rm.canonical_group_key(
                                    base_name, file_path.name, titles, did)
                                if stray:
                                    self._version_strays.append((file_path.name, did))
                        except Exception:
                            key = base_name
                    # Use long path for stat on Windows
                    if os.name == 'nt':
                        path_str = str(file_path.resolve())
                        if not path_str.startswith('\\\\?\\'):
                            path_str = '\\\\?\\' + path_str
                        stat_path = Path(path_str)
                        modified_time = datetime.fromtimestamp(stat_path.stat().st_mtime)
                    else:
                        modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)

                    # Use "NO_VERSION" as placeholder if no version found
                    version_str = version if version else "NO_VERSION"
                    file_groups[key].append((file_path, version_str, modified_time))
            except (OSError, PermissionError):
                # Skip files that cause errors
                continue

        return file_groups

    @staticmethod
    def _pdf_ver_key(item):
        """PDF 组内版本新旧键 (越大越新): (阶段, 版本号, mtime)。版本号优先于 mtime
        (Dropbox 同步会改 mtime), item=(path, version_str, mtime); 数字=IFC 阶段(更新)、
        字母=IFR 阶段、NO_VERSION 最旧; mtime 仅平手。SSOT: mainv3 §4.3。"""
        v = (item[1] or '').strip().upper()
        if v.isdigit():
            return (2, int(v), item[2])
        if v[:1].isalpha() and v != 'NO_VERSION':
            return (1, ord(v[0]) - ord('A') + 1, item[2])
        return (0, -1, item[2])

    def identify_old_versions(self, file_groups: Dict[str, List[Tuple[Path, str, datetime]]], ss_folder: Path) -> List[Tuple[Path, Path, str]]:
        """
        Identify files that should be moved to SS folder.

        Returns:
            List of (source_path, dest_path, reason) for files to move
        """
        files_to_move: List[Tuple[Path, Path, str]] = []

        for base_name, files in file_groups.items():
            # Only process groups with multiple files (multiple versions)
            if len(files) <= 1:
                continue

            # 版本号优先, mtime 仅平手 (_pdf_ver_key) — 见 CLAUDE.md「Incremental Check: NO mtime」
            files_sorted = sorted(files, key=self._pdf_ver_key, reverse=True)

            # Keep the newest file, mark others for moving
            newest_file = files_sorted[0]

            for file_path, version, modified_time in files_sorted[1:]:
                dest_path = ss_folder / file_path.name
                reason = f"较旧版本 (最新: {newest_file[0].name})"
                files_to_move.append((file_path, dest_path, reason))

        return files_to_move

    def identify_ifc_in_ifr_client(self, target_dir: Path, ss_folder: Path) -> List[Tuple[Path, Path, str]]:
        """识别 IFR(Client) 目录中的 IFC 文件，标记移动到 SS。"""
        files_to_move = []

        # 只在 IFR(Client) 子目录中启用
        if "IFR(Client)" not in str(target_dir):
            return files_to_move

        try:
            pdf_files = list(target_dir.glob("*.pdf"))
        except (OSError, PermissionError):
            return files_to_move

        for pdf_file in pdf_files:
            stem = pdf_file.stem
            is_ifc = bool(_RE_IFC.search(stem))
            is_ifr = bool(_RE_IFR.search(stem))

            # IFC 文件且不是 IFR → 标记移动
            if is_ifc and not is_ifr:
                dest_path = ss_folder / pdf_file.name
                reason = "IFC文件不应在IFR(Client)目录中"
                files_to_move.append((pdf_file, dest_path, reason))

        return files_to_move

    def move_files(self, files_to_move: List[Tuple[Path, Path, str]], ss_folder: Path) -> int:
        """
        Move files to SS folder.

        Returns:
            Number of files moved
        """
        if not files_to_move:
            return 0

        # Helper function for long path on Windows
        def to_long_path(path: Path) -> str:
            if os.name == 'nt':
                path_str = str(path.resolve())
                if not path_str.startswith('\\\\?\\'):
                    path_str = '\\\\?\\' + path_str
                return path_str
            return str(path)

        # Create SS folder if needed
        if not self.dry_run:
            try:
                ss_folder.mkdir(exist_ok=True)
            except (OSError, PermissionError):
                # Try with long path
                try:
                    Path(to_long_path(ss_folder)).mkdir(exist_ok=True)
                except Exception as e:
                    print(f"      [!] 无法创建 SS 文件夹: {e}")
                    return 0

        ss_name = ss_folder.name
        moved_count = 0

        for source, dest, reason in files_to_move:
            if self.dry_run:
                print(f"      [预览] {source.name} -> {ss_name}/")
            else:
                try:
                    # Handle existing file in SS folder
                    if dest.exists():
                        # Add timestamp to avoid overwrite
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        new_name = f"{dest.stem}_{timestamp}{dest.suffix}"
                        dest = ss_folder / new_name

                    # Use long path format on Windows
                    shutil.move(to_long_path(source), to_long_path(dest))
                    print(f"      [v] 已移动: {source.name} -> {ss_name}/")
                    moved_count += 1
                except Exception as e:
                    print(f"      [!] 移动失败: {source.name} - {e}")

        return moved_count

    def analyze_directory(self, target_dir: Path) -> Dict:
        """
        Analyze a directory without making changes.

        Returns:
            Analysis results
        """
        result = {
            "exists": target_dir.exists(),
            "total_files": 0,
            "multi_version_groups": {},
            "files_to_move": 0,
            "ifc_files": [],
            "ifc_files_count": 0
        }

        if not target_dir.exists():
            return result

        file_groups = self.scan_files(target_dir)
        result["total_files"] = sum(len(files) for files in file_groups.values())

        multi_version_groups = {k: v for k, v in file_groups.items() if len(v) > 1}
        result["multi_version_groups"] = multi_version_groups
        result["files_to_move"] = sum(len(v) - 1 for v in multi_version_groups.values())

        # 检测 IFR(Client) 目录中的 IFC 文件
        ss_folder = self._find_ss_folder(target_dir)
        ifc_files = self.identify_ifc_in_ifr_client(target_dir, ss_folder)
        result["ifc_files"] = ifc_files
        result["ifc_files_count"] = len(ifc_files)
        result["files_to_move"] += len(ifc_files)

        return result

    def process_directory(self, target_dir: Path, show_details: bool = True) -> Dict[str, int]:
        """
        Process a single directory.

        Returns:
            Statistics dictionary
        """
        stats = {"scanned": 0, "groups": 0, "moved": 0}

        if not target_dir.exists():
            if show_details:
                print(f"    [!] 目录不存在，跳过")
            return stats

        ss_folder = self._find_ss_folder(target_dir)

        # Scan files
        self._version_strays = []
        file_groups = self.scan_files(target_dir)
        if self._version_strays and show_details:
            print(f"    [i] {len(self._version_strays)} 个文件对不上客户清单标题，"
                  f"单独保留不合并 (可能编号有误): "
                  f"{', '.join(n for n, _ in self._version_strays[:5])}"
                  f"{' …' if len(self._version_strays) > 5 else ''}")
        total_files = sum(len(files) for files in file_groups.values())
        stats["scanned"] = total_files

        if show_details:
            print(f"    找到 {total_files} 个 PDF 文件")

        # Find groups with multiple versions
        multi_version_groups = {k: v for k, v in file_groups.items() if len(v) > 1}
        stats["groups"] = len(multi_version_groups)

        # 检测 IFR(Client) 目录中的 IFC 文件
        ifc_files = self.identify_ifc_in_ifr_client(target_dir, ss_folder)

        if not multi_version_groups and not ifc_files:
            if show_details:
                print(f"    [v] 没有重复版本文件")
            self._preview_name_format(target_dir, set(), show_details)
            return stats

        if multi_version_groups:
            if show_details:
                print(f"    发现 {len(multi_version_groups)} 组多版本文件:")

                # Display groups
                for base_name, files in multi_version_groups.items():
                    files_sorted = sorted(files, key=self._pdf_ver_key, reverse=True)  # 与决策同序
                    print(f"\n      基础文件: {base_name}")

                    for i, (file_path, version, modified_time) in enumerate(files_sorted):
                        status = "[保留]" if i == 0 else "[->SS]"
                        time_str = modified_time.strftime('%Y-%m-%d %H:%M')
                        print(f"        {status} {file_path.name} (版本:{version}, {time_str})")

        # Identify and move old versions
        files_to_move = self.identify_old_versions(file_groups, ss_folder)

        if files_to_move:
            if show_details:
                print(f"\n    移动 {len(files_to_move)} 个旧版本到 SS 文件夹:")
            moved_count = self.move_files(files_to_move, ss_folder)
            stats["moved"] = moved_count

        # 处理 IFC 文件
        if ifc_files:
            if show_details:
                print(f"\n    发现 {len(ifc_files)} 个 IFC 文件在 IFR(Client) 目录中:")
                for src, dst, reason in ifc_files:
                    print(f"        [->SS] {src.name} ({reason})")
            moved_ifc = self.move_files(ifc_files, ss_folder)
            stats["moved"] += moved_ifc

        moved_srcs = {str(s) for s, _d, _r in files_to_move} | {str(s) for s, _d, _r in ifc_files}
        self._preview_name_format(target_dir, moved_srcs, show_details)
        return stats

    def _preview_name_format(self, target_dir: Path, moved_srcs: set, show_details: bool):
        """DRY-RUN PREVIEW of identity-preserving filename-format fixes (case/space/
        underscore/Rev only — register_membership.normalize_filename_format) for the
        files that STAY. Preview ONLY: renaming client deliverables needs an explicit
        per-project go. Description→register-canonical is intentionally NOT done here
        (reconcile-gated + higher-risk). Mirror of VersionManager._preview_name_format
        in ifr_automation_v10.py."""
        try:
            import register_membership as _rm
        except Exception:
            return
        props = []
        try:
            for f in target_dir.glob('*.pdf'):
                if not f.is_file() or str(f) in moved_srcs:
                    continue
                new = _rm.normalize_filename_format(f.name)
                if new != f.name:
                    props.append((f.name, new))
        except (OSError, PermissionError):
            return
        if props and show_details:
            print(f"    [格式预览] {len(props)} 个文件名可统一格式 (仅预览, 批准后才执行):")
            for old, new in props[:8]:
                print(f"        {old}  ->  {new}")
            if len(props) > 8:
                print(f"        … 另 {len(props) - 8} 个")

    def process_project(self, project_path: Path, show_details: bool = True) -> Dict[str, int]:
        """
        Process all target directories in a project.

        Returns:
            Statistics dictionary
        """
        project_stats = {"scanned": 0, "groups": 0, "moved": 0}

        for subdir in self.TARGET_SUBDIRS:
            target_dir = project_path / subdir

            if show_details:
                # Shorten display path
                short_path = subdir.split("\\")[-1] if "\\" in subdir else subdir
                print(f"\n  [{short_path}]")

            stats = self.process_directory(target_dir, show_details)

            project_stats["scanned"] += stats["scanned"]
            project_stats["groups"] += stats["groups"]
            project_stats["moved"] += stats["moved"]

        return project_stats


# =============================================================================
# Native Version Manager v3.1 (v5.0: 按扩展名分组保留最新文件)
# =============================================================================

# --- v3.0 新增常量 ---

# 扫描根目录配置
# (相对于 project_path 的路径, 扫描深度, 显示名)
# depth=1: 直接子文件夹就是文档文件夹
#   e.g. Native/GG31-C-PLN-001_Civil Site Layout Plan_Rev/
# depth=2: 需要先进一层分类目录
#   e.g. Reports/Civil & Structure/GG31-C-RPT-001_xxx/
SCAN_ROOTS = [
    ("Design/Engineering/1. Drawings/1. Native", 1, "1. Native"),
    ("Design/Engineering/2. Calcs & Reports/Reports", 2, "Reports"),
    ("Design/Engineering/2. Calcs & Reports/Schedule", 1, "Schedule"),
]

# 扫描时跳过的子文件夹名 (不区分大小写)
_SKIP_FOLDER_NAMES = {
    'ss', 'superseded', 'superceded',
    'bom', 'appendix', 'approved',
}
# 以这些前缀开头的文件夹也跳过 (不区分大小写)
_SKIP_FOLDER_PREFIXES = ('sy supply',)

# 不参与版本管理的扩展名 (v3: 不再忽略 .xlsx/.txt/.csv — Reports 里这些是正式文件)
_IGNORE_EXT = {'.dwl', '.dwl2', '.err', '.log', '.tmp', '.lnk', '.db', '.ini'}

# Revision patterns for file classification
_RE_IFC = re.compile(
    r'[_\s-](?:[Rr]ev|[Rr])(\d+)(?:[_\s-]?IFC)?(?=[_.\s]|$)',
    re.IGNORECASE
)
_RE_IFR = re.compile(
    r'[_\s-](?:[Rr]ev|[Rr])([A-Z])(?=[_.\s]|$)',
    re.IGNORECASE
)
# AS BUILT marker: '_AB' / '-AB' / ' AB' (optionally '-exp') or 'AS BUILT'.
# Checked BEFORE IFC/IFR so an AS BUILT file like '...Rev1...-AB.dwg' is NOT
# misclassified as IFC and renamed to _IFC (brownfield LMS hazard).
_RE_AB = re.compile(
    r'(?:[_\s-]AB(?:[_\s-]?exp)?(?=[_.\s-]|$)|AS[_\s-]?BUILT)',
    re.IGNORECASE
)


@dataclass
class DwgFile:
    """A versioned file with parsed metadata."""
    path: Path
    filename: str
    rev_type: str        # 'IFR', 'IFC', or 'OTHER'
    revision: str        # e.g. 'A', 'B', '0', '1'
    mtime: datetime
    bak_path: Optional[Path] = None

    @property
    def rev_display(self) -> str:
        if self.rev_type == 'IFR':
            return f"Rev{self.revision}"
        elif self.rev_type == 'IFC':
            return f"Rev{self.revision}_IFC"
        elif self.rev_type == 'AB':
            return f"Rev{self.revision}_AB" if self.revision else "AS BUILT"
        return "no-rev"


@dataclass
class NativeFileAction:
    """An action to perform on a file (rename/move)."""
    source: Path
    dest: Path
    action: str   # 'rename', 'move_to_ss'
    reason: str


@dataclass
class FolderRelocation:
    """A planned folder relocation action."""
    source: Path
    dest: Path
    reason: str
    action: str   # 'move', 'merge', 'warn'


@dataclass
class NativeFolderResult:
    """Result of processing a single document folder."""
    folder_name: str
    folder_path: Path
    doc_id: str
    description: str
    dwg_files: List[DwgFile] = field(default_factory=list)
    actions: List[NativeFileAction] = field(default_factory=list)
    kept_ifr: Optional[DwgFile] = None
    kept_ifc: Optional[DwgFile] = None
    kept_ifr_all: List[DwgFile] = field(default_factory=list)  # v5.0: 每种扩展名的最新 IFR
    kept_ifc_all: List[DwgFile] = field(default_factory=list)  # v5.0: 每种扩展名的最新 IFC
    ab_files: List[DwgFile] = field(default_factory=list)      # v5.1: AS BUILT (never renamed)


def _classify_dwg(filename: str) -> Tuple[str, str]:
    """Classify a filename as AB (AS BUILT), IFR, IFC, or OTHER.

    AB is checked FIRST: AS BUILT deliverables often carry a numeric 'Rev<n>'
    that would otherwise match the IFC pattern, causing them to be renamed to
    '_IFC' (corrupting the deliverable). AB files are kept out of the IFR/IFC
    rename+supersede logic entirely.
    """
    stem = Path(filename).stem
    if _RE_AB.search(stem):
        mr = re.search(r'[_\s-](?:[Rr]ev|[Rr])\s*(\d+)', stem)
        return ('AB', mr.group(1) if mr else '')
    m = _RE_IFC.search(stem)
    if m:
        return ('IFC', m.group(1))
    m = _RE_IFR.search(stem)
    if m:
        return ('IFR', m.group(1).upper())
    return ('OTHER', '')


def _rev_sort_key(f) -> Tuple[int, object]:
    """版本新旧排序键 (越大越新): (版本号, mtime)。版本号优先于 mtime —— Dropbox 同步会改
    mtime, 纯 mtime 会把更新的修订误判为旧版、错降进 SS。IFR=字母 A<B<C (A→1), IFC=数字 0<1<2;
    mtime 仅在版本号相同时平手。SSOT: mainv3 §4.3 + 记忆 feedback_current_version_by_mtime。"""
    rev = (getattr(f, 'revision', '') or '')
    rt = getattr(f, 'rev_type', '')
    if rt == 'IFC':
        try:
            rank = int(re.sub(r'\D', '', rev) or -1)
        except (ValueError, TypeError):
            rank = -1
    elif rt == 'IFR':
        rank = (ord(rev[0].upper()) - ord('A') + 1) if rev[:1].isalpha() else 0
    else:
        rank = -1
    return (rank, getattr(f, 'mtime', None))


def _parse_folder_name(folder_name: str) -> Tuple[str, str]:
    """从文件夹名中尽力提取 doc_id 和 description。

    策略：
    1. 尝试用已知 doc_id 模式匹配 (GG 前缀 或 50023-XX-NNN 格式)
    2. 匹配失败则以第一个 '_' 分割
    3. 全部失败返回 (folder_name, '')
    """
    # 宽松的 doc_id 模式 (不要求特定前缀)
    m = re.match(r'^([A-Z0-9][\w-]+-\d{3})[_\s](.+?)(?:[_\s]Rev.*)?$', folder_name, re.IGNORECASE)
    if m:
        return (m.group(1), m.group(2).strip())

    # fallback: 用第一个 '_' 分割
    if '_' in folder_name:
        parts = folder_name.split('_', 1)
        return (parts[0], parts[1].rstrip().rstrip('_Rev').strip())

    return (folder_name, '')


def _make_standard_dwg_name(doc_id: str, description: str, rev_type: str,
                            revision: str, ext: str = '.dwg') -> str:
    """Generate a standard filename. IFC gets _IFC suffix."""
    if rev_type == 'IFC':
        return f"{doc_id}_{description}_Rev{revision}_IFC{ext}"
    return f"{doc_id}_{description}_Rev{revision.upper()}{ext}"


def _belongs_to_doc(filename: str, doc_id: Optional[str]) -> bool:
    """True if filename plausibly belongs to doc_id. Used to AVOID moving foreign
    base/survey XREF drawings (e.g. 'BaseDrawings.dwg', '53666-...-SV-...dwg') out
    of a doc folder — relocating an active XREF target breaks the host's xref link.
    Conservative: unknown doc_id → treat as foreign (don't move)."""
    if not doc_id:
        return False
    stem = Path(filename).stem.lower()
    did = doc_id.lower()
    if did in stem:
        return True
    # also accept the discipline-seq tail, e.g. 'ga-001' from '50023-ga-001'
    m = re.search(r'([a-z]{1,3}-?\d{3})$', did)
    return bool(m and m.group(1) in stem)


def _to_long_path(path: Path) -> Path:
    """Convert path to long path format on Windows (>260 char support)."""
    if os.name == 'nt':
        path_str = str(path.resolve())
        if not path_str.startswith('\\\\?\\'):
            path_str = '\\\\?\\' + path_str
        return Path(path_str)
    return path


class NativeVersionManager:
    """Manage file versions across Native, Reports, and Schedule folders (v3.0)."""

    def __init__(self, project_path: str, dry_run: bool = True,
                 ifc_folder_mode: bool = False):
        self.project_path = Path(project_path)
        self.dry_run = dry_run
        # v5.1 (SSOT mainv3 §4.1, 2026-06-04 milestone-folder convention):
        # when True, older IFC versions go to a dedicated 'IFC/' subfolder
        # (not 'SS/'), and AS BUILT files are routed to 'Rev.{N} - AB/'.
        self.ifc_folder_mode = ifc_folder_mode
        self.results: List[Tuple[str, NativeFolderResult]] = []

    @property
    def native_root(self) -> Path:
        """Backward-compatible: returns Native path."""
        return self.project_path / "Design" / "Engineering" / "1. Drawings" / "1. Native"

    # --- v3.0: folder discovery ---

    def find_doc_folders(self, folder_filter: Optional[str] = None,
                         scope: str = 'all') -> List[Tuple[str, Path]]:
        """发现所有文档文件夹。

        Returns: [(display_group, folder_path), ...]
        """
        results = []

        for rel_path, depth, display_name in SCAN_ROOTS:
            # --scope 过滤
            if scope != 'all':
                if scope == 'native' and 'Native' not in rel_path:
                    continue
                if scope == 'reports' and 'Reports' not in rel_path:
                    continue
                if scope == 'schedule' and 'Schedule' not in rel_path:
                    continue

            root = self.project_path / rel_path
            if not root.exists():
                continue

            if depth == 1:
                # 直接子文件夹就是文档文件夹
                candidates = self._list_doc_dirs(root)
            elif depth == 2:
                # 先遍历分类目录，再找文档文件夹
                candidates = []
                for category_dir in sorted(root.iterdir()):
                    if category_dir.is_dir() and not self._should_skip(category_dir.name):
                        candidates.extend(self._list_doc_dirs(category_dir))
            else:
                candidates = []

            # 关键字过滤
            for folder in candidates:
                if folder_filter and folder_filter.lower() not in folder.name.lower():
                    continue
                results.append((display_name, folder))

        return results

    def _list_doc_dirs(self, parent: Path) -> List[Path]:
        """列出 parent 下所有非排除文件夹，按名称排序。"""
        dirs = []
        try:
            for item in sorted(parent.iterdir(), key=lambda p: p.name):
                if item.is_dir() and not self._should_skip(item.name):
                    dirs.append(item)
        except (OSError, PermissionError):
            pass
        return dirs

    @staticmethod
    def _should_skip(name: str) -> bool:
        """判断文件夹名是否应跳过。"""
        if name.startswith('.'):
            return True
        low = name.lower()
        if low in _SKIP_FOLDER_NAMES:
            return True
        for prefix in _SKIP_FOLDER_PREFIXES:
            if low.startswith(prefix):
                return True
        return False

    # --- v3.0: superseded folder detection ---

    @staticmethod
    def _find_or_create_ss_folder(folder: Path, create: bool = False) -> Path:
        """在 folder 下查找已有的 superseded 文件夹，或返回默认 SS/ 路径。

        查找顺序 (不区分大小写): SS -> SUPERSEDED -> Superceded
        如果都不存在且 create=True，创建 SS/。
        """
        try:
            for item in folder.iterdir():
                if item.is_dir() and item.name.lower() in ('ss', 'superseded', 'superceded'):
                    return item
        except (OSError, PermissionError):
            pass

        ss_path = folder / 'SS'
        if create:
            _to_long_path(ss_path).mkdir(exist_ok=True)
        return ss_path

    # --- v3.0: scan all file types ---

    def scan_files(self, folder: Path) -> List[DwgFile]:
        """扫描文件夹中的所有工作文件（排除忽略列表和 superseded 文件夹）。"""
        files = []
        bak_map: Dict[str, Path] = {}

        # superseded 文件夹名
        ss_names = {'ss', 'superseded', 'superceded'}

        try:
            all_items = list(folder.iterdir())
        except (OSError, PermissionError):
            return []

        # 收集 .bak 文件映射
        for f in all_items:
            if f.is_file() and f.suffix.lower() == '.bak':
                bak_map[f.stem.lower()] = f

        for f in all_items:
            if not f.is_file():
                continue
            # 跳过 superseded 子文件夹中的文件 (不应出现，但防御性处理)
            if f.parent.name.lower() in ss_names:
                continue
            ext = f.suffix.lower()
            if ext in _IGNORE_EXT or ext == '.bak':
                continue

            rev_type, revision = _classify_dwg(f.name)  # 分类逻辑对所有文件类型通用
            try:
                mtime = datetime.fromtimestamp(_to_long_path(f).stat().st_mtime)
            except (OSError, PermissionError):
                mtime = datetime.min

            files.append(DwgFile(
                path=f, filename=f.name,
                rev_type=rev_type, revision=revision,
                mtime=mtime, bak_path=bak_map.get(f.stem.lower()),
            ))
        return files

    def _reg_titles_for(self, folder: Path) -> Dict[str, List[str]]:
        """Client *Deliverables List* titles for register-canonical grouping — the
        shared cross-track oracle (register_membership). SOFT dep: {} → grouping
        stays byte-identical. Cached. Mirror of NativeVersionManager._reg_titles_for
        in ifr_automation_v10.py (same SSOT)."""
        cache = getattr(self, '_reg_titles_cache', None)
        if cache is None:
            cache = self._reg_titles_cache = {}
        k = str(folder)
        if k not in cache:
            try:
                import register_membership as _rm
                cache[k] = _rm.load_register_titles(folder)
            except Exception:
                cache[k] = {}
        return cache[k]

    def _title_subkey(self, f: DwgFile, reg: Dict[str, List[str]]):
        """Register-canonical title sub-key to SPLIT genuinely-different drawings
        sharing a doc-id folder (CA-001 = Trench Alignment + Cable Route GA), so
        'newest per ext' never supersedes a real 2nd drawing. '' when no register
        (byte-identical); title token when matched; ('STRAY', name) to isolate a
        title matching no register name."""
        if not reg:
            return ''
        try:
            import register_membership as _rm
            did = _rm.extract_doc_id(f.filename)
            titles = reg.get((did or '').upper())
            if not titles:
                return ''
            _k, matched, stray = _rm.canonical_group_key('', f.filename, titles, did)
            if stray:
                self._folder_strays.append(f.filename)
                return ('STRAY', f.filename)
            self._folder_titles.add(_rm.norm_title(matched))
            return _rm.norm_title(matched)
        except Exception:
            return ''

    def process_folder(self, folder: Path) -> NativeFolderResult:
        """Process a single document folder: classify, decide keeps, plan actions.

        v5.0 改进:
        - 每种扩展名(.docx, .pdf, .dwg, .xlsx 等)各自保留最新版本
          (之前只保留整个 IFR/IFC 类别中 mtime 最新的单个文件，
           导致最新 .docx 被移到 SS)
        - OTHER 文件(无版本号)只在存在同扩展名的版本化文件时才移到 SS
          (之前把唯一的 Calculation.xlsx 等辅助文件也移到了 SS)
        """
        doc_id, description = _parse_folder_name(folder.name)
        result = NativeFolderResult(
            folder_name=folder.name, folder_path=folder,
            doc_id=doc_id or folder.name, description=description or '',
        )
        files = self.scan_files(folder)
        result.dwg_files = files
        if not files:
            return result

        ifr_files = [f for f in files if f.rev_type == 'IFR']
        ifc_files = [f for f in files if f.rev_type == 'IFC']
        ab_files = [f for f in files if f.rev_type == 'AB']
        other_files = [f for f in files if f.rev_type == 'OTHER']
        result.ab_files = ab_files
        ss_folder = self._find_or_create_ss_folder(folder)  # v3.0: 自动检测
        # v5.1: in milestone-folder mode, older IFC → dedicated 'IFC/' subfolder
        ifc_dest = (folder / 'IFC') if self.ifc_folder_mode else ss_folder

        # Register-canonical title split (soft): a doc-id folder holding TWO real
        # drawings must NOT collapse to one "newest". '' subkey when no register →
        # byte-identical grouping. (mirror of the engine NativeVersionManager)
        reg = self._reg_titles_for(folder)
        self._folder_strays = []
        self._folder_titles = set()
        _sub = lambda f: self._title_subkey(f, reg)

        # v5.0: IFR — 按 (扩展名 × 图名) 分组，每组保留最新版 (版本号优先, mtime 仅平手)
        if ifr_files:
            ifr_by_ext: dict[tuple, list] = {}
            for f in ifr_files:
                ifr_by_ext.setdefault((f.path.suffix.lower(), _sub(f)), []).append(f)
            for _k, group in ifr_by_ext.items():
                group_sorted = sorted(group, key=_rev_sort_key, reverse=True)
                result.kept_ifr_all.append(group_sorted[0])
                for old in group_sorted[1:]:
                    self._plan_move(result, old, ss_folder, "older IFR revision")
            result.kept_ifr = max(result.kept_ifr_all, key=_rev_sort_key)

        # v5.0: IFC — 按 (扩展名 × 图名) 分组，每组保留最新版 (版本号优先, mtime 仅平手)
        if ifc_files:
            ifc_by_ext: dict[tuple, list] = {}
            for f in ifc_files:
                ifc_by_ext.setdefault((f.path.suffix.lower(), _sub(f)), []).append(f)
            for _k, group in ifc_by_ext.items():
                group_sorted = sorted(group, key=_rev_sort_key, reverse=True)
                result.kept_ifc_all.append(group_sorted[0])
                for old in group_sorted[1:]:
                    action_type = 'move_to_ifc' if self.ifc_folder_mode else 'move_to_ss'
                    self._plan_move(result, old, ifc_dest, "older IFC revision",
                                    action_type=action_type)
            result.kept_ifc = max(result.kept_ifc_all, key=_rev_sort_key)

        # v5.1: AS BUILT — never renamed/superseded. In milestone-folder mode,
        # route to the per-drawing 'Rev.{N} - AB/' folder; otherwise leave in place.
        if ab_files and self.ifc_folder_mode:
            ab_folder = self._ab_folder_path(folder, ab_files)
            for ab in ab_files:
                self._plan_move(result, ab, ab_folder, "AS BUILT -> AB folder",
                                action_type='move_to_ab')

        # v5.0: OTHER — 仅当存在同扩展名的版本化文件时才移到 SS
        # (保护唯一的辅助计算文件如 Calculation.xlsx)
        # v5.1: 额外保护 — 不搬不属于本 doc-id 的文件 (base/survey XREF 底图，
        #       搬走会断 xref 链接，如 'BaseDrawings.dwg' / '53666-...-SV-...dwg')
        for other in other_files:
            ext = other.path.suffix.lower()
            has_versioned_same_ext = any(
                f.path.suffix.lower() == ext and f.rev_type != 'OTHER'
                for f in files
            )
            if has_versioned_same_ext and _belongs_to_doc(other.filename, doc_id):
                self._plan_move(result, other, ss_folder, "unversioned/legacy (versioned copy exists)")

        # v5.0: 对所有保留的文件执行重命名（不只是主文件）。AS BUILT 不在此列表，永不改名。
        # 当客户清单确认本文件夹含 ≥2 张不同图时跳过：标准名只用文件夹描述，会把两张图
        # 重命名成同名互相覆盖 (数据丢失)。
        multi_title = len(self._folder_titles) > 1
        if doc_id and description and not multi_title:
            for kept in result.kept_ifr_all:
                self._plan_rename(result, kept, doc_id, description)
            for kept in result.kept_ifc_all:
                self._plan_rename(result, kept, doc_id, description)
        if multi_title:
            print(f"    [i] {folder.name}: 客户清单确认含 {len(self._folder_titles)} 张不同图，"
                  f"各自独立保留最新版、跳过按文件夹名的标准化重命名 (避免同名覆盖)")
        if self._folder_strays:
            print(f"    [i] {folder.name}: {len(self._folder_strays)} 个文件对不上客户清单标题，"
                  f"单独保留不合并 (可能编号有误): {', '.join(self._folder_strays[:5])}")

        return result

    @staticmethod
    def _ab_folder_path(folder: Path, ab_files: List[DwgFile]) -> Path:
        """Return the per-drawing AS BUILT folder: reuse an existing 'Rev.* - AB'
        if present, else default to 'Rev.{N} - AB' (N = AS BUILT rev, default 1)."""
        try:
            for item in folder.iterdir():
                if item.is_dir() and re.match(r'rev\.?\s*\d+\s*-\s*ab$', item.name, re.IGNORECASE):
                    return item
        except (OSError, PermissionError):
            pass
        revs = [f.revision for f in ab_files if f.revision]
        n = revs[0] if revs else '1'
        return folder / f"Rev.{n} - AB"

    def _plan_move(self, result: NativeFolderResult, dwg: DwgFile,
                   dest_folder: Path, reason: str, action_type: str = 'move_to_ss'):
        result.actions.append(NativeFileAction(
            source=dwg.path, dest=dest_folder / dwg.filename,
            action=action_type, reason=reason
        ))
        if dwg.bak_path and dwg.bak_path.exists():
            result.actions.append(NativeFileAction(
                source=dwg.bak_path, dest=dest_folder / dwg.bak_path.name,
                action=action_type, reason=f".bak follows {dwg.filename}"
            ))

    def _plan_rename(self, result: NativeFolderResult, dwg: DwgFile,
                     doc_id: str, description: str):
        ext = dwg.path.suffix  # v3.0: 保持原始扩展名
        standard = _make_standard_dwg_name(doc_id, description, dwg.rev_type,
                                           dwg.revision, ext)
        if dwg.filename != standard:
            result.actions.append(NativeFileAction(
                source=dwg.path, dest=dwg.path.parent / standard,
                action='rename', reason=f"标准化 -> {standard}"
            ))
            if dwg.bak_path and dwg.bak_path.exists():
                bak_std = _make_standard_dwg_name(doc_id, description,
                                                  dwg.rev_type, dwg.revision, '.bak')
                if dwg.bak_path.name != bak_std:
                    result.actions.append(NativeFileAction(
                        source=dwg.bak_path, dest=dwg.bak_path.parent / bak_std,
                        action='rename', reason=f".bak 标准化 -> {bak_std}"
                    ))

    def process_all(self, folder_filter: Optional[str] = None,
                    scope: str = 'all') -> List[Tuple[str, NativeFolderResult]]:
        """处理所有文档文件夹。返回 [(display_group, result), ...]"""
        doc_folders = self.find_doc_folders(folder_filter, scope)
        results = []
        for group_name, folder in doc_folders:
            result = self.process_folder(folder)
            results.append((group_name, result))
        self.results = results
        return results

    def execute_actions(self, results: List[Tuple[str, NativeFolderResult]]) -> Dict[str, int]:
        """Execute planned actions. v3.0: uses _find_or_create_ss_folder."""
        stats = {'renamed': 0, 'moved': 0, 'errors': 0, 'ss_created': 0}

        for _group_name, result in results:
            if not result.actions:
                continue

            moves = [a for a in result.actions if a.action == 'move_to_ss']
            if moves:
                ss_folder = self._find_or_create_ss_folder(result.folder_path, create=True)
                if not ss_folder.exists():
                    try:
                        _to_long_path(ss_folder).mkdir(exist_ok=True)
                        stats['ss_created'] += 1
                    except (OSError, PermissionError):
                        try:
                            ss_folder.mkdir(exist_ok=True)
                            stats['ss_created'] += 1
                        except Exception as e:
                            print(f"      [!] 无法创建 SS: {e}")
                            stats['errors'] += 1
                            continue
                # Update move destinations to use the detected ss_folder
                for action in moves:
                    action.dest = ss_folder / action.source.name

            # v5.1: ensure IFC/ and Rev.N-AB/ destination dirs exist (not redirected)
            for action in result.actions:
                if action.action in ('move_to_ifc', 'move_to_ab'):
                    try:
                        _to_long_path(action.dest.parent).mkdir(parents=True, exist_ok=True)
                    except (OSError, PermissionError):
                        try:
                            action.dest.parent.mkdir(parents=True, exist_ok=True)
                        except Exception as e:
                            print(f"      [!] 无法创建 {action.dest.parent.name}: {e}")
                            stats['errors'] += 1

            for action in result.actions:
                try:
                    src = str(_to_long_path(action.source))
                    dst_path = action.dest
                    if _to_long_path(dst_path).exists():
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        dst_path = dst_path.parent / f"{dst_path.stem}_{ts}{dst_path.suffix}"
                    dst = str(_to_long_path(dst_path))

                    shutil.move(src, dst)
                    if action.action == 'rename':
                        stats['renamed'] += 1
                    else:
                        stats['moved'] += 1
                except Exception as e:
                    print(f"      [!] 失败 {action.source.name}: {e}")
                    stats['errors'] += 1

        return stats


# =============================================================================
# Folder Relocator — detect & fix misplaced structural folders
# =============================================================================

# Canonical level-1 folders under "1. Drawings/"
_L1_PATTERNS = [
    (re.compile(r'^1[\.\s]*Native', re.IGNORECASE),        '1. Native'),
    (re.compile(r'^2[\.\s]*IFR[_\s]*internal', re.IGNORECASE), '2. IFR_internal'),
    (re.compile(r'^3[\.\s]*IFR\s*\(Client\)', re.IGNORECASE), '3. IFR(Client)'),
    (re.compile(r'^4[\.\s]*IFC\s*\(Client\)', re.IGNORECASE), '4. IFC(Client)'),
]

# Canonical level-2 folders (children of 3. IFR(Client)/)
_L2_PATTERNS = [
    (re.compile(r'^1[\.\s]*Drawing', re.IGNORECASE),       '1.Drawing'),
    (re.compile(r'^2[\.\s]*Reports?', re.IGNORECASE),      '2.Reports'),
    (re.compile(r'^3[\.\s]*Deliverables?', re.IGNORECASE), '3.Deliverables'),
]

# Level-1 folders that are "flat" (should NOT contain structural subfolders)
_FLAT_L1 = {'1. Native', '2. IFR_internal', '4. IFC(Client)'}

# Expected children of 3. IFR(Client)/
_IFR_CLIENT_CHILDREN = {'1.Drawing', '2.Reports', '3.Deliverables'}


class FolderRelocator:
    """Detect and relocate misplaced structural folders under 1. Drawings/."""

    def __init__(self, project_path: Path, dry_run: bool = True):
        self.project_path = project_path
        self.drawings_root = project_path / "Design" / "Engineering" / "1. Drawings"
        self.dry_run = dry_run

    @staticmethod
    def _match_l1(name: str) -> Optional[str]:
        """Match a folder name to a canonical level-1 name, or None."""
        for pat, canonical in _L1_PATTERNS:
            if pat.match(name):
                return canonical
        return None

    @staticmethod
    def _match_l2(name: str) -> Optional[str]:
        """Match a folder name to a canonical level-2 name, or None."""
        for pat, canonical in _L2_PATTERNS:
            if pat.match(name):
                return canonical
        return None

    def _ifr_client_path(self) -> Path:
        """Return the path to 3. IFR(Client)/, detecting actual folder name."""
        try:
            for item in self.drawings_root.iterdir():
                if item.is_dir() and self._match_l1(item.name) == '3. IFR(Client)':
                    return item
        except (OSError, PermissionError):
            pass
        return self.drawings_root / '3. IFR(Client)'

    def scan(self) -> List[FolderRelocation]:
        """Scan 1. Drawings/ and return a list of planned relocations."""
        relocations: List[FolderRelocation] = []

        if not self.drawings_root.exists():
            return relocations

        ifr_client = self._ifr_client_path()

        # --- Pass 1: direct children of 1. Drawings/ (depth 0) ---
        try:
            depth0_items = sorted(self.drawings_root.iterdir(), key=lambda p: p.name)
        except (OSError, PermissionError):
            return relocations

        for item in depth0_items:
            if not item.is_dir():
                continue

            l1_match = self._match_l1(item.name)
            l2_match = self._match_l2(item.name)

            if l2_match:
                # Level-2 subfolder directly under 1. Drawings/ → move to IFR(Client)
                dest = ifr_client / l2_match
                relocations.append(FolderRelocation(
                    source=item, dest=dest,
                    reason=f"Level-2 文件夹 '{item.name}' 不应直接在 1. Drawings/ 下",
                    action='merge' if dest.exists() else 'move',
                ))
                continue

            if not l1_match:
                continue

            # --- Pass 2: children of each level-1 folder (depth 1) ---
            try:
                depth1_items = sorted(item.iterdir(), key=lambda p: p.name)
            except (OSError, PermissionError):
                continue

            for child in depth1_items:
                if not child.is_dir():
                    continue

                child_l1 = self._match_l1(child.name)
                child_l2 = self._match_l2(child.name)

                if child_l1:
                    # Level-1 nested inside another level-1 → move back to Drawings
                    dest = self.drawings_root / child_l1
                    relocations.append(FolderRelocation(
                        source=child, dest=dest,
                        reason=f"Level-1 文件夹 '{child.name}' 被误放在 '{item.name}/' 内",
                        action='merge' if dest.exists() else 'move',
                    ))
                elif child_l2 and l1_match in _FLAT_L1:
                    # Level-2 subfolder inside a flat level-1 → move to IFR(Client)
                    dest = ifr_client / child_l2
                    relocations.append(FolderRelocation(
                        source=child, dest=dest,
                        reason=f"'{child.name}' 不应在 flat 文件夹 '{item.name}/' 内",
                        action='merge' if dest.exists() else 'move',
                    ))
                elif child_l2 and l1_match == '3. IFR(Client)':
                    # Check for level-2 nested inside another level-2
                    self._scan_deep(child, item, ifr_client, relocations)

        # --- Pass 3: warn about missing expected children of IFR(Client) ---
        if ifr_client.exists():
            existing_children = set()
            try:
                for c in ifr_client.iterdir():
                    if c.is_dir():
                        m = self._match_l2(c.name)
                        if m:
                            existing_children.add(m)
            except (OSError, PermissionError):
                pass

            for expected in _IFR_CLIENT_CHILDREN:
                if expected not in existing_children:
                    relocations.append(FolderRelocation(
                        source=ifr_client / expected,
                        dest=ifr_client / expected,
                        reason=f"3. IFR(Client)/ 缺少子文件夹 '{expected}'",
                        action='warn',
                    ))

        return relocations

    def _scan_deep(self, folder: Path, parent_l1: Path,
                   ifr_client: Path, relocations: List[FolderRelocation]):
        """Check for structural folders nested 2+ levels deep."""
        try:
            children = sorted(folder.iterdir(), key=lambda p: p.name)
        except (OSError, PermissionError):
            return

        for child in children:
            if not child.is_dir():
                continue
            child_l1 = self._match_l1(child.name)
            child_l2 = self._match_l2(child.name)

            if child_l1:
                dest = self.drawings_root / child_l1
                relocations.append(FolderRelocation(
                    source=child, dest=dest,
                    reason=f"Level-1 '{child.name}' 嵌套在 '{folder.name}/' 内 (深层错位)",
                    action='merge' if dest.exists() else 'move',
                ))
            elif child_l2:
                dest = ifr_client / child_l2
                relocations.append(FolderRelocation(
                    source=child, dest=dest,
                    reason=f"Level-2 '{child.name}' 嵌套在 '{folder.name}/' 内 (深层错位)",
                    action='merge' if dest.exists() else 'move',
                ))

    def execute(self, relocations: List[FolderRelocation]) -> Dict[str, int]:
        """Execute planned relocations. Returns stats dict."""
        stats = {'moved': 0, 'merged': 0, 'warnings': 0, 'errors': 0}

        for rel in relocations:
            if rel.action == 'warn':
                stats['warnings'] += 1
                continue

            try:
                src = _to_long_path(rel.source)
                dst = _to_long_path(rel.dest)

                if rel.action == 'move':
                    # Destination doesn't exist — simple move
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(src), str(dst))
                    print(f"  [v] 已移动: {rel.source.name} -> {rel.dest}")
                    stats['moved'] += 1

                elif rel.action == 'merge':
                    # Destination exists — move children one by one
                    dst.mkdir(parents=True, exist_ok=True)
                    moved_children = 0
                    try:
                        for child in src.iterdir():
                            child_dst = dst / child.name
                            if child_dst.exists():
                                # Add timestamp suffix to avoid collision
                                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                                if child.is_file():
                                    child_dst = dst / f"{child.stem}_{ts}{child.suffix}"
                                else:
                                    child_dst = dst / f"{child.name}_{ts}"
                            shutil.move(str(child), str(child_dst))
                            moved_children += 1
                    except Exception as e:
                        print(f"  [!] 合并部分失败 {rel.source.name}: {e}")
                        stats['errors'] += 1

                    # Remove empty source folder
                    try:
                        if src.exists() and not any(src.iterdir()):
                            src.rmdir()
                    except (OSError, PermissionError):
                        pass

                    print(f"  [v] 已合并: {rel.source.name} -> {rel.dest} ({moved_children} 项)")
                    stats['merged'] += 1

            except Exception as e:
                print(f"  [!] 失败: {rel.source.name} - {e}")
                stats['errors'] += 1

        return stats


class InteractiveUI:
    """Interactive user interface for Version Manager."""

    def __init__(self):
        self.root_path: Optional[Path] = None
        self.manager: Optional[VersionManager] = None

    def clear_screen(self):
        """Clear console screen."""
        os.system('cls' if os.name == 'nt' else 'clear')

    def print_header(self):
        """Print application header."""
        print("\n" + "=" * 70)
        print("            版本文件管理工具 - Version Manager v3.0")
        print("           自动识别并移动旧版本文件到 SS 文件夹")
        print("      PDF (IFR/IFC) + Native/Reports/Schedule (all files)")
        print("=" * 70)

    def print_separator(self):
        """Print separator line."""
        print("-" * 70)

    def detect_root_path(self) -> Optional[Path]:
        """Try to detect the project root path."""
        possible_roots = [
            Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"),
            Path.home() / "Dropbox" / "Projects" / "Project (EPC)",
        ]

        for root in possible_roots:
            if root.exists():
                return root

        return None

    def get_root_path(self) -> Optional[Path]:
        """Get root path from user or auto-detect."""
        detected = self.detect_root_path()

        if detected:
            print(f"\n检测到项目根目录:")
            print(f"  {detected}")
            choice = input("\n使用此目录? [Y/n]: ").strip().lower()
            if choice in ('', 'y', 'yes'):
                return detected

        print("\n请输入项目根目录路径:")
        print("（可以直接拖拽文件夹到此窗口）")
        path_str = input("> ").strip().strip('"')

        if path_str:
            path = Path(path_str)
            if path.exists():
                return path
            else:
                print(f"\n[!] 路径不存在: {path}")

        return None

    def show_main_menu(self):
        """Display main menu and return choice."""
        print("\n主菜单:")
        print()
        print("  --- PDF 版本管理 (IFR_internal / IFR(Client)) ---")
        print("  [1] 扫描所有项目 (预览)")
        print("  [2] 处理所有项目")
        print("  [3] 选择单个项目处理")
        print()
        print("  --- Native/Reports/Schedule 版本管理 ---")
        print("  [5] Native/Reports/Schedule 版本管理")
        print()
        print("  --- 结构修复 ---")
        print("  [6] 文件夹归位 (检测错位的结构文件夹)")
        print()
        print("  [4] 修改根目录")
        print("  [0] 退出")
        print()
        self.print_separator()

        return input("\n请选择 [0-6]: ").strip()

    def scan_all_projects_preview(self):
        """Scan all projects and show preview."""
        print("\n正在扫描项目...")

        projects = self.manager.scan_for_projects()

        if not projects:
            print("\n[!] 未找到任何项目")
            return

        print(f"\n找到 {len(projects)} 个项目:\n")
        self.print_separator()

        total_to_move = 0

        for i, project in enumerate(projects, 1):
            print(f"\n[{i}] {project.name}")
            print(f"    路径: {project}")

            # Analyze each subdirectory
            for subdir in VersionManager.TARGET_SUBDIRS:
                target_dir = project / subdir
                short_name = subdir.split("\\")[-1]

                analysis = self.manager.analyze_directory(target_dir)

                if analysis["exists"]:
                    if analysis["files_to_move"] > 0:
                        print(f"    {short_name}: {analysis['total_files']} 文件, {analysis['files_to_move']} 待移动")
                        total_to_move += analysis["files_to_move"]
                    elif analysis["total_files"] > 0:
                        print(f"    {short_name}: {analysis['total_files']} 文件, [v] 无重复")

        self.print_separator()
        print(f"\n汇总: 共 {total_to_move} 个旧版本文件待移动")

        input("\n按 Enter 返回主菜单...")

    def process_all_projects(self):
        """Process all projects with confirmation."""
        print("\n正在扫描项目...")

        projects = self.manager.scan_for_projects()

        if not projects:
            print("\n[!] 未找到任何项目")
            input("\n按 Enter 返回...")
            return

        # First show preview
        print(f"\n找到 {len(projects)} 个项目")
        print("\n将处理以下目录:")
        for subdir in VersionManager.TARGET_SUBDIRS:
            print(f"  - {subdir}")

        print("\n" + "!" * 70)
        print("  警告: 此操作将移动所有项目中的旧版本文件到 SS 文件夹")
        print("!" * 70)

        confirm = input("\n确认执行? 输入 'YES' 确认: ").strip()
        if confirm != 'YES':
            print("\n已取消操作")
            input("\n按 Enter 返回...")
            return

        # Process each project
        self.print_separator()
        total_stats = {"scanned": 0, "groups": 0, "moved": 0}

        for i, project in enumerate(projects, 1):
            print(f"\n[{i}/{len(projects)}] 处理项目: {project.name}")
            self.print_separator()

            stats = self.manager.process_project(project)

            total_stats["scanned"] += stats["scanned"]
            total_stats["groups"] += stats["groups"]
            total_stats["moved"] += stats["moved"]

        # Summary
        print("\n" + "=" * 70)
        print("                         处理完成")
        print("=" * 70)
        print(f"  处理项目数: {len(projects)}")
        print(f"  扫描文件数: {total_stats['scanned']}")
        print(f"  多版本组数: {total_stats['groups']}")
        print(f"  移动文件数: {total_stats['moved']}")

        input("\n按 Enter 返回主菜单...")

    def select_single_project(self):
        """Select and process a single project."""
        print("\n正在扫描项目...")

        projects = self.manager.scan_for_projects()

        if not projects:
            print("\n[!] 未找到任何项目")
            input("\n按 Enter 返回...")
            return

        # Show project list
        print(f"\n找到 {len(projects)} 个项目:\n")

        for i, project in enumerate(projects, 1):
            print(f"  [{i}] {project.name}")

        print(f"\n  [0] 返回主菜单")

        self.print_separator()

        choice = input("\n请选择项目 [0-" + str(len(projects)) + "]: ").strip()

        try:
            idx = int(choice)
            if idx == 0:
                return
            if 1 <= idx <= len(projects):
                project = projects[idx - 1]
                self.process_single_project(project)
        except ValueError:
            print("\n[!] 无效选择")
            input("\n按 Enter 返回...")

    def process_single_project(self, project: Path):
        """Process a single project with detailed preview and confirmation."""
        self.clear_screen()
        self.print_header()

        print(f"\n项目: {project.name}")
        print(f"路径: {project}")

        self.print_separator()

        # Preview first
        print("\n预览分析:")

        total_to_move = 0
        details = []

        for subdir in VersionManager.TARGET_SUBDIRS:
            target_dir = project / subdir
            short_name = subdir.split("\\")[-1]

            analysis = self.manager.analyze_directory(target_dir)

            if analysis["exists"] and analysis["files_to_move"] > 0:
                total_to_move += analysis["files_to_move"]
                details.append((short_name, target_dir, analysis))

                print(f"\n  [{short_name}]")
                print(f"    PDF 文件: {analysis['total_files']}")
                print(f"    多版本组: {len(analysis['multi_version_groups'])}")
                if analysis.get('ifc_files_count', 0) > 0:
                    print(f"    IFC 文件: {analysis['ifc_files_count']}")
                print(f"    待移动: {analysis['files_to_move']}")

                # Show details of each group
                for base_name, files in analysis['multi_version_groups'].items():
                    files_sorted = sorted(files, key=VersionManager._pdf_ver_key, reverse=True)  # 与决策同序
                    print(f"\n      基础文件: {base_name}")
                    for i, (file_path, version, modified_time) in enumerate(files_sorted):
                        status = "[保留]" if i == 0 else "[->SS]"
                        time_str = modified_time.strftime('%Y-%m-%d %H:%M')
                        print(f"        {status} {file_path.name}")
                        print(f"               版本: {version}, 修改: {time_str}")

                # Show IFC files found in IFR(Client)
                for src, dst, reason in analysis.get('ifc_files', []):
                    print(f"\n      [->SS] {src.name} ({reason})")

        if total_to_move == 0:
            print("\n[v] 此项目没有需要移动的旧版本文件")
            input("\n按 Enter 返回...")
            return

        self.print_separator()
        print(f"\n共 {total_to_move} 个文件将被移动到 SS 文件夹")

        print("\n选择操作:")
        print("  [1] 执行移动")
        print("  [2] 仅预览 (dry-run)")
        print("  [0] 取消返回")

        choice = input("\n请选择: ").strip()

        if choice == '0':
            return
        elif choice == '2':
            # Dry run
            self.manager.dry_run = True
            print("\n[预览模式] 以下为预览结果，不会实际移动文件:")
            self.manager.process_project(project)
            self.manager.dry_run = False
        elif choice == '1':
            # Confirm execution
            confirm = input("\n确认执行? [Y/n]: ").strip().lower()
            if confirm in ('', 'y', 'yes'):
                print("\n正在处理...")
                stats = self.manager.process_project(project)

                print("\n" + "=" * 70)
                print("                         处理完成")
                print("=" * 70)
                print(f"  扫描文件: {stats['scanned']}")
                print(f"  多版本组: {stats['groups']}")
                print(f"  移动文件: {stats['moved']}")
            else:
                print("\n已取消操作")

        input("\n按 Enter 返回主菜单...")

    def native_dwg_menu(self):
        """Native/Reports/Schedule version management sub-menu (v3.0)."""
        self.clear_screen()
        self.print_header()
        print("\n  === Native/Reports/Schedule 版本管理 ===")

        # Need a specific project path (not root)
        print("\n选择项目:")
        projects = self.manager.scan_for_projects()

        if not projects:
            print("\n[!] 未找到任何项目")
            input("\n按 Enter 返回...")
            return

        for i, project in enumerate(projects, 1):
            print(f"  [{i}] {project.name}")
        print(f"\n  [0] 返回主菜单")

        choice = input(f"\n请选择项目 [0-{len(projects)}]: ").strip()
        try:
            idx = int(choice)
            if idx == 0:
                return
            if 1 <= idx <= len(projects):
                project = projects[idx - 1]
            else:
                return
        except ValueError:
            return

        self.clear_screen()
        self.print_header()
        print(f"\n  项目: {project.name}")

        native_mgr = NativeVersionManager(str(project), dry_run=True)

        # v3.0: 选择范围
        print("\n  扫描范围:")
        print("    [1] 全部 (Native + Reports + Schedule)")
        print("    [2] 仅 Native")
        print("    [3] 仅 Reports")
        print("    [4] 仅 Schedule")
        scope_choice = input("\n  请选择 [1-4, 默认=1]: ").strip() or '1'
        scope_map = {'1': 'all', '2': 'native', '3': 'reports', '4': 'schedule'}
        scope = scope_map.get(scope_choice, 'all')

        # Ask for folder filter
        folder_filter = input("\n  筛选文件夹 (留空=全部, 或输入关键字如 PLN-001): ").strip() or None

        # Process (dry-run preview)
        results = native_mgr.process_all(folder_filter, scope)

        if not results:
            print("\n  没有找到文档文件夹")
            input("\n按 Enter 返回...")
            return

        # v3.0: 按 group_name 分组显示
        total_actions = sum(len(r.actions) for _, r in results)
        folders_with_files = sum(1 for _, r in results if r.dwg_files)

        print(f"\n  [预览] 扫描 {len(results)} 个文件夹, "
              f"{folders_with_files} 个有文件, {total_actions} 个动作")
        self.print_separator()

        for group_name, group_results in groupby(results, key=lambda x: x[0]):
            group_list = list(group_results)
            group_folders_with_files = [r for _, r in group_list if r.dwg_files]
            print(f"\n  === {group_name} === ({len(group_folders_with_files)} folders)")

            for _, result in group_list:
                if not result.dwg_files:
                    continue
                print(f"    {result.folder_name}")
                print(f"      doc_id: {result.doc_id}  |  {len(result.dwg_files)} files")

                for kept in result.kept_ifr_all:
                    print(f"      [保留 IFR] {kept.filename}"
                          f"  ({kept.rev_display}, "
                          f"{kept.mtime.strftime('%Y-%m-%d %H:%M')})")
                for kept in result.kept_ifc_all:
                    print(f"      [保留 IFC] {kept.filename}"
                          f"  ({kept.rev_display}, "
                          f"{kept.mtime.strftime('%Y-%m-%d %H:%M')})")

                for action in result.actions:
                    if action.action == 'rename':
                        print(f"      [重命名] {action.source.name}")
                        print(f"            -> {action.dest.name}")
                    elif action.action == 'move_to_ss':
                        print(f"      [->{action.dest.parent.name}/] {action.source.name}  ({action.reason})")

        self.print_separator()
        renames = sum(1 for _, r in results for a in r.actions if a.action == 'rename')
        moves = sum(1 for _, r in results for a in r.actions if a.action == 'move_to_ss')
        print(f"\n  汇总: {renames} 个重命名, {moves} 个移动到 SS/")

        if total_actions == 0:
            print("\n  [v] 所有文件已是最新标准格式")
            input("\n按 Enter 返回...")
            return

        print("\n  选择操作:")
        print("    [1] 执行")
        print("    [0] 返回 (不执行)")

        execute_choice = input("\n  请选择: ").strip()
        if execute_choice == '1':
            confirm = input("\n  确认执行? 输入 'YES': ").strip()
            if confirm == 'YES':
                native_mgr.dry_run = False
                stats = native_mgr.execute_actions(results)
                print(f"\n  完成: {stats['renamed']} 重命名, {stats['moved']} 移动, "
                      f"{stats['ss_created']} SS/ 创建, {stats['errors']} 错误")
            else:
                print("\n  已取消")

        input("\n按 Enter 返回主菜单...")

    def folder_relocate_menu(self):
        """Folder relocation sub-menu: detect and fix misplaced structural folders."""
        self.clear_screen()
        self.print_header()
        print("\n  === 文件夹归位 (检测错位的结构文件夹) ===")

        # Select project
        print("\n选择项目:")
        projects = self.manager.scan_for_projects()

        if not projects:
            print("\n[!] 未找到任何项目")
            input("\n按 Enter 返回...")
            return

        for i, project in enumerate(projects, 1):
            print(f"  [{i}] {project.name}")
        print(f"\n  [0] 返回主菜单")

        choice = input(f"\n请选择项目 [0-{len(projects)}]: ").strip()
        try:
            idx = int(choice)
            if idx == 0:
                return
            if 1 <= idx <= len(projects):
                project = projects[idx - 1]
            else:
                return
        except ValueError:
            return

        self.clear_screen()
        self.print_header()
        print(f"\n  项目: {project.name}")
        print(f"  路径: {project}")

        # Scan for misplaced folders
        relocator = FolderRelocator(project, dry_run=True)
        drawings_root = relocator.drawings_root

        if not drawings_root.exists():
            print(f"\n  [!] 1. Drawings/ 不存在: {drawings_root}")
            input("\n按 Enter 返回...")
            return

        print(f"\n  扫描: {drawings_root}")
        relocations = relocator.scan()

        if not relocations:
            print("\n  [v] 未检测到错位的文件夹，结构正确")
            input("\n按 Enter 返回...")
            return

        # Display results
        self.print_separator()
        actionable = [r for r in relocations if r.action != 'warn']
        warnings = [r for r in relocations if r.action == 'warn']

        if actionable:
            print(f"\n  检测到 {len(actionable)} 个错位文件夹:\n")
            for i, rel in enumerate(actionable, 1):
                action_str = '移动' if rel.action == 'move' else '合并'
                print(f"    [{i}] [{action_str}] {rel.source.name}")
                print(f"         从: {rel.source.parent}")
                print(f"         到: {rel.dest}")
                print(f"         原因: {rel.reason}")

        if warnings:
            print(f"\n  警告 ({len(warnings)}):")
            for rel in warnings:
                print(f"    [!] {rel.reason}")

        if not actionable:
            print("\n  没有需要移动的文件夹 (仅有警告)")
            input("\n按 Enter 返回...")
            return

        self.print_separator()
        print(f"\n  共 {len(actionable)} 个文件夹待归位")
        print("\n  选择操作:")
        print("    [1] 执行归位")
        print("    [0] 返回 (不执行)")

        execute_choice = input("\n  请选择: ").strip()
        if execute_choice == '1':
            confirm = input("\n  确认执行? 输入 'YES': ").strip()
            if confirm == 'YES':
                relocator.dry_run = False
                stats = relocator.execute(actionable)
                print(f"\n  完成: {stats['moved']} 移动, {stats['merged']} 合并, "
                      f"{stats['warnings']} 警告, {stats['errors']} 错误")
            else:
                print("\n  已取消")

        input("\n按 Enter 返回主菜单...")

    def run(self):
        """Run the interactive UI."""
        self.clear_screen()
        self.print_header()

        # Get root path
        self.root_path = self.get_root_path()
        if not self.root_path:
            print("\n[!] 未设置有效的根目录")
            input("\n按 Enter 退出...")
            return

        self.manager = VersionManager(str(self.root_path))

        # Main loop
        while True:
            self.clear_screen()
            self.print_header()

            print(f"\n当前根目录: {self.root_path}")
            self.print_separator()

            choice = self.show_main_menu()

            if choice == '0':
                print("\n再见!")
                break
            elif choice == '1':
                self.scan_all_projects_preview()
            elif choice == '2':
                self.process_all_projects()
            elif choice == '3':
                self.select_single_project()
            elif choice == '4':
                self.root_path = self.get_root_path()
                if self.root_path:
                    self.manager = VersionManager(str(self.root_path))
            elif choice == '5':
                self.native_dwg_menu()
            elif choice == '6':
                self.folder_relocate_menu()


def _load_project_path() -> Optional[str]:
    """Load project path from config.json."""
    cfg_file = Path(__file__).parent / 'config.json'
    if not cfg_file.exists():
        return None
    try:
        with open(cfg_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        recent = config.get('recent_projects', [])
        if recent:
            return recent[0]
        return config.get('default_root_path')
    except Exception:
        return None


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Version Manager v3.0 - PDF + Native/Reports/Schedule 版本管理',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('--native', action='store_true',
                        help='Native/Reports/Schedule 模式 (管理文档文件夹)')
    parser.add_argument('--relocate', action='store_true',
                        help='文件夹归位模式 (检测错位的结构文件夹)')
    parser.add_argument('--execute', action='store_true',
                        help='执行操作 (默认 dry-run 预览)')
    parser.add_argument('--project', type=str, default=None,
                        help='项目路径 (自动从 config.json 读取)')
    parser.add_argument('--folder', type=str, default=None,
                        help='仅处理匹配的文件夹 (关键字过滤)')
    parser.add_argument('--scope', type=str, default='all',
                        choices=['native', 'reports', 'schedule', 'all'],
                        help='扫描范围 (默认 all)')
    parser.add_argument('--ifc-folder', action='store_true',
                        help='v5.1 milestone 分筐 (SSOT §4.1): 旧 IFC → IFC/ (而非 SS/), '
                             'AS BUILT → Rev.N - AB/。默认关闭 (维持旧 SS/ 行为)')

    # If no args, run interactive mode
    if len(sys.argv) == 1:
        ui = InteractiveUI()
        ui.run()
        return

    args = parser.parse_args()

    if args.relocate:
        # CLI Folder Relocation mode
        project_path = args.project or _load_project_path()
        if not project_path:
            print("错误: 未指定项目路径。使用 --project 或在 config.json 中设置")
            sys.exit(1)

        project_path = Path(project_path)
        if not project_path.exists():
            print(f"错误: 路径不存在: {project_path}")
            sys.exit(1)

        dry_run = not args.execute
        mode_str = "[预览]" if dry_run else "[执行]"

        print(f"\n{'=' * 70}")
        print(f"  文件夹归位 {mode_str}")
        print(f"  项目: {project_path.name}")
        print(f"{'=' * 70}")

        relocator = FolderRelocator(project_path, dry_run=dry_run)

        if not relocator.drawings_root.exists():
            print(f"\n  [!] 1. Drawings/ 不存在: {relocator.drawings_root}")
            sys.exit(1)

        relocations = relocator.scan()

        if not relocations:
            print("\n  [v] 未检测到错位的文件夹，结构正确")
            sys.exit(0)

        actionable = [r for r in relocations if r.action != 'warn']
        warnings = [r for r in relocations if r.action == 'warn']

        if actionable:
            print(f"\n  检测到 {len(actionable)} 个错位文件夹:\n")
            for i, rel in enumerate(actionable, 1):
                action_str = '移动' if rel.action == 'move' else '合并'
                print(f"    [{i}] [{action_str}] {rel.source.name}")
                print(f"         从: {rel.source.parent}")
                print(f"         到: {rel.dest}")
                print(f"         原因: {rel.reason}")

        if warnings:
            print(f"\n  警告 ({len(warnings)}):")
            for rel in warnings:
                print(f"    [!] {rel.reason}")

        if actionable and not dry_run:
            confirm = input("\n  确认执行? 输入 'YES': ").strip()
            if confirm == 'YES':
                stats = relocator.execute(actionable)
                print(f"\n  完成: {stats['moved']} 移动, {stats['merged']} 合并, "
                      f"{stats['warnings']} 警告, {stats['errors']} 错误")
            else:
                print("  已取消")
        elif actionable and dry_run:
            print(f"\n  使用 --execute 执行操作")

    elif args.native:
        # CLI Native/Reports/Schedule mode
        project_path = args.project or _load_project_path()
        if not project_path:
            print("错误: 未指定项目路径。使用 --project 或在 config.json 中设置")
            sys.exit(1)

        project_path = Path(project_path)
        if not project_path.exists():
            print(f"错误: 路径不存在: {project_path}")
            sys.exit(1)

        dry_run = not args.execute
        mode_str = "[预览]" if dry_run else "[执行]"
        scope_str = f" (scope: {args.scope})" if args.scope != 'all' else ""

        print(f"\n{'=' * 70}")
        print(f"  Native/Reports/Schedule 版本管理 {mode_str}{scope_str}")
        print(f"  项目: {project_path.name}")
        print(f"{'=' * 70}")

        mgr = NativeVersionManager(str(project_path), dry_run=dry_run,
                                   ifc_folder_mode=args.ifc_folder)
        if args.ifc_folder:
            print("  [v5.1] milestone 分筐模式: 旧 IFC → IFC/, AS BUILT → Rev.N - AB/, "
                  "AS BUILT 永不改名, 跳过 xref 底图")

        results = mgr.process_all(folder_filter=args.folder, scope=args.scope)
        if not results:
            print("\n  没有找到文档文件夹")
            sys.exit(0)

        # v3.0: 按 group_name 分组显示
        total_actions = sum(len(r.actions) for _, r in results)

        for group_name, group_results in groupby(results, key=lambda x: x[0]):
            group_list = list(group_results)
            group_folders_with_files = [r for _, r in group_list if r.dwg_files]
            print(f"\n  === {group_name} === ({len(group_folders_with_files)} folders)")

            for _, result in group_list:
                if not result.dwg_files:
                    continue
                print(f"\n    {result.folder_name}")
                print(f"      doc_id: {result.doc_id}  |  {len(result.dwg_files)} files")
                for kept in result.kept_ifr_all:
                    print(f"      [保留 IFR] {kept.filename}")
                for kept in result.kept_ifc_all:
                    print(f"      [保留 IFC] {kept.filename}")
                if not mgr.ifc_folder_mode:
                    for ab in result.ab_files:
                        print(f"      [保留 AS BUILT] {ab.filename}")
                for action in result.actions:
                    if action.action == 'rename':
                        print(f"      [重命名] {action.source.name} -> {action.dest.name}")
                    else:
                        print(f"      [->{action.dest.parent.name}/] {action.source.name}")

        renames = sum(1 for _, r in results for a in r.actions if a.action == 'rename')
        ss_n  = sum(1 for _, r in results for a in r.actions if a.action == 'move_to_ss')
        ifc_n = sum(1 for _, r in results for a in r.actions if a.action == 'move_to_ifc')
        ab_n  = sum(1 for _, r in results for a in r.actions if a.action == 'move_to_ab')
        print(f"\n  汇总: {renames} 重命名, {ss_n} →SS/, {ifc_n} →IFC/, {ab_n} →Rev.N-AB/")

        if not dry_run and total_actions > 0:
            confirm = input("\n  确认执行? 输入 'YES': ").strip()
            if confirm == 'YES':
                stats = mgr.execute_actions(results)
                print(f"\n  完成: {stats['renamed']} 重命名, {stats['moved']} 移动, "
                      f"{stats['ss_created']} SS/ 创建, {stats['errors']} 错误")
            else:
                print("  已取消")
        elif dry_run and total_actions > 0:
            print(f"\n  使用 --execute 执行操作")
    else:
        # Show help for unrecognized args
        parser.print_help()


if __name__ == "__main__":
    main()
