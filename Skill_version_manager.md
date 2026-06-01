# Skill: Version Manager v4 — IFC/IFR 分离 + 灵活的工作文件版本管理

> **NOTE**: Version Manager functionality has been **merged into `ifr_automation_v10.py`**
> as part of the unified Engineering Pipeline. This file (`version_manager_v5.py`) is
> kept as a standalone fallback. See `Skill_ifr_automation_v6.md` for the merged docs.
>
> **文件**: `version_manager_v5.py` (standalone fallback)
> **合并到**: `ifr_automation_v10.py` (classes: `VersionManager`, `NativeVersionManager`, `FolderRelocator`)
> **当前版本**: v5.0

---

## 零、v4.0 新增：IFC 文件检测（PDF 模式）

### 问题
`IFR(Client)/2.Reports` 目录中 IFC 和 IFR 文件混在一起。IFC 文件因 `_IFC` 后缀导致 base_name 不同，形成独立的 singleton 组，不会被 `identify_old_versions()` 标记为旧版本。

### 新增方法: `identify_ifc_in_ifr_client(target_dir, ss_folder)`

| 属性 | 说明 |
|------|------|
| 所在类 | `VersionManager` |
| 触发条件 | `target_dir` 路径中含 `IFR(Client)` |
| 使用正则 | `_RE_IFC` / `_RE_IFR`（与 Native 模式共用） |
| 判定逻辑 | `is_ifc and not is_ifr` → 移动到 SS |
| 返回值 | `List[Tuple[Path, Path, str]]` — 同 `identify_old_versions()` 格式 |

### 修改的方法

| 方法 | 变化 |
|------|------|
| `analyze_directory()` | 新增 `ifc_files` / `ifc_files_count` 字段，调用 `identify_ifc_in_ifr_client()` |
| `process_directory()` | 在旧版本处理后追加 IFC 文件检测和移动 |
| `process_single_project()` (UI) | 预览中显示 IFC 文件数量和详情 |

---

## 一、问题总结 (v2.0 → v3.0 的改进)

| 问题 | v2.0 现状 | v3.0 目标 |
|------|-----------|-----------|
| 扫描范围 | 只扫 `1. Native/` | 扫 Native + Reports + Schedule |
| 文件夹匹配 | 正则 `_RE_FOLDER_NAME` 要求 `_Rev` 结尾 + GG 前缀 | 不依赖正则发现文件夹，排除列表外全部纳入 |
| Doc ID | 只支持 `GG\d{2}-[CE]-XXX-\d{3}` | 尽力提取，失败则用文件夹名 |
| 文件类型 | 只处理 `.dwg` | 忽略列表外的所有文件 |
| SS 文件夹 | 硬编码 `SS/` | 自动检测 `SS/` / `SUPERSEDED/` / `Superceded/` |

---

## 二、改动清单

### 2.1 新增常量：`SCAN_ROOTS`

替代 `native_root` 属性的单一路径。

```python
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
```

### 2.2 新增常量：排除文件夹名 + 忽略扩展名

```python
# 扫描时跳过的子文件夹名 (不区分大小写)
_SKIP_FOLDER_NAMES = {
    'ss', 'superseded', 'superceded',
    'bom', 'appendix', 'approved',
}
# 以这些前缀开头的文件夹也跳过 (不区分大小写)
_SKIP_FOLDER_PREFIXES = ('sy supply',)

# 不参与版本管理的扩展名
_IGNORE_EXT = {'.dwl', '.dwl2', '.err', '.log', '.tmp', '.lnk', '.db', '.ini'}
```

> **注意**: v2 的 `_NATIVE_IGNORE_EXT` 包含 `.xlsx`、`.txt`、`.csv`，v3 不再忽略它们 — Reports 里这些是正式文件。

### 2.3 删除 `_RE_FOLDER_NAME` 正则

不再使用正则匹配文件夹名来发现文件夹。改为：

```python
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
```

### 2.4 新函数：`find_doc_folders()` (替代 `find_rev_folders()`)

```python
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
```

### 2.5 新函数：`_find_or_create_ss_folder()`

```python
@staticmethod
def _find_or_create_ss_folder(folder: Path, create: bool = False) -> Path:
    """在 folder 下查找已有的 superseded 文件夹，或返回默认 SS/ 路径。

    查找顺序 (不区分大小写): SS → SUPERSEDED → Superceded
    如果都不存在且 create=True，创建 SS/。
    """
    for item in folder.iterdir():
        if item.is_dir() and item.name.lower() in ('ss', 'superseded', 'superceded'):
            return item

    ss_path = folder / 'SS'
    if create:
        _to_long_path(ss_path).mkdir(exist_ok=True)
    return ss_path
```

### 2.6 修改 `scan_dwg_files()` → `scan_files()`

- 重命名为 `scan_files()` (不再限于 dwg)
- 跳过 `_IGNORE_EXT` 中的扩展名
- 跳过 superseded 子文件夹中的文件
- 分类逻辑 (`_classify_dwg`) 应用于所有文件类型（IFR/IFC 按文件名中的版本号判断）

```python
def scan_files(self, folder: Path) -> List[DwgFile]:
    """扫描文件夹中的所有工作文件（排除忽略列表和 superseded 文件夹）。"""
    files = []
    bak_map: Dict[str, Path] = {}

    # 找到 superseded 文件夹以便跳过
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
```

> **可选**: 将 `DwgFile` 重命名为 `VersionedFile`，但为减少改动量可暂不改。

### 2.7 修改 `process_folder()`

主要变更：
- 调用 `_parse_folder_name()` 替代 `_parse_rev_folder_name()`
- 调用 `scan_files()` 替代 `scan_dwg_files()`
- 调用 `_find_or_create_ss_folder()` 替代硬编码 `folder / 'SS'`
- 版本分类逻辑 (IFR/IFC/OTHER, keep newest, move old) **完全不变**

```python
def process_folder(self, folder: Path) -> NativeFolderResult:
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
    other_files = [f for f in files if f.rev_type == 'OTHER']
    ss_folder = self._find_or_create_ss_folder(folder)  # 自动检测

    # ... 下方版本分类逻辑完全不变 ...
```

### 2.8 修改 `process_all()`

```python
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
```

### 2.9 修改 `execute_actions()`

- 使用 `_find_or_create_ss_folder(folder, create=True)` 替代硬编码 `folder / 'SS'`
- 其余逻辑不变

### 2.10 CLI 参数新增 `--scope`

```python
parser.add_argument('--scope', type=str, default='all',
                    choices=['native', 'reports', 'schedule', 'all'],
                    help='扫描范围 (默认 all)')
```

使用示例：
```bash
python version_manager.py --native --scope reports
python version_manager.py --native --scope native --folder PLN-001
```

### 2.11 UI 分组显示

CLI 输出和交互模式 [5] 的预览按 `display_group` 分组：

```
  === 1. Native === (15 folders)
    GG31-C-PLN-001_Civil Site Layout Plan_Rev
      [保留 IFR] GG31-C-PLN-001 Civil Site Plan_rB.dwg
      [->SS/] GG31-Site Plan - 11544 Augusta Hwy revK.dwg

  === Reports === (5 folders, via Civil & Structure)
    GG31-C-RPT-001_MV Power Station Foundation Report
      [保留] GG31-C-RPT-001_...Rev0_IFC.docx
      [->SUPERSEDED/] GG31-C-RPT-001_...RevA.docx
```

**实现方式**: `itertools.groupby()` 按 group_name 分组

```python
from itertools import groupby

for group_name, group_results in groupby(results, key=lambda x: x[0]):
    group_list = list(group_results)
    folders_with_files = [r for _, r in group_list if r.dwg_files]
    print(f"\n  === {group_name} === ({len(folders_with_files)} folders)")
    for _, result in group_list:
        if not result.dwg_files:
            continue
        print(f"    {result.folder_name}")
        print(f"      doc_id: {result.doc_id}  |  {len(result.dwg_files)} files")
        if result.kept_ifr:
            print(f"      [保留 IFR] {result.kept_ifr.filename}")
        if result.kept_ifc:
            print(f"      [保留 IFC] {result.kept_ifc.filename}")
        for action in result.actions:
            if action.action == 'rename':
                print(f"      [重命名] {action.source.name} -> {action.dest.name}")
            else:
                print(f"      [->{action.dest.parent.name}/] {action.source.name}")
```

---

## 二bis、v5.0 修复：浏览器下载重复文件归组

### 问题
浏览器下载同名文件时自动添加 `(1)`, `(2)` 后缀（如 `_RevB (1).pdf`）。`extract_base_name_and_version()` 将 `(1)` 保留在 base_name 中，导致 `_RevB.pdf` 和 `_RevB (1).pdf` 被归为不同组，version manager 不会清理旧副本。

### 修复
在 `extract_base_name_and_version()` 中，提取版本号前先去掉 `\s*\(\d+\)\s*$` 后缀：
```python
name_without_ext = re.sub(r'\s*\(\d+\)\s*$', '', name_without_ext)
```

### 效果
| 文件 | 修复前 base | 修复后 base |
|------|------------|------------|
| `_RevB.pdf` | `...Block Diagram.pdf` | `...Block Diagram.pdf` |
| `_RevB (1).pdf` | `...Block Diagram (1).pdf` | `...Block Diagram.pdf` |

两个文件归入同一组，`identify_old_versions()` 按修改时间保留最新版。

---

## 二ter、v5.0 新增：Client Sharepoint Coverage + SS 自动检测

### TARGET_SUBDIRS 扩展
`VersionManager.TARGET_SUBDIRS` 新增两个 Client Sharepoint 目录:
```python
TARGET_SUBDIRS = [
    r"Design\Engineering\1. Drawings\2. IFR_internal",
    r"Design\Engineering\1. Drawings\3. IFR(Client)\1.Drawing",
    r"Design\Engineering\1. Drawings\3. IFR(Client)\2.Reports",
    r"Design\Engineering\1. Drawings\3. IFR(Client)\3.Deliverables",
    r"Design\Engineering\1. Drawings\4. IFC(Client)",              # NEW v5.1
    r"Design\Engineering\13. Client Sharepoint\1.IFR\1.Report",
    r"Design\Engineering\13. Client Sharepoint\1.IFR\2.Drawing",
]
```

### `_find_ss_folder()` 自动检测
Client Sharepoint 使用 `Superseded` 而非 `SS`。新增静态方法自动检测:
```python
@staticmethod
def _find_ss_folder(target_dir: Path) -> Path:
    for item in target_dir.iterdir():
        if item.is_dir() and item.name.lower() in ('ss', 'superseded', 'superceded'):
            return item
    return target_dir / 'SS'  # default
```

替换了 `analyze_directory()` 和 `process_directory()` 中所有 `target_dir / "SS"` 硬编码。

**应用于**: `ifr_automation_v10.py` 和 `version_manager_v5.py` 两个文件。

### v5.1: IFC(Client) 版本管理
`TARGET_SUBDIRS` 新增 `4. IFC(Client)` 目录，用于清理旧 IFC 修订版（如 `_Rev0_IFC.pdf` 在 `_Rev1_IFC.pdf` 存在时归档到 SS/）。

**安全保护**: `identify_ifc_in_ifr_client()` 第1958行 `if "IFR(Client)" not in str(target_dir): return` — `4. IFC(Client)` 不含 "IFR(Client)" 字符串，不会误触发 IFC→SS 逻辑。

**分组原理**: `extract_base_name_and_version()` 从 `_Rev0_IFC.pdf` 提取 version=`0`，base_name=`..._IFC.pdf`；`_Rev1_IFC.pdf` 提取 version=`1`，same base_name → 分为一组 → `identify_old_versions()` 按 mtime 保留最新。

**Deliverable 联动**: `/ifr` Stage 2 (VM) 清理旧版 → Stage 4 `scan_ifc_folder()` 只读到最新 IFC revision → Excel 自动更新。

**应用于**: `ifr_automation_v10.py` 和 `version_manager_v5.py` 两个文件。

---

## 三、不改动的部分

| 组件 | 说明 |
|------|------|
| `VersionManager` 类 | v4 新增 `identify_ifc_in_ifr_client()`，v5 修复 download-duplicate 归组 + Client Sharepoint Coverage；其余 PDF 版本管理逻辑不变 |
| `InteractiveUI` 菜单结构 | 保持 [0]-[5] 选项，只改 [5] 内部处理逻辑 |
| `config.json` 格式 | 不变 |
| `_classify_dwg()` | IFR/IFC/OTHER 分类逻辑不变 |
| `_make_standard_dwg_name()` | 标准命名逻辑不变 |
| `_to_long_path()` | Windows 长路径支持不变 |
| `_RE_IFC` / `_RE_IFR` | 版本号正则不变 |

---

## 四、数据流

```
用户选择项目
     │
     ▼
find_doc_folders(scope, filter)
     │
     ├── SCAN_ROOTS[0]: Native/ (depth=1)
     │     └── 直接列出子文件夹 → 排除 _SKIP 名单 → 关键字过滤
     ├── SCAN_ROOTS[1]: Reports/ (depth=2)
     │     └── 先列分类目录 → 再列子文件夹 → 排除 → 过滤
     └── SCAN_ROOTS[2]: Schedule/ (depth=1)
           └── 同 depth=1 逻辑
     │
     ▼
对每个文件夹: process_folder()
     │
     ├── _parse_folder_name() → doc_id, description (尽力提取)
     ├── scan_files() → 所有非忽略文件，分类 IFR/IFC/OTHER
     ├── _find_or_create_ss_folder() → 检测已有 SS/SUPERSEDED/Superceded
     └── 版本逻辑: 保留最新 IFR + 最新 IFC，旧版 → SS
     │
     ▼
按 group_name 分组显示预览
     │
     ▼
用户确认 → execute_actions()
```

---

## 五、验证方式

```bash
# 1. 扫描全部范围 (dry-run)
python version_manager.py --native
# 预期: 看到 === 1. Native ===, === Reports ===, === Schedule === 分组

# 2. 关键字过滤
python version_manager.py --native --folder PLN-001
# 预期: 只显示文件夹名包含 PLN-001 的结果

# 3. 范围过滤
python version_manager.py --native --scope reports
# 预期: 只扫描 Reports 下的文件夹

# 4. 交互模式
python version_manager.py
# 选 [5] → 预期看到分组显示

# 5. 执行
python version_manager.py --native --execute
# 预期: 确认后执行移动，SUPERSEDED/Superceded 文件夹被正确复用
```

---

## 六、迁移注意事项

1. **`native_root` 属性**: 可保留为向后兼容（指向 Native 路径），但 `find_doc_folders()` 不依赖它
2. **`_RE_FOLDER_NAME` 正则**: 删除，改用 `_parse_folder_name()`
3. **`_NATIVE_IGNORE_EXT`**: 替换为 `_IGNORE_EXT`（去掉 `.xlsx`/`.txt`/`.csv`）
4. **`process_all()` 返回类型变化**: `List[NativeFolderResult]` → `List[Tuple[str, NativeFolderResult]]`，所有调用处需适配
5. **交互模式 [5]** 的显示逻辑需适配新的分组返回值
