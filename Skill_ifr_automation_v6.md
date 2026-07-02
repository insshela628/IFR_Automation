# Skill: Engineering Pipeline v7→v10 — IFR Sync + Version Management + Sharepoint Sync + Deliverable + IFC Conversion

> **文件**: `ifr_automation_v10.py` (升级自 v7, 新增 IFC conversion)
> **版本**: v10.0
> **配置**: `config.json` (新增 `deliverable` section)

---

## 零、v10 IFC Stamp 规格 (2026-03-19)

### 双印章布局
`_stamp_via_com_draw()` 在每个 title block 右下方绘制两个印章框：
1. **下方**: FOR CONSTRUCTION (rect 110.5 x 17.7, text h=7.0, Arial Narrow Bold)
2. **上方**: DRAWINGS TO BE PRINTED IN COLOUR (rect 110.5 x 26.0, text h=5.5, 两行)
3. **间距**: 2.0 (按比例缩放)

### 关键参数
| 参数 | 值 | 说明 |
|------|-----|------|
| `_REF_TB_WIDTH` | 841.0 | 参考图框宽度 (A1) |
| `_REF_TB_HEIGHT` | 594.0 | 参考图框高度 |
| `_REF_RECT_W` | 110.511 | FOR CONSTRUCTION 框宽 |
| `_REF_RECT_H` | 17.745 | FOR CONSTRUCTION 框高 |
| `_REF_COLOUR_RECT_H` | 26.0 | COLOUR 框高 (两行文字) |
| `_REF_COLOUR_GAP` | 2.0 | 两框间距 |
| `_REF_COLOUR_TEXT_H` | 5.5 | COLOUR 文字高度 (小于 FOR CONSTRUCTION) |
| `_REF_X_RIGHT_OFFSET` | 29.779 | 框右边距离图框右边 |
| `_REF_Y_BOTTOM` | 73.259 | 框底部距离图框底部 |

### COM 安全规则
- 所有 `.TextString` 读写: `_safe_get_text()` / `_safe_set_text()` (含 `_com_retry`)
- 所有 `GetAttributes()` / `TagString`: 在 `_get_attrs_dict()` 中用 `_com_retry`
- 所有 `AddLightWeightPolyline` / `AddMText`: 用 `_com_retry`
- 文档打开后: 验证 `ModelSpace.Count` + `Layouts.Count` + 2s settle
- Title block 搜索: 最多重试 3 次

### XREF 处理
- 输出路径 > 240 chars → SaveAs 到 temp 目录
- SaveAs 前绑定所有 XREF: `blk.Bind(False)` (Insert bind, 保留原始图层名)
- 防止 PUBLISH 从 temp 目录打开时 XREF 路径断裂

---

## 一、v7 架构概览

### 合并后的单一脚本结构

```
ifr_automation_v7.py
  ├── Shared Utilities (to_long_path, UIHelper, ConfigManager, IFC/IFR regex)
  ├── IFR Sync Module (v6原有，不动)
  │     ProjectValidator, ProjectScanner, DrawingCollector, IFRAutomation
  ├── Version Management Module (从 version_manager_v4.py 合并)
  │     VersionManager (PDF), NativeVersionManager, FolderRelocator
  ├── Deliverable Module (NEW)
  │     DeliverableManager, DeliverableLayout, FormatChangeWarning
  ├── Pipeline Orchestrator (NEW)
  │     PipelineOrchestrator — sequences 4 stages per project
  └── Unified Interactive Menu + CLI
        UnifiedInteractiveMode — 10 options
```

### Pipeline 流程 (per project)

```
1. IFR Sync → mirror PDFs to IFR(Client) from IFR_internal + source dirs
2. Version Management → clean old versions to SS (incl. Client Sharepoint + IFC(Client))
3. Sharepoint Sync → archive approved + sync IFR_internal → Client Sharepoint
   3a. archive_approved_in_ifr_client(): 归档 -Approved 文件 + Sharepoint反向反馈
   3b. sync_to_sharepoint(): Drawing源=IFR_internal, Reports源=IFR(Client)/2.Reports
4. Deliverable Cross-Check & Update → compare files vs Excel, update IFC status
```

### 三条同步链

```
        ┌──────────────┐
        │ IFR_internal │  (已版本管理，唯一)
        └──────┬───────┘
               │
    Stage 1    │    Stage 3
   ┌───────────┼───────────────┐
   ▼                           ▼
┌──────────────────┐   ┌──────────────────────┐
│ IFR(Client)/     │   │ Client Sharepoint/   │
│ 1.Drawing        │   │ 1.IFR/2.Drawing      │
│ 2.Reports        │   │ 1.IFR/1.Report       │
└────────┬─────────┘   └──────────┬───────────┘
         │     Stage 3 反向反馈    │
         │◄───────────────────────┘
         │  Sharepoint已审批 → 归档到 Approved to IFC/
         ▼
   Approved to IFC/
```

### Bot 命令作用域

| 命令 | 调用 | 包含的阶段 |
|------|------|-----------|
| `/ifr` | `IFRAutomation.process_project()` | Stage 1 only (IFR Sync) |
| `/pipeline` | `PipelineOrchestrator.run_pipeline()` | Stage 1-4 全部 (含归档+Sharepoint同步) |
| `/deliverable` | `DeliverableManager.cross_check()` | Stage 4 only |
| `/ifc` | `IFCTransmittalManager.run()` | IFC 转换 + transmittal + deliverable 更新 |
| `/panel_ifc` | `PanelIFCManager.run()` / `ApprovedIFCManager.run()` | 多页 panel IFC / Approved IFC 流程 |
| `/issue_register` | `IssueRegisterManager.run()` | 独立：三维度 responder 分配（不属于 pipeline 4 阶段） |

---

## 二、Interactive Menu

```
[1] 完整流程 (IFR Sync + Version Mgmt + Sharepoint Sync + Deliverable)  ← 新增
[2] 仅 IFR 同步
[3] 仅版本管理 (PDF)
[4] 仅交付物检查 & 更新                                ← 新增
[5] Native/Reports/Schedule 版本管理
[6] 文件夹归位
[7] 复制 IFC(Client) 到目标
[8] 仅验证项目结构
[9] 配置 / 日志
[0] 退出
```

---

## 三、Deliverable Cross-Check (新增功能)

### 3.1 Dynamic Excel Layout Detection

自动扫描 rows 1-20，寻找连续的 `Revision | Submission Date | Status` 列头。
- **Warnertown**: K=Revision, L=Submission Date, M=Status (row 6)
- **LMS**: L=Revision, M=Submission Date, N=Status (row 9)
- 检测失败 → raise `FormatChangeWarning`

### 3.2 Cross-Check Dimension 1: Item Presence

| 情况 | 处理 |
|------|------|
| 文件夹有，Excel没有 | 自动插入新行 (doc_id + description + revision)，黄色高亮 |
| Excel有，文件夹没有 | 仅报告 (可能是 Reserved 项目) |

### 3.3 Cross-Check Dimension 2: Revision Updates

| 检查 | 处理 |
|------|------|
| 文件夹 revision > Excel revision | 更新 revision 列 + status 列 → 'Submitted'，黄色高亮 |
| Submission Date | **不自动更新** (用户根据高亮手动填写)。日期格式 `dd/mm/yy` |
| 前次高亮 | 每次更新前清除上次的黄色高亮 |

### 3.3a IFC Status Cross-Check

`scan_ifc_folder()` 扫描5个来源确定 IFC 状态：

| 来源 | Revision 类型 |
|------|-------------|
| `4. IFC(Client)/` | 数字 (Rev0, Rev1) |
| `13. Client Sharepoint/1.IFR/1.Report/Approved to IFC/` | 字母 (RevA, RevB) |
| `13. Client Sharepoint/1.IFR/2.Drawing/Approved to IFC/` | 字母 |
| `3. IFR(Client)/1.Drawing/Approved to IFC/` | 字母 |
| `3. IFR(Client)/2.Reports/Approved to IFC/` | 字母 |

**K列 = numerical IFC revision**: 一旦 doc-ID 出现在 `4. IFC(Client)/`，K列必须写入数字版本号 (0, 1, 2, 3...)，替换旧的字母版本号。Rev0 有效，必须写入。Guard 条件仅为 `if ifc_rev is not None`。

**`/ifc` 命令必须始终更新 deliverable**: `IFCTransmittalManager.run()` 基于 ALL 当前 IFC 文件（`scan_ifc_files()` grouped dict）更新 deliverable Excel，不受 `new_files` 是否为空的影响。`new_files` 检查只控制 transmittal 生成。

**Deliverable 同步**: `/ifc` 更新 deliverable Excel 后，必须同步到所有位置：
- `8. Deliverables/`（主文件位置）
- `3. IFR(Client)/3.Deliverables`（PRIMARY_SYNC）
- `13. Client Sharepoint/1.IFR/3.Deliverables`（SECONDARY_SYNC）

### 3.4 File Version Management

1. Increment file revision: `1.9 → 2.0` (after .9 bump major), `14 → 15`
2. Update "Last Updated" date cell
3. Update filename to match new revision
4. Copy old version to `SUPERSEDED/`, then delete (Dropbox-safe)
5. Sync to primary target + secondary target (if exists)

---

## 四、CLI 新增参数

| 参数 | 说明 |
|------|------|
| `--pipeline` | 运行完整管线 (IFR sync + version mgmt + deliverable) |
| `--stages ifr_sync version_mgmt deliverable` | 指定管线阶段 |
| `--deliverable-only` | 仅运行交付物检查 + 更新 |
| `--deliverable-check-only` | 仅运行交付物检查 (报告，不更新) |
| `--version-mgmt` | 仅运行 PDF 版本管理 |
| `--native` | Native/Reports/Schedule 版本管理 |
| `--scope native\|reports\|schedule\|all` | --native 的范围 |
| `--folder KEYWORD` | --native 的文件夹过滤 |
| `--execute` | 执行操作 (--native/--version-mgmt 默认 dry-run) |

---

## 五、Doc ID 提取

```
GG31-C-PLN-001_Civil Site Layout_RevB.pdf → GG31-C-PLN-001
50023-RC-300_Foundation Report_Rev D.pdf  → 50023-RC-300
COMMS &AUX Cable Route_GG31-E-PLN-003_RevA.pdf → GG31-E-PLN-003  (非标准命名)
```

支持 `XX-X-XXX-NNN` (GG) 和 `NNNNN-XX-NNN` (LMS) 格式。

**两步提取策略**:
1. `match()` — doc-ID 在文件名开头 (标准命名)
2. `search()` — doc-ID 在文件名任意位置 (非标准/不准确命名)

**核心规则**: FILE NO 一致即匹配，不依赖 FILE NAME 是否准确。

---

## 六、Revision 比较逻辑

- Letter: A < B < C (ord comparison)
- Number: 0 < 1 < 2 (integer comparison)
- Skip: revision 为空 或 status 为 "N/A" / "Reserved"

---

## 七、错误处理

| 场景 | 处理 |
|------|------|
| Dropbox 文件锁 | 指数退避重试 (3次) |
| Excel 临时锁文件 (`~$...xlsx`) | 跳过 |
| 合并单元格 | layout detection 跳过合并区域 |
| openpyxl 未安装 | 优雅降级，提示安装 |

---

## 八、Config 新增

```json
{
  "deliverable": {
    "source_folders": [...],
    "deliverable_path": "Design/Engineering/1. Drawings/3. IFR(Client)/3.Deliverables",
    "primary_sync": "...",
    "secondary_sync": "Design/Engineering/13. Client Sharepoint/1.IFR/3.Deliverables",
    "highlight_color": "FFFF99",
    "auto_detect_layout": true
  }
}
```

---

## 九、Sharepoint Sync 详细规格

### 归档触发 (`archive_approved_in_ifr_client`)

| 触发条件 | 说明 |
|----------|------|
| `-Approved` 后缀 | 文件名含 `-Approved` (本地触发) |
| Sharepoint 反向反馈 | doc-ID 已在 Client Sharepoint 的 `Approved to IFC/` 或 `-Approved` 文件中 |

归档后 `IFR(Client)/1.Drawing` 和 `2.Reports` 只保留 pending 文件，员工无需查看 Sharepoint 即可知审批状态。

### 同步源映射

| 源 | 目标 | 原因 |
|----|------|------|
| `2. IFR_internal` | `13. Client Sharepoint/1.IFR/2.Drawing` | 已版本管理、doc-ID 唯一 |
| `IFR(Client)/2.Reports` | `13. Client Sharepoint/1.IFR/1.Report` | Reports 无 IFR_internal 对应 |

Fallback: `IFR_internal` 不存在时回退到 `IFR(Client)/1.Drawing`

### Version Manager 覆盖范围

`VersionManager.TARGET_SUBDIRS` 包括:
- `IFR_internal`, `IFR(Client)/1.Drawing`, `2.Reports`, `3.Deliverables`
- `4. IFC(Client)` — 旧 IFC 修订版 (如 Rev0 被 Rev1 取代时) 归档到 SS/
- `Client Sharepoint/1.IFR/1.Report`, `2.Drawing`

`_find_ss_folder()` 自动检测 `SS`/`Superseded`/`Superceded` (Client Sharepoint 使用 `Superseded`)

**IFC(Client) 安全保护**: `identify_ifc_in_ifr_client()` 有 `"IFR(Client)" not in str(target_dir)` 检查，不会把 `4. IFC(Client)/` 中的 IFC 文件误判为"不该在此目录"

---

## 十、复盘教训 (2026-03-19)

以下是开发过程中发现的问题和修复，必须在未来开发中避免重犯。

### 10.1 日期格式必须匹配 Excel 实际格式
- **问题**: 代码写入 `YYYY-MM-DD` 但 Excel 中使用 `dd/mm/yy`
- **修复**: `strftime('%d/%m/%y')`
- **规则**: 写入 Excel 前，先检查目标 Excel 的现有日期格式，保持一致

### 10.2 IFC 状态必须覆盖所有来源目录
- **问题**: `scan_ifc_folder()` 只扫描 `4. IFC(Client)/`，Reports 的 IFC 状态永远不会被反映
- **修复**: 扩展到 5 个目录（含 Client Sharepoint Approved to IFC 和 IFR(Client) Approved to IFC）
- **规则**: 任何与"状态判定"相关的扫描，必须列举所有可能的文件来源，不能只查一个目录

### 10.3 IFC numerical revision 必须写入 K 列（含 Rev0）
- **问题**: 旧逻辑 `rev != 0` 导致 Rev0 的 IFC 文件不写入 K 列，deliverable Excel 状态不更新
- **修复**: Guard 改为 `if ifc_rev is not None`（移除 `!= 0`）
- **规则**: `4. IFC(Client)/` 下的所有文件，K 列必须写入 numerical revision (0, 1, 2, 3...)，替换旧字母版本号

### 10.4 SS 文件夹名不能硬编码
- **问题**: 代码硬编码 `target_dir / "SS"` 但 Client Sharepoint 使用 `Superseded`
- **修复**: `_find_ss_folder()` 自动检测 SS/Superseded/Superceded
- **规则**: 文件夹名有变体时，必须先扫描再决定，不能假设命名

### 10.5 同步源应选择已版本管理的目录
- **问题**: 初始设计从 `IFR(Client)/1.Drawing` 同步到 Sharepoint，但该目录含重复文件和旧版本
- **修复**: 改用 `IFR_internal` 作为 Drawing 源
- **规则**: 同步链的源应选择"最干净"的目录（已去重、已版本管理）

### 10.6 内外双向同步必须考虑反向反馈
- **问题**: 内→外同步做了（IFR_internal→Sharepoint），但外→内的审批状态没回传，员工看不到审批结果
- **修复**: `archive_approved_in_ifr_client()` 检查 Sharepoint 审批状态，反向归档 IFR(Client) 中的文件
- **规则**: 任何双向镜像关系，必须同时实现 outbound sync 和 inbound feedback

---

## 十一、Issue Register Responder Assignment (独立模块，非 Pipeline 阶段)

### 定位
`IssueRegisterManager` 是**独立执行器**，与 `IFCTransmittalManager` 类似——不在 `PipelineOrchestrator` 的 4 阶段流程内，由 bot 命令 `/issue_register` 单独触发。

### 输入输出
- **Input**: `3. IFR(Client)/{prefix}-Design Review Comments Register_Rev{X}.xlsx`（Sheet: `Master Register`）
- **Output**: `{original}_updated.xlsx`（浅蓝=自动填充，黄色=冲突标记）
- **Target Column**: M（Responder）

### 三维度审查法（跨项目通用）

| 维度 | 来源 | 精确度 | 角色 |
|------|------|--------|------|
| **1. Role** | 工程师岗位描述 | 最低 | 兜底默认 |
| **2. Email Allocation** | 协调邮件 / `Team-Allocation.md` 内 json fence | 中 | 项目级 canonical |
| **3. Title Block** | 最新 Rev IFC PDF → DRN/DES/CHK/APP | **最高** | Ground truth |

**冲突规则**: 3 > 2 > 1。维度 3 缺失（无 IFC PDF）→ 回退 2 → 1。通用签名（`ACE`/`AW`/空）过滤，降级看 DES/CHK。

### Project-Specific Override

⚠️ **Dropbox 永远不放任何团队/responder 配置**（全公司可见）。配置只在 D: 盘私人 vault：

```
D:\3.Career\obsidian-vault\04-Work-SOP\Projects\{code}-*\Team-Allocation.md
```

md 文件里嵌入一个 ```json``` fence block——这就是**运行时配置**。
`IssueRegisterManager._extract_allocation_json_from_md()` 扫描所有 ```json``` fence，
返回第一个同时包含 `allocation` 和 `special_rules` 字段的块。
人类阅读的说明和机器读的配置在同一个 md 文件内（单一源，无 .json + .md 同步漂移）。

项目 code 从 register 文件名自动识别（`50023-Design Review...` → `50023`；`GG31-...` → `GG31`）。D: 盘找不到对应项目目录（或 md 里没有合规 json fence）→ 退化为只用维度 3（title block）；维度 3 也拿不到 → M 列留空等人工。

Schema（md 里 ```json fence 内容）：

```json
{
  "allocation": {
    "<prefix>": {"responder": "XX", "note": "..."}
  },
  "special_rules": {
    "<doc-id-substr>": {"responder": "XX", "note": "..."}
  },
  "notes": {
    "<responder>": "<annotation, e.g. HOLD>"
  }
}
```

- Prefix 提取：`[A-Z]{2,3}-\d{3}` pattern（如 `50023-EA-301` → `EA`）
- Special rules 用 substring 匹配 doc-id（如 `EA-300` 命中 `50023-EA-300`）

### Title-Block 提取

`_extract_titleblock_fields()` 工作流程：
1. 扫描 `4. IFC(Client)/*.pdf`，提取每个文件的 DRN/DES/CHK/APP
2. 先看最后一页（title block 通常在右下角），再看第一页
3. `pdfplumber.extract_tables()` → 找含 `DRN`/`DES`/`CHK`/`APP` 列头的表
4. 从最新 REV 行开始逆序遍历，返回第一个有数据的行
5. 优先级：DES > CHK > DRN（跳过 `ACE`/`AW`/空）
6. **Graceful degradation**: `pdfplumber` 未安装时，维度 3 静默跳过，回退到维度 2

### Excel 逻辑

- A=Status, B=DocNum, I=Comment, J=Severity, M=**Responder**, R=Closeout, W=Closeout(post-workshop)
- Skip：A != "Open"，或 W 已填 + M 已填
- 空 M → 自动填充 + 浅蓝高亮
- M 已填但与三维度结果冲突 → 保留原值 + 黄色高亮（提醒 coordinator 复查）

### 团队分工存储位置

⚠️ **C 盘 memory 只保留去敏方法论**（所有用户可见），具体成员名字/项目分工在 D 盘。
⚠️ **Dropbox 永远不放任何团队/responder 配置**（全公司可见）。

- 通用 SOP: `D:\3.Career\obsidian-vault\04-Work-SOP\PM-Methodology\Issue-Register-Responder-Workflow.md`
- 项目专用（单一源——人类阅读 + 运行时配置）: `D:\3.Career\obsidian-vault\04-Work-SOP\Projects\{code}-{name}\Team-Allocation.md` — 人类说明 + 嵌入 ```json``` fence block；运行时由 `IssueRegisterManager._find_d_drive_allocation()` 自动定位，`_extract_allocation_json_from_md()` 提取 json 块

### 为什么独立于 Pipeline

1. **触发频率不同**：Pipeline 每次 IFR/IFC 更新都跑；Issue Register 仅在客户返回新 Rev Comment Register 时跑（低频）
2. **输入不同**：Pipeline 处理 drawings/reports；Issue Register 处理 client-returned Excel register
3. **外部依赖**：需要 `pdfplumber`（IFR/IFC pipeline 不需要）
4. **数据流向相反**：Pipeline 是内→外（发给客户）；Issue Register 是外→内（客户反馈回来）


## 十二、Robust AS BUILT Batch Run (AutoCAD 自检 + QA 闭环)

> **何时用**：批量跑 AS BUILT 转换（Coleambally2 / Warnertown 等多 DWG）。
> AutoCAD 只有 COM 接口、可能无限卡死 —— **绝不要用裸 loop 直接驱动**。用下面的防卡死
> 运行器，它把每个文件隔离、卡死时自动重启 AutoCAD。相关事实见 memory
> `feedback-autocad-com-health`、`feedback-stamp-alignment`。

### Step 1 — 开跑前：AutoCAD 在跑吗 / 卡死了吗？
- `tasklist | findstr acad` —— 列出存活的 AutoCAD。
- **卡死 vs 工作中**：看 acad.exe 的 CPU（隔几秒采样两次 `(Get-Process acad).CPU`）：
  - CPU **平在 ~0% 不动** + 没有新文件产出 → **卡死**（模态对话框 / PUBLISH 卡 / RPC 拒绝）。
  - CPU **在涨** → 正在工作，别动它。
- **不要信 Python 控制台**：子进程输出是缓冲的，"看着不动"≠卡死。唯一可靠的进度信号是
  **输出目录里有新的 AB DWG/PDF 出现**。
- 开跑前若已有僵死 acad.exe，先按 Step 4 清掉，让批次从干净状态开始。

### Step 2 — 跑防卡死批量运行器
```
python run_ab_batch_safe.py <project> [docid-or-srcdir-filter] [--timeout N]
```
- `<project>`：`cole2` | `warnertown` | `lms`（或完整项目路径）。
- 可选 filter：doc-id 或 source 文件夹名的子串（如 `PLN-005`）。
- `--timeout N`：每文件硬超时秒数（默认 240；大的多页 DWG 提到 ~300）。
- 例：
  - `python run_ab_batch_safe.py cole2`          —— 全部
  - `python run_ab_batch_safe.py cole2 PLN-005`  —— 单个 doc-id
  - `python run_ab_batch_safe.py cole2 --timeout 300`

**运行器的保证**（所以不用守着）：
- 每个 source 一个**子进程** + 硬超时 → 单个卡死文件拖不垮整批。
- **每个文件前先 kill AutoCAD** → 每次转换都从同样的干净状态开始
  （把多个文件链在一个 warm 实例上会让 COM 变脏、大多数转换失败 —— 已踩坑验证）。
- UTF-8 子进程 IO（中文日志不再 cp1252 崩溃）。
- 超时 → 杀 AutoCAD、记 `timeout`、继续下一个。
- 末尾打印 `=== DONE: ok=.. (warn=..) fail=.. ===` 并列出每个 FAIL/WARN 的 doc-id。

### Step 3 — QA 是自动的（闭环），只需读结果
- 每文件：`_convert_with_qa_retry` → `_qa_validate_ab_pdf`（PyMuPDF/fitz）→ 自动重试最多 3 次，
  仍不过则升级。另外 `_run_post_batch_qa` 扫描**所有**输出 PDF。
- 运行器汇总里：`PASS`（通过）、`WARN`（成功但 QA 标记了问题——看该行）、`FAIL`。
- QA 判据定义在 CLAUDE.md → "AS BUILT Post-Conversion QA"，不要重新实现。
- **印章压图重叠触发必须含线条(linework)**，不止文字/表格格——接线图/单线图几乎全是裸线、无表格格，只查文字/格会漏判印章压线（LMS EL-002）。
- **LMS 专用路径 `asbuilt_revclean_lms.py` 现已补上印章压图 QA 门**：`publish_pdf` 后调 `qa_stamp_overlap()`，复用主引擎 detector（`AsBuiltManager.__new__` 绕开 COM `__init__`），只读 FLAG，命中即打印 `[需人工检查] … 印章压图面内容` 并把该次 run 标为不干净，绝不改动 PDF。
- **只做 QA / 验证的 agent 必须只读**：用 fitz 打开 PDF 即可，**绝不驱动 AutoCAD**
  （不转换、不 SaveAs、不碰 COM）—— 否则可能引发第二次卡死 + 孤儿 acad.exe。

### Step 4 — 手动卡死恢复（仅当你在运行器之外手动开过 AutoCAD、或留下孤儿时）
```
taskkill /F /IM acad.exe          # 强杀全部 AutoCAD
tasklist | findstr acad           # 确认一个不剩
```
然后重启 / 重跑批次。
- **绝不在转换事务进行中途杀掉并留下孤儿 acad.exe** —— 下次运行会报
  "Invalid execution context" / "RPC server unavailable"。重跑前务必确认 `tasklist` 干净。
  （运行器已经做了 kill-before-each，所以这步只针对手动运行。）

相关文件：`run_ab_batch_safe.py`、`ifr_automation_v10.py`、`CLAUDE.md`、memory
`feedback-autocad-com-health` / `feedback-stamp-alignment` / `project-asbuilt-qa-loop`。
