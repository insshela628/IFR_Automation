# IFR/IFC Automation — V6 Project Directory

## Project Overview
AutoCAD IFR sync and IFC conversion automation for solar farm engineering projects.

## Current Scripts
| File | Purpose |
|------|---------|
| `ifr_automation_v10.py` | Main pipeline (~9500+ lines). IFR sync, Sharepoint sync, IFC conversion, deliverable Excel, panel IFC batch, approved IFC |
| `test_stamp_v2.py` | Reference implementation for v2 stamp style (rect + MText) |
| `test_stamp_live.py` | Live AutoCAD stamp test — no bot restart needed |
| `test_ifc_dryrun.py` | Dry-run test for IFCManager |
| `test_panel_ifc_dryrun.py` | Dry-run test for PanelIFCManager |
| `test_approved_ifc_dryrun.py` | Dry-run test for ApprovedIFCManager |
| `inspect_standard_frame.py` | Inspect Tatua_Standard_Frame.dwg block structure |
| `inspect_ifr_block.py` | Inspect IFR block definition details |
| `inspect_attrs.py` | List all title block attributes from a panel DWG (discovered ACE_TitleBlock_CHINT) |
| `test_ifc_v9.py` | **Reusable** automated IFC test: frame measure → TB update → stamp → SaveAs → PDF. Multi-DWG per project, multi-TB per DWG. Project-configurable (`python test_ifc_v9.py warnertown`, `tatua`, etc). Logs to `test_logs/` |
| `inspect_bot_ifc.py` | Inspect bot IFC output DWGs — IFC_STAMP entities, COLOUR MTEXT, IFR blocks, viewport freeze status |
| `test_ab_dryrun.py` | Dry-run test for AsBuiltManager — scans all 3 projects (LMS, Warnertown, Tatua) |
| `test_ab_stamp_warnertown.py` | Live stamp test on Warnertown standard frame DWG — visual verification |

## Key Classes in ifr_automation_v10.py
| Class | Purpose |
|-------|---------|
| `IFRAutomation` | Core IFR file sync (drawings, reports, schedules) |
| `DeliverableManager` | Deliverable Excel update, cross-check, highlight |
| `VersionManager` | SS cleanup, file versioning |
| `IFCStampMixin` | Shared IFC stamp logic (COM draw: MText + bold rect, proportional positioning) |
| `IFCManager` | Single-page IFC conversion (inherits IFCStampMixin) |
| `PanelIFCManager` | Multi-page panel IFC conversion + PUBLISH PDF (inherits IFCStampMixin) |
| `ApprovedIFCManager` | Detect approved IFR → archive → locate Native DWG → IFC convert (inherits IFCStampMixin) |
| `IssueRegisterManager` | Three-angle responder assignment for Design Review Comments / RFI Register (pdfplumber title-block extraction, writes `_updated.xlsx`) |
| `AsBuiltManager` | Multi-project AS BUILT conversion (inherits IFCManager). Auto-detects native/ab_output paths + SaveAs mode. Two modes: subfolder (`Rev.N - AB/` + XREF bind, for Warnertown) and same-dir (save alongside IFC source, no XREF bind, for Coleambally2). Overrides: `_remove_ifc_stamp` (QA + geometry-based cleanup), `_stamp_via_com_draw` (unified proportional thick-border stamp for ALL TBs), `_update_title_block` (AS BUILT rev row). Supports LMS, Warnertown, Coleambally (Tatua), Coleambally2 |
| `PipelineOrchestrator` | 6-stage pipeline: Health Check → IFR Sync → Version Mgmt → IFC Transmittal → Sharepoint Sync → Deliverable. IFC Transmittal (Stage 3) calls `IFCTransmittalManager.run_for_pipeline()` — scan+dedup only, passes `ifc_map` to Stage 5 for unified Excel write |

## Top-Level Constraint
**All rules are cross-project.** Switching between ANY projects (current or future) must NOT require re-explaining any previously learned behavior. Rules are written as universal constraints — never "for project X" or "for project Y". Project-specific details (title block names, tag lists, folder paths) belong in the Projects reference section below, separate from the rules.

## Critical Rules (universal — apply to ALL projects)

### PDF Export (MANDATORY pattern — no exceptions)
- NEVER use `-EXPORTPDF` (silently fails)
- NEVER use per-page Plot + merge (produces incomplete PDFs)
- ALWAYS use PUBLISH + DSD batch approach
- **DSD must include ALL non-Model layouts WITH content** — skip empty layouts (≤1 entity = only default viewport → blank page in PDF). Iterate `doc.Layouts`, collect every layout where `Name != 'model'` AND `layout.Block.Count > 1`:
```python
layout_names = []
for layout in doc.Layouts:
    if layout.Name.lower() != 'model':
        if layout.Block.Count <= 1:
            continue  # empty layout → blank page
        layout_names.append(layout.Name)
for lname in layout_names:
    dsd_lines.extend([f'[DWF6Sheet:{stem}-{lname}]', f'DWG={dwg}', f'Layout={lname}', ...])
```
- This applies to `_publish_single_pdf` (IFCManager) AND `_publish_group_pdf` (PanelIFCManager)
- **XREF binding before temp SaveAs**: When path > 240 chars, DWG is saved to temp dir first. All XREFs must be bound (`blk.Bind(False)`) BEFORE SaveAs, otherwise relative XREF paths break → missing lines/content in PDF. Insert bind (`False`) preserves original layer names.
- **XREF binding for AS BUILT (subfolder mode only)**: `convert_to_ab` and `convert_multi_to_ab` bind XREFs before SaveAs ONLY when `_save_in_source_dir=False` (subfolder mode — DWG saved to `Rev.N - AB/`, different directory than source). When `_save_in_source_dir=True` (same-dir mode — Coleambally2), DWG stays in source folder so XREF paths remain valid, no binding needed.
- **AS BUILT SaveAs modes**: Auto-detected by `_AB_OUTPUT_CANDIDATES` match. `5. As Built/3. As Built Client` → subfolder mode (Warnertown). `5. As Built/` (flat) → same-dir mode (Coleambally2). `6.AS Built` → subfolder mode (LMS/Tatua).
- **Failure mode if XREF not bound**: E-PLN-003 (252 chars) exported with missing cable route lines because XREFs couldn't resolve from temp dir. Warnertown AS BUILT: 10 DWGs reported success but 0 PDFs created because XREFs in subfolder couldn't resolve.
- **Unicode path workaround**: PUBLISH SendCommand silently fails on non-ASCII paths (e.g. `√` in directory names). `_publish_single_pdf` detects non-ASCII via `str(path).isascii()`, copies DWG to temp dir, builds DSD with temp paths, moves PDF back after export. Applied to `_publish_single_pdf` only (other methods already operate from controlled output dirs).
- **PDF wait**: CMDACTIVE=0 doesn't guarantee plot engine has flushed to disk. All 3 PUBLISH methods (`_publish_single_pdf`, `_publish_group_pdf`, `_publish_ab_group_pdf`) poll with file-size-stable check (2s intervals, up to 60s) instead of blind `sleep(3)`.

### IFR Stamp Removal
- `_remove_ifc_stamp()`: pre-scan `doc.Blocks` definitions to find IFR/IFC stamp block names (blocks with ≤10 entities containing "FOR REVIEW" or "FOR CONSTRUCTION" MText, or name containing 'IFR'/'IFC'), then delete matching INSERT references
- **Pass 0**: ModelSpace INSERT scan + **Pass 0b**: PaperSpace block reference scan (Pass 0 SelectionSet only searches active space — PaperSpace block references missed without 0b)
- Use exact text matching via `_strip_mtext_formatting()` (not substring) — substring match would delete "Drawing to be Printed in colour"
- All SelectionSet names must use timestamps to avoid stale name collisions
- **Failure mode if only 'IFR' checked**: IFC-named stamp blocks (e.g. `IFC_STAMP_*`) not caught → FOR CONSTRUCTION leftover in AS BUILT output (BLD-001)
- **AsBuiltManager `_remove_stamps_by_geometry`**: catch-all safety net that removes stamp entities near title blocks by spatial detection. Does NOT skip any layer (IFC_STAMP/QA included) — catches stamps that passes 1-3 missed (e.g. manually-drawn stamps from pre-bot IFC conversions). Uses substring matching for MText (not exact) since entities are already spatially filtered to the stamp zone.
- **Whitespace-insensitive phrase match (`_plain_matches_stamp_phrase`)**: some pre-bot stamps render the status as a single run with NO space (C-PLN-005 "CROSS SECTION" had a QA-layer MText `FORCONSTRUCTION`). Plain `'FOR CONSTRUCTION' in text` substring-matched False → survived BOTH `_remove_qa_layer_stamps`/`_check_stamp_entity` AND QA #3, shipping as a page-5 leftover. Fix matches `_QA_STAMP_PHRASES` with ALL whitespace stripped from both sides (strict superset). Safe because both removal sites are scoped to QA-layer / the stamp zone AND to STANDALONE Text/MText — never title-block ATTRIBUTES, so the live "ISSUED FOR CONSTRUCTION" revision-row value is untouched. QA check #3 mirrors this: `FOR\s*CONSTRUCTION` (not `\s+`) with a `(?<!ISSUED)` lookbehind guard. Diagnose such leftovers with a filtered SelectionSet (`DXF 1 = *CONSTRUCTION*`), never by iterating the 24k-entity ModelSpace.
- **Pass 3 resilience**: PaperSpace iteration uses per-layout try/except so one layout failure doesn't skip the rest (previously entire loop was wrapped in single try/except → one COM error skipped all subsequent layouts, e.g. SEC-02 9-page all with leftover FOR CONSTRUCTION)
- **Failure mode if geometry cleanup skips IFC_STAMP layer**: Pass 1 only searches active space; Pass 3 may fail silently on COM error → FOR CONSTRUCTION on IFC_STAMP layer in PaperSpace survives all removal passes → leftover in AS BUILT output (SEC-02, CFG-001)

### IFC Revision History
- `preserve_ifr=True` **(default for ALL projects)**: KEEP all existing IFR revision rows, add IFC row AFTER the last IFR row. **Idempotent**: if an IFC row already exists (DESCRIPTION contains "CONSTRUCTION"), overwrite it in-place instead of appending a new row. Clean up any duplicate IFC rows above target_row (from previous bug).
- `preserve_ifr=False`: Rev0 clears all IFR rows. Rev1+ keeps existing IFC rows. Use for clean IFC-only revision tables. Currently no project uses this.
- **CRITICAL**: `_update_title_block()` must be idempotent — calling it multiple times on the same title block must produce the same result as calling it once. This prevents duplicate "FOR CONSTRUCTION" rows when `_find_all_title_blocks()` returns the same entity twice or when conversion is re-run.

### Revision-Table Column Alignment (AS BUILT)
- Some title-block definitions place a revision row's attribute slots at a slightly different X than the rows above (Warnertown `GPA`/`ACE-Wanertown_Siyuan`: row-5 PROJECT slot X=286.6 vs rows 1-4 X=281.8 → the AS BUILT row's `GG31` rendered crooked, shifted right). Prior conversions only filled rows 1-4 so it never showed; the bot's AS BUILT row lands on the misaligned slot.
- `AsBuiltManager._update_title_block()` → `_align_attr_x()`: after writing the AS BUILT row, copy each attribute's `InsertionPoint.X` + `TextAlignmentPoint.X` from the row directly below (preserving own Y), then `Update()`. No-op for already-straight columns; keeps every revision-table column vertically aligned. Cosmetic + best-effort (COM failures swallowed, never breaks conversion).

### Multi-Sheet DWG Title Blocks
- A single DWG can have MULTIPLE title block references — in ModelSpace AND/OR in separate PaperSpace layouts
- **CRITICAL**: `SelectionSet.Select(5)` only searches the ACTIVE space, NOT all layouts
- `_find_all_title_blocks()` uses TWO passes:
  1. SelectionSet in ModelSpace (handles 400K+ entity DWGs safely)
  2. Direct iteration of EACH PaperSpace layout's Block (PaperSpace layouts are small, safe to iterate)
- **PaperSpace priority**: Collect MS and PS results into FOUR separate lists (`_ms_tbs`, `_ms_fallback`, `_ps_tbs`, `_ps_fallback`). Priority: PS exact > PS fallback > MS exact > MS fallback. Stamps drawn in PaperSpace are always visible in PDF (no viewport freeze issues). Stamps drawn in ModelSpace may be outside PaperSpace viewport's visible area.
- `_remove_ifc_stamp()` also cleans ALL PaperSpace layouts (Pass 3: `_is_stamp_polyline()` identifies closed rectangles ~111x18, aspect ratio 3-15, at x>500 y<200), not just active space
- `convert_to_ifc()` must update EVERY title block (revision, personnel, date) and add stamp near EACH one
- Call `_remove_ifc_stamp()` ONCE before the loop, then `_stamp_via_com_draw()` per title block
- `_add_ifc_stamp()` is for single-TB convenience only (it calls remove + draw internally)
- **Failure mode if only active space searched**: multi-layout DWGs like PLN-004 (2 PaperSpace layouts) only get 1 title block found → page 1 missing FOR CONSTRUCTION
- **Failure mode if PaperSpace TBs not prioritized**: C-PLN-003's TBs (`Coleamablly`, `GPA`) are in PaperSpace — stamps drawn in ModelSpace are outside viewport view → invisible in PDF

### IFC Stamp
- COM draw only (MText + bold rect on `IFC_STAMP` layer). No block modify.
- **FOR CONSTRUCTION** box always drawn by `_stamp_via_com_draw()`
- **COLOUR box** logic (3 outcomes via `_scan_has_colour()` + `_check_colour_overlap()`):
  1. `has_colour=True` → original DWG already has COLOUR text → keep as-is, fix typo only, do NOT redraw
  2. `has_colour=False` + overlap detected → existing drawings/text in COLOUR area → skip to avoid overlap
  3. `has_colour=False` + no overlap → draw COLOUR box (rect + MText "DRAWINGS TO BE PRINTED IN COLOUR")
- `_scan_has_colour(doc)`: **Two-pass scan** — Pass A: SelectionSet MTEXT in ModelSpace (safe for 400K+), Pass B: direct iteration of each PaperSpace layout. Fixes COULOUR→COLOUR typos in-place. MUST search both spaces — SelectionSet mode 5 only searches active space, PaperSpace COLOUR MText is invisible to it.
- `_check_colour_overlap(doc, ..., layout_name=None)`: When layout_name is PaperSpace, iterates layout block directly (bounding box overlap check). When ModelSpace, uses SelectionSet crossing window (mode 1) with ±15 unit expansion. Skips IFC_STAMP/DEFPOINTS/ASHADE layers and viewports.
- `_ensure_colour_has_border(doc, ...)`: When `has_colour=True`, checks if existing COLOUR MText has an enclosing rectangle. If not, draws one at calculated stamp position on IFC_STAMP layer. Called from both IFCStampMixin and AsBuiltManager `_stamp_via_com_draw`.
- **Failure mode if `_scan_has_colour` only searches active space**: PaperSpace COLOUR text missed → `has_colour=False` → duplicate COLOUR box drawn → stamp overlap (PLN-002, PLN-005)
- **Raster overlap auto-fix (viewport-aware, AS BUILT)**: `_check_colour_overlap` is COM-based and CANNOT see ModelSpace content shown THROUGH a PaperSpace viewport — so a stamp can land on construction notes that are only visible in the published PDF (C-PLN-006 page 1: COLOUR/AS BUILT boxes drawn over the NOTES block). After `convert_to_ab` publishes, `_raster_fix_stamp_overlaps(acad, dwg, pdf)` RENDERS the PDF (PyMuPDF), per page detects the two red stamp boxes and measures foreign BLACK ink inside them, and if a box is occupied moves that page's layout's `IFC_STAMP` group to the nearest clear right-edge slot (smallest shift; notes-filled sheets only have room ABOVE the notes), then republishes. **Per-layout**: only the overlapping page moves (C-PLN-006 page 1 moved up; page 2 stayed at the standard bottom-right). **No-op for clean pages** (`_scan_pdf_stamp_overlaps` returns empty → early return → zero regression, the common case). paper-shift = pdf-shift ÷ scale, scale = COLOUR red-rect pdf-height ÷ COLOUR polyline paper-height. Page→layout via `_publish_layout_order` (TabOrder, same order `_publish_single_pdf` emits pages). Closed-loop (re-scan up to 2 passes). Republish must happen WHILE the short junction is alive (long AB path >256 → PUBLISH rejects the real path; `doc.FullName` is the junction alias).
- Position INSIDE frame bottom-right, scaled proportionally from reference measurements
- Cover sheet (page `00`): NO stamp
- **Black border**: Both polyline and MText explicitly set `color = 7` (black). Layer color also set to 7.
- **Viewport thaw**: After creating stamp in ModelSpace, thaw `IFC_STAMP` layer in ALL PaperSpace viewports (viewport layer freeze overrides global state — stamp invisible if frozen in viewport). Uses `GetFrozenLayers()`/`PutFrozenLayers()` with empty VARIANT array fallback + VPLAYER SendCommand fallback. **Skip for PaperSpace stamps** — entities drawn directly in PaperSpace are always visible in PDF, no viewport thaw needed.
- **Failure mode if viewport thaw omitted for ModelSpace stamps**: Stamp drawn in ModelSpace but invisible through PaperSpace viewport → Layout1 missing FOR CONSTRUCTION

### Typo Correction
- `_fix_known_typos(doc)` runs during IFC conversion (both IFCManager and PanelIFCManager)
- Uses SelectionSet with MTEXT filter (safe for large DWGs)
- **Case-insensitive** matching via `re.sub(re.IGNORECASE)` — handles `Coulour`, `COULOUR`, `coulour`
- Known fixes: `Coulour` → `Colour`
- Add new typos to `IFCStampMixin._TYPO_FIXES` list

### IFC Filename Sanitization
- `_build_ifc_filename()` strips `()` and `&` from description (in addition to `<>:"/\|?*`)
- **Why**: AutoCAD `-PUBLISH` command parsing breaks on `()` and `&` in DWG filenames
- Failure mode: PUBLISH fails silently → no PDF output

### Conversion Success = PDF Created
- `result['success']` in `convert_to_ifc`, `convert_to_ab`, `convert_multi_to_ab` MUST depend on `pdf_ok` (PDF actually exists on disk)
- DWG-only conversion (SaveAs succeeded but PUBLISH failed) is NOT success — user needs the PDF
- **Failure mode if success unconditional**: bot reports "转换=10, 失败=0" but 0 PDFs created → user thinks conversion worked, wastes time investigating missing files
- Previously (before 2026-05-25 fix): `result['success'] = True` was set regardless of pdf_ok

### Incremental Check
- NO mtime comparison (unreliable with Dropbox/cloud sync — Dropbox touches mtime on sync)
- Only skip if a **PDF** with matching doc-ID exists in `4. IFC(Client)/` — that's the final deliverable
- Native folder IFC DWG alone is NOT sufficient (DWG may exist from a failed conversion where PDF export didn't complete)
- User can force re-conversion via `force_doc_ids` or bot "强制重转" button
- **When code changes require re-export**: use "强制重转" for affected doc-IDs, or delete old PDFs from `4. IFC(Client)/`
- **Failure mode if ignored**: files get re-converted unnecessarily, or incomplete conversions get skipped

### Doc-ID Matching (FILE NO vs FILE NAME)
- **Rule**: Match by FILE NO (doc-ID) only, even if FILE NAME (description part) is inaccurate or different
- `_extract_doc_id_standalone()` uses two-pass strategy:
  1. `match()` — doc-ID at start of filename (standard naming)
  2. `search()` — doc-ID anywhere in filename (non-standard/inaccurate naming)
- `_find_native_folder()` matches folders by doc-ID only, not by description
- **Failure mode if ignored**: files like `GG31-E-PLN-003_COMMS &AUX Cable...` get silently skipped when description doesn't match standard naming
- When doc-ID extraction fails, log a visible warning (⚠) so user knows which files are being missed

### IFC Stamp on ALL Sheets (FOR CONSTRUCTION)
- `convert_to_ifc()` must apply FOR CONSTRUCTION to EVERY sheet in a multi-sheet DWG
- `_find_all_title_blocks()` finds ALL title block instances via SelectionSet
- Loop over ALL title blocks: update revision description + add visual stamp near each
- Log each sheet's stamp status (`Sheet N/M: 更新 title block + FOR CONSTRUCTION`)
- If `block_ref` or `space` is None for any sheet, log warning instead of silently skipping

### IFC Filename Normalization
- `normalize_ifc_filenames()` runs at start of `batch_convert_approved()` — adds `_IFC` suffix to files missing it
- Any file in `4. IFC(Client)/` is treated as IFC regardless of `_IFC` suffix

### Version Manager — Download Duplicate Handling
- `extract_base_name_and_version()` strips browser download-duplicate suffixes `(1)`, `(2)` etc. from filenames BEFORE extracting base name and version
- Without this, `_RevB.pdf` and `_RevB (1).pdf` get different base names and are treated as separate files instead of duplicates
- `identify_old_versions()` then correctly groups them and keeps the newest by modification time
- Applied in both `ifr_automation_v10.py` and standalone `version_manager_v5.py`

### Version Manager — Client Sharepoint Coverage
- `VersionManager.TARGET_SUBDIRS` includes `13. Client Sharepoint/1.IFR/1.Report` and `2.Drawing`
- `_find_ss_folder()` auto-detects existing `SS`/`Superseded`/`Superceded` folder (Client Sharepoint uses `Superseded`)
- Old versions in Client Sharepoint are moved to `Superseded/` just like in `IFR(Client)`
- Applied in both `ifr_automation_v10.py` and standalone `version_manager_v5.py`

### Version Manager — IFC(Client) + Client Sharepoint Coverage
- **Three entry points** for IFC version cleanup (all cover `4. IFC(Client)/`, `13. Client Sharepoint/1.IFR/2.Drawing/`, AND `13. Client Sharepoint/2.IFC/`):
  1. `/ifr` Stage 2: `VersionManager` via `TARGET_SUBDIRS` (includes all 3 paths)
  2. `/pipeline` Stage 3: `IFCTransmittalManager.run_for_pipeline()` — scan + dedup only (no Excel write, no state update). `ifc_map` passed to Stage 5 for unified deliverable update.
  3. `/ifc`: `IFCTransmittalManager.run()` Step 1 — `VersionManager.process_directory()` on all `VERSION_MANAGED_PATHS`, then Step 2b doc-ID based dedup as fallback. Step 5b syncs IFC PDFs to Client Sharepoint.
  4. `/panel_ifc` (Approved IFC) Step 4b: `IFCTransmittalManager` scan + deduplicate on `4. IFC(Client)/`, then sync to `13. Client Sharepoint/2.IFC/` + version cleanup
- **IFC PDF sync**: `_sync_ifc_files_to_sharepoint()` copies PDFs from `4. IFC(Client)/` to all `sync_targets` (currently `13. Client Sharepoint/2.IFC/`). Idempotent: skips files with matching size+mtime.
- **Transmittal config**: Per-project header via `config.json` `project_overrides` keyed by folder name. `_get_project_number()` supports GG-style (`GG-31`) and 5-digit formats.
- `IFCTransmittalManager._extract_doc_id()` delegates to `_extract_doc_id_standalone()` — supports GG, LMS, TSF, and generic patterns (previously only LMS → GG31 files silently skipped)
- `VersionManager.TARGET_SUBDIRS` includes `4. IFC(Client)` — `identify_ifc_in_ifr_client()` has `"IFR(Client)" not in str(target_dir)` guard, does NOT misfire
- SS/ folder auto-created if it doesn't exist
- Applied in both `ifr_automation_v10.py` and standalone `version_manager_v5.py`
- **Failure mode if ignored**: old IFC revisions accumulate, deliverable Excel may show stale revision numbers

### IFR Report Sync — Latest Version Only
- `_mirror_reports()` must only sync the **latest version** per doc-ID to `IFR(Client)/2.Reports`
- Source folders (`Reports/Electrical/`, `Reports/Civil & Structure/`, etc.) may contain multiple revisions of the same document in subfolders (e.g. `RevA.pdf` + `RevB.pdf` in the same folder, old version not yet in SS/)
- **Three-phase approach**:
  1. Collect all candidate PDFs with mtime
  2. Group by doc-ID via `_extract_doc_id_standalone()`, keep only the newest per group
  3. Dedup by filename + copy
- Files without extractable doc-ID are kept as-is (no grouping)
- **Failure mode if ignored**: outdated revisions get synced to IFR(Client), cluttering the client-facing directory

### Sharepoint Sync — IFR_internal / IFR(Client) → Client Sharepoint
- `sync_to_sharepoint()` copies non-approved IFR files to `13. Client Sharepoint/1.IFR/`
- **Source mapping** (note swapped numbers):
  - `2. IFR_internal` → `13. Client Sharepoint/1.IFR/2.Drawing` (version-managed, unique per doc-ID)
  - `IFR(Client)/2.Reports` → `13. Client Sharepoint/1.IFR/1.Report`
- **Why IFR_internal for Drawing**: already has version management + uniqueness per doc-ID; avoids syncing duplicates or outdated revisions from IFR(Client)/1.Drawing
- **Fallback**: if `IFR_internal` doesn't exist, falls back to `IFR(Client)/1.Drawing`
- **Pre-sync archiving** (`archive_approved_in_ifr_client()`): two archive triggers:
  1. File has `-Approved` suffix in IFR(Client) (local trigger)
  2. File's doc-ID already approved in Client Sharepoint AND file revision ≤ highest approved revision (reverse feedback — revision-aware)
- **Why revision-aware reverse feedback**: employees see only pending files in `IFR(Client)/`; approved items move to `Approved to IFC/`. But if RevA is Approved and RevB is a new IFR submission, RevB must NOT be archived — it needs to flow through to Client Sharepoint for review.
- **Failure mode if NOT revision-aware**: RevA Approved → reverse feedback archives RevB (newer) → sync skips RevB → client never sees new submission
- **Skip rules** (any match → skip):
  1. File has `-Approved` suffix → managed by `ApprovedIFCManager`
  2. File's doc-ID already approved AND source revision ≤ approved revision → already approved or older (e.g. RevA Approved, RevA source → skip; RevB source → **sync** for new client review)
  3. Exact filename already exists in target with same size/mtime → already synced
- **Why revision-aware skip**: RevA may be Approved, but RevB is a new submission — must be synced for client to review. Previous behavior skipped ALL revisions once any version was Approved, blocking new IFR submissions from reaching Client Sharepoint.
- **Pipeline position**: Stage 3 of 4 (after Version Management, before Deliverable Cross-Check)
- Uses existing `_should_copy()` for idempotent copy (size + mtime comparison)

### Deliverable Cross-Check — Date Format & IFC Status
- **Date format**: `strftime('%d/%m/%y')` — Excel uses `dd/mm/yy` (e.g. `19/03/26`), NOT `YYYY-MM-DD`
- **Date MUST be written** for ALL update types: revision mismatches, IFC status, doc-ID corrections. Previously revision mismatches didn't write date → stale dates persisted.
- **Status MUST be set** for ALL matched items with revisions. If folder revision matches Excel but status is empty → set "Submitted". Previously only flagged when `folder_rev > excel_rev`.
- **SOURCE_FOLDERS** includes `IFR(Client)/2.Reports` as fallback — catches reports where source folder only has `.docx` but PDF was placed directly in IFR(Client).
- **Subfolder skip list**: `('ss', 'superseded', 'superceded', 'approved to ifc')` — prevents scanning archived/approved files.
- **IFC status sources** (`scan_ifc_folder()`):
  1. `4. IFC(Client)/` — numeric revisions (e.g. Rev0, Rev1)
  2. `13. Client Sharepoint/1.IFR/1.Report/Approved to IFC/` — letter revisions (e.g. RevA, RevB)
  3. `13. Client Sharepoint/1.IFR/2.Drawing/Approved to IFC/` — letter revisions
  4. `3. IFR(Client)/1.Drawing/Approved to IFC/` — letter revisions
  5. `3. IFR(Client)/2.Reports/Approved to IFC/` — letter revisions
- **IFCTransmittalManager.update_deliverable_ifc_status()**: Write revision to `layout.rev_col` (NOT hardcoded column 12). Write date to `layout.date_col`. Use green fill on K-M. ALL numerical revisions including 0 are written.
- **K column = numerical IFC revision**: Once a doc-ID appears in `4. IFC(Client)/`, its K column MUST show the numerical revision (0, 1, 2, 3...), replacing the old letter revision. Rev0 is valid and must be written — do NOT guard with `rev != 0`.
- **`/ifc` command MUST always update deliverable**: `IFCTransmittalManager.run()` builds `all_ifc_map` from ALL current IFC files (via `scan_ifc_files()` grouped dict), NOT just `new_files` from `collect_new_ifc_files()`. The `new_files` check only gates transmittal generation (Step 5+), NOT deliverable update (Step 3). Previously, when all files were already in `ifc_state.json`, `new_files` was empty → early return skipped deliverable update entirely.
- **`/deliverable` bot display MUST show IFC status**: `format_deliverable_summary()` includes IFC status count. `format_deliverable_details()` shows each doc-ID with old status → Approved IFC + revision.
- **Failure mode if only source 1 scanned**: Reports never appear in `4. IFC(Client)/`, so their IFC status is never reflected in deliverable Excel
- **Failure mode if rev col hardcoded**: `update_deliverable_ifc_status()` wrote integer revision (0) to date column (L) → Excel showed `0/01/00`
- **Failure mode if deliverable gated by new_files**: All IFC files already transmitted → `new_files` empty → deliverable never updated → Excel shows stale IFR status for items already at IFC

### AS BUILT Post-Conversion QA (MANDATORY)
- `_qa_validate_ab_pdf()`: Auto-runs after every successful AB PDF export. Uses PyMuPDF (fitz) to check:
  1. Page count matches expected (from title block count or page DWG count)
  2. No duplicate stamps (COLOUR count > 1 per page)
  3. No leftover IFC stamps ("FOR CONSTRUCTION" text)
  4. No leftover IFR stamps ("ISSUED FOR REVIEW" text)
  5. No phantom pages (different dimensions from majority — indicates Model tab export)
  6. AS BUILT stamp present on every page
  7. Stamp does NOT overlap design content (the headline cross-project invariant — see below)
- Results stored in `result['qa_warnings']`, logged inline, summarized in `batch_convert`
- **Critical**: QA must detect ALL issues that were previously found manually. If QA passes, the output is production-ready.

#### Stamp detection — colour-/tag-/project-agnostic (mechanics + numbers)
The cross-project INVARIANTS live in skill `asbuilt-ifr-stamp-standard` (QA criteria + "two cures"). This is the V6 implementation + the tuneable numbers (numbers stay HERE, never in the skill).
- **`_detect_stamp_boxes(page)`**: a stamp box = a `get_drawings()` rect with `fill is not None` (the cw=2.0 border renders FILLED — red on colour CTBs, BLACK on mono CTBs like C-PLN-003) that ENCLOSES a stamp word (`BUILT`/`COLOUR`/`COLOR`). Size window `0.04–0.30·pw` × `0.012–0.12·ph`; dedup near-coincident rects. Replaces the old red-only `_detect_stamp_red_rects` (mono stamps were invisible to it) and the old zone+size guess (title-block detail/revision-table CELLS share the zone → false "未对齐", esp. when the stamp OVERLAPS a table so the cell encloses the stamp text; filled-only excludes thin-stroked cells).
- **`_stamp_overlaps_content(page, boxes)`** (geometry, colour-independent — raster ink-counting breaks on mono CTBs where the stamp's own black text reads as "ink"): overlap = (a) a FOREIGN text word (not in `_STAMP_OWN_WORDS`) centred inside a box, OR (b) a foreign cell rect (`fill=None`) covering **>15%** of a box. Foreign rect must be CELL-sized in BOTH dims (`>0.02·pw & >0.01·ph`, but `<0.5·pw & <0.35·ph`) so sheet borders / viewport frames (large in both dims) don't 100%-intersect and false-flag (GAD-003).
- **Check #7 (overlap gate)**: if `_stamp_overlaps_content` is true on the final PDF → warn `印章压住图面内容 — 自动避让无空位 [需人工检查]`. Contains `需人工检查` → escalates (no retry; re-converting reruns the same relocator). NEVER relax this to force a pass.
- **Placement check (#6 rewritten)**: bottom-right is DEFAULT but NOT required — a content-avoiding relocation is legitimate. Flag only OFF-sheet, or FLOATING mid-sheet (no box edge within `_EDGE_BAND=0.08` of any page edge). `_EDGE_BAND` is the tuneable number.

#### Stamp-vs-content auto-relocation (cure #2 "move the stamp")
- **`_scan_pdf_stamp_overlaps`**: trigger = `_stamp_overlaps_content` (geometric). Then grid-search candidate group anchors, ordered by proximity to the BOTTOM-LEFT page corner (the conventional empty zone), first candidate whose every box is raster-clear wins. Clear = `_black_frac < _RASTER_CLEAR_FRAC` (0.004) measured on the empty target (no stamp there → no own-ink confound). Returns 2-D `{dx,dy}`.
- **`_raster_fix_stamp_overlaps`** (single-page path, `max_iter=5`): translate the whole `IFC_STAMP` group by `(dx,dy)/scale` (scale = stamp-box PDF height ÷ paper polyline height; PDF +y is down → negate for paper), Save, republish, re-scan. 5 iters lets scale/landing error self-correct (each scan re-measures actual position).
- **Verified (NSW153 Coleambally #2)**: C-PLN-003 / C-PLN-005 stamps relocated to clear lower-left, QA PASS; E-BLD-001 (dense block diagram, no on-edge clear slot) correctly FAILs `[需人工检查]`. Genuinely multi-source (`is_multi_page`, `len(dwgs)>1`) drawings go through `convert_multi_to_ab`, which currently does NOT call the relocator (publishes via `_publish_ab_group_pdf`) — relocation there is a known gap; QA still catches the overlap.

#### Revision-row fallback cell — pitch-based vertical placement
- A missing `{n}{suffix}` slot on the AS BUILT row is drawn as standalone text (`_draw_cell_fallback`). Its Y is the same-column reference cell's align-Y translated up by the row PITCH (median dY per row-number over all same-column pairs), NOT a single global `target_y` read off another column. Mixing a left-justified cell's baseline-Y with a centre-justified cell's middle-Y floated the drawn glyph half a line high (the 'ACE' drift on C-PLN-007 / GAD-001 / GAD-004). Pitch is a like-for-like difference → convention-independent; anchoring on the same column keeps the baseline consistent.

#### Structured faults + remediation playbook (presentation layer — must NOT alter verdicts)
- **`qa_playbook.json`** is the SSOT DATA for fault remediation: each fault code → `match` substrings, `severity` (挡出图/可重试/提示/基建), `auto_tried`, `fix_checklist`. Add/edit a fault HERE only (code reads it; this markdown + the skill only POINT here, never copy the checklists — 铁律 2/3).
- **`qa_faults.py`** is PARSE-ONLY and engine-independent: `classify()` / `build_fault()` / `faults_from_result()` turn the flat QA warning STRINGS the engine already emits into structured objects for the bot / a future UI. It does NOT run QA/COM/geometry and NEVER feeds back into PASS/FAIL/escalate/retry — `_qa_validate_ab_pdf` and `_QA_ESCALATE_KEYWORDS` remain the sole verdict authority. The bot's `_render_faults()` (telegram_bot_v4) renders each fault as a ☐ checklist; lazy-imports `qa_faults`, falls back to plain error lines on any failure.
- **Regression gate**: `golden/test_qa_regression.py` (offline, read-only — calls the real `_qa_validate_ab_pdf` via `__new__`, no AutoCAD) asserts per-PDF warnings+verdict match the frozen `golden/qa_verdicts.json`. Run it after ANY change that could touch QA; re-baseline only by deliberately running `golden/gen_qa_golden.py`. This is what locks the user-approved gold standard (26 PDFs: 25 PASS + E-BLD-001 ESCALATE).

### PDF Layout Filtering (Phantom Page Prevention)
- `_publish_single_pdf()` skips layouts with no plot configuration (`ConfigName` empty or 'None') — phantom layouts that gain entities after XREF bind but have no real content
- This is IN ADDITION to the existing entity count check (`≤1 entity = empty`)
- **Failure mode if not filtered**: XREF bind adds entities to previously-empty layouts → DSD includes them → PDF has extra blank Model-tab-sized pages (E-PLN-004: 2 pages instead of 1)

### Other
- **Bot flow**: No interactive prompts. COM mode default (no COM/UTB selection menu)
- **Imports**: `telegram_bot_v4.py` imports from `ifr_automation_v10` (not v7/v8/v9). Bot picks up fixes to VersionManager/IFCManager automatically without bot code changes.
- **Personnel fallback**: Pages without revision rows inherit `group_personnel` from first page with data
- **Sheet numbering**: `SHEET_NO` and `TOTAL_SHEETS` filled automatically in `ACE_TitleBlock_CHINT`
- **Large DWG safety**: NEVER iterate ModelSpace/PaperSpace entity-by-entity. Use SelectionSet with filters. DWGs can have 400K+ entities.
- **COM resilience**: `_com_retry()` for all COM calls. After failure: close all docs + reset COM. Lock files: attempt `unlink()` first.
- **Attribute access safety**: NEVER access `.TextString` directly on COM attribute objects. Always use `_safe_get_text(attr)` for reads and `_safe_set_text(attr, value)` for writes — both wrap `_com_retry()`. `_get_attrs_dict()` also wraps `GetAttributes()` and `TagString` in `_com_retry()`. Failure mode: "Call was rejected by callee" crash during title block update.

### Document Open Robustness
- After `Documents.Open()`, wait for BOTH `ModelSpace.Count` AND `Layouts.Count` to succeed (up to 20s)
- Extra 2s settle time after document load (first open in session is slowest)
- Title block search: retry up to 3 times with increasing wait (3s, 6s) if first attempt fails
- PaperSpace title block search: retry up to 3 times with increasing wait (COM "Call rejected" on first attempt is common)
- XREF binding before temp SaveAs: when output path > 240 chars, bind all XREFs (`blk.Bind(False)`) to prevent missing content
- After XREF bind: `doc.Regen(1)` + 2s settle before SaveAs (binding changes document state)
- **Failure mode if skipped**: first conversion in session fails because AutoCAD not fully ready; PLN-004 "Add.Select" error

### Long SOURCE Path (>260 chars / MAX_PATH) — universal
- A source DWG path > 260 chars breaks TWO things; both are now handled in code:
  1. **Scan invisibility**: `pathlib.Path.is_file()`/`.exists()` do a full `os.stat` on the long path → `FileNotFoundError` → returns False → the DWG is silently dropped from the scan. Fix: `_find_latest_ifc_source` enumerates with `os.scandir` (dirent-based `is_file()`, never re-stats the long path). NEVER gate DWG enumeration on `Path.is_file()`.
  2. **Open failure**: AutoCAD COM `Documents.Open` cannot open a >~256-char path and **rejects the `\\?\` prefix** ("Invalid file name"). Fix: `AsBuiltManager._shortpath_open_target()` exposes the DWG's PARENT folder via a short `mklink /J` directory **junction** in TEMP and opens through it — relative XREFs still resolve (unlike copying the lone DWG to temp). `rmdir` removes only the junction link, never the target. Cleaned up in `finally`.
- **Not a Dropbox/online-only problem** — online-only placeholders open fine if the path is short (C-PLN-008, 253 chars, was online-only yet converted). The discriminator is purely path LENGTH. Don't chase cloud-hydration red herrings.
- Output long paths (SaveAs) already handled separately via temp short path + `to_long_path` move (see below / >240 rule).
- Surfaced 2026-06-04: 5 of 6 Warnertown Civil & Structure report-folder drawings (paths 262-287 chars) failed to open until the junction fix; the 6th (253) worked.

### SaveAs Resilience (backup plans — NEVER just abort)
- **6 fallback strategies** in sequence if SaveAs fails:
  1. Normal `doc.SaveAs(path)` with `_com_retry`
  2. `doc.Save()` first (flush XREF bind state), then `doc.SaveAs(path)`
  3. `doc.SaveAs(path, 61)` with explicit acNative DWG format
  4. If using temp path, try direct long path instead
  5. `SendCommand('_SAVEAS\n\n{path}\n')` via command line
  6. Last resort: `doc.Save()` in-place + PUBLISH PDF directly from open document
- Clean up pre-existing file at save_path before each attempt
- **Failure mode before fix**: E-PLN-003 XREF bind succeeded but SaveAs failed ("Open.SaveAs") with no fallback → conversion aborted with stamps+TB already done but no output

### Lock File Resilience
- `.dwl`/`.dwl2` lock files are **unavoidable** — NEVER abort just because they exist
- **Strategy**: try delete (3 retries) → check if DWG already open in AutoCAD → reuse open document → if stale lock, open anyway
- Both IFCManager and PanelIFCManager use same approach
- **Failure mode before fix**: PLN-004 locked → script aborted immediately

### IFR Stamp — Universal Across Projects
- IFR stamp spec and position are **identical across ALL projects** (current and future)
- Only **scale/proportion** varies per project (title block size differs → stamp position shifts)
- Scale issue was identified and fixed via test files on 2026-03-16 (proportional scaling from reference frame)
- **Validation required**: when adding a new project, verify stamp appears in correct position by test-converting one DWG first

### Revision Tag Format — Universal but Flexible
- Current format across ALL projects: `{row}REV`, `{row}DATE`, `{row}DESCRIPTION`, `{row}{PERSONNEL_TAG}` (e.g. `1REV`, `2DATE`, `3DESCRIPTION`)
- This may change in future projects — code must NOT hardcode tag patterns, use project config

### Reference Frame Dimensions — Must Verify Per Project
- Proportional stamp scaling uses `_REF_TB_WIDTH` and `_REF_TB_HEIGHT` from reference frame
- Tatua: confirmed 841×594 from `Tatua_Standard_Frame.dwg`
- Warnertown: confirmed 841×594 from `ACE_Standard_Frame_wanertown_UPDATED.dwg` (verified 2026-03-19)
- **Stamp coordinates must be RELATIVE to title block bbox**, not absolute from origin. Formula: `stamp_pos = tb_min + REF_OFFSET * scale`. Tatua works with absolute coords only because its TB happens to be near origin; Warnertown TB is at (2739, 43).

## Projects
- **Warnertown BESS** (SA):
  - Title block: `ACE-Wanertown_Siyuan` (ModelSpace + PaperSpace, 69 attrs, 6 revision rows)
  - Tags: DRAWN/CHECK/ENGINEER/QA/PROJECT (plus SUBJECT, DRAWINGNUMBER per row)
  - Reference frame: `ACE_Standard_Frame_wanertown_UPDATED by MG.dwg` in `1. Native/`
  - No built-in IFR stamp block (stamps are standalone entities per drawing)
  - Stamp geometry: proportional scaling from 841×594 reference (verified 2026-03-19, same as Tatua)
  - AS BUILT: `AsBuiltManager` auto-detects paths. Native: `Design/Engineering/1. Drawings/1. Native`. AB output: `Design/Engineering/1. Drawings/5. As Built/3. As Built Client`. 13 doc-IDs detected, 11 existing AB Rev1.
  - **AS BUILT also covers report-folder drawings** (`_EXTRA_DRAWING_SOURCES`): native DWGs that live inside `2. Calcs & Reports/Reports/Civil & Structure` (and `…/Electrical`) per-report folders, NOT in `1. Native/`. The folder is named by the REPORT doc-ID but the DWG carries the DRAWING doc-ID (e.g. folder `GG31-C-RPT-001` holds drawing `GG31-C-PLN-006`) → doc-ID/description parsed from the DWG filename, deduped against the main Native scan. 6 Civil drawings converted 2026-06-04 (C-PLN-006/008/009/010/013 + C-RPT-004), all REV 1, QA clean. Required the long-path junction fix (see "Long SOURCE Path").
  - AS BUILT stamp verified on real DWG (BLD-001, 1689×1192.9 = 2× standard frame): CW=3.16, text=11.07 — proportional scaling works
  - `preserve_ifr=True`
- **Tatua Solar Farm** (NZ, Coleambally):
  - Standard frame: `Coleamablly` (ModelSpace), tags: DESIGNED/DRAWN/APPROVED/PROJECT
  - Has built-in 'IFR' block with stamp text
  - Panel design: `ACE_TitleBlock_CHINT` (PaperSpace, 45 attrs), tags: DESIGNED/DRAWN/APPROVED/PROJECT, plus `SHEET_NO`, `TOTAL_SHEETS`, 4 revision rows
  - Reference frame: 841×594 (verified)
  - Folder names: `☆` prefix + 2-digit doc-ID suffixes (e.g. `☆TSF-EN-CIL-DRG-01`)
  - AS BUILT: `AsBuiltManager` auto-detects paths. Native: `1. Drawings/1. Native` (no `Design/Engineering/` prefix). AB output: `1. Drawings/6.AS Built`. 2 doc-IDs detected.
  - `preserve_ifr=True`
- **LMS / NAWMA BESS** (SA):
  - Doc-ID prefix: `50023-`
  - Dropbox path: `Project (EPC)/2.SA/LMS/Design/Engineering/`
  - Client: LMS Energy
  - Deliverable Excel layout: L=Revision, M=Submission Date, N=Status (row 9)
  - Brownfield project (existing substation expansion) — higher complexity
  - Design Review Comment Register: `50023-Design Review Comments Register_Rev{X}.xlsx` in `3. IFR(Client)/`
  - RFI Register: `LMS NAWMA BESS Project RFI_{X}.xlsx` in `3. IFR(Client)/`
  - Title blocks: `Coleamablly` (43 attrs) and `Riverina_tellhow` (43 attrs) — both in `AsBuiltManager.LMS_TITLE_BLOCKS`
  - Tags: DESIGNED/DRAWN/CHECK/APPROVED/PROJECT (same as Tatua)
  - AS BUILT conversion: `AsBuiltManager` inherits `IFCStampMixin._stamp_via_com_draw` (841×594 ratio system) — same frame as FOR CONSTRUCTION, guaranteed alignment with COLOUR stamp.
  - Multi-page verified: EL-001 (3 PaperSpace layouts 01/02/03, Riverina_tellhow TB) — all 3 sheets stamped + PDF exported (982 KB)
- **Coleambally #2** (NSW):
  - Doc-ID prefix: `NSW153-`
  - Dropbox path: `Project (EPC)/1.NSW/Coleambally #2/Design/Engineering/`
  - Title block: likely `Coleamablly` (same as Tatua/LMS)
  - Tags: DESIGNED/DRAWN/CHECK/APPROVED/PROJECT
  - IFC DWGs: flat in native folder with `_IFC` suffix and LETTER revisions (e.g. `_RevF_IFC.dwg`)
  - AS BUILT: **subfolder mode** (standardized to match Warnertown). `3. As Built Client/` created 2026-05-29. DWG to `Rev.N - AB/` inside Native doc-ID folders, PDF to `5. As Built/3. As Built Client/`. XREF binding enabled.
  - IFC conversions are pre-bot (2025-05) — stamps may be on non-standard layers, see [[pre-bot-ifc-awareness]]
  - Native cleaned 2026-05-29: 69 .bak + 5 .dwl + 4 old IFC DWG + 2 Copy DWG moved/deleted
  - VersionManager cleaned 2026-05-29: IFR_internal 5 + IFC(Client) 2 old revisions → SUPERSEDED/
  - 25 IFC DWGs in Native, 10 existing AB PDFs in `3. As Built Client/`
  - `preserve_ifr=True`

## Approved IFC Flow (`/panel_ifc` without args)
- **Trigger**: `/panel_ifc` no args → project selection → approved mode. `/panel_ifc <folder>` → existing panel flow.
- **Scan**: `13. Client Sharepoint/1.IFR/{1.Report, 2.Drawing}/` for `-Approved` suffix PDFs
- **Archive**: Move detected files to `Approved to IFC/` subfolder; merge with already-archived files
- **Native mapping**: Extract doc-ID → find matching folder in `1. Native/` (flat or nested)
- **DWG cleanup**: Keep latest IFR DWG + IFC DWGs, move outdated to SS/
- **Convert**: Delegate to `IFCManager.convert_to_ifc()` (stamp + SaveAs + PDF export)
- **IFC version cleanup** (Step 4b): After conversion, `VersionManager.process_directory()` on `4. IFC(Client)/` — moves old revisions (e.g. `_Rev0_IFC.pdf` superseded by `_Rev1_IFC.pdf`) to `SS/`. Runs BEFORE deliverable update so Excel reflects latest revision.
- **Incremental**: PDF-only check in `4. IFC(Client)/` (no mtime); skipped items can be force-reconverted via bot button
- **Deliverable**: Auto-update Excel with "Approved IFC" status after conversion

## Issue Register Responder Assignment (Cross-Project Workflow)

When a client returns a Design Review Comment Register / RFI Register, the Responder column (M) needs to be populated for all Open items. Automated via `IssueRegisterManager` class in `ifr_automation_v10.py` and `/issue_register` Telegram bot command.

**Trigger:** New revision of Comment Register received in `3. IFR(Client)/`
**Input:** `{prefix}-Design Review Comments Register_Rev{X}.xlsx` → Sheet: `Master Register`
**Action:** Apply three-angle review → auto-fill Col M
**Output:** `{original}_updated.xlsx` with light blue (auto-filled) + yellow (conflict) highlights

### Three-Angle Review (MANDATORY — universal across projects)

| 维度 | 来源 | 精确度 | 角色 |
|------|------|--------|------|
| **1. Role** | Engineer position description | Lowest | Fallback default only |
| **2. Email Allocation** | Project coordinator email / `Team-Allocation.md` (json fence) | Medium | Project-level canonical |
| **3. Title Block** | Latest Rev IFC PDF → DRN/DES/CHK/APP | **Highest** | Ground truth |

**Conflict resolution**: Dimension 3 wins > 2 wins > 1. Generic stamps (`ACE`, `AW`, empty) filtered — degrade to DES/CHK. When dimension 3 absent (no IFC PDF), fall back to 2 → 1.

**Example of why this matters**: SLD (EL-001) is traditionally a Senior EE domain → dim 1/2 = CP. But title block showed DES=RV → dim 3 overrides, correct responder is RV. Relying on dim 1/2 alone routes comments to the wrong person.

### Excel Column Map

A=Status, B=DocNum, I=Comment, J=Severity (A/B/C/D), M=**Responder** (fill this), R=Closeout, W=Closeout(post-workshop)

### Skip Rules
- Col A != "Open" → skip
- Col W has closeout AND Col M already filled → skip (already handled)
- Empty Col M → auto-fill + light blue
- Existing Col M disagrees with resolved responder → keep value, add yellow fill (coordinator review)

### Project-Specific Allocation

**Allocation is NEVER stored in Dropbox** (company-visible). `IssueRegisterManager` loads
per-project config from the D: drive private vault only:

```
D:\3.Career\obsidian-vault\04-Work-SOP\Projects\{code}-*\Team-Allocation.md
```

The md file contains ONE fenced ```json``` block — this is the runtime config.
`IssueRegisterManager._extract_allocation_json_from_md()` scans for fenced json
blocks and picks the first one containing both `allocation` and `special_rules`
keys. Human notes and the machine-readable config live in the same md file
(single source of truth — no sync drift between .json + .md).

Project code is auto-detected from the register filename (`50023-Design Review...` → `50023`,
`GG31-...` → `GG31`). If no D: drive md (or json fence inside it) matches, the manager falls
back to title-block extraction (dim 3) only — cells with no title-block data are left blank
for manual fill.

Schema (inside the md's ```json fence):

```json
{
  "allocation":     {"<prefix>": {"responder": "XX", "note": "..."}},
  "special_rules":  {"<doc-id>": {"responder": "XX", "note": "..."}},
  "notes":          {"<responder>": "<annotation, e.g. HOLD>"}
}
```

Prefix matching uses `[A-Z]{2,3}-\d{3}` pattern in doc-id (e.g. `50023-EA-301` → prefix `EA`). Special rules check substring containment against the doc-id (e.g. `EA-300` matches `50023-EA-300`).

### Title-Block Extraction (dimension 3)

`IssueRegisterManager._extract_titleblock_fields()`:
1. Open PDF via `pdfplumber` (try/except — skip dimension 3 if not installed)
2. Scan last page first (title blocks live in bottom-right), then first page
3. `extract_tables()` → find table containing `DRN`/`DES`/`CHK`/`APP` column headers
4. Walk rows bottom-up until a populated REV row; return {DES, DRN, CHK, APP, rev}
5. Prefer DES > CHK > DRN (skip generic `ACE`/`AW`/empty)

**Failure mode if pdfplumber unavailable**: Dimension 3 silently skipped, allocation falls through to dimension 2. Log: `[IssueRegister] pdfplumber unavailable — dimension 3 skipped`

### Team Allocation Storage

- **Generic SOP**: `D:\3.Career\obsidian-vault\04-Work-SOP\PM-Methodology\Issue-Register-Responder-Workflow.md`
- **Per-project (single source of truth — human reference + runtime config)**: `D:\3.Career\obsidian-vault\04-Work-SOP\Projects\{code}-{name}\Team-Allocation.md` — human-readable notes AND a fenced ```json``` block parsed by `IssueRegisterManager` at runtime. No separate .json file (eliminates sync drift).
- **Dropbox**: NEVER stores `Team-Allocation.md` or any team/responder metadata (company-visible path).
- **C: drive memory** contains only the desensitized methodology (no member names, no project-specific allocation — those live on D: since C: is visible to other users)

## Knowledge Capture (MANDATORY — auto-fire after every non-trivial feature)

When a significant technical challenge is solved during implementation, **automatically** summarize and classify it to the Obsidian vault at `D:\3.Career\obsidian-vault\` — do NOT wait for the user to ask.

### Rules
1. **One note per domain folder** — if a breakthrough spans multiple domains (e.g. AI-Tech + Work-SOP), create one note in each with `[[cross-link]]` in Related section
2. **Naming**: `YYYY-MM-DD-Topic-Slug.md` (date prefix, no counters, no collision risk)
3. **Frontmatter**: follow each folder's schema (see `README.md` in each folder)
4. **AI Tech Stack section is MANDATORY** — every note must explain:
   - What AI tools were used at **development time** (e.g. Claude Code CLI as architect/test-engineer/implementer)
   - What runs at **runtime** (e.g. python-telegram-bot, rule engine, zero LLM calls)
   - The distinction matters: "AI-assisted development" ≠ "AI-powered runtime"
5. **Domain folders**:
   - `02-AI-Tech/Agent/` — technical patterns (state machines, prompt engineering, agent architecture)
   - `04-Work-SOP/Lessons-Learned/` — workflow improvements, process automation insights
   - Other folders as appropriate (check existing structure)
6. **No tech debt**: no central index to maintain, no counters, no orphan configs. Dataview queries in README.md auto-aggregate.

### What qualifies as "significant"
- New bot capability (e.g. conversational state machine, interactive review)
- Non-obvious architecture decision (e.g. zero-LLM runtime, PaperSpace-aware stamps)
- Hard debugging breakthrough (e.g. XREF binding before SaveAs, viewport thaw)
- Cross-project pattern that will recur

## Previous Versions
Stored in `SS/` subfolder. Key reference: `SS/ifr_automation_v8.py` has the original PUBLISH + DSD implementation.
