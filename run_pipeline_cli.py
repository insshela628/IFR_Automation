#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
run_pipeline_cli.py — headless CLI for the IFR pipeline.

Purpose: give the design panel (drag_window.py, a different repo/env) a
subprocess entry to the SAME pipeline the Telegram bot runs, WITHOUT importing
telegram / apscheduler. Reuses the bot's exact construction
(_scan_projects / PipelineOrchestrator) against the engine ifr_automation_v10.

Usage:
  python run_pipeline_cli.py --list
  python run_pipeline_cli.py "<keyword>"            # dry-run preview (safe)
  python run_pipeline_cli.py "<keyword>" --execute  # real run (moves files → Superseded/)
  python run_pipeline_cli.py "<keyword>" --stages deliverable asbuilt_version   # subset
  python run_pipeline_cli.py "<keyword>" --refresh-approved   # ONLY latest-wins Approved to IFC swap (real)
  python run_pipeline_cli.py "<keyword>" --ifc-status [--execute]  # internal DLV ← issued IFC PDFs
  python run_pipeline_cli.py "<keyword>" --replot-stale [--execute] [--only KEY ...]  # stale IFC → fresh PDF
  python run_pipeline_cli.py "<keyword>" --sync-13 [--execute]   # QA-gated 13 Client Sharepoint overwrite

--replot-stale: every stale IFC that may be redone at the SAME rev (submissions ledger: never sent,
or changed within 24h of sending = contingency) is brought back in line: designer-edited IFC DWG →
re-published as-is; IFR moved → reconverted; then the 13. Client Sharepoint copy is overwritten
after QA. Later changes → skipped, Rev bump = PM. Old PDF + change map in _state/<project>/replot/.
Refuses while AutoCAD is open. PM 2026-11-12: always first to fix.

--sync-13: 4. IFC(Client) PDFs whose same-name 13. Client Sharepoint copy differs → QA against the
submissions ledger (never sent / changed within 24h of sending; content QA; same page count) →
overwrite the 13 copy (old one backed up to _state/<project>/replot/). Else reported, untouched.

--ifc-status: new IFC PDF in 4. IFC(Client) → its INTERNAL DLV row becomes Approved IFC
+ numeric rev + date, stale to-do comment cleared. Idempotent (unchanged rows untouched).
Never rebuilds the client DLV — a PDF in the folder is not a transmittal. Driven by
auto_drain's IFC-PDF watch every round (PM 2026-11-09: 出了 IFC PDF 就要更新台账).

--refresh-approved: the one real write the daily patrol (auto_drain ④.8) is allowed. Runs
IFRAutomation.refresh_approved_to_latest alone (PM 2026-10-22: "以后不能这样一个个问了，你要自动
换成最新"): Approved to IFC (internal + client) → latest IFR, old copy → SS, IFR_internal
old letters → Superseded. No AutoCAD, no IFC copy to Sharepoint (those stay human-triggered).

Stage subset (`--stages`): `PipelineOrchestrator` has always accepted a `stages`
list; only this CLI never exposed it. Note it is a MEMBERSHIP FILTER, not a
sequence -- run_pipeline checks `if 'x' in self.stages` in a hardcoded order, so
you can choose WHICH stages run but never their order.

Output: human-readable progress to stdout; last line is a machine marker
  RESULT_JSON: {...}   (parsed by drag_window)
Exit: 0 ok, 2 no match / ambiguous, 1 error.
"""
import sys
import json
from pathlib import Path

# UTF-8 stdio BEFORE importing the engine: importing ifr_automation_v10 triggers
# colorama.init(), which wraps stdout and writes through the original stream. On
# Windows that stream defaults to cp1252 → printing '—'/'✓'/中文 (the preview /
# result output) raises UnicodeEncodeError and kills the child. Reconfiguring the
# underlying stream to UTF-8 first means colorama wraps an already-UTF-8 stream.
# (drag_window reads the pipe as UTF-8, so both ends agree.)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_HERE = Path(__file__).parent.resolve()
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from ifr_automation_v10 import (  # engine only — no telegram/apscheduler
    ConfigManager,
    IFRAutomation,
    PipelineOrchestrator,
    format_pipeline_result,
)


# Mirror of PipelineOrchestrator's default stage list (ifr_automation_v10.py).
# Kept here only to VALIDATE --stages input; the engine remains the SSOT for the
# order and for what each stage does. A stage added there and not here is simply
# not selectable from this CLI -- it never silently runs the wrong thing.
_VALID_STAGES = ["health_check", "ifr_sync", "version_mgmt", "ifc_transmittal",
                 "asbuilt_version", "sharepoint_sync", "deliverable"]


def _get_config() -> ConfigManager:
    return ConfigManager(config_path=_HERE / "config.json")


def _scan_projects():
    config = _get_config()
    automation = IFRAutomation(
        root_path=config.get("default_root_path", str(_HERE)),
        config=config, interactive=True,
    )
    by_region = automation.scan_projects()
    out = []
    for projs in by_region.values():
        out.extend(projs)
    out.sort(key=lambda p: p.project_name)
    return out


def _find(projects, keyword):
    kw = keyword.lower()
    return [p for p in projects if kw in p.project_name.lower()]


def _emit(status, **extra):
    """Final machine-readable marker for drag_window."""
    print("RESULT_JSON: " + json.dumps({"status": status, **extra},
                                       ensure_ascii=False, default=str))


def main(argv):
    args = [a for a in argv if a]
    execute = "--execute" in args
    do_list = "--list" in args
    refresh_only = "--refresh-approved" in args
    ifc_status = "--ifc-status" in args
    replot = "--replot-stale" in args
    only = None
    if "--only" in args:
        i = args.index("--only")
        only = set()
        for a in args[i + 1:]:
            if a.startswith("--"):
                break
            only.add(a.upper())
        args = args[:i] + args[i + 1 + len(only):]
    # --stages s1 s2 ... : everything after the flag until the next --flag.
    stages = None
    if "--stages" in args:
        i = args.index("--stages")
        stages = []
        for a in args[i + 1:]:
            if a.startswith("--"):
                break
            stages.append(a)
        bad = [s for s in stages if s not in _VALID_STAGES]
        if bad or not stages:
            print(f"[ERROR] --stages 取值不合法: {bad or '(空)'}; 可选: {' '.join(_VALID_STAGES)}")
            _emit("error", message="bad stages", bad=bad)
            return 2
        args = args[:i] + [a for a in args[i + 1 + len(stages):]]
    kw = " ".join(a for a in args if not a.startswith("--")).strip()

    try:
        projects = _scan_projects()
    except Exception as e:
        print(f"[ERROR] 扫描项目失败: {e}")
        _emit("error", stage="scan", message=str(e))
        return 1

    if do_list:
        for p in projects:
            print(p.project_name)
        _emit("ok", count=len(projects), names=[p.project_name for p in projects])
        return 0

    if not kw:
        print("[ERROR] 需要项目关键词 (或 --list)")
        _emit("error", message="no keyword")
        return 2

    matches = _find(projects, kw)
    if not matches:
        print(f"[ERROR] 没有项目匹配 '{kw}'。可用: " +
              ", ".join(p.project_name for p in projects[:20]))
        _emit("no_match", keyword=kw)
        return 2
    if len(matches) > 1:
        print(f"[ERROR] '{kw}' 匹配到多个项目，请更精确:")
        for p in matches:
            print("  - " + p.project_name)
        _emit("ambiguous", keyword=kw, matches=[p.project_name for p in matches])
        return 2

    project = matches[0]
    if "--sync-13" in args:
        try:
            from ifr_automation_v10 import sharepoint_reconcile
            r = sharepoint_reconcile(Path(project.project_path), execute=execute)
        except Exception as e:
            _emit("error", stage="sync_13", project=project.project_name, message=str(e))
            return 1
        for k in ("overwritten", "would_overwrite", "refused"):
            for x in r.get(k) or []:
                print(f"  [{k}] {x}")
        _emit("ok", project=project.project_name, dry_run=not execute, sync_13=r)
        return 0
    if replot:
        try:
            from ifr_automation_v10 import ApprovedIFCManager
            r = ApprovedIFCManager(Path(project.project_path), dry_run=not execute).replot_stale(only)
        except Exception as e:
            _emit("error", stage="replot_stale", project=project.project_name, message=str(e))
            return 1
        for k in ("replotted", "reconverted", "sharepoint", "skipped", "errors"):
            for x in r.get(k) or []:
                print(f"  [{k}] {x}")
        for d in r.get("diffs") or []:
            print(f"  [diff] {d['file']}: 每页变化% {d['pages']} {d.get('pages_count', '')}")
        _emit("error" if r.get("errors") else "ok", project=project.project_name,
              dry_run=not execute, **r)
        return 1 if r.get("errors") else 0
    if ifc_status:
        # Internal DLV only: rows whose doc has a current IFC PDF in 4. IFC(Client) (or an
        # Approved-to-IFC report) → Approved IFC + rev + date. No Sharepoint/client write.
        try:
            from ifr_automation_v10 import DeliverableManager, IFCTransmittalManager
            pp = Path(project.project_path)
            ifc_map = DeliverableManager(pp, dry_run=True).scan_ifc_folder()
            tm = IFCTransmittalManager(pp, config=_get_config(), dry_run=not execute)
            n = tm.update_deliverable_ifc_status(ifc_map, sync=False)
        except Exception as e:
            _emit("error", stage="ifc_status", project=project.project_name, message=str(e))
            return 1
        if ifc_map and not getattr(tm, "last_excel", None) and not getattr(tm, "last_error", None):
            print("  ⚠ 找不到 DLV (8. Deliverables / 3.Deliverables) — 本项目没写, 不是没变化")
            _emit("no_dlv", project=project.project_name, ifc_pdfs=len(ifc_map))
            return 0
        if getattr(tm, "last_error", None):
            _emit("error", stage="ifc_status", project=project.project_name,
                  message=tm.last_error)
            return 1
        for c in tm.last_changes:
            print(f"  {c['doc_id']}: {c['old_status'] or '(空)'} Rev{c['old_rev'] or '-'}"
                  f" → Approved IFC Rev{c['new_rev']}"
                  + (f"  (清掉 comment: {c['cleared_comment']})" if c['cleared_comment'] else ""))
        print(f"{'已写' if execute else '预览'} {n} 行 → "
              f"{getattr(tm, 'last_excel', None) and tm.last_excel.name}")
        _emit("ok", project=project.project_name, dry_run=not execute, updated=n,
              changes=tm.last_changes,
              excel=str(getattr(tm, "last_excel", "") or ""),
              ifc_pdfs=sorted(v.get("filename", "") for v in ifc_map.values()))
        return 0
    if refresh_only:
        try:
            cfg = _get_config()
            au = IFRAutomation(root_path=cfg.get("default_root_path", str(_HERE)),
                               config=cfg, dry_run=False, interactive=False)
            r = au.refresh_approved_to_latest(Path(project.project_path))
        except Exception as e:
            _emit("error", stage="refresh_approved", project=project.project_name, message=str(e))
            return 1
        for d in r.get("details", []):
            print("  " + d)
        _emit("ok", project=project.project_name, refresh_approved=r)
        return 0
    mode = "执行 (真实落盘, 可逆→Superseded)" if execute else "预览 (dry-run, 不动文件)"
    print("=" * 60)
    print(f"  IFR Pipeline — {mode}")
    print(f"  项目: {project.project_name}")
    if stages:
        print(f"  Stages: {' '.join(stages)}  (子集; 顺序由引擎写死, 此处只选不排)")
    print(f"  路径: {project.project_path}")
    print("=" * 60)
    sys.stdout.flush()

    try:
        pipeline = PipelineOrchestrator(_get_config(), dry_run=not execute,
                                        stages=stages)
        result = pipeline.run_pipeline(Path(project.project_path), project)
    except Exception as e:
        import traceback
        traceback.print_exc()
        _emit("error", stage="run", project=project.project_name, message=str(e))
        return 1

    try:
        # format_pipeline_result emits Telegram-HTML (<b>/<code>…); the GUI Text
        # pane renders no HTML, so strip tags at this GUI-only boundary.
        import re as _re
        print(_re.sub(r"</?[a-zA-Z][^>]*>", "", format_pipeline_result(result)))
    except Exception:
        print(json.dumps(result, ensure_ascii=False, default=str, indent=2))

    print("\n" + ("=" * 60))
    print("  ✓ 完成" + ("" if execute else " (预览，未改动文件)"))
    # `result` carries every stage's structured counts. Emitting it lets an
    # unattended caller (auto_drain patrol slot) build its own verdicts without
    # scraping the human-readable stdout above.
    _emit("ok", project=project.project_name, dry_run=not execute,
          stages=stages or _VALID_STAGES, result=result)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
