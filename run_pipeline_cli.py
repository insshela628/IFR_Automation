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

Output: human-readable progress to stdout; last line is a machine marker
  RESULT_JSON: {...}   (parsed by drag_window)
Exit: 0 ok, 2 no match / ambiguous, 1 error.
"""
import sys
import json
from pathlib import Path

_HERE = Path(__file__).parent.resolve()
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from ifr_automation_v10 import (  # engine only — no telegram/apscheduler
    ConfigManager,
    IFRAutomation,
    PipelineOrchestrator,
    format_pipeline_result,
)


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
    print("RESULT_JSON: " + json.dumps({"status": status, **extra}, ensure_ascii=False))


def main(argv):
    args = [a for a in argv if a]
    execute = "--execute" in args
    do_list = "--list" in args
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
    mode = "执行 (真实落盘, 可逆→Superseded)" if execute else "预览 (dry-run, 不动文件)"
    print("=" * 60)
    print(f"  IFR Pipeline — {mode}")
    print(f"  项目: {project.project_name}")
    print(f"  路径: {project.project_path}")
    print("=" * 60)
    sys.stdout.flush()

    try:
        pipeline = PipelineOrchestrator(_get_config(), dry_run=not execute)
        result = pipeline.run_pipeline(Path(project.project_path), project)
    except Exception as e:
        import traceback
        traceback.print_exc()
        _emit("error", stage="run", project=project.project_name, message=str(e))
        return 1

    try:
        print(format_pipeline_result(result))
    except Exception:
        print(json.dumps(result, ensure_ascii=False, default=str, indent=2))

    print("\n" + ("=" * 60))
    print("  ✓ 完成" + ("" if execute else " (预览，未改动文件)"))
    _emit("ok", project=project.project_name, dry_run=not execute)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
