# -*- coding: utf-8 -*-
"""Structured QA fault layer (①) — PARSE-ONLY, engine-independent.

Turns the flat QA warning STRINGS the engine already produces into structured
fault objects {code, page, severity, title, message, auto_tried, fix_checklist},
by matching against the data-driven playbook (`qa_playbook.json`, the SSOT).

CRITICAL CONTRACT — this must NOT change the gold standard:
  * It does NOT run conversion, QA, AutoCAD, or any geometry. It only reads
    strings that are ALREADY decided (PASS/FAIL is set upstream, untouched).
  * It NEVER feeds anything back into the PASS/FAIL / escalate / retry logic.
    It is a pure DOWNSTREAM presentation layer for the bot (③) / a future UI.
  * The engine (`ifr_automation_v10.py`) imports nothing from here and is not
    modified. Both the batch runner and the bot may import this freely.

The fault CODE → remediation mapping is NOT hardcoded here — it is loaded from
`qa_playbook.json` so a new fault is added by editing data only (铁律 2/3).
"""
import re
import json
from pathlib import Path

_PLAYBOOK_PATH = Path(__file__).resolve().parent / "qa_playbook.json"
_PAGE_RE = re.compile(r"Page\s+(\d+)", re.IGNORECASE)

_playbook_cache = None


def _playbook():
    global _playbook_cache
    if _playbook_cache is None:
        _playbook_cache = json.loads(_PLAYBOOK_PATH.read_text(encoding="utf-8"))
    return _playbook_cache


def classify(warning: str) -> str:
    """Map a raw QA warning string to a fault code (registry order, first hit).
    Returns 'UNKNOWN' if nothing matches."""
    faults = _playbook().get("faults", {})
    for code, spec in faults.items():
        for sub in spec.get("match", []):
            if sub in warning:
                return code
    return "UNKNOWN"


def build_fault(warning: str, doc_id: str = None) -> dict:
    """Parse one warning string into a structured fault object. Pure."""
    code = classify(warning)
    if code == "UNKNOWN":
        spec = _playbook().get("unknown", {})
    else:
        spec = _playbook()["faults"][code]
    page_m = _PAGE_RE.search(warning)
    return {
        "doc_id": doc_id,
        "code": code,
        "page": int(page_m.group(1)) if page_m else None,
        "severity": spec.get("severity", "提示"),
        "title": spec.get("title", code),
        "message": warning,
        "auto_tried": list(spec.get("auto_tried", [])),
        "fix_checklist": list(spec.get("fix_checklist", [])),
    }


# Severity ordering for sorting (most urgent first).
_SEV_ORDER = {"挡出图": 0, "基建": 1, "可重试": 2, "提示": 3}


def faults_from_result(result: dict) -> list:
    """Extract structured faults from a conversion result dict WITHOUT touching
    it. Reads the strings the engine already populated:
      - result['qa_warnings']      (PDF-level QA)
      - result['tb_qa_warnings']   (title-block QA, non-blocking)
      - result['errors']           (escalation / conversion-failure strings)
    De-dups by (code, page) — clean qa_warnings/tb_qa first, then errors, so an
    errors entry that merely RE-WRAPS a qa_warning (e.g. 'QA失败(需人工介入): …')
    is dropped while a genuinely error-only fault (timeout, SaveAs fail) is kept.
    Sorts by severity then page."""
    doc_id = result.get("doc_id")
    seen, faults = set(), []
    # Accept both shapes: the engine result (qa_warnings/tb_qa_warnings) and the
    # bot-mapped result (post_qa). Clean warning lists first, then errors.
    for key in ("qa_warnings", "post_qa", "tb_qa_warnings", "errors"):
        for w in (result.get(key) or []):
            if not isinstance(w, str):
                continue
            f = build_fault(w, doc_id=doc_id)
            sig = (f["code"], f["page"])
            if sig in seen:
                continue
            seen.add(sig)
            faults.append(f)
    faults.sort(key=lambda f: (_SEV_ORDER.get(f["severity"], 9),
                               f["page"] if f["page"] is not None else 0))
    return faults
