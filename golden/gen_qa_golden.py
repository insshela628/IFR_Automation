# -*- coding: utf-8 -*-
"""Generate / refresh the QA golden-master snapshot.

Characterization test fixture (golden master) that PINS the current,
user-approved QA behaviour so later changes (e.g. structured fault objects,
playbook wiring) cannot silently regress it.

KEY DESIGN — independence from the thing it guards:
  * It calls the REAL `AsBuiltManager._qa_validate_ab_pdf` (same source as the
    engine → zero fidelity drift), but via `__new__` so NO project config /
    AutoCAD is needed. Fully OFFLINE, read-only.
  * It treats QA as a black box: freezes {pdf -> warnings + verdict}. The
    verdict (PASS / ESCALATE / RETRYABLE) is recomputed here from the SAME
    `_QA_ESCALATE_KEYWORDS` the engine uses, so a flip in categorisation is
    also caught.
  * Keyed by PDF FILENAME, not doc_id (two C-PLN-002 scopes share a doc_id).
  * Intentional FAILs are part of the baseline (E-BLD-001 must escalate,
    a clean sheet must PASS) — that is what locks "25/27 + 2 known manual".

Run ONLY to (re)baseline AFTER you have deliberately accepted a new golden
state. Never auto-run. `python golden/gen_qa_golden.py`
"""
import sys, json, hashlib
from pathlib import Path

V6 = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(V6))

# Output PDFs under review (the live, user-approved deliverables).
PDF_DIR = Path(
    r"C:\Users\jilin\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\1.NSW\NSW 153 - Coleambally #2\Design\Engineering"
    r"\1. Drawings\5. As Built\3. As Built Client")

GOLDEN = Path(__file__).resolve().parent / "qa_verdicts.json"

_DOC_ID = __import__("re").compile(r"(NSW153-[A-Z]+-[A-Z0-9]+(?:-[0-9]+)?)")


def doc_id_of(name: str) -> str:
    m = _DOC_ID.search(name)
    return m.group(1) if m else name


def make_manager():
    """A no-__init__ AsBuiltManager: only the pure QA methods are exercised."""
    from ifr_automation_v10 import AsBuiltManager
    return AsBuiltManager.__new__(AsBuiltManager), AsBuiltManager


def verdict_of(warnings, escalate_keywords):
    if not warnings:
        return "PASS"
    if any(kw in w for w in warnings for kw in escalate_keywords):
        return "ESCALATE"
    return "RETRYABLE"


def collect():
    mgr, cls = make_manager()
    escalate_kw = cls._QA_ESCALATE_KEYWORDS
    pdfs = sorted(p for p in PDF_DIR.glob("*AS BUILT.pdf")
                  if "☆" not in p.name)
    snap = {}
    for p in pdfs:
        warns = mgr._qa_validate_ab_pdf(p, doc_id_of(p.name), expected_pages=None)
        snap[p.name] = {
            "doc_id": doc_id_of(p.name),
            "size": p.stat().st_size,
            "warnings": list(warns),
            "verdict": verdict_of(warns, escalate_kw),
        }
    return snap


if __name__ == "__main__":
    snap = collect()
    GOLDEN.write_text(json.dumps(snap, ensure_ascii=False, indent=2),
                      encoding="utf-8")
    n = len(snap)
    npass = sum(1 for v in snap.values() if v["verdict"] == "PASS")
    nesc = sum(1 for v in snap.values() if v["verdict"] == "ESCALATE")
    nret = sum(1 for v in snap.values() if v["verdict"] == "RETRYABLE")
    print(f"GOLDEN written: {GOLDEN}")
    print(f"  {n} PDFs  |  PASS={npass}  ESCALATE={nesc}  RETRYABLE={nret}")
    for name, v in snap.items():
        if v["verdict"] != "PASS":
            print(f"  [{v['verdict']}] {name}")
            for w in v["warnings"]:
                print(f"       - {w}")
