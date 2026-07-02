# -*- coding: utf-8 -*-
"""Generate / refresh the QA golden-master snapshot — EVERY MAPPABLE PROJECT.

Characterization test fixture (golden master) that PINS the current,
user-approved QA behaviour so later changes (structured fault objects, playbook
wiring, stamp-overlap / title-block detectors) cannot silently regress it.

KEY DESIGN — independence from the thing it guards:
  * Calls the REAL `AsBuiltManager._qa_validate_ab_pdf` (same source as the
    engine → zero fidelity drift), via `__new__` so NO project config / AutoCAD
    is needed. Fully OFFLINE, read-only.
  * Treats QA as a black box: freezes {key -> warnings + verdict}. The verdict
    (PASS / ESCALATE / RETRYABLE) is recomputed from the SAME
    `_QA_ESCALATE_KEYWORDS` the engine uses, so a categorisation flip is caught.
  * Intentional FAILs are part of the baseline (a deliverable that currently
    overlaps must ESCALATE; a clean sheet must PASS). The gate guards against
    CODE-induced drift, not against known-defective deliverables.

CROSS-PROJECT COVERAGE (user: "all project ... 目前 Dropbox EPC 路径下可以映射到的所有
项目" — every project reachable under Project (EPC), NOT only those with a
populated "5. As Built" folder):
  * Auto-discovers EVERY project by its drawings tree ("1. Drawings" under
    Project (EPC), excluding 6.Archived / templates), regardless of the
    Design/Engineering vs Engineering/Design layout or nesting (1.Completed/…).
  * Registers every project — including ones with ZERO As Built deliverables yet
    (recorded in coverage.json) — so a project that later gains deliverables is
    picked up automatically and none is silently missed.
  * Within each project's drawings tree, snapshots real AS BUILT deliverable PDFs
    (squished-"ASBUILT" name test; excludes -AB/superseded/☆/_FIXED_REVIEW), and
    EXCLUDES duplicate-copy areas (Share to/with Client, Handover, Submit/
    Submission, Feedback, Client Template, IFR/IFC, Native, archive/backup/old).
  * Keyed by the path relative to Project (EPC) → globally unique, human-readable,
    collision-free across projects.
  * Online-only Dropbox placeholders are force-materialised before scanning so the
    corpus is stable regardless of local cache state.

Run ONLY to (re)baseline AFTER deliberately accepting a new golden state.
Never auto-run. `python golden/gen_qa_golden.py`
"""
import sys, json
from pathlib import Path

V6 = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(V6))

PROJECTS_ROOT = Path(
    r"C:\Users\jilin\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)")

# Top-level entries never scanned (archived, templates, org charts).
_SKIP_TOP = ("6.Archived", "### PROJECT TEMPLATE", "### GGE Org. Chart")

# Duplicate-copy / non-deliverable path segments (case-insensitive substring on
# any path part BELOW the drawings root) — keep only the canonical As Built copy.
_DUP_SEG = ("superseded", "share to client", "shared with client", "handover",
            "submit", "submission", "feedback", "client template",
            "ifr", "ifc", "native", "old", "archive", "backup", "wip", "draft")

GOLDEN = Path(__file__).resolve().parent / "qa_verdicts.json"
COVERAGE = Path(__file__).resolve().parent / "coverage.json"

import re as _re
_DOC_ID = _re.compile(r"([A-Z0-9]{2,}-[A-Z]{1,4}-[A-Z0-9]+(?:-[0-9]+)?)")


def doc_id_of(name: str) -> str:
    m = _DOC_ID.search(name.upper())
    return m.group(1) if m else name


def _is_deliverable(p: Path) -> bool:
    """A real AS BUILT deliverable PDF (any spacing: 'AS BUILT'/'AsBuilt'), NOT an
    old '-AB' variant, a ☆ marker, or a review artifact."""
    if p.suffix.lower() != ".pdf":
        return False
    st = p.stem
    if "☆" in st:
        return False
    up = st.upper()
    if "FIXED_REVIEW" in up or up.endswith("_REVIEW"):
        return False
    return "ASBUILT" in up.replace(" ", "").replace("_", "").replace("-", "")


def _project_label(drawings_dir: Path) -> str:
    """'Region/Project' (or nested) label from a '1. Drawings' path, stripping the
    trailing [Design/Engineering | Engineering/Design] wrapper."""
    rel = list(drawings_dir.relative_to(PROJECTS_ROOT).parts)
    cut = rel.index("1. Drawings")
    head = rel[:cut]
    while head and head[-1] in ("Design", "Engineering"):
        head.pop()
    return "/".join(head)


def discover_projects():
    """{project_label: [drawings_dir, ...]} for every mappable project."""
    projects = {}
    for dw in PROJECTS_ROOT.glob("*/**/1. Drawings"):
        if not dw.is_dir():
            continue
        parts = dw.relative_to(PROJECTS_ROOT).parts
        if parts and parts[0] in _SKIP_TOP:
            continue
        projects.setdefault(_project_label(dw), []).append(dw)
    return projects


def _deliverables_in(drawings_dir: Path):
    out = []
    for p in drawings_dir.rglob("*.pdf"):
        rel_parts = [x.lower() for x in p.relative_to(drawings_dir).parts]
        if any(any(d in seg for d in _DUP_SEG) for seg in rel_parts):
            continue
        if _is_deliverable(p):
            out.append(p)
    return out


def _materialise(p: Path) -> bool:
    """Force a Dropbox online-only placeholder to fully hydrate (read all bytes)."""
    try:
        with open(p, "rb") as fh:
            fh.read()
        return True
    except Exception:
        return False


def make_manager():
    from ifr_automation_v10 import AsBuiltManager
    return AsBuiltManager.__new__(AsBuiltManager), AsBuiltManager


def verdict_of(warnings, escalate_keywords):
    if not warnings:
        return "PASS"
    if any(kw in w for w in warnings for kw in escalate_keywords):
        return "ESCALATE"
    return "RETRYABLE"


def collect():
    """{key -> {doc_id,size,warnings,verdict}} over EVERY project's deliverables.
    Also writes coverage.json (all projects incl. zero-deliverable) as a side
    effect record — the snapshot itself only holds files that exist."""
    import fitz
    mgr, cls = make_manager()
    escalate_kw = cls._QA_ESCALATE_KEYWORDS
    projects = discover_projects()
    snap = {}
    coverage = {}
    for proj in sorted(projects):
        seen = set()
        for dw in projects[proj]:
            for p in _deliverables_in(dw):
                if p in seen:
                    continue
                seen.add(p)
        coverage[proj] = len(seen)
        for p in sorted(seen):
            key = p.relative_to(PROJECTS_ROOT).as_posix()
            try:
                fitz.open(str(p)).close()
            except Exception:
                if not _materialise(p):
                    snap[key] = {"doc_id": doc_id_of(p.name), "size": -1,
                                 "warnings": ["ONLINE_ONLY_UNREADABLE"],
                                 "verdict": "SKIPPED"}
                    continue
            warns = mgr._qa_validate_ab_pdf(p, doc_id_of(p.name), expected_pages=None)
            snap[key] = {
                "doc_id": doc_id_of(p.name),
                "size": p.stat().st_size,
                "warnings": list(warns),
                "verdict": verdict_of(warns, escalate_kw),
            }
    try:
        COVERAGE.write_text(json.dumps(coverage, ensure_ascii=False, indent=2),
                            encoding="utf-8")
    except Exception:
        pass
    return snap


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    snap = collect()
    GOLDEN.write_text(json.dumps(snap, ensure_ascii=False, indent=2),
                      encoding="utf-8")
    n = len(snap)
    npass = sum(1 for v in snap.values() if v["verdict"] == "PASS")
    nesc = sum(1 for v in snap.values() if v["verdict"] == "ESCALATE")
    nret = sum(1 for v in snap.values() if v["verdict"] == "RETRYABLE")
    nskip = sum(1 for v in snap.values() if v["verdict"] == "SKIPPED")
    coverage = json.loads(COVERAGE.read_text(encoding="utf-8")) if COVERAGE.exists() else {}
    print(f"GOLDEN written: {GOLDEN}")
    print(f"  {n} deliverable PDFs across {len(coverage)} projects  |  "
          f"PASS={npass}  ESCALATE={nesc}  RETRYABLE={nret}  SKIPPED={nskip}")
    print("  coverage (every mappable project; 0 = no As Built yet):")
    for pr in sorted(coverage):
        print(f"    {coverage[pr]:3}  {pr}")
    print("  non-PASS:")
    for name, v in snap.items():
        if v["verdict"] != "PASS":
            print(f"  [{v['verdict']}] {name}")
            for w in v["warnings"]:
                print(f"       - {w}")
