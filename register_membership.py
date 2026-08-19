# -*- coding: utf-8 -*-
"""Cross-track register-membership oracle (SHARED SSOT).

The client *Deliverables List* is the TITLE authority every version/collision
gate consults to adjudicate same-doc-id / different-title clashes WITHOUT a
synonym dictionary — the doc-id is the join key, the register maps it to its
canonical Doc Name(s):

  • a doc-id the register lists under ONE title → every file under it is that one
    drawing, so a filename that only drifted (SLD vs Single Line Diagram) is
    reconciled and versioned together, NOT frozen for a human;
  • a doc-id the register lists under SEVERAL titles (50023-CA-001 = Trench
    Alignment + Cable Route GA) → genuinely several drawings, each versioned
    independently;
  • a title matching NO register name → flagged stray, never merged/renamed
    (guards the normalize-Rule-3 data-loss trap);
  • no register for the doc-id → caller falls back to a conservative wholesale
    defer.

This module is the SSOT for that logic so the four tracks stay identical:
  - AS BUILT deliverable PDFs   (AsBuiltManager.supersede_ab_deliverables)
  - IFR/IFC deliverable PDFs     (VersionManager)
  - Native DWG working files     (NativeVersionManager)
  - the external bot fault layer

Pure/offline: no AutoCAD, no COM. The register loader is a SOFT dependency on
the Doc-Control keystone parser (deliverable_list.parse_deliverable_list) — it
returns an empty dict on any failure so callers degrade gracefully.
"""
import re
from pathlib import Path
from difflib import SequenceMatcher
from typing import Dict, List, Optional

# ── title-signature extraction (mirrors ifr_automation_v10._ab_title_*) ──────
# NB: the rev alternative accepts a LETTER rev ('_RevA', 'Rev B') as well as a
# numeric one. Letter revs are the IFR/draft scheme, so a Native doc-folder holds
# both — while this regex only stripped digits, '..._RevA.dwg' kept 'REVA' inside
# its title signature and looked like a DIFFERENT DRAWING from '..._Rev1_IFC.dwg'.
# That fake collision made the gate declare "N distinct titles share doc-id" and
# defer the whole folder to manual forever (21 files on GG-31). Keep it in step
# with normalize_filename_format's _RE_REV, which has always accepted letters.
_RE_AB_SIG_STRIP = re.compile(
    r'(?:^|[\s_\-(])(?:rev\.?\s*(?:\d+(?:\.\d+)?|[A-Za-z])|as[ _]?built'
    r'|ab(?:-exp)?|ifc|sheet\s*\d+|exp)(?=$|[\s_\-).])', re.IGNORECASE)


def title_text(stem: str, doc_id: Optional[str]) -> str:
    """Human drawing-title portion of a filename stem — doc-id + rev + AB/IFC/exp/
    sheet markers removed, separators → single spaces, WORD BOUNDARIES KEPT (for
    token reconciliation against a register title)."""
    sig = stem
    if doc_id:
        sig = re.sub(re.escape(doc_id), ' ', sig, flags=re.IGNORECASE)
    prev = None
    while prev != sig:                       # collapse adjacent markers
        prev = sig
        sig = _RE_AB_SIG_STRIP.sub(' ', sig)
    return re.sub(r'\s+', ' ', re.sub(r'[_\-]', ' ', sig)).strip()


def title_signature(stem: str, doc_id: Optional[str]) -> str:
    """Collapsed, punctuation-free, upper-cased title signature. Two files sharing
    a signature are the SAME drawing at different revs; differing signatures are
    kept apart (so two different drawings wrongly sharing a doc-id never merge)."""
    return re.sub(r'[^A-Za-z0-9]', '', title_text(stem, doc_id)).upper()


# Words carrying no drawing-identity (mirrors the Doc-Control keystone _TITLE_STOP
# so both sides tokenize a title the same way). IDENTITY tokens = len≥3 words left.
_TITLE_STOP = frozenset({
    "the", "and", "for", "plan", "layout", "design", "rev", "as", "built", "ab",
    "ifc", "diagram", "box", "general", "arrangement", "power", "station",
    "nawma", "bess", "new", "existing", "drawing", "sheet", "details", "detail"})


def title_tokens(s: Optional[str]) -> frozenset:
    return frozenset(t for t in re.findall(r"[a-z0-9]+", (s or "").lower())
                     if len(t) >= 3 and t not in _TITLE_STOP)


def norm_title(s: Optional[str]) -> str:
    """Punctuation/space-insensitive, upper-cased title key (same shape as
    title_signature, so the two are directly comparable)."""
    return re.sub(r'[^A-Za-z0-9]', '', s or '').upper()


def initials(s: Optional[str]) -> str:
    """First letter of each word → acronym key ('Single Line Diagram'→'SLD')."""
    return ''.join(w[0] for w in re.findall(r'[A-Za-z0-9]+', s or '')).upper()


def title_reconciles(file_title: str, reg_title: str) -> bool:
    """True iff a file's title denotes the SAME drawing as a register title.
    `file_title` may be the spaced title_text OR a collapsed signature. Same-
    drawing is decided (all auto-resolved, NO human step) by ANY of:
      • fuzzy ≥ 0.85 on the normalized text (case/space/underscore drift), OR
      • one is the ACRONYM of the other (SLD ≡ Single Line Diagram), OR
      • identity-token containment / overlap — the file's significant words are a
        subset of the register title's (or vice-versa), or Jaccard ≥ 0.6 (real
        registers are wordier: 'Trench Alignment' vs 'Trench Alignment Layout
        Plan', so pure-fuzzy under-matches).
    A genuinely different drawing sharing a doc-id reconciles to NEITHER register
    title → the caller flags it instead of collapsing it."""
    a, b = norm_title(file_title), norm_title(reg_title)
    if not a or not b:
        return False
    if a == b:
        return True
    if SequenceMatcher(None, a, b).ratio() >= 0.85:
        return True
    # acronym-of, either direction (a collapsed side has no separators, so compare
    # the SHORT normalized form to the LONG side's word-initials).
    if len(a) <= len(b):
        if len(a) >= 2 and a == initials(reg_title):
            return True
    elif len(b) >= 2 and b == initials(file_title):
        return True
    ta, tb = title_tokens(file_title), title_tokens(reg_title)
    if ta and tb and (ta <= tb or tb <= ta
                      or len(ta & tb) / len(ta | tb) >= 0.6):
        return True
    return False


def match_register_title(file_title: str, reg_titles: List[str]) -> Optional[str]:
    """First register title that reconciles with `file_title`, else None (stray)."""
    return next((t for t in (reg_titles or [])
                 if title_reconciles(file_title, t)), None)


# ── doc-id extraction (mirrors ifr_automation_v10._RE_DOC_ID_* / _extract_doc_id
#    _standalone; the engine passes its own richer extractor, v5 uses this) ────
_RE_DOC_ID_GG = re.compile(r'(GG\d{2}-[A-Z]-[A-Z]{3}-\d{3})', re.IGNORECASE)
_RE_DOC_ID_LMS = re.compile(r'(\d{5}-[A-Z]{2}-\d{3})', re.IGNORECASE)
_RE_DOC_ID_TSF = re.compile(r'(TSF-[A-Z]{2}-[A-Z]{3}-\w+-\d{2})', re.IGNORECASE)
_RE_DOC_ID_GENERIC = re.compile(r'([A-Z0-9][\w]+-[A-Z]+-[A-Z]*-?\d{3})', re.IGNORECASE)
_DOC_ID_PATS = (_RE_DOC_ID_GG, _RE_DOC_ID_LMS, _RE_DOC_ID_TSF, _RE_DOC_ID_GENERIC)


def extract_doc_id(filename: str) -> Optional[str]:
    """Structured doc-id from a filename, or None. Match-at-start first (standard
    naming), then search-anywhere (inaccurate names — FILE-NO rule). NO greedy
    'everything before _Rev' fallback: a None answer means 'no clean doc-id', so
    the caller groups by base_name as today (safe)."""
    stem = Path(filename).stem
    for pat in _DOC_ID_PATS:
        m = pat.match(stem)
        if m:
            return m.group(1)
    for pat in _DOC_ID_PATS:
        m = pat.search(stem)
        if m:
            return m.group(1)
    return None


def canonical_group_key(base_name: str, filename: str,
                        reg_titles: Optional[List[str]],
                        doc_id: Optional[str] = None):
    """Register-canonical grouping key for a versioning/collision gate. Returns
    (key, matched_title, is_stray):
      • no register for this doc-id (reg_titles falsy) → (base_name, None, False)
        — BYTE-IDENTICAL to today's base_name grouping (zero-regression default);
      • file title reconciles to a register title → ('<DOCID>::<TITLE><ext>',
        title, False) — so a re-worded revision (SLD vs Single Line Diagram) of
        the same drawing shares a key and the older one supersedes (fixes the PDF/
        Native UNDER-merge);
      • register present but NO title matches → (base_name, None, True) — STRAY:
        keep the file in its own group (never merged/moved), surface the flag.
    Two different register titles under one doc-id (CA-001 Trench vs Cable Route)
    yield two different keys → they stay split (the collision is preserved)."""
    if not reg_titles:
        return base_name, None, False
    ext = Path(filename).suffix
    ftext = title_text(Path(filename).stem, doc_id)
    matched = match_register_title(ftext, reg_titles)
    if matched is None:
        return base_name, None, True
    did = (doc_id or '').upper()
    return f"{did}::{norm_title(matched)}{ext}", matched, False


# ── register loading (soft dep on the Doc-Control keystone parser) ───────────
_REGISTER_CACHE: Dict[str, Dict[str, List[str]]] = {}
# Folders (relative to an ancestor) that can hold the client Deliverables List.
# '8. Deliverables' / '3.Deliverables' matter because the CURRENT list often lives
# there rather than beside the AS BUILT PDFs — GG-31's rev6.3 does.
_AB_DIR_NAMES = ('5. As Built', '6.AS Built', '6. As Built', '5.As Built',
                 '8. Deliverables', '8.Deliverables',
                 'Design/Engineering/8. Deliverables',
                 '3. IFR(Client)/3.Deliverables')
# Real lists are named '… Deliverables *Document* List …', so a literal
# '*Deliverables List*' glob matched nothing and the title oracle silently
# degraded to "no register" on every GG project.
_LIST_GLOBS = ('*Deliverable*List*.xlsx', '*DLV*.xlsx')


def _rank_list(p: Path):
    nm = p.name.lower()
    try:
        mt = p.stat().st_mtime
    except OSError:
        mt = 0.0
    return (('comment' in nm or 'ace' in nm), mt)


def find_deliverables_list(start: Path) -> Optional[Path]:
    """Walk UP from `start` looking for the client *Deliverables List*.xlsx —
    checked in each ancestor directly AND inside its `5. As Built/` (the engine's
    canonical home). Returns the best-ranked candidate (comment/ace-tagged +
    newest wins), or None."""
    start = Path(start).resolve()
    ancestors = [start] + list(start.parents)
    for anc in ancestors[:9]:
        cands: List[Path] = []
        for base in (anc, *[anc / d for d in _AB_DIR_NAMES]):
            if base.is_dir():
                for g in _LIST_GLOBS:
                    cands += [c for c in base.glob(g)
                              if not c.name.startswith('~$')]
                cands = list(dict.fromkeys(cands))   # globs overlap
        if cands:
            return max(cands, key=_rank_list)
    return None


def _parse_register(list_path: Path) -> Dict[str, List[str]]:
    """{DOC_ID_UPPER: [doc_title, ...]} via the Doc-Control keystone parser."""
    import sys as _sys
    ea_root = Path(__file__).resolve().parents[2] / "L02-knowledge" / "energy-agent"
    if ea_root.is_dir() and str(ea_root) not in _sys.path:
        _sys.path.insert(0, str(ea_root))
    from modules.review.deliverable_list import parse_deliverable_list
    reg: Dict[str, List[str]] = {}
    for r in parse_deliverable_list(str(list_path)):
        did = (r.get('doc_number') or '').upper()
        title = (r.get('doc_title') or '').strip()
        if did and title:
            reg.setdefault(did, [])
            if title not in reg[did]:
                reg[did].append(title)
    return reg


def load_register_titles(start: Path) -> Dict[str, List[str]]:
    """{DOC_ID_UPPER: [title, ...]} from the nearest client *Deliverables List*
    above `start`. Cached per resolved list path. SOFT dependency: empty dict if
    the list or parser is absent (caller then falls back to conservative defer)."""
    list_path = find_deliverables_list(Path(start))
    if not list_path:
        return {}
    key = str(list_path.resolve())
    if key in _REGISTER_CACHE:
        return _REGISTER_CACHE[key]
    try:
        reg = _parse_register(list_path)
    except Exception as e:
        print(f"    (客户清单标题权威不可用，碰撞闸降级为保守整体挂起: {e})")
        reg = {}
    _REGISTER_CACHE[key] = reg
    return reg


# ── filename-format normalization (identity-preserving; safe everywhere) ─────
_RE_MULTISPACE = re.compile(r'[ \t]+')
# leading sep (^ / space / _ / -) captured so it normalizes to a single space;
# trailing lookahead stops at a non-alnum so 'Rev1' and 'RevA' both catch.
_RE_REV = re.compile(r'(^|[ _\-])R(?:ev)?\.?\s*([0-9]+[A-Za-z]?|[A-Za-z])(?![A-Za-z0-9])',
                     re.IGNORECASE)
_RE_AB_SUFFIX = re.compile(r'[ _\-]*as[ _\-]?built\b', re.IGNORECASE)


def normalize_filename_format(name: str) -> str:
    """IDENTITY-PRESERVING filename tidy — case/space/underscore/Rev only. Does
    NOT touch the description words, so it is safe to apply everywhere (no chance
    of collapsing a genuinely-different drawing). Description→register-canonical
    is a SEPARATE, reconcile-gated step (see canonicalize_description).

    Rules: 'REV 1'/'rev.1'/'Rev1' → 'Rev 1'; '_AS BUILT'/'as built' → '_AS BUILT'
    (single, uppercased); collapse runs of spaces; trim. Extension preserved."""
    p = Path(name)
    stem, ext = p.stem, p.suffix
    had_ab = bool(_RE_AB_SUFFIX.search(stem))
    stem = _RE_AB_SUFFIX.sub('', stem)                       # strip, re-add once
    stem = _RE_REV.sub(
        lambda m: (' ' if m.group(1) else '') + f"Rev {m.group(2).upper()}", stem)
    stem = _RE_MULTISPACE.sub(' ', stem).strip(' _-')
    if had_ab:
        stem = f"{stem}_AS BUILT"
    return f"{stem}{ext}"
