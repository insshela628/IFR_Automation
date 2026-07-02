# -*- coding: utf-8 -*-
"""Offline test for the shared register-membership oracle (register_membership.py).
No AutoCAD, no client folders, no Excel — pure logic. Mirrors the 5 collision-gate
cases proven in test_supersede_oracle.py plus filename-format normalization."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import register_membership as R

results = []


def check(name, got, want):
    ok = got == want
    print(f"[{'PASS' if ok else 'FAIL'}] {name}")
    if not ok:
        print(f"    got={got!r} want={want!r}")
    results.append(ok)


# ── title reconciliation (the membership oracle) ─────────────────────────────
# 1. SLD ≡ Single Line Diagram (acronym-of) → same drawing.
check("SLD acronym-reconciles Single Line Diagram",
      R.title_reconciles("SLD", "Single Line Diagram"), True)

# 2. CA-001 Trench Alignment ⊂ Trench Alignment Layout Plan (token containment).
check("Trench Alignment token-reconciles wordier register title",
      R.title_reconciles("Trench Alignment", "Trench Alignment Layout Plan"), True)

# 3. Trench Alignment vs Cable Route GA → genuinely different (NEITHER matches).
check("Trench Alignment does NOT reconcile Cable Route GA",
      R.title_reconciles("Trench Alignment", "Cable Route GA"), False)

# 4. Stray: Zeta Section matches neither register name.
check("Stray Zeta Section does NOT reconcile Alpha Plan",
      R.title_reconciles("Zeta Section", "Alpha Plan"), False)

# 5. Case/space/underscore drift → fuzzy reconciles.
check("case/space drift fuzzy-reconciles",
      R.title_reconciles("site  plan", "Site Plan"), True)

# match_register_title picks the right title / returns None for stray.
reg = ["Trench Alignment Layout Plan", "Cable Route GA"]
check("match picks Trench for Trench file",
      R.match_register_title("Trench Alignment", reg), "Trench Alignment Layout Plan")
check("match picks Cable Route for Cable file",
      R.match_register_title("Cable Route", reg), "Cable Route GA")
check("match returns None for stray",
      R.match_register_title("Zeta Section", reg), None)

# ── title_signature: same drawing → same sig; different → different ──────────
check("SLD sig == title_text-based sig for same file",
      R.title_signature("50023-EL-001 REV 0 SLD_AS BUILT", "50023-EL-001"), "SLD")
check("Single Line Diagram collapses to SINGLELINEDIAGRAM",
      R.title_signature("50023-EL-001 REV 1 Single Line Diagram_AS BUILT",
                        "50023-EL-001"), "SINGLELINEDIAGRAM")

# ── filename-format normalization (identity-preserving) ──────────────────────
check("REV 1 → Rev 1, AS BUILT re-cased",
      R.normalize_filename_format("50023-EL-001 REV 1 single line diagram_as built.pdf"),
      "50023-EL-001 Rev 1 single line diagram_AS BUILT.pdf")
check("rev.0 + double spaces collapse",
      R.normalize_filename_format("50023-CA-001  rev.0   Trench Alignment.pdf"),
      "50023-CA-001 Rev 0 Trench Alignment.pdf")
check("RevA letter rev preserved",
      R.normalize_filename_format("GG31-C-PLN-006_RevA.dwg"),
      "GG31-C-PLN-006 Rev A.dwg")
# description words are NEVER altered by format norm (safety).
check("description words untouched (no canonicalization)",
      R.normalize_filename_format("50023-EL-001 Rev 1 SLD_AS BUILT.pdf"),
      "50023-EL-001 Rev 1 SLD_AS BUILT.pdf")

print()
print("ALL PASS" if all(results) else "SOME FAILED")
sys.exit(0 if all(results) else 1)
