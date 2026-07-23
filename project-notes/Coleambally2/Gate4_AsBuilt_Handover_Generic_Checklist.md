# Gate 4 — AS BUILT / Handover Readiness Generic Checklist (Tech-DD proof)

Purpose: a reusable pre-submission gate so AS BUILT packs survive technical due-diligence
without being tripped by the kind of nit-picking EE raised on Coleambally #2. Derived from
the Cole2 July-2026 punchlist. Promote to SOP (cross-project) once validated.

## A. Stamp & title block (the automated layer — already solid on Cole2)
- [ ] Two-box stamp on **every sheet**: `AS BUILT` + `DRAWINGS TO BE PRINTED IN COLOUR`, aligned (≤20 pt), boxes match each other (same fill/weight).
- [ ] No leftover `FOR CONSTRUCTION` / `ISSUED FOR REVIEW` text.
- [ ] Revision table top row = AS BUILT with correct REV / DATE / DES / DRN / APP — **no blank personnel**.
- [ ] Revision number consistent between filename, title-block REV cell, and the deliverables register.

## B. Narrative / notes tense (the gap this punchlist exposed — NOT yet automated)
- [ ] **No future/imperative tense.** Scan every NOTE for: `SHALL`, `TO BE CONFIRMED ON SITE`, `TO BE COORDINATED` (future), `TO RETICULATE`, `PROPOSED`, `TBC`. AS BUILT = statement of installed fact (past/present).
- [ ] "Read in conjunction with supplier drawing / to be provided" style notes resolved or reworded.
- [ ] Reference-drawing cross-refs valid (each referenced dwg number exists and is current rev).

## C. Content reflects as-installed
- [ ] **No temporary construction elements** shown (laydown, site offices, temp buildings, construction compound).
- [ ] Base map = **latest as-built site layout**, not an early design base.
- [ ] Dimensions/coordinates = as-measured, not "to be confirmed."
- [ ] Any clash/encroachment flagged by client (e.g. water tank vs APZ) resolved and reflected.
- [ ] Missing sheets present (e.g. a "Notes — Sheet 1" that was dropped).

## D. Deliverable completeness (by ROW, not by artifact type)
- [ ] Every deliverables-list line has its artifact — drawings **and** reports, datasheets, certs, FAT/COC.
- [ ] Support documents collected: easements (executed deed + plan + title search), warranties, equipment datasheets, COCs, commissioning/test reports.
- [ ] "Manifest says present" ≠ present — physically confirm each file is in the pack.

## E. Filing hygiene
- [ ] Client-facing folder is **flat, PDF-only, current revision only**.
- [ ] Native tree clean: one loose current master + `Rev.N - AB\` snapshot + one `Superseded\` (no naming churn like both `Rev1_AB` and `RevE_AB`).
- [ ] Each punchlist item has traceable evidence in its own support folder.

## F. Automation opportunity (feed the spine, machine-first)
- The B-class (tense/narrative) and C-class (temp elements / stale base) defects are the recurring, nit-pick-magnet classes. Build a **DWG note-lint** (runs on the DWG via COM/ODA — PDF text is vectorised and NOT extractable) that flags future-tense tokens and known temp-layer names before any AS BUILT export. This converts a manual client-review round-trip into a pre-submission auto-gate.
