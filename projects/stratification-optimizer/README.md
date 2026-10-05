# SCOPE — Stratification, Configuration, Optimization, and Profiling Engine

SCOPE is a Python-based desktop tool I built to automate stratified sampling,
recruitment planning, and targeted case-swap workflows for standardized test
development (e.g., BASC-4 Narrowbands, with CELF-6 and WISC-6 profiles in
progress). It replaces ad hoc, spreadsheet-driven stratification processes
with a reproducible, profile-driven engine and desktop GUI.

## What it does

- **Sample Optimization** — uses CP-SAT constraint solving to select a
  stratified sample from a candidate pool that meets demographic, clinical,
  and OSEP targets within specified tolerances.
- **Recruitment Planning** — reconciles research-plan targets into integer
  recruitment Ns and reports current usable cases vs. remaining need.
- **Targeted Case Swap / OSEP / Clinical Adjustments** — validates and applies
  approved case replacements through an audited, lock-preserving swap engine
  without re-running the full optimizer.
- **Profile-driven reporting** — generates one-way/two-way/three-way
  demographic tables, OSEP/clinical workbooks, TEID lists, and QA output from
  versioned Test Profiles rather than hard-coded assessment logic.

## Why

Standardization studies for psychological/educational assessments need
representative, demographically balanced samples that satisfy many
overlapping quotas (age, sex, race/ethnicity, parent education, region,
clinical status, etc.). Manually balancing these in spreadsheets is slow and
error-prone. SCOPE formalizes the rules into versioned profiles and solves the
selection problem with a constraint solver, producing auditable outputs
(swap logs, run manifests) suitable for data-science/psychometric review.

## Contents of this folder

- `source/` — trimmed source tree from the SCOPE project: `scope/` (engine and
  legacy CP-SAT implementation), `gui/` (desktop app and profile manager),
  `scripts/` (CLI entry point), `profiles/`, `misc/` (build scripts and
  internal docs), and `tests/`. Build artifacts, virtual environments, and
  large generated/cache folders are intentionally excluded.
- `data/` — example outputs from a real stratification run on the BASC-4
  Narrowbands **Mood** scale (post-stratification pool, selected cases,
  demographic/OSEP tables, swap tables, and an updated TEID list), included
  here as a sample of the tool's real-world output.

## Run from source

```
python gui/app.py            # desktop GUI
python scripts/run_scope.py  # CLI / scripted workflow
```

See `source/misc/README.md` for full architecture and workflow details.
