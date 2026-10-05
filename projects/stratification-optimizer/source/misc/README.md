# SCOPE

Sample Configuration, Optimization, and Profiling Engine.

SCOPE has one execution path: both the GUI and `scripts/run_scope.py` call
`scope.engine.run`. The existing SOE preparation and CP-SAT implementation is
bundled under `scope/legacy` as the BASC regression baseline; assessment choices
belong in versioned Test Profiles, not separate GUI or optimizer copies.

The main desktop shell uses five clickable pages. Readiness still controls
actions, but users may visit Output or Results early to see an explanation or
preview:

`Sample Data & Task → Operation Setup → Review → Output → Results`

The user first selects the assessment, uploads participant-level sample data,
then chooses Targeted Case Swaps, Targeted OSEP / Clinical Adjustments, or
CP-SAT. Recruitment/Sample Planning is intentionally packaged as the separate
`SCOPE_Sample_Planning.exe`. The GUI does not use legacy Part A/B/C terminology.

Projects can be saved, opened, and selected from Recent Projects without users
editing JSON. Project controls use progressive disclosure: only sources for the
selected workflow and Test Profile appear, and optional profile-specific sources
stay collapsed until requested. The Run screen and contextual bottom status area show stages,
progress, and elapsed time. The status area stays hidden until validation or a
run needs it. Small animal companions appear only in that status area; successful
runs retain the professional message **Complete! Everything is ready for review.**
Every successful run opens Results with verified output paths, direct report/file
buttons, an Open Run Folder action, and New Run.

Actions available to users:

- Optimize Sample: prepare, optimize, profile, and publish outputs.
- Recruitment Planning: reconcile Research Plan targets into integer recruitment
  Ns, then show current usable cases and remaining need. This is a count-level
  CP-SAT problem; no person-level cases are selected.
- Targeted Case Swap: validate and apply approved replacements without CP-SAT.
- Targeted OSEP / Clinical Adjustments: apply specific group-aware case changes
  through the same audited, lock-preserving swap engine.
- Profile / QA Sample: regenerate profile tables without changing membership.

## Test Profiles

A Test Profile contains schema mappings, assessment structure, rules, and
workflow capabilities. It also declares each workflow's canonical input roles,
labels, file types, requirement status, help text, and source adapter. The GUI
dynamically renders these requirements instead of showing one global set of
upload boxes. The coworker-facing GUI shows only validated profiles available
for the selected workflow. BASC-4 Narrowbands is the validated regression baseline. CELF-6 and
WISC-6 carry their validated assessment metadata (forms, ages, demographic
fields, eligibility, and clinical/OSEP mappings). Their catalog entries expose
Recruitment Planning and Targeted Case Swap while continuing to block Sample
Optimization until each validated assessment optimizer is connected to the
shared SCOPE engine.

Profile creation and editing are administrator functions. Use
`python gui/profile_manager.py` or the separately packaged
`SCOPE_Profile_Manager.exe` to create, edit, discover resources for, validate,
and install versioned profile specifications. Normal SCOPE reads approved
profiles but cannot modify them. Both standalone applications use the same
installed-profile catalog under the user's local application data directory.
The Profile Manager begins with **Add Assessment**, **Import Profile**, and
**Open Existing Profile**. Add Assessment supports four paths: international /
adapted assessment, adapt an existing profile, build from any available source
documents, or configure manually. Source documents are optional evidence; a
Research Plan is not a global requirement. Drafts, Needs Review profiles, and Validated profiles are
stored under `%LOCALAPPDATA%\Pearson\SCOPE\profiles` in standalone builds
(`user_profiles` when running from source), and user-installed versions take
precedence over bundled defaults with the same profile ID.

The editor is a guided workflow with progressive disclosure — **Overview →
Sources → Sample Rules → Demographics → Targets → Clinical/OSEP → Optimization
Rules → Output Tables → Validate** — instead of one flat list of mostly empty
document fields. An at-a-glance banner always answers four questions: which
cases are included, which demographic targets are being optimized, which
tolerances are used, and where each rule came from (inherited from the parent
profile, from assessment sources, manual/RD decision, or SCOPE default).

The **Sample Rules** step edits structured eligibility filters (for example
`Study = STDZ`) that the generic CP-SAT engine applies directly to the
candidate universe. The **Targets** step separates the *source* target format
from the *optimization/reporting* format: a fully crossed cell-by-cell census
workbook (for example `WAIS 5 EN CDN Targets.xlsx`) is imported as cells,
SCOPE derives the marginal/group targets, and the **Optimization Rules** step
turns profile tolerances such as Region ±10% into Range rules for the CP-SAT
objective. Raw crossed cells become objective rows only when the profile
explicitly keeps them.

Profile validation now reports core structural coherence separately from
workflow readiness. Targeted Case Swap, Targeted OSEP/Clinical Adjustments,
Recruitment Planning, Sample Optimization, and Reporting are evaluated
independently as Ready, Needs Review, or Not Configured. International
adaptations retain inherited values for review and do not silently treat them
as confirmed for a new market.
The SCOPE v2 Profile Manager authors separate Sample Planning, Targeted Case
Swap, Targeted OSEP/Clinical Adjustment, and CP-SAT capability flags; defines
task input roles; stores profile-owned targeted-change locks; and declares the
project output choices `teid_lists`, `tables`, and `both`. Loading and resaving
an older descriptor preserves assessment-specific nested settings while adding
the v2 membership-source definitions required by the current SCOPE apps.

## Saved projects

Saved `.scope` files are JSON internally and are grouped into `workflow`, `test_profile`,
`project_configuration`, canonical `inputs`, `canonical_targets`, `output`, and
`engine_version`. Case-universe sources, current sample, target sources,
recruitment inputs, and targeted-swap inputs remain distinct. It remains
loadable by `scripts/run_scope.py`, so a GUI-created project can be reproduced
through the Python workflow.

## Targeted-swap inputs

Targeted Case Swap accepts separate current-selected and available/candidate
TEID membership sources. The same workbook can supply both roles when its
`TEID_List` sheet contains `Selected TEIDs` and `Available Pool TEIDs` columns.
TEID-only membership lists can be mapped independently to an original dataset
for selected cases and an updated dataset for available candidates; a shared
participant dataset remains available as a fallback. Participant data may be
`.xlsx`, `.sas7bdat`, or `.csv`.

Overlap between the selected membership and a refreshed candidate source is
expected. Current selected membership is authoritative: SCOPE retains those
TEIDs as selected and excludes them from the effective add pool. It still
rejects duplicates within an individual source, missing IDs, attempts to add a
currently selected TEID, and swaps that violate profile-owned form, age-group,
or sex locks.

## Profile-driven reports

Each Test Profile owns report-template references and the available one-way,
two-way, and three-way demographic tables. Sample Optimization lets users
select the exact tables for each run and stores those choices in the `.scope`
project. Results exposes those deliverables as profile-driven previews before a run.
Recruitment Planning selects an approved Field Research workbook by Test
Profile and Study Type, preserves its existing workbook design, and populates
Target N/%, Current N/%, Remaining N, and Diff. Sample Optimization renders the
selected demographic sheets from the approved `TEMPLATE_1WAY`,
`TEMPLATE_2WAY`, and `TEMPLATE_3WAY` layouts and also writes OSEP/clinical workbooks in
addition to the final TEID list and QA output. Targeted Case Swap writes four
distinct deliverables based on the Output choice: refreshed Selected Cases and
Available Pool TEID lists, refreshed demographic/OSEP/clinical tables, or both.
The Swap Log and Run Manifest are always retained. Demographic workbooks show
target and sample counts/percentages without a derived difference column or row.

## Recruitment-plan workbook

Use the GUI's **Create recruitment-plan template** button, then complete the
`Recruitment Targets` worksheet. Each row is a marginal target with:

- `Dimension` (for example, Age Band, Sex, Ethnicity, OSEP, or Language)
- `Category`
- either `Target_Pct` (0–100) or `Target_N`
- optional `Tolerance_N`

SCOPE solves all listed dimensions together, writes the intersection-level
recruitment cells, the requested-versus-planned marginal targets, and the
remaining recruitment need. It also writes a reusable `SCOPE Targets` sheet so
Sample Optimization can consume the exact canonical targets approved during
Recruitment Planning. An optional progress workbook can contain either
`Dimension`, `Category`, and `Actual_N` columns or individual case rows with
columns matching the target dimensions.

Run from source with `python gui/app.py`, or build the standalone app with
`powershell -ExecutionPolicy Bypass -File misc/Build_Scope_EXE.ps1`. Build the
separate Sample Planning app with
`misc/Build_Scope_Sample_Planning_EXE.ps1`. Build the administrator tool with
`misc/Build_Scope_Profile_Manager_EXE.ps1`.
