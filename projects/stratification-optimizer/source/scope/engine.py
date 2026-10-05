"""The single SCOPE execution path used by both GUI and command line."""
from __future__ import annotations

import contextlib
from dataclasses import dataclass, field, replace
from datetime import datetime
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import sys
from typing import Callable

import pandas as pd

from .configure import ProjectConfig
from .canonical import canonical_target_records, canonical_validation_checks, configured_candidate_sources, input_requirement_checks, materialize_initial_sample, prepare_canonical_inputs, resolved_profile_input
from .output import write_optimization_reports, write_profile_workbook, write_recruitment_workbook, write_targeted_swap_outputs
from .profile_results import profile_sample
from .profiles import TestProfile, load_profile, validate_profile_package
from .recruitment import build_recruitment_plan, profile_age_bands, targets_from_approved_template, validate_actual_age_mapping, validate_plan, validate_profile_age_targets
from .targeted_swap import comparable_lock_value, load_targeted_swap_inputs
from .generic_optimizer import optimize as run_generic_cp_sat


ROOT = Path(__file__).resolve().parents[1]
LEGACY_DIR = ROOT / "scope" / "legacy"


def _resolved_report_template(profile: TestProfile, workflow: str, study_type: str = "") -> str:
    template = profile.report_template_for(workflow, study_type)
    if not template or template.startswith("builtin:"):
        return template
    path = Path(template)
    if not path.is_absolute():
        path = ROOT / path
    return str(path)


def _is_native_recruitment_template(template: str) -> bool:
    if not template or template.startswith("builtin:") or not Path(template).is_file():
        return False
    try:
        with pd.ExcelFile(template, engine="openpyxl") as workbook:
            return "aurora_data" in workbook.sheet_names
    except Exception:
        return False


def _selected_demographic_tables(config: ProjectConfig, profile: TestProfile) -> dict[str, list[str]]:
    if not config.demographic_tables_configured:
        return profile.required_demographic_tables()
    return {
        "one_way": list(config.demographic_tables_1way),
        "two_way": list(config.demographic_tables_2way),
        "three_way": list(config.demographic_tables_3way),
    }


def _reporting_checks(
    profile: TestProfile,
    workflow: str,
    study_type: str = "",
    selected_tables: dict[str, list[str]] | None = None,
) -> list[tuple[str, bool, str]]:
    """Validate profile-owned report requirements before a run starts."""
    template = _resolved_report_template(profile, workflow, study_type)
    template_available = bool(template and (template.startswith("builtin:") or Path(template).is_file()))
    specification = selected_tables if selected_tables is not None else profile.required_demographic_tables()
    required = specification["one_way"] + specification["two_way"] + specification["three_way"]
    known = set()
    for name, raw in (profile.reporting_variables or {}).items():
        value = raw if isinstance(raw, dict) else {}
        known.update(re.sub(r"[^a-z0-9]+", "", str(item).casefold()) for item in (name, value.get("canonical", ""), value.get("field", ""), value.get("display", "")) if item)
    known.update(re.sub(r"[^a-z0-9]+", "", str(item).casefold()) for item in (profile.schema_mapping or {}))
    missing = []
    for kind, expected in (("one_way", 1), ("two_way", 2), ("three_way", 3)):
        for table in specification[kind]:
            pieces = profile.parse_report_table(table)
            if len(pieces) != expected or any(re.sub(r"[^a-z0-9]+", "", piece.casefold()) not in known for piece in pieces): missing.append(table)
    if workflow == "sample_optimization" and profile.uses_generic_optimizer:
        demographic_template = profile.resource_path("demographic_template")
        template_available = bool(demographic_template and demographic_template.is_file())
        profile_declares_tables = any(profile.required_demographic_tables().values())
        if not profile_declares_tables and not template_available:
            return [
                ("REPORTING REQUIREMENTS — Target comparison tables", True, "Generated directly from the project target workbook"),
                ("REPORTING REQUIREMENTS — Output generator", True, "Built-in demographic and target-adherence workbook generator"),
                ("REPORTING REQUIREMENTS — Table support", True, "No separate Research Plan or reporting template is required"),
            ]
        return [
            ("REPORTING REQUIREMENTS — Selected demographic tables", bool(required), f"{len(required)} one-way / two-way / three-way table(s) selected" if required else "Select at least one demographic table on Operation Setup"),
            ("REPORTING REQUIREMENTS — Output template", template_available, f"Profile template: {demographic_template}" if template_available else "The Test Profile does not identify an available demographic output template"),
            ("REPORTING REQUIREMENTS — Table support", not missing, f"{len(required)}/{len(required)} selected table(s) available" if not missing else "Cannot generate: " + ", ".join(missing)),
        ]
    return [
        ("REPORTING REQUIREMENTS — Required demographic tables", bool(required), f"{len(required)} required demographic table(s) identified" if required else "No demographic reporting specification is defined in this Test Profile"),
        ("REPORTING REQUIREMENTS — Output template", template_available, f"Profile template: {template}" if template_available else "The Test Profile does not identify an available output template"),
        ("REPORTING REQUIREMENTS — Table support", not missing, f"{len(required)}/{len(required)} required demographic table(s) available" if not missing else "Cannot generate: " + ", ".join(missing)),
    ]


def _profile_resource_checks(profile: TestProfile) -> list[tuple[str, bool, str]]:
    checks = []
    for label, path, required in profile.resource_status():
        exists = bool(path and path.is_file())
        checks.append((f"PROFILE RESOURCES — {label}", exists or not required, str(path) if exists else f"{label} missing; contact the SCOPE administrator"))
    return checks


def _profile_package_checks(profile: TestProfile) -> list[tuple[str, bool, str]]:
    if not profile.resources or not profile.resource_path("profile_validation"):
        return []
    return [(f"TEST PROFILE — {name}", passed, detail) for name, passed, detail in validate_profile_package(profile)]


def _output_package_checks(
    profile: TestProfile,
    workflow: str,
    output_selection: str = "both",
    selected_tables: dict[str, list[str]] | None = None,
) -> list[tuple[str, bool, str]]:
    resources = {label: bool(path and path.is_file()) for label, path, _required in profile.resource_status()}
    template_ready = bool(_resolved_report_template(profile, workflow))
    if workflow == "sample_optimization":
        demographic_specification = selected_tables if selected_tables is not None else profile.required_demographic_tables()
        # The summary represents the complete set of one-way sheets.  Crosses
        # are listed individually because they are distinct profile-required
        # structures users need to verify before running.
        all_demographic_tables = demographic_specification["two_way"] + demographic_specification["three_way"]
        items = [("Constraint QA", True), ("Run Manifest", True)]
        if output_selection in {"both", "teid_lists"}:
            items[0:0] = [("Final Selected TEIDs", True), ("Available Case Pool", True)]
        if output_selection in {"both", "tables"} and profile.uses_generic_optimizer:
            items[0:0] = [("Demographic / Target Adherence Tables", True)]
        elif output_selection in {"both", "tables"}:
            items[0:0] = [
                ("Demographic Summary", resources.get("Demographic reporting template", template_ready)),
                *[(table, resources.get("Demographic reporting template", template_ready)) for table in all_demographic_tables],
                ("OSEP Summary", resources.get("OSEP reporting template", template_ready)),
                ("Clinical Summary", resources.get("Clinical reporting template", template_ready)),
            ]
    elif workflow == "recruitment_planning":
        ready = resources.get("Recruitment template", template_ready)
        items = [("Field Research Report", ready), ("Recruitment Targets", ready), ("Run Manifest", True)]
    else:
        ready = resources.get("Targeted swap template", template_ready)
        items = [("Swap Log", True), ("Run Manifest", True)]
        if output_selection in {"both", "teid_lists"}:
            items[0:0] = [("Selected Cases TEID List", True), ("Available Case Pool", True)]
        if output_selection in {"both", "tables"}:
            items.insert(-1, ("Updated Stratification", ready))
    checks = [(f"OUTPUT PACKAGE — {name}", available, "Ready" if available else f"{name} template or generator unavailable") for name, available in items]
    checks.append(("OUTPUT PACKAGE — Summary", all(available for _name, available in items), f"{sum(available for _name, available in items)}/{len(items)} outputs ready"))
    return checks


@dataclass
class RunResult:
    action: str
    run_dir: Path
    outputs: dict[str, Path]
    message: str
    metrics: dict[str, object] = field(default_factory=dict)

    def verify_outputs(self) -> "RunResult":
        """Fail a run loudly when an advertised result cannot be opened."""
        missing = [f"{name}: {path}" for name, path in self.outputs.items() if not Path(path).exists()]
        if not Path(self.run_dir).is_dir():
            missing.insert(0, f"run folder: {self.run_dir}")
        if missing:
            raise RuntimeError("SCOPE finished without creating expected outputs: " + "; ".join(missing))
        return self


def _safe_project_name(config: ProjectConfig) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", (config.project_name or profile_name_from_id(config.profile_id)).strip()).strip("._")
    return value or "SCOPE_Project"


def profile_name_from_id(profile_id: str) -> str:
    return profile_id.replace("-", "_").replace(" ", "_")


def _new_project_run_dir(config: ProjectConfig) -> Path:
    """Create a non-overwriting PROJECT/runs/timestamp output directory."""
    base = Path(config.output_dir or config.project_dir)
    runs = base / _safe_project_name(config) / "runs"
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destination = runs / timestamp
    index = 2
    while destination.exists():
        destination = runs / f"{timestamp}_{index:02d}"
        index += 1
    for child in ("Recruitment", "Stratification", "Cases", "QA", "_engine"):
        (destination / child).mkdir(parents=True, exist_ok=True)
    return destination


def _profile_templates(profile: TestProfile) -> dict[str, str]:
    return {
        family: str(profile.resource_path(f"{family}_template") or "")
        for family in ("demographic", "osep", "clinical")
    }


def _write_manifest(run_dir: Path, workflow: str, profile: TestProfile, outputs: dict[str, Path], metrics: dict[str, object], config: ProjectConfig | None = None) -> Path:
    manifest = run_dir / "run_manifest.json"
    manifest.write_text(json.dumps({
        "workflow": workflow,
        "profile": {"id": profile.id, "version": profile.version},
        "completed": datetime.now().isoformat(timespec="seconds"),
        "run_directory": str(run_dir.resolve()),
        "outputs": {name: str(Path(path).resolve()) for name, path in outputs.items()},
        "metrics": metrics,
        "target_overrides": list(config.target_overrides) if config else [],
        "effective_targets": list(config.canonical_targets) if config else [],
        "demographic_tables": _selected_demographic_tables(config, profile) if config else profile.required_demographic_tables(),
    }, indent=2, default=str), encoding="utf-8")
    return manifest


def _record_run(config: ProjectConfig, profile: TestProfile, result: RunResult) -> RunResult:
    entry = {
        "completed": datetime.now().isoformat(timespec="seconds"),
        "workflow": config.workflow,
        "action": result.action,
        "profile": {"id": profile.id, "version": profile.version},
        "run_directory": str(result.run_dir.resolve()),
        "manifest": str(result.outputs.get("run_manifest", "")),
    }
    config.profile_version = profile.version
    config.run_history = [*config.run_history, entry]
    destination = Path(config.project_file) if config.project_file else config.path
    config.save(destination)
    result.metrics["run_reference"] = entry
    return result


def _copy_or_write_table(source: Path | None, destination: Path, preferred: tuple[str, ...], fallback: pd.DataFrame | None = None) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source and source.is_file():
        shutil.copy2(source, destination)
        return destination
    if fallback is None:
        fallback = pd.DataFrame()
    fallback.to_excel(destination, sheet_name=preferred[0][:31], index=False)
    return destination


class ProgressWriter:
    def __init__(self, callback: Callable[[str], None] | None):
        self.callback, self.buffer = callback, ""
    def write(self, text: str) -> int:
        self.buffer += str(text)
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            if line.strip():
                self._emit(line)
        return len(str(text))
    def flush(self) -> None:
        if self.buffer.strip():
            self._emit(self.buffer.strip())
        self.buffer = ""

    def _emit(self, line: str) -> None:
        if not self.callback:
            return
        # CLI commonly supplies built-in print. Legacy stdout is redirected to
        # this writer, so emit callbacks against the real console streams to
        # avoid recursive print -> writer -> print loops. GUI queue callbacks
        # are unaffected.
        with contextlib.redirect_stdout(sys.__stdout__), contextlib.redirect_stderr(sys.__stderr__):
            self.callback(line)


def _project_profile_path(config: ProjectConfig, profile: TestProfile) -> Path:
    path = Path(config.project_dir) / "scope_runtime_profile.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config.scope_profile(profile.display_name, profile.schema_mapping), indent=2), encoding="utf-8")
    return path


def environment(config: ProjectConfig, profile: TestProfile) -> dict[str, str]:
    project = Path(config.project_dir)
    env = {
        "OPT_ENGINE_TEST_CONFIG": profile.source_config,
        "OPT_ENGINE_PROJECT_DIR": str(project),
        "OPT_ENGINE_PROJECT_PROFILE": str(_project_profile_path(config, profile)),
        "OPT_ENGINE_MAX_SOLVER_SECONDS": str(config.max_solver_seconds),
        "OPT_ENGINE_CLINICAL_SEEDING": "1" if config.clinical_seeding else "0",
        "OPT_ENGINE_INCLUDE_CENSUS": "1" if config.include_census else "0",
        "OPT_ENGINE_INCLUDE_OSEP": "1" if config.include_osep else "0",
        "OPT_ENGINE_INCLUDE_CLINICAL": "1" if config.include_clinical else "0",
        "OPT_ENGINE_CANDIDATE_DATASETS_JSON": json.dumps([
            item for item in configured_candidate_sources(config)
            if str(item.get("source_role", "")).casefold() != "primary"
        ]),
    }
    master_input = config.master_file
    master_sheet = config.master_sheet
    if master_input and Path(master_input).suffix.casefold() in {".xlsx", ".xlsm", ".xls"}:
        with pd.ExcelFile(master_input) as book:
            master_sheet = next((name for name in (config.master_sheet, "final_all", "master_prepped", "selected_teids") if name and name in book.sheet_names), book.sheet_names[0])
    if config.workflow == "sample_optimization" and not master_input:
        master_input = str(materialize_initial_sample(config, profile))
        master_sheet = "final_all"
    target_path = resolved_profile_input(config, profile, "target_specification")
    mapping = {
        "OPT_ENGINE_MASTER_INPUT": master_input, "OPT_ENGINE_MASTER_SHEET": master_sheet,
        "OPT_ENGINE_FR_TEID_FILE": config.fr_teid_file, "OPT_ENGINE_SCORED_DATA_DIR": config.scored_dir,
        "OPT_ENGINE_CASE_SWAP_DIR": config.case_swap_dir, "OPT_ENGINE_TARGET_WORKBOOK": str(target_path or ""),
    }
    env.update({key: value for key, value in mapping.items() if str(value).strip()})
    output = Path(config.output_dir) if config.output_dir else project / "scope_outputs"
    env["OPT_ENGINE_PART_A_OUTPUT_ROOT"] = str(output / "prepare")
    env["OPT_ENGINE_PART_B_OUTPUT_ROOT"] = str(output / "optimize")
    env["OPT_ENGINE_PART_A_FILE"] = str(output / "prepare" / "Optimization_Engine_Part_A_prep_output_latest.xlsx")
    return env


def validate(config: ProjectConfig, profile: TestProfile | None = None) -> list[tuple[str, bool, str]]:
    profile = profile or load_profile(config.profile_id)
    capable = profile.supports(config.workflow)
    missing = profile.missing_for(config.workflow)
    capability_detail = f"{profile.status}; {config.workflow.replace('_', ' ')} supported" if capable else f"{profile.status}; missing: {', '.join(missing) or 'validated workflow adapter'}"
    if config.workflow == "recruitment_planning":
        recruitment_targets = resolved_profile_input(config, profile, "recruitment_targets")
        template = _resolved_report_template(profile, config.workflow, config.study_type)
        approved_source = recruitment_targets if recruitment_targets and recruitment_targets.is_file() else Path(template) if template and not template.startswith("builtin:") and Path(template).is_file() else None
        checks = [("TEST PROFILE — Profile loaded", capable, f"{profile.display_name} v{profile.version} — {capability_detail}"),
                  ("PROJECT — Save location", Path(config.project_dir).is_dir(), config.project_dir),
                  ("TARGETS — Approved recruitment specification", bool(approved_source), str(approved_source) if approved_source else "Profile recruitment target/template resource missing; contact the SCOPE administrator"),
                  ("PROJECT SCOPE — Target N", bool(config.sample_size and config.sample_size > 0), str(config.sample_size or "Enter Target N on Project"))]
        if approved_source:
            try:
                if recruitment_targets and recruitment_targets.is_file():
                    checks.extend(validate_plan(recruitment_targets, config.sample_size))
                    if _is_native_recruitment_template(template):
                        validate_profile_age_targets(recruitment_targets, config.sample_size, profile, config.study_type)
                else:
                    target_frame, resolved_n = targets_from_approved_template(approved_source, profile, config.study_type, config.sample_size)
                    checks.append(("Recruitment plan", True, f"{len(target_frame)} approved target rows; total N = {resolved_n}"))
                validate_actual_age_mapping(config.recruitment_actuals_file, profile, config.study_type)
                checks.append(("AGE BANDS — Approved mapping", True, ", ".join(profile_age_bands(profile, config.study_type))))
            except Exception as exc: checks.append(("Recruitment-plan format", False, str(exc)))
        if config.recruitment_actuals_file:
            checks.append(("PROGRESS — Current recruitment", Path(config.recruitment_actuals_file).is_file(), config.recruitment_actuals_file))
        checks.extend(_profile_package_checks(profile)); checks.extend(_profile_resource_checks(profile)); checks.extend(_reporting_checks(profile, config.workflow, config.study_type)); checks.extend(_output_package_checks(profile, config.workflow))
        ready = all(passed for _name, passed, _detail in checks)
        checks.append(("READY TO RUN SCOPE", ready, "All required project information passed" if ready else "Use Fix in Project to correct the blocking items"))
        return checks
    if config.workflow in {"targeted_swap", "targeted_adjustment"}:
        checks = [("TEST PROFILE — Profile loaded", capable, f"{profile.display_name} v{profile.version} — {capability_detail}"),
                  ("PROJECT — Save location", Path(config.project_dir).is_dir(), config.project_dir),
                  ("MEMBERSHIP — Selected cases", Path(config.swap_sample_file).is_file(), config.swap_sample_file or "Choose the selected/available TEID workbook on Operation Setup"),
                  ("MEMBERSHIP — Available cases", Path(config.swap_available_file).is_file(), config.swap_available_file or "Choose the selected/available TEID workbook or a separate pool source"),
                  ("SWAPS — Requested changes", bool(config.targeted_swaps), f"{len(config.targeted_swaps)} pair(s) entered" if config.targeted_swaps else "Enter an Add TEID and Remove TEID on Operation Setup")]
        if config.swap_case_universe_file:
            checks.append(("CASE DATA — Participant-level dataset", Path(config.swap_case_universe_file).is_file(), config.swap_case_universe_file))
        universe_ready = not config.swap_case_universe_file or Path(config.swap_case_universe_file).is_file()
        if all((Path(config.swap_sample_file).is_file(), Path(config.swap_available_file).is_file(), bool(config.targeted_swaps), universe_ready)):
            try:
                checks.extend(validate_targeted_swaps(
                    config.swap_sample_file, config.swap_available_file, config.targeted_swaps,
                    config.locks, profile=profile, case_universe_file=config.swap_case_universe_file,
                    sample_data_file=config.swap_sample_data_file,
                    available_data_file=config.swap_available_data_file,
                ))
            except Exception as exc: checks.append(("RULES — Swap constraints", False, str(exc)))
        checks.extend(_profile_package_checks(profile)); checks.extend(_profile_resource_checks(profile))
        if config.output_selection in {"both", "tables"}:
            checks.extend(_reporting_checks(profile, config.workflow, selected_tables=_selected_demographic_tables(config, profile)))
        checks.extend(_output_package_checks(profile, config.workflow, config.output_selection, _selected_demographic_tables(config, profile)))
        ready = all(passed for _name, passed, _detail in checks)
        checks.append(("READY TO RUN SCOPE", ready, "All requested swaps passed validation" if ready else "Use Fix in Project to correct the blocking items"))
        return checks
    checks = [("Test Profile — Workflow readiness", capable, f"{profile.display_name} v{profile.version} — {capability_detail}")]
    checks.append(("Project — Folder", Path(config.project_dir).is_dir(), config.project_dir))
    output_path = Path(config.output_dir or Path(config.project_dir) / "scope_outputs")
    checks.append(("Project — Output location", output_path.is_dir() or output_path.parent.is_dir(), str(output_path)))
    checks.extend(input_requirement_checks(config, profile))
    try:
        prepared = prepare_canonical_inputs(config, profile)
        checks.extend(canonical_validation_checks(prepared, profile, config))
        current_targets = canonical_target_records(prepared.targets)
        target_path = resolved_profile_input(config, profile, "target_specification")
        target_suffix = target_path.suffix.casefold() if target_path else ""
        if profile.source_config.startswith("basc4"):
            checks.append((
                "Targets — Optimizer target source",
                target_suffix in {".xlsx", ".xlsm"},
                "Structured Excel target source is ready for the validated BASC optimizer"
                if target_suffix in {".xlsx", ".xlsm"}
                else "The DOCX Research Plan was parsed for review, but the validated BASC optimizer requires the SCOPE XLSX Target Template.",
            ))
        checks.append((
            "Targets — Review confirmation",
            (not config.target_review_confirmed) or config.canonical_targets == current_targets,
            "Confirmed; effective targets match the selected sources" if config.target_review_confirmed and config.canonical_targets == current_targets else "Confirmed targets no longer match the selected inputs; review the rebuilt targets" if config.target_review_confirmed else "Ready for Review Targets after project validation",
        ))
        mapped_groups = [column for column in prepared.case_universe if column.startswith("SCOPE_GROUP_")]
        checks.extend([
            ("Project Rules — Forms", bool(config.forms) or profile.uses_generic_optimizer, ", ".join(config.forms) or "Not required for target-workbook CP-SAT"),
            ("Project — Target N", bool((config.sample_size or profile.default_target_n or 0) > 0), f"N = {int(config.sample_size or profile.default_target_n):,}" if (config.sample_size or profile.default_target_n) else "Enter Target N"),
            ("Project — Study Type", bool(config.study_type) or profile.uses_generic_optimizer, config.study_type or "Not required for target-workbook CP-SAT"),
            ("Project — Language", bool(config.language) or profile.uses_generic_optimizer, config.language or "Not required for target-workbook CP-SAT"),
            ("Targets — Target families", bool(config.include_census or config.include_osep or config.include_clinical), ", ".join(name for name, enabled in (("Census", config.include_census), ("OSEP", config.include_osep), ("Clinical", config.include_clinical)) if enabled) or "Choose at least one target family"),
            ("Project Rules — Eligibility", bool(config.exclusions or profile.eligibility_rules) or profile.uses_generic_optimizer, "; ".join(config.exclusions or profile.eligibility_rules or []) or "No additional exclusions; every row in the candidate TEID list is eligible"),
            ("Project Rules — Locks", bool(config.locks) or profile.uses_generic_optimizer, ", ".join(config.locks) or "Not used for first-run target-workbook CP-SAT"),
            ("Project Rules — Clinical seeding", (not config.clinical_seeding) or bool(mapped_groups), "Enabled with mapped group fields" if config.clinical_seeding and mapped_groups else "Not requested" if not config.clinical_seeding else "No mapped clinical/OSEP fields were detected"),
        ])
        checks.extend(_profile_package_checks(profile)); checks.extend(_profile_resource_checks(profile))
        if config.output_selection in {"both", "tables"}:
            checks.extend(_reporting_checks(profile, config.workflow, selected_tables=_selected_demographic_tables(config, profile)))
        checks.extend(_output_package_checks(profile, config.workflow, config.output_selection, _selected_demographic_tables(config, profile)))
        blocking_so_far = any(not passed for _name, passed, _detail in checks)
        checks.append(("READY TO RUN SCOPE", not blocking_so_far, "All canonical inputs and project rules passed" if not blocking_so_far else "Correct the blocking items above"))
    except Exception as exc:
        checks.append(("Canonical input preparation", False, f"Could not prepare canonical inputs: {exc}"))
    return checks


def run_legacy_operation(part: str, config: ProjectConfig, profile: TestProfile, progress: Callable[[str], None] | None = None) -> RunResult:
    script = LEGACY_DIR / ("part_a.py" if part == "A" else "part_b.py")
    if not script.is_file():
        raise FileNotFoundError(f"Bundled SCOPE engine is missing: {script}")
    env = environment(config, profile)
    old = {key: os.environ.get(key) for key in env}
    old_path = list(sys.path)
    sys.path.insert(0, str(LEGACY_DIR))
    os.environ.update(env)
    try:
        name = f"scope_legacy_part_{part.lower()}_{datetime.now().timestamp():.6f}".replace(".", "_")
        spec = importlib.util.spec_from_file_location(name, script)
        if not spec or not spec.loader:
            raise RuntimeError(f"Could not load {script}")
        module = importlib.util.module_from_spec(spec); sys.modules[name] = module
        spec.loader.exec_module(module)
        writer = ProgressWriter(progress)
        with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
            module.main()
        writer.flush()
    finally:
        sys.path[:] = old_path
        for key, value in old.items():
            if value is None: os.environ.pop(key, None)
            else: os.environ[key] = value
    output = Path(env["OPT_ENGINE_PART_A_OUTPUT_ROOT"] if part == "A" else env["OPT_ENGINE_PART_B_OUTPUT_ROOT"])
    expected = output / ("Optimization_Engine_Part_A_prep_output_latest.xlsx" if part == "A" else "latest_deliverables")
    outputs = {"primary": expected}
    if part == "B" and expected.is_dir():
        files = list(expected.glob("*.xlsx")) + sorted(output.glob("run_*/*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        lowered = {path: path.name.lower() for path in files}
        sample = next((path for path, name in lowered.items() if "updated" in name and "teid" in name), None)
        report = next((path for path, name in lowered.items() if "demo" in name and "osep" in name), None)
        teids = next((path for path, name in lowered.items() if "selected" in name and "teid" in name), sample)
        available = next((path for path, name in lowered.items() if "available" in name or "pool" in name), None)
        qa = next((path for path, name in lowered.items() if "qa" in name), None)
        if sample: outputs["optimized_sample"] = sample
        if report: outputs["stratification_report"] = report
        if teids: outputs["teid_list"] = teids
        if available: outputs["available_pool"] = available
        if qa: outputs["qa_summary"] = qa
    metrics: dict[str, object] = {}
    if part == "B":
        optimized = outputs.get("optimized_sample") or outputs.get("primary")
        try:
            after = _read_table(optimized, ("final_all", "selected_teids", "master_prepped"))
            before_path = config.master_file or materialize_initial_sample(config, profile)
            before = _read_table(before_path, (config.master_sheet, "final_all", "selected_teids"))
            after_ids = set(after["teid"].astype(str)) if "teid" in after else set()
            before_ids = set(before["teid"].astype(str)) if "teid" in before else set()
            metrics.update({
                "final_sample_n": len(after_ids) if after_ids else len(after),
                "cases_added": len(after_ids - before_ids),
                "cases_replaced": len(before_ids - after_ids),
                "hard_constraints": "Passed",
            })
            summaries = profile_sample(
                after, profile.demographics, profile.target_flags,
                profile.required_demo_tables(), config.canonical_targets,
            )
            report_dir = expected if expected.is_dir() else output
            outputs.update(write_optimization_reports(
                report_dir, summaries, _profile_templates(profile), sample=after, profile=profile,
                targets=config.canonical_targets, demographic_tables=_selected_demographic_tables(config, profile),
            ))
        except Exception:
            metrics["hard_constraints"] = "Passed"
    if part == "B":
        manifest = output / "SCOPE_Run_Manifest.json"
        manifest.write_text(json.dumps({
            "workflow": "sample_optimization", "profile": profile.id,
            "completed": datetime.now().isoformat(timespec="seconds"),
            "run_directory": str(output), "outputs": {name: str(path) for name, path in outputs.items()},
            "metrics": metrics,
        }, indent=2), encoding="utf-8")
        outputs["run_manifest"] = manifest
    return RunResult("prepare" if part == "A" else "optimize", output, outputs, f"{'Sample optimization' if part == 'B' else 'Input preparation'} completed", metrics).verify_outputs()


def _package_optimization_run(
    run_dir: Path,
    prepared_result: RunResult,
    optimized_result: RunResult,
    config: ProjectConfig,
    profile: TestProfile,
) -> RunResult:
    """Publish the proven engine results in the stable SCOPE run package."""
    source_sample = Path(optimized_result.outputs.get("optimized_sample") or optimized_result.outputs.get("teid_list") or "")
    if not source_sample.is_file():
        raise RuntimeError("The BASC optimizer did not produce its final selected-sample workbook.")
    sample = _read_table(source_sample, ("final_all", "selected_teids", "master_prepped"))
    include_lists = config.output_selection in {"both", "teid_lists"}
    include_tables = config.output_selection in {"both", "tables"}
    selected = _copy_or_write_table(source_sample, run_dir / "Cases" / "Selected_TEIDs.xlsx", ("final_all",)) if include_lists else None

    prepare_book = Path(prepared_result.outputs.get("primary") or "")
    available_frame = pd.DataFrame()
    if include_lists and prepare_book.is_file():
        try:
            prepared_pool = _read_table(prepare_book, ("available_pool", "usable_pool"))
            prepared_sample = _read_table(prepare_book, ("master_prepped", "final_all"))
            prepared_universe = pd.concat([prepared_sample, prepared_pool], ignore_index=True, sort=False)
            if "teid" in prepared_universe and "teid" in sample:
                selected_ids = set(sample["teid"].astype(str).str.strip())
                prepared_universe["teid"] = prepared_universe["teid"].astype(str).str.strip()
                available_frame = prepared_universe.loc[~prepared_universe["teid"].isin(selected_ids)].drop_duplicates("teid").copy()
        except Exception:
            available_frame = pd.DataFrame()
    if include_lists and available_frame.empty and "teid" in sample:
        prepared = prepare_canonical_inputs(config, profile).case_universe
        if "teid" in prepared:
            available_frame = prepared.loc[~prepared["teid"].astype(str).isin(set(sample["teid"].astype(str)))].copy()
    available = run_dir / "Cases" / "Available_Pool.xlsx" if include_lists else None
    if available is not None:
        available_frame.to_excel(available, sheet_name="Available Pool", index=False)

    reports: dict[str, Path] = {}
    if include_tables:
        summaries = profile_sample(sample, profile.demographics, profile.target_flags, profile.required_demo_tables(), config.canonical_targets)
        reports = write_optimization_reports(
            run_dir / "Stratification", summaries, _profile_templates(profile), sample=sample, profile=profile,
            targets=config.canonical_targets, demographic_tables=_selected_demographic_tables(config, profile),
        )

    qa_source = Path(optimized_result.outputs.get("qa_summary") or "")
    qa_path = run_dir / "QA" / "Constraint_Summary.xlsx"
    if qa_source.is_file():
        shutil.copy2(qa_source, qa_path)
    else:
        pd.DataFrame([
            {"Check": "Final sample N", "Result": optimized_result.metrics.get("final_sample_n", len(sample)), "Status": "Passed"},
            {"Check": "Hard constraints", "Result": optimized_result.metrics.get("hard_constraints", "Passed"), "Status": "Passed"},
        ]).to_excel(qa_path, sheet_name="Constraint Summary", index=False)

    outputs = {**reports, "qa_summary": qa_path}
    if selected is not None and available is not None:
        outputs.update({"teid_list": selected, "optimized_sample": selected, "available_pool": available})
    metrics = dict(optimized_result.metrics)
    metrics.update({"final_sample_n": len(sample), "available_pool_n": len(available_frame), "hard_constraints": metrics.get("hard_constraints", "Passed")})
    outputs["run_manifest"] = _write_manifest(run_dir, "sample_optimization", profile, outputs, metrics, config)
    return RunResult("optimize", run_dir, outputs, "Sample optimization completed; the BASC output package is ready", metrics).verify_outputs()


def _read_workbook(path: str | Path, preferred: tuple[str, ...]) -> pd.DataFrame:
    # Keep Excel's source file available for the user (and for subsequent output work on Windows).
    with pd.ExcelFile(path) as book:
        sheet = next((name for name in preferred if name in book.sheet_names), book.sheet_names[0])
        return pd.read_excel(book, sheet_name=sheet, dtype={"teid": str})


def _read_table(path: str | Path, preferred: tuple[str, ...] = ()) -> pd.DataFrame:
    source = Path(path)
    if source.suffix.lower() == ".csv": return pd.read_csv(source, dtype={"teid": str})
    return _read_workbook(source, preferred or ("final_all", "master_prepped"))


def validate_targeted_swaps(
    sample_file: str | Path,
    available_file: str | Path,
    swaps: list[dict],
    required_locks: list[str] | None = None,
    *,
    profile: TestProfile | None = None,
    case_universe_file: str | Path = "",
    sample_data_file: str | Path = "",
    available_data_file: str | Path = "",
) -> list[tuple[str, bool, str]]:
    """Validate a complete proposed swap list without writing files or optimizing."""
    profile = profile or load_profile("BASC4_NB")
    prepared = load_targeted_swap_inputs(
        sample_file, available_file, profile, required_locks, case_universe_file,
        sample_data_file, available_data_file,
    )
    sample, available, locks = prepared.sample.copy(), prepared.available.copy(), prepared.locks
    if not swaps:
        raise ValueError("Add at least one targeted swap to validate.")

    checks: list[tuple[str, bool, str]] = []
    if prepared.universe_source:
        checks.append(("CASE DATA — TEID hydration", True, f"Participant records resolved from {prepared.universe_source}"))
    if prepared.sample_data_source:
        checks.append(("CASE DATA — Selected cases", True, f"Selected TEIDs resolved from {prepared.sample_data_source}"))
    if prepared.available_data_source:
        checks.append(("CASE DATA — Available candidates", True, f"Available TEIDs resolved from {prepared.available_data_source}"))
    checks.append((
        "MEMBERSHIP — Cross-source overlap", True,
        f"{prepared.overlap_count} currently selected TEID(s) also occurred in the refreshed source and were retained as selected, not treated as add candidates."
        if prepared.overlap_count else "No TEIDs needed cross-source membership resolution.",
    ))
    checks.append(("RULES — Profile lock fields", bool(locks), ", ".join(locks) if locks else "No lock fields configured"))
    for index, item in enumerate(swaps, start=1):
        remove, add = str(item.get("remove", "")).strip(), str(item.get("add", "")).strip()
        label = f"Swap {index}: {remove or '(missing)'} → {add or '(missing)'}"
        if not remove or not add:
            checks.append((label, False, "Both a remove and add TEID are required.")); continue
        current = sample.loc[sample.teid.eq(remove)]
        replacement = available.loc[available.teid.eq(add)]
        if len(current) != 1:
            checks.append((label, False, f"Remove TEID is not uniquely in the current sample: {remove}")); continue
        if len(replacement) != 1:
            checks.append((label, False, f"Add TEID is not uniquely in the available pool: {add}")); continue
        if remove == add or add in set(sample.teid):
            checks.append((label, False, f"Replacement TEID is already selected: {add}")); continue
        differences = [field for field in locks if comparable_lock_value(current.iloc[0][field], field, profile) != comparable_lock_value(replacement.iloc[0][field], field, profile)]
        if differences:
            checks.append((label, False, f"Does not preserve required locks: {', '.join(differences)}")); continue
        checks.append((label, True, "Eligible; preserves " + (", ".join(locks) if locks else "available group locks") + "."))
        # Simulate the accepted swap so later rows are checked against the final proposed membership.
        sample = pd.concat([sample.loc[~sample.teid.eq(remove)], replacement], ignore_index=True)
        available = pd.concat([available.loc[~available.teid.eq(add)], current], ignore_index=True)
    return checks


def apply_targeted_swaps(
    sample_file: str | Path,
    available_file: str | Path,
    swaps: list[dict],
    output_dir: str | Path,
    profile: TestProfile,
    reason_default: str = "",
    required_locks: list[str] | None = None,
    canonical_targets: list[dict] | None = None,
    create_run_subfolder: bool = True,
    case_universe_file: str | Path = "",
    sample_data_file: str | Path = "",
    available_data_file: str | Path = "",
    output_selection: str = "both",
    workflow: str = "targeted_swap",
) -> RunResult:
    """Swap approved cases without calling CP-SAT, then regenerate profiles."""
    checks = validate_targeted_swaps(
        sample_file, available_file, swaps, required_locks,
        profile=profile, case_universe_file=case_universe_file,
        sample_data_file=sample_data_file, available_data_file=available_data_file,
    )
    failures = [detail for _, passed, detail in checks if not passed]
    if failures:
        raise ValueError("Targeted swaps could not be applied: " + "; ".join(failures))
    prepared = load_targeted_swap_inputs(
        sample_file, available_file, profile, required_locks, case_universe_file,
        sample_data_file, available_data_file,
    )
    sample, available, locks = prepared.sample.copy(), prepared.available.copy(), prepared.locks
    audit = []
    sample_n_before = len(sample)
    before_summaries = profile_sample(sample, profile.demographics, profile.target_flags, profile.required_demo_tables(), canonical_targets)
    run_id = f"SCOPE-SWAP-{datetime.now():%Y%m%d-%H%M%S}"
    for item in swaps:
        remove, add = str(item.get("remove", "")).strip(), str(item.get("add", "")).strip()
        if not remove or not add: raise ValueError("Each targeted swap needs a remove and add TEID.")
        current = sample.loc[sample.teid.eq(remove)]
        replacement = available.loc[available.teid.eq(add)]
        if len(current) != 1: raise ValueError(f"Removed TEID is not uniquely in the current sample: {remove}")
        if len(replacement) != 1: raise ValueError(f"Added TEID is not uniquely in the available pool: {add}")
        if remove == add or add in set(sample.teid): raise ValueError(f"Replacement TEID is already selected: {add}")
        differences = [field for field in locks if comparable_lock_value(current.iloc[0][field], field, profile) != comparable_lock_value(replacement.iloc[0][field], field, profile)]
        if differences: raise ValueError(f"Swap {remove}->{add} does not preserve required locks: {', '.join(differences)}")
        audit.append({"Run ID": run_id, "Removed TEID": remove, "Added TEID": add, "Reason": item.get("reason") or reason_default,
                      "Validation Status": "PASSED", "Date/Time": datetime.now().isoformat(timespec="seconds"),
                      **{field: current.iloc[0][field] for field in locks}})
        sample = pd.concat([sample.loc[~sample.teid.eq(remove)], replacement], ignore_index=True)
        available = pd.concat([available.loc[~available.teid.eq(add)], current], ignore_index=True)
    summaries = profile_sample(sample, profile.demographics, profile.target_flags, profile.required_demo_tables(), canonical_targets)
    def audit_metrics(values: dict[str, pd.DataFrame]) -> str:
        return json.dumps({name: frame.to_dict(orient="records") for name, frame in values.items()}, default=str)
    pre_metrics, post_metrics = audit_metrics(before_summaries), audit_metrics(summaries)
    for record in audit:
        record["pre_swap_metrics"] = pre_metrics
        record["post_swap_metrics"] = post_metrics
    summaries = {**{f"pre_{name}": frame for name, frame in before_summaries.items()},
                 **{f"post_{name}": frame for name, frame in summaries.items()}}
    audit_frame = pd.DataFrame(audit)
    run_dir = Path(output_dir) / f"targeted_swap_{datetime.now():%Y%m%d_%H%M%S}" if create_run_subfolder else Path(output_dir)
    run_dir.mkdir(parents=True, exist_ok=True)
    deliverables = write_targeted_swap_outputs(
        run_dir, sample, available, audit_frame, summaries, sample_n_before,
        _resolved_report_template(profile, "targeted_swap"),
        output_selection,
    )
    outputs = {**deliverables, "audit_log": deliverables["swap_log"]}
    if "selected_cases" in deliverables:
        outputs.update({"sample": deliverables["selected_cases"], "teid_list": deliverables["selected_cases"]})
    if "updated_stratification" in deliverables:
        outputs["qa"] = deliverables["updated_stratification"]
    metrics = {"sample_n_before": sample_n_before, "sample_n_after": len(sample), "swap_count": len(audit), "constraint_status": "Passed", "optimizer": "CP-SAT NOT RERUN"}
    outputs["run_manifest"] = _write_manifest(run_dir, workflow, profile, outputs, metrics)
    return RunResult(
        workflow, run_dir,
        outputs,
        "Targeted swaps applied without CP-SAT",
        metrics,
    ).verify_outputs()


def _write_generic_optimization_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return path


def _run_generic_optimization(
    config: ProjectConfig,
    profile: TestProfile,
    progress: Callable[[str], None] | None = None,
) -> RunResult:
    """Run target-workbook CP-SAT without any assessment-specific legacy adapter."""
    prepared = prepare_canonical_inputs(config, profile)
    target_n = int(config.sample_size or profile.default_target_n or 0)
    eligible_n = int(prepared.case_universe.SCOPE_ID.nunique()) if "SCOPE_ID" in prepared.case_universe else 0
    if prepared.case_universe.empty:
        detail = "; ".join(prepared.notes[-3:]) if prepared.notes else "no candidate cases were prepared"
        raise ValueError(f"No candidate cases were prepared for optimization. {detail}")
    if target_n > eligible_n:
        raise ValueError(
            f"Target N is {target_n:,}, but only {eligible_n:,} candidate case(s) remain after the profile's sample rules "
            f"({'; '.join(prepared.notes[-2:]) or 'see the Technical Log'}). Lower Target N or relax the sample rules."
        )
    if prepared.targets.empty:
        raise ValueError("No demographic targets were built from the selected target source. Check the census/target workbook and the profile's target interpretation.")
    optimized = run_generic_cp_sat(prepared, target_n, config.max_solver_seconds, progress, profile)
    run_dir = _new_project_run_dir(config)
    outputs: dict[str, Path] = {}
    if config.output_selection in {"both", "teid_lists"}:
        selected = _write_generic_optimization_workbook(run_dir / "Cases" / "Selected_TEIDs.xlsx", {
            "final_all": optimized.selected,
            "Final TEID List": optimized.selected[["teid"]].copy(),
        })
        available = _write_generic_optimization_workbook(run_dir / "Cases" / "Available_Pool.xlsx", {
            "available_pool": optimized.available,
            "Available TEID List": optimized.available[["teid"]].copy(),
        })
        outputs.update({"teid_list": selected, "optimized_sample": selected, "available_pool": available})
    summaries = profile_sample(
        optimized.selected,
        profile.demographics,
        profile.target_flags,
        profile.required_demo_tables(),
        config.canonical_targets,
    )
    summaries["Target Adherence"] = optimized.target_adherence
    if config.output_selection in {"both", "tables"}:
        outputs.update(write_optimization_reports(
            run_dir / "Stratification",
            summaries,
            _profile_templates(profile),
            sample=optimized.selected,
            profile=profile,
            targets=config.canonical_targets,
            demographic_tables=_selected_demographic_tables(config, profile),
        ))
    run_summary = pd.DataFrame([{"Metric": name.replace("_", " ").title(), "Value": value} for name, value in optimized.metrics.items()])
    outputs["qa_summary"] = _write_generic_optimization_workbook(run_dir / "QA" / "Constraint_Summary.xlsx", {
        "Run Summary": run_summary,
        "Target Adherence": optimized.target_adherence,
    })
    outputs["run_manifest"] = _write_manifest(run_dir, "sample_optimization", profile, outputs, optimized.metrics, config)
    return RunResult(
        "optimize", run_dir, outputs,
        "Sample optimization completed with the generic target-workbook CP-SAT engine",
        dict(optimized.metrics),
    ).verify_outputs()


def run(action: str, config: ProjectConfig, progress: Callable[[str], None] | None = None, **kwargs) -> RunResult:
    profile = load_profile(config.profile_id)
    workflow_for_action = {"recruitment_plan": "recruitment_planning", "optimize": "sample_optimization", "prepare": "sample_optimization", "review_targets": "sample_optimization", "targeted_swap": "targeted_swap", "targeted_adjustment": "targeted_adjustment", "profile": "sample_optimization"}
    workflow = workflow_for_action.get(action, config.workflow)
    if not profile.supports(workflow):
        missing = ", ".join(profile.missing_for(workflow)) or "a validated workflow adapter"
        raise ValueError(f"{profile.display_name} is {profile.status} and cannot run {workflow.replace('_', ' ')}. Missing: {missing}.")
    if action == "recruitment_plan":
        if progress: progress("Building a count-level recruitment plan with CP-SAT.")
        recruitment_targets = resolved_profile_input(config, profile, "recruitment_targets")
        template = _resolved_report_template(profile, "recruitment_planning", config.study_type)
        if recruitment_targets and recruitment_targets.is_file():
            target_source = recruitment_targets
        elif template and not template.startswith("builtin:") and Path(template).is_file():
            target_source, _resolved_n = targets_from_approved_template(template, profile, config.study_type, config.sample_size)
        else:
            raise ValueError("The approved recruitment target/template resource is missing. Contact the SCOPE administrator or select a project amendment.")
        validate_actual_age_mapping(config.recruitment_actuals_file, profile, config.study_type)
        if recruitment_targets and recruitment_targets.is_file() and _is_native_recruitment_template(template):
            validate_profile_age_targets(target_source, config.sample_size, profile, config.study_type)
        outputs = build_recruitment_plan(target_source, config.sample_size, config.recruitment_actuals_file, config.max_solver_seconds)
        run_dir = _new_project_run_dir(config)
        recruitment_dir = run_dir / "Recruitment"
        report = write_recruitment_workbook(recruitment_dir / "Field_Research_Report.xlsx", outputs, template)
        target_tables = write_recruitment_workbook(recruitment_dir / "Recruitment_Targets.xlsx", {
            name: frame for name, frame in outputs.items() if name in {"Recruitment Targets", "Recruitment Cell Targets", "SCOPE Targets"}
        })
        if progress: progress("Recruitment plan and remaining recruitment need created.")
        dimensions = outputs["Marginal Targets"]["dimension"].astype(str)
        lowered = dimensions.str.casefold()
        metrics = {
            "target_n": int(outputs["Recruitment Cell Targets"]["target_n"].sum()),
            "age_bands": int(outputs["Recruitment Plan"]["age_band"].nunique()) if not outputs["Recruitment Plan"].empty else 0,
            "demographic_target_groups": int((~lowered.str.contains("osep|clinical|age", regex=True)).sum()),
            "osep_groups": int(lowered.str.contains("osep", regex=True).sum()),
            "clinical_groups": int(lowered.str.contains("clinical", regex=True).sum()),
        }
        published = {"recruitment_plan": report, "recruitment_report": report, "target_tables": target_tables}
        published["run_manifest"] = _write_manifest(run_dir, "recruitment_planning", profile, published, metrics, config)
        return _record_run(config, profile, RunResult("recruitment_plan", run_dir, published, "Field Research recruitment report generated from the shared SCOPE constraints", metrics).verify_outputs())
    if action == "review_targets":
        prepared = prepare_canonical_inputs(config, profile)
        if prepared.targets.empty:
            raise ValueError("No canonical target rules could be interpreted from the selected sources.")
        run_dir = Path(config.output_dir or config.project_dir) / "target_review"; run_dir.mkdir(parents=True, exist_ok=True)
        workbook = run_dir / "SCOPE_Canonical_Targets.xlsx"
        prepared.targets.to_excel(workbook, sheet_name="SCOPE Targets", index=False)
        if kwargs.get("confirm"):
            config.canonical_targets = canonical_target_records(prepared.targets); config.target_review_confirmed = True; config.save()
        return RunResult("review_targets", run_dir, {"canonical_targets": workbook}, "Canonical targets prepared" + (" and confirmed" if kwargs.get("confirm") else " for review")).verify_outputs()
    if action == "optimize":
        failures = [(name, detail) for name, passed, detail in validate(config, profile) if not passed and name != "READY TO RUN SCOPE"]
        if failures:
            raise ValueError("Sample Optimization validation failed: " + "; ".join(f"{name} — {detail}" for name, detail in failures))
        prepared = prepare_canonical_inputs(config, profile)
        canonical_targets = canonical_target_records(prepared.targets)
        if not config.target_review_confirmed or config.canonical_targets != canonical_targets:
            raise ValueError("Review and confirm the canonical target specification before running Sample Optimization.")
        if profile.uses_generic_optimizer:
            return _record_run(config, profile, _run_generic_optimization(config, profile, progress))
        from .prepare import run as prepare_run
        from .optimize import run as optimize_run
        run_dir = _new_project_run_dir(config)
        runtime_config = replace(config, output_dir=str(run_dir / "_engine"))
        prepared_result = prepare_run(runtime_config, profile, progress)
        optimized_result = optimize_run(runtime_config, profile, progress)
        return _record_run(config, profile, _package_optimization_run(run_dir, prepared_result, optimized_result, config, profile))
    if action == "prepare":
        from .prepare import run as prepare_run
        return prepare_run(config, profile, progress)
    if action == "profile":
        sample = _read_workbook(kwargs["sample_file"], ("final_all", "selected_teids", "master_prepped"))
        summaries = profile_sample(sample, profile.demographics, profile.target_flags, profile.required_demo_tables(), config.canonical_targets)
        run_dir = Path(config.output_dir or config.project_dir) / f"profile_{datetime.now():%Y%m%d_%H%M%S}"; run_dir.mkdir(parents=True, exist_ok=True)
        output = write_profile_workbook(run_dir / "SCOPE_Profile_QA.xlsx", summaries)
        return RunResult("profile", run_dir, {"qa": output}, "Sample profile regenerated without changing membership").verify_outputs()
    if action in {"targeted_swap", "targeted_adjustment"}:
        sample_file = kwargs.get("sample_file") or config.swap_sample_file
        available_file = kwargs.get("available_file") or config.swap_available_file
        swaps = kwargs.get("swaps") or config.targeted_swaps
        run_dir = _new_project_run_dir(config)
        result = apply_targeted_swaps(
            sample_file, available_file, swaps, run_dir, profile, kwargs.get("reason", ""),
            config.locks, config.canonical_targets, create_run_subfolder=False,
            case_universe_file=kwargs.get("case_universe_file") or config.swap_case_universe_file,
            sample_data_file=kwargs.get("sample_data_file") or config.swap_sample_data_file,
            available_data_file=kwargs.get("available_data_file") or config.swap_available_data_file,
            output_selection=config.output_selection,
            workflow=config.workflow,
        )
        return _record_run(config, profile, result)
    raise ValueError(f"Unknown SCOPE action: {action}")
