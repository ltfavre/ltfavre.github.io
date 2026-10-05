"""Validated, user-facing SCOPE test profiles."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import shutil
import sys

import yaml


ROOT = Path(__file__).resolve().parents[1]
PROFILES_DIR = ROOT / "profiles"
USER_PROFILES_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "Pearson" / "SCOPE" / "profiles"


@dataclass(frozen=True)
class InputRequirement:
    """One profile-owned source needed to construct canonical workflow inputs."""

    role: str
    label: str
    group: str
    status: str
    kind: str = "file"
    accepted_types: tuple[str, ...] = ()
    purpose: str = ""
    adapter: str = "canonical_table"

    @property
    def required(self) -> bool:
        return self.status.casefold() == "required"


@dataclass(frozen=True)
class TestProfile:
    id: str
    display_name: str
    version: str
    source_config: str
    forms: list[str]
    demographics: list[str]
    target_flags: list[str]
    runnable: bool = True
    status: str = "Validated"
    capabilities: dict[str, bool] | None = None
    missing: dict[str, list[str]] | None = None
    schema_mapping: dict[str, str] | None = None
    age_structure: list[str] | None = None
    osep_mappings: dict[str, object] | None = None
    eligibility_rules: list[str] | None = None
    test_ids: list[str] | None = None
    languages: list[str] | None = None
    study_types: list[str] | None = None
    target_families: list[str] | None = None
    default_target_n: int | None = None
    workflow_inputs: dict[str, list[InputRequirement]] | None = None
    reporting: dict[str, object] | None = None
    data_specification: dict[str, object] | None = None
    sampling_specification: dict[str, object] | None = None
    output_specification: dict[str, object] | None = None
    resources: dict[str, object] | None = None
    reporting_variables: dict[str, object] | None = None
    package_dir: str = ""
    catalog_visibility: str = "user"
    edition: str = ""
    market: str = ""
    description: str = ""
    source_documents: list[dict[str, object]] | None = None
    provenance: dict[str, object] | None = None
    readiness: dict[str, object] | None = None
    adaptation: dict[str, object] | None = None
    profile_kind: str = "standard"
    sample_filters: list[dict[str, object]] = field(default_factory=list)
    target_tolerances: dict[str, object] = field(default_factory=dict)
    target_interpretation: dict[str, object] = field(default_factory=dict)
    field_origins: dict[str, str] = field(default_factory=dict)
    last_modified: str = ""
    descriptor_path: str = ""

    @property
    def user_visible(self) -> bool:
        return self.catalog_visibility.casefold() not in {"administrator", "admin", "hidden"}

    @property
    def optimizer_engine(self) -> str:
        return str(self.source_config or (self.sampling_specification or {}).get("engine_config", "")).strip()

    @property
    def uses_generic_optimizer(self) -> bool:
        return self.optimizer_engine.casefold() in {"generic_cp_sat", "generic-cp-sat", "target_workbook_cp_sat"}

    def supports(self, workflow: str) -> bool:
        aliases = {
            "targeted_swap": "targeted_swap",
            "targeted_adjustment": "targeted_swap",
            "sample_optimization": "sample_optimization",
            "recruitment_planning": "recruitment_planning",
        }
        values = self.capabilities or {}
        canonical = aliases.get(workflow, workflow)
        enabled = bool(values[workflow]) if workflow == "targeted_adjustment" and workflow in values else bool(values.get(canonical, self.runnable if canonical != "recruitment_planning" else False))
        saved = self.readiness or {}
        state = saved.get(workflow, saved.get(canonical))
        if isinstance(state, dict):
            state = state.get("status")
        # A validated profile is authoritative: leftover adaptation-review
        # readiness must not re-gate a workflow the profile already enables.
        if self.status.casefold() == "validated":
            return enabled
        return enabled and (not state or str(state).casefold() == "ready")

    def missing_for(self, workflow: str) -> list[str]:
        canonical = "targeted_swap" if workflow == "targeted_adjustment" else workflow
        values = self.missing or {}
        stored = values.get(workflow, values.get(canonical, []))
        if stored:
            return list(stored)
        readiness = (self.readiness or {}).get(workflow, (self.readiness or {}).get(canonical, {}))
        return [str(value) for value in readiness.get("missing", [])] if isinstance(readiness, dict) else []

    def requirements_for(self, workflow: str) -> list[InputRequirement]:
        canonical = "targeted_swap" if workflow == "targeted_adjustment" else workflow
        values = self.workflow_inputs or {}
        return list(values.get(workflow, values.get(canonical, [])))

    def report_template(self, workflow: str) -> str:
        key = {
            "recruitment_planning": "recruitment_template",
            "sample_optimization": "optimization_template",
            "targeted_swap": "targeted_swap_template",
            "targeted_adjustment": "targeted_swap_template",
        }.get(workflow, "")
        return str((self.reporting or {}).get(key, ""))

    def recruitment_spec(self, study_type: str = "") -> dict[str, object]:
        """Return the profile-owned recruitment report definition for a study type."""
        recruitment = (self.reporting or {}).get("recruitment") or {}
        if not isinstance(recruitment, dict):
            recruitment = {}
        wanted = re.sub(r"[^a-z0-9]+", "", study_type.casefold())
        for name, specification in recruitment.items():
            if re.sub(r"[^a-z0-9]+", "", str(name).casefold()) == wanted and isinstance(specification, dict):
                return dict(specification)
        for name, specification in recruitment.items():
            normalized = re.sub(r"[^a-z0-9]+", "", str(name).casefold())
            if wanted and (wanted in normalized or normalized in wanted) and isinstance(specification, dict):
                return dict(specification)
        return {}

    def report_template_for(self, workflow: str, study_type: str = "") -> str:
        value = self.recruitment_spec(study_type).get("template") if workflow == "recruitment_planning" else None
        value = value or self.report_template(workflow)
        if not value:
            return ""
        path = Path(str(value))
        if not path.is_absolute() and not str(value).startswith("builtin:"):
            package_candidate = Path(self.package_dir) / path
            root_candidate = ROOT / path
            path = package_candidate if package_candidate.exists() or not root_candidate.exists() else root_candidate
        return str(path)

    def output_previews(self, workflow: str, study_type: str = "") -> list[dict[str, object]]:
        """Build user previews from the installed profile's reporting specification."""
        deliverables = (self.output_specification or {}).get("workflow_deliverables") or {}
        canonical = "targeted_swap" if workflow == "targeted_adjustment" else workflow
        declared = deliverables.get(workflow, deliverables.get(canonical, [])) if isinstance(deliverables, dict) else []
        if workflow == "recruitment_planning":
            specification = self.recruitment_spec(study_type)
            tables = list(specification.get("required_tables") or self.required_report_tables("demographics"))
            template = self.report_template_for(workflow, study_type)
            return [{
                "title": "Field Research Report", "kind": "workbook",
                "description": f"Approved format for {self.display_name} • {study_type or 'profile default'}",
                "tables": tables, "template": template,
            }]
        required = (self.reporting or {}).get("required_tables") or {}
        if workflow == "sample_optimization":
            if self.uses_generic_optimizer and not any(self.required_demographic_tables().values()):
                return [dict(item) if isinstance(item, dict) else {"title": str(item), "kind": "workbook"} for item in declared] or [
                    {"title": title, "kind": "workbook"} for title in (
                        "Final Selected TEIDs", "Available Case Pool", "Demographic / Target Adherence Tables", "QA / Constraint Summary",
                    )
                ]
            previews = []
            demographic = self.required_demographic_tables(); template = str(self.resource_path("demographic_template") or "")
            for key, title, template_sheet in (
                ("one_way", "One-Way Demographic Tables", "TEMPLATE_1WAY"),
                ("two_way", "Two-Way Demographic Tables", "TEMPLATE_2WAY"),
                ("three_way", "Three-Way Demographic Tables", "TEMPLATE_3WAY"),
            ):
                if demographic[key]: previews.append({"title": title, "kind": "demographic_template", "tables": demographic[key], "template": template, "template_sheet": template_sheet})
            if required.get("osep"): previews.append({"title": "OSEP Tables", "kind": "table", "tables": list(required["osep"])})
            if required.get("clinical"): previews.append({"title": "Clinical Tables", "kind": "table", "tables": list(required["clinical"])})
            if declared:
                for item in declared:
                    preview = dict(item) if isinstance(item, dict) else {"title": str(item), "kind": "workbook"}
                    title = str(preview.get("title", "")).casefold()
                    if (("demographic" in title and "table" in title)
                            or any(family in title for family in ("demographic_tables", "osep_tables", "osep tables", "clinical_tables", "clinical tables"))):
                        continue
                    previews.append(preview)
            else:
                previews.extend({"title": title, "kind": "file"} for title in ("Final Selected TEIDs", "Available Case Pool", "QA / Constraint Summary"))
            return previews
        if declared:
            return [dict(item) if isinstance(item, dict) else {"title": str(item), "kind": "workbook"} for item in declared]
        if workflow in {"targeted_swap", "targeted_adjustment"}:
            return [{"title": title, "kind": "file", "description": "CP-SAT NOT RERUN"} for title in (
                "Updated Selected Cases TEID List", "Updated Available Case Pool", "Swap Log", "Refreshed Stratification Tables")]
        return []

    def required_demo_tables(self) -> list[str]:
        required = (self.reporting or {}).get("required_tables") or {}
        if not isinstance(required, dict):
            return []
        return [str(value) for value in required.get("demo_cross_tabs", [])]

    @staticmethod
    def parse_report_table(value: str) -> list[str]:
        return [part.strip() for part in re.split(r"\s*(?:×|\b[xX]\b)\s*", str(value)) if part.strip()]

    def required_demographic_tables(self) -> dict[str, list[str]]:
        required = (self.reporting or {}).get("required_tables") or {}
        if not isinstance(required, dict): required = {}
        return {
            "one_way": [str(value) for value in required.get("demographics", required.get("one_way", []))],
            "two_way": [str(value) for value in required.get("demo_cross_tabs", required.get("two_way", []))],
            "three_way": [str(value) for value in required.get("demo_three_way", required.get("three_way", []))],
        }

    def required_report_tables(self, family: str) -> list[str]:
        required = (self.reporting or {}).get("required_tables") or {}
        if not isinstance(required, dict):
            return []
        return [str(value) for value in required.get(family, [])]

    def resource_path(self, name: str) -> Path | None:
        value = (self.resources or {}).get(name)
        if not value:
            return None
        path = Path(str(value))
        if not path.is_absolute():
            package_candidate = Path(self.package_dir) / path
            root_candidate = ROOT / path
            path = package_candidate if package_candidate.exists() or not root_candidate.exists() else root_candidate
        return path

    def resource_status(self) -> list[tuple[str, Path | None, bool]]:
        labels = {
            "research_specification": "Research specification",
            "approved_targets": "Default CP-SAT target workbook",
            "census_targets": "Census targets",
            "osep_specification": "OSEP specification",
            "recruitment_targets": "Recruitment target specification",
            "recruitment_template": "Recruitment template",
            "demographic_template": "Demographic reporting template",
            "osep_template": "OSEP reporting template",
            "clinical_template": "Clinical reporting template",
            "targeted_swap_template": "Targeted swap template",
            "profile_validation": "Profile validation specification",
        }
        required = set((self.output_specification or {}).get("required_resources", (self.resources or {}).keys()))
        names = [name for name in labels if name in (self.resources or {}) or name in required]
        return [(labels.get(name, name.replace("_", " ").title()), self.resource_path(name), name in required) for name in names]


def _from_dict(data: dict, package_dir: str | Path = "", descriptor_path: str | Path = "") -> TestProfile:
    required = ("id", "display_name", "version")
    missing = [name for name in required if not str(data.get(name, "")).strip()]
    if missing:
        raise ValueError("SCOPE test profile is missing: " + ", ".join(missing))
    runnable = bool(data.get("runnable", True))
    capabilities = data.get("capabilities") or {
        "recruitment_planning": False,
        "sample_optimization": runnable,
        "targeted_swap": runnable,
        "targeted_adjustment": runnable,
    }
    if not isinstance(data.get("workflow_inputs") or {}, dict):
        raise ValueError("workflow_inputs must be an object grouped by workflow.")
    try:
        workflow_inputs = {
            str(workflow): [
                InputRequirement(
                    role=str(item["role"]), label=str(item["label"]),
                    group=str(item.get("group", "INPUTS")), status=str(item.get("status", "required")),
                    kind=str(item.get("kind", "file")),
                    accepted_types=tuple(str(value) for value in item.get("accepted_types", [])),
                    purpose=str(item.get("purpose", "")), adapter=str(item.get("adapter", "canonical_table")),
                )
                for item in items
            ]
            for workflow, items in (data.get("workflow_inputs") or {}).items()
        }
    except (KeyError, TypeError) as exc:
        raise ValueError(f"A workflow input is malformed: {exc}") from exc
    data_spec = dict(data.get("data_specification") or {})
    sampling_spec = dict(data.get("sampling_specification") or {})
    reporting_spec = dict(data.get("reporting_specification") or data.get("reporting") or {})
    output_spec = dict(data.get("output_specification") or {})
    raw_filters = data.get("sample_filters", sampling_spec.get("sample_filters")) or []
    raw_tolerances = data.get("target_tolerances", sampling_spec.get("target_tolerances")) or {}
    raw_interpretation = data.get("target_interpretation", sampling_spec.get("target_interpretation")) or {}
    raw_origins = data.get("field_origins", (data.get("provenance") or {}).get("field_origins")) or {}
    modified = str(data.get("last_modified", ""))
    descriptor = Path(descriptor_path) if descriptor_path else None
    if not modified and descriptor and descriptor.is_file():
        modified = datetime.fromtimestamp(descriptor.stat().st_mtime, timezone.utc).isoformat()
    return TestProfile(
        id=str(data["id"]), display_name=str(data["display_name"]), version=str(data["version"]),
        source_config=str(data.get("source_config", sampling_spec.get("engine_config", ""))), forms=[str(value) for value in data.get("forms", data_spec.get("forms", []))],
        demographics=[str(value) for value in data.get("demographics", data_spec.get("demographics", []))],
        target_flags=[str(value) for value in data.get("target_flags", sampling_spec.get("target_flags", []))], runnable=runnable,
        status=str(data.get("status", "Validated" if runnable else "Draft")),
        capabilities={str(key): bool(value) for key, value in capabilities.items()},
        missing={str(key): [str(value) for value in values] for key, values in (data.get("missing") or {}).items()},
        schema_mapping={str(key): str(value) for key, value in (data.get("schema_mapping") or data_spec.get("variable_map") or {}).items()},
        age_structure=[str(value) for value in data.get("age_structure", data_spec.get("age_structure", []))],
        osep_mappings=data.get("osep_mappings") or data_spec.get("osep_mappings") or {},
        eligibility_rules=[str(value) for value in data.get("eligibility_rules", sampling_spec.get("eligibility_rules", []))],
        test_ids=[str(value) for value in data.get("test_ids", data_spec.get("test_ids", []))],
        languages=[str(value) for value in data.get("languages", data_spec.get("languages", []))],
        study_types=[str(value) for value in data.get("study_types", data_spec.get("study_types", []))],
        target_families=[str(value) for value in data.get("target_families", sampling_spec.get("target_families", ["Census", "OSEP", "Clinical"]))],
        default_target_n=int(data.get("default_target_n", sampling_spec.get("default_target_n"))) if data.get("default_target_n", sampling_spec.get("default_target_n")) not in (None, "") else None,
        workflow_inputs=workflow_inputs,
        reporting=reporting_spec, data_specification=data_spec, sampling_specification=sampling_spec,
        output_specification=output_spec, resources=dict(data.get("resources") or {}),
        reporting_variables=dict(data.get("reporting_variables") or data_spec.get("reporting_variables") or reporting_spec.get("variables") or {}),
        package_dir=str(package_dir),
        catalog_visibility=str(data.get("catalog_visibility", "user")),
        edition=str(data.get("edition", "")),
        market=str(data.get("market", data.get("locale", data.get("country", "")))),
        description=str(data.get("description", "")),
        source_documents=[dict(value) for value in (data.get("source_documents") or []) if isinstance(value, dict)],
        provenance=dict(data.get("provenance") or {}), readiness=dict(data.get("workflow_readiness") or {}),
        adaptation=dict(data.get("adaptation") or {}),
        profile_kind=str(data.get("profile_kind", "standard") or "standard"),
        sample_filters=[dict(value) for value in raw_filters if isinstance(value, dict)],
        target_tolerances={str(key): (dict(value) if isinstance(value, dict) else value) for key, value in raw_tolerances.items()} if isinstance(raw_tolerances, dict) else {},
        target_interpretation=dict(raw_interpretation) if isinstance(raw_interpretation, dict) else {},
        field_origins={str(key): str(value) for key, value in raw_origins.items()} if isinstance(raw_origins, dict) else {},
        last_modified=modified,
        descriptor_path=str(descriptor_path),
    )


@dataclass(frozen=True)
class ProfileLoadError:
    path: Path
    assessment_name: str
    message: str


@dataclass
class ProfileCatalog:
    profiles: list[TestProfile] = field(default_factory=list)
    errors: list[ProfileLoadError] = field(default_factory=list)
    searched_directories: list[Path] = field(default_factory=list)

    @property
    def user_profiles(self) -> list[TestProfile]:
        return [profile for profile in self.profiles if profile.user_visible]


def profile_directories() -> list[Path]:
    """Resolve profile roots consistently in source and PyInstaller builds."""
    # User-installed profiles intentionally precede bundled defaults so an
    # administrator's saved update to an existing assessment is discoverable.
    candidates = [USER_PROFILES_DIR, PROFILES_DIR]
    if getattr(sys, "frozen", False):
        candidates.insert(1, Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent)) / "profiles")
    unique: list[Path] = []
    for path in candidates:
        resolved = path.resolve()
        if resolved not in unique:
            unique.append(resolved)
    return unique


def _error_name(path: Path, raw: str) -> str:
    match = __import__("re").search(r'["\']display_name["\']\s*:\s*["\']([^"\']+)', raw)
    return match.group(1) if match else path.stem.replace("_", " ")


def discover_profile_catalog(directories: list[str | Path] | None = None) -> ProfileCatalog:
    """Discover every descriptor and retain actionable parse/validation errors."""
    roots = [Path(value).resolve() for value in directories] if directories is not None else profile_directories()
    catalog = ProfileCatalog(searched_directories=roots)
    seen: set[str] = set()
    for folder in roots:
        if not folder.is_dir():
            continue
        descriptors = [
            *sorted(folder.glob("*/profile.yaml")), *sorted(folder.glob("*/profile.yml")),
            *sorted(folder.glob("*/profile.json")), *sorted(folder.glob("*.json")),
        ]
        for path in descriptors:
            raw = ""
            try:
                raw = path.read_text(encoding="utf-8")
                payload = yaml.safe_load(raw) if path.suffix.casefold() in {".yaml", ".yml"} else json.loads(raw)
                if not isinstance(payload, dict):
                    raise ValueError("descriptor must contain an object")
                profile = _from_dict(payload, path.parent, path)
                if profile.id in seen:
                    continue
                catalog.profiles.append(profile)
                seen.add(profile.id)
            except Exception as exc:
                catalog.errors.append(ProfileLoadError(path, _error_name(path, raw), str(exc)))
    return catalog


def list_profiles() -> list[TestProfile]:
    return discover_profile_catalog().profiles


def load_profile(profile_id: str) -> TestProfile:
    for profile in list_profiles():
        if profile.id == profile_id:
            return profile
    raise ValueError(f"SCOPE test profile was not found: {profile_id}")


WORKFLOW_LABELS = {
    "targeted_swap": "Targeted Case Swap",
    "targeted_adjustment": "Targeted OSEP / Clinical Adjustments",
    "recruitment_planning": "Recruitment Planning",
    "sample_optimization": "Sample Optimization",
    "reporting": "Reporting",
}


def _workflow_result(enabled: bool, missing: list[str], ready_detail: str) -> dict[str, object]:
    if not enabled:
        return {"status": "NOT CONFIGURED", "detail": "This workflow has not been enabled for the profile.", "missing": []}
    if missing:
        return {"status": "NEEDS REVIEW", "detail": "Missing information: " + ", ".join(missing), "missing": missing}
    return {"status": "READY", "detail": ready_detail, "missing": []}


def workflow_readiness(profile: TestProfile) -> dict[str, dict[str, object]]:
    """Evaluate each SCOPE workflow independently from core profile validity."""
    capabilities = profile.capabilities or {}
    mapping = profile.schema_mapping or {}
    mapping_names = set(mapping) | {str(value) for value in mapping.values()}
    reporting = profile.reporting or {}
    required_tables = reporting.get("required_tables") or {}
    has_report_tables = isinstance(required_tables, dict) and any(bool(value) for value in required_tables.values())
    sampling = profile.sampling_specification or {}
    inputs = profile.workflow_inputs or {}
    locks = [str(value) for value in sampling.get("default_locks", [])] or ["form", "age_group", "sex"]
    inherited_review = {key for key, value in (((profile.adaptation or {}).get("review") or {}).items()) if value == "review"}

    def inherited_missing(relevant: set[str]) -> list[str]:
        pending = sorted(inherited_review & relevant)
        return ["confirmation of inherited " + ", ".join(value.replace("_", " ") for value in pending)] if pending else []

    targeted_missing = []
    if not mapping.get("teid"):
        targeted_missing.append("TEID variable mapping")
    if not profile.forms:
        targeted_missing.append("assessment forms")
    if not locks:
        targeted_missing.append("targeted-change locks")
    elif any(lock not in mapping_names for lock in locks):
        targeted_missing.append("variable mappings for all targeted-change locks")
    if not inputs.get("targeted_swap"):
        targeted_missing.append("selected/available membership input definitions")
    targeted_missing.extend(inherited_missing({"locale", "structure", "age_structure", "mappings", "eligibility"}))
    targeted_enabled = bool(capabilities.get("targeted_swap", profile.runnable))

    adjustment_missing = list(targeted_missing)
    families = {str(value).casefold() for value in (profile.target_families or [])}
    group_mappings = profile.osep_mappings or {}
    if not (families & {"osep", "clinical"}) and not group_mappings:
        adjustment_missing.append("OSEP or clinical category definitions")
    if not (required_tables.get("osep") or required_tables.get("clinical")):
        adjustment_missing.append("OSEP or clinical reporting definitions")
    adjustment_missing.extend(inherited_missing({"clinical_osep", "reporting"}))
    adjustment_enabled = bool(capabilities.get("targeted_adjustment", capabilities.get("targeted_swap", False)))

    recruitment_missing = []
    if not profile.age_structure:
        recruitment_missing.append("age-band structure")
    if not profile.demographics:
        recruitment_missing.append("demographic categories")
    if not profile.study_types:
        recruitment_missing.append("study types")
    recruitment = reporting.get("recruitment") or {}
    if not inputs.get("recruitment_planning"):
        recruitment_missing.append("recruitment target input definition")
    if not recruitment and not reporting.get("recruitment_template"):
        recruitment_missing.append("recruitment reporting definition")
    recruitment_missing.extend(inherited_missing({"locale", "structure", "age_structure", "demographics", "targets", "reporting"}))

    optimization_missing = []
    generic_optimizer = profile.uses_generic_optimizer
    requirements = [
        (mapping.get("teid"), "TEID variable mapping"),
        (profile.demographics, "demographic categories"),
        (profile.optimizer_engine, "optimizer/engine configuration"),
        (inputs.get("sample_optimization"), "optimization input definitions"),
    ]
    if not generic_optimizer:
        requirements[1:1] = [
            (profile.forms, "assessment forms"),
            (profile.age_structure, "age structure"),
            (profile.eligibility_rules or profile.sample_filters, "eligibility/exclusion rules or sample filters"),
        ]
    for value, label in requirements:
        if not value:
            optimization_missing.append(label)
    relevant_review = {"locale", "demographics", "targets", "mappings"} if generic_optimizer else {"locale", "structure", "age_structure", "demographics", "clinical_osep", "eligibility", "targets", "mappings", "reporting"}
    optimization_missing.extend(inherited_missing(relevant_review))

    reporting_missing = []
    if not has_report_tables:
        reporting_missing.append("required report tables")
    if not profile.reporting_variables and not profile.demographics:
        reporting_missing.append("reporting variable definitions")
    if profile.uses_generic_optimizer and has_report_tables and not profile.resource_path("demographic_template"):
        reporting_missing.append("a demographic output template for the report tables")
    reporting_missing.extend(inherited_missing({"demographics", "clinical_osep", "reporting"}))
    reporting_enabled = bool(has_report_tables or profile.reporting_variables)

    return {
        "targeted_swap": _workflow_result(targeted_enabled, targeted_missing, "Case identity, membership inputs, forms, and lock fields are configured."),
        "targeted_adjustment": _workflow_result(adjustment_enabled, adjustment_missing, "Targeted-change rules and OSEP/clinical reporting are configured."),
        "recruitment_planning": _workflow_result(bool(capabilities.get("recruitment_planning", False)), recruitment_missing, "Age bands, targets, study types, and recruitment reporting are configured."),
        "sample_optimization": _workflow_result(bool(capabilities.get("sample_optimization", False)), optimization_missing, "Candidate ID/demographic mappings, target-workbook inputs, and the CP-SAT engine are configured." if generic_optimizer else "Case schema, eligibility rules, targets, and optimizer inputs are configured."),
        "reporting": _workflow_result(reporting_enabled, reporting_missing, "Reporting variables and required tables are configured."),
    }


def profile_validation_summary(profile: TestProfile) -> dict[str, object]:
    """Return human-readable section validation plus independent workflow readiness."""
    mapping = profile.schema_mapping or {}
    reporting = profile.reporting or {}
    required_tables = reporting.get("required_tables") or {}
    has_report_tables = isinstance(required_tables, dict) and any(bool(value) for value in required_tables.values())
    adaptation_review = (profile.adaptation or {}).get("review") or {}

    def section(key: str, label: str, complete: bool, missing: str, *, optional: bool = False) -> dict[str, str]:
        state = "COMPLETE" if complete else "NOT CONFIGURED" if optional else "NEEDS REVIEW"
        detail = "Configured" if complete else missing
        review_value = adaptation_review.get(key)
        if review_value == "review":
            state, detail = "NEEDS REVIEW", "Inherited value requires confirmation for this market/locale."
        elif review_value == "not_applicable":
            state, detail = "NOT APPLICABLE", "Marked not applicable by the profile author."
        return {"key": key, "label": label, "status": state, "detail": detail}

    identity_ok = bool(profile.id.strip() and profile.display_name.strip() and profile.version.strip())
    sections = [
        {"key": "identity", "label": "Identity", "status": "COMPLETE" if identity_ok else "ERROR", "detail": "Assessment name, profile ID, and version are defined." if identity_ok else "Assessment name, profile ID, and version are required."},
        section("locale", "Identity & Locale", bool(profile.market), "Market/locale needs confirmation."),
        section("structure", "Forms / Languages", bool(profile.forms and profile.languages), "Forms and languages need configuration or confirmation."),
        section("age_structure", "Age Structure", bool(profile.age_structure), "Age bands are not configured."),
        section("mappings", "Case Schema / Variable Mapping", bool(mapping.get("teid")), "The TEID variable mapping is missing."),
        section("demographics", "Demographic Categories", bool(profile.demographics), "Demographic categories are not configured."),
        section("clinical_osep", "OSEP / Clinical Categories", bool(profile.osep_mappings or profile.target_flags), "No OSEP or clinical definitions are configured.", optional=True),
        section("eligibility", "Sample Filters / Eligibility", bool(profile.eligibility_rules or profile.sample_filters), "Sample filter and eligibility rules are not configured.", optional=True),
        section("targets", "Sample Targets / Tolerances", bool(profile.target_families and (profile.default_target_n or profile.target_tolerances or profile.target_interpretation or (profile.sampling_specification or {}).get("target_tolerances"))), "Sample targets or tolerances are not configured.", optional=True),
        section("reporting", "Reporting Definitions", has_report_tables, "Required reporting tables are not configured.", optional=True),
        section("sources", "Source / Provenance", bool(profile.source_documents or profile.provenance), "No source documents or manual provenance have been recorded.", optional=True),
    ]
    workflows = workflow_readiness(profile)
    errors = [item for item in sections if item["status"] == "ERROR"]
    reviews = [item for item in sections if item["status"] == "NEEDS REVIEW"]
    return {
        "overall": "INVALID" if errors else "VALID",
        "core_valid": not errors,
        "needs_review": bool(reviews or any(item["status"] == "NEEDS REVIEW" for item in workflows.values())),
        "sections": sections,
        "workflows": workflows,
        "sources": list(profile.source_documents or []),
    }


def validate_profile_package(profile: TestProfile) -> list[tuple[str, bool, str]]:
    """Backward-compatible tuple validation focused on profile integrity."""
    summary = profile_validation_summary(profile)
    checks: list[tuple[str, bool, str]] = [
        (str(item["label"]).upper(), item["status"] != "ERROR", str(item["detail"]))
        for item in summary["sections"]
    ]
    missing_resources = [label for label, path, required in profile.resource_status() if required and not (path and path.is_file())]
    checks.append(("OUTPUT SPECIFICATION — Resources", not missing_resources, "All explicitly required resources exist" if not missing_resources else "Missing required files: " + ", ".join(missing_resources)))
    readiness = summary["workflows"]
    ready_count = sum(item["status"] == "READY" for item in readiness.values())
    checks.append(("WORKFLOW READINESS", True, f"{ready_count} workflow(s) ready; optional workflows are evaluated independently"))
    checks.append(("PROFILE STATUS", profile.status.casefold() == "validated" and bool(summary["core_valid"]) and not missing_resources, f"{profile.display_name} v{profile.version} — {profile.status}"))
    return checks


def validate_profile_data(data: dict, package_dir: str | Path = "") -> tuple[TestProfile | None, list[tuple[str, bool, str]]]:
    """Validate an editable descriptor using the same rules as normal SCOPE."""
    try:
        profile = _from_dict(data, package_dir)
    except Exception as exc:
        return None, [("PROFILE IDENTITY", False, str(exc))]
    return profile, validate_profile_package(profile)


def install_profile_package(data: dict, resource_files: dict[str, str | Path] | None = None) -> TestProfile:
    """Install/update a user profile and copy its approved reusable resources."""
    profile = _from_dict(data)
    destination = USER_PROFILES_DIR / profile.id
    destination.mkdir(parents=True, exist_ok=True)
    resources = dict(data.get("resources") or {})
    for role, source_value in (resource_files or {}).items():
        source = Path(source_value)
        if not source.is_file():
            continue
        resource_dir = destination / "resources"
        resource_dir.mkdir(parents=True, exist_ok=True)
        target = resource_dir / source.name
        shutil.copy2(source, target)
        resources[role] = str(target.relative_to(destination))
    payload = json.loads(json.dumps(data)); payload["resources"] = resources
    copied_documents = []
    for document in payload.get("source_documents") or []:
        if not isinstance(document, dict):
            continue
        item = dict(document)
        source = Path(str(item.get("path", "")))
        if source.is_file():
            source_dir = destination / "sources"
            source_dir.mkdir(parents=True, exist_ok=True)
            target = source_dir / source.name
            try:
                same_file = source.resolve() == target.resolve()
            except OSError:
                same_file = False
            if not same_file:
                shutil.copy2(source, target)
            item["path"] = str(target.relative_to(destination))
        copied_documents.append(item)
    payload["source_documents"] = copied_documents
    payload["last_modified"] = datetime.now(timezone.utc).isoformat()
    reporting = payload.get("reporting_specification") or payload.get("reporting") or {}
    recruitment = reporting.get("recruitment") or {}
    for study_type, specification in recruitment.items():
        if not isinstance(specification, dict):
            continue
        source = Path(str(specification.get("template", "")))
        if source.is_file():
            template_dir = destination / "templates" / "recruitment"
            template_dir.mkdir(parents=True, exist_ok=True)
            target = template_dir / source.name
            shutil.copy2(source, target)
            specification["template"] = str(target.relative_to(destination))
    descriptor = destination / "profile.yaml"
    descriptor.write_text(yaml.safe_dump(payload, sort_keys=False, allow_unicode=True), encoding="utf-8")
    return _from_dict(payload, destination, descriptor)


def save_user_profile(data: dict) -> TestProfile:
    """Save an extensibility profile outside bundled, versioned profiles."""
    payload = json.loads(json.dumps(data))
    payload["last_modified"] = datetime.now(timezone.utc).isoformat()
    profile = _from_dict(payload)
    USER_PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    path = USER_PROFILES_DIR / f"{profile.id}.json"
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return _from_dict(payload, path.parent, path)


# ---------------------------------------------------------------------------
# Progressive-disclosure profile configuration (sample filters, target
# interpretation, optimization tolerances, and rule provenance).  Everything
# below is pure python so the pandas-free Profile Manager build can use it.
# ---------------------------------------------------------------------------

ORIGIN_LABELS = {
    "inherited": "Inherited from parent profile",
    "assessment": "From assessment sources",
    "manual": "Manual / RD decision",
    "default": "SCOPE default",
}

PROFILE_KINDS = ("standard", "international_adaptation")

SAMPLE_FILTER_FIELDS = (
    ("study_type", "Study"),
    ("form", "Form"),
    ("language", "Language"),
    ("sex", "Sex"),
    ("region", "Region"),
    ("ethnicity", "Ethnicity"),
    ("education", "Education (PED/SED)"),
    ("age_group", "Age group"),
)

TARGET_SOURCE_FORMATS = (
    ("crossed_cells", "Fully crossed cell-by-cell targets (e.g. a census workbook)"),
    ("marginal_rows", "Marginal / one-row-per-group targets"),
    ("scope_template", "SCOPE target template"),
    ("mixed", "Mixed / multiple formats"),
)

TARGET_CELLS_AS = (
    ("report_only", "Report only — optimize the derived marginal/group targets"),
    ("objective", "Keep cells as CP-SAT objective rows as well"),
    ("exclude", "Exclude cell rows entirely"),
)

TARGET_DIMENSION_ALIASES = {
    "age_group": ("target age band", "age band", "age group", "agegrp"),
    "sex": ("sex", "gender"),
    "region": ("region", "reg"),
    "ethnicity": ("ethnicity", "eth", "race ethnicity"),
    "education": ("ped/sed", "ped", "sed", "education", "education level"),
}

DIMENSION_DISPLAY = {
    "age_group": "Age Group",
    "sex": "Sex",
    "region": "Region",
    "ethnicity": "Ethnicity",
    "education": "Education",
}


def _norm(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().casefold())


def dimension_semantic(value: object) -> str:
    """Resolve a dimension label (canonical or workbook alias) to its semantic name."""
    key = _norm(value)
    if not key:
        return ""
    for semantic, aliases in TARGET_DIMENSION_ALIASES.items():
        if key == _norm(semantic) or key in {_norm(alias) for alias in aliases}:
            return semantic
    return ""


def normalize_sample_filter(item: dict) -> dict:
    """Normalize one structured profile sample-filter rule."""
    field = str(item.get("field", "")).strip()
    key = _norm(field)
    semantic = next((name for name, _label in SAMPLE_FILTER_FIELDS if _norm(name) == key), "") or dimension_semantic(field) or key
    label = str(item.get("label", "")).strip() or dict(SAMPLE_FILTER_FIELDS).get(semantic, field.replace("_", " ").title())
    operator = str(item.get("operator", "include")).strip().casefold()
    if operator not in {"include", "exclude"}:
        operator = "include"
    raw_values = item.get("values", [])
    if isinstance(raw_values, str):
        raw_values = [value.strip() for value in re.split(r"[,;]", raw_values)]
    values = [str(value).strip() for value in raw_values if str(value or "").strip()]
    origin = str(item.get("origin", "manual")).strip().casefold() or "manual"
    return {
        "field": semantic or field,
        "label": label,
        "operator": operator,
        "values": values,
        "origin": origin,
        "note": str(item.get("note", "")).strip(),
    }


def profile_sample_filters(profile: TestProfile) -> list[dict]:
    """Return the profile's structured sample filter / eligibility rules."""
    return [normalize_sample_filter(item) for item in (profile.sample_filters or [])]


def profile_study_type_defaults(profile: TestProfile) -> list[str]:
    """Study types the profile includes by default via its sample rules."""
    values: list[str] = []
    for rule in profile_sample_filters(profile):
        if rule.get("field") == "study_type" and rule.get("operator") == "include":
            for value in rule.get("values", []):
                if value not in values:
                    values.append(value)
    return values


def format_sample_filter(item: dict) -> str:
    """Human-readable mirror of a structured sample filter, stored in eligibility_rules."""
    normalized = normalize_sample_filter(item)
    joined = ", ".join(normalized["values"])
    verb = "=" if normalized["operator"] == "include" else "≠"
    text = f"{normalized['label']} {verb} {joined}"
    return f"{text} — {normalized['note']}" if normalized["note"] else text


def normalize_optimization_rule(category: str, spec: object) -> dict:
    """Normalize one optimization rule for a demographic category or group.

    A rule may carry tolerances (tolerance_percent / tolerance_n) or an explicit
    target count (target_n, with applies_to = "all" or "group" + group label).
    """
    if isinstance(spec, (int, float)):
        spec = {"tolerance_percent": spec}
    spec = dict(spec) if isinstance(spec, dict) else {}
    category_text = str(category).strip()
    if " | " in category_text:
        category_text, _, grouped = category_text.partition(" | ")
        category_text = category_text.strip()
        spec.setdefault("applies_to", "group")
        spec.setdefault("group", grouped.strip())
    rule = str(spec.get("rule", "Target")).strip().title() or "Target"
    if rule not in {"Target", "Range", "Minimum", "Maximum", "Exact"}:
        rule = "Target"
    semantic = dimension_semantic(category_text) or _norm(category_text)
    applies_to = str(spec.get("applies_to", "all")).strip().casefold() or "all"
    if applies_to not in {"all", "group"}:
        applies_to = "all"
    group = str(spec.get("group", "")).strip() if applies_to == "group" else ""
    return {
        "category": category_text,
        "semantic": semantic,
        "applies_to": applies_to,
        "group": group,
        "rule": rule,
        "target_n": spec.get("target_n", ""),
        "tolerance_percent": spec.get("tolerance_percent", ""),
        "tolerance_n": spec.get("tolerance_n", ""),
        "priority": str(spec.get("priority", "Demographic")).strip() or "Demographic",
        "origin": str(spec.get("origin", "manual")).strip().casefold() or "manual",
        "note": str(spec.get("note", "")).strip(),
    }


def profile_optimization_rules(profile: TestProfile) -> list[dict]:
    """Return normalized per-category tolerance/priority rules for optimization."""
    return [normalize_optimization_rule(category, spec) for category, spec in (profile.target_tolerances or {}).items()]


def describe_profile_optimization_rules(profile: TestProfile) -> str:
    """One-line summary of the profile's optimization tolerances / target Ns."""
    parts: list[str] = []
    for rule in profile_optimization_rules(profile):
        scope = "all groups" if rule.get("applies_to") == "all" else f"only {rule.get('group')}"
        if rule.get("target_n") not in (None, ""):
            parts.append(f"{rule['category']} ({scope}) = {rule['target_n']} N")
        elif rule.get("tolerance_percent") not in (None, ""):
            parts.append(f"{rule['category']} ({scope}) ±{rule['tolerance_percent']}%")
        elif rule.get("tolerance_n") not in (None, ""):
            parts.append(f"{rule['category']} ({scope}) ±{rule['tolerance_n']} N")
        else:
            parts.append(f"{rule['category']} ({scope}) {rule['rule']}")
    return "; ".join(parts)


def describe_target_interpretation(profile: TestProfile) -> str:
    """One-line summary of how the profile turns target sources into the objective."""
    interpretation = profile.target_interpretation or {}
    if not interpretation:
        return ""
    labels = dict(TARGET_SOURCE_FORMATS)
    source_format = str(interpretation.get("source_format", ""))
    derived = [DIMENSION_DISPLAY.get(dimension_semantic(value) or str(value), str(value)) for value in interpretation.get("derive_marginals") or []]
    cells = str(interpretation.get("cells_as", "report_only"))
    cells_label = {"report_only": "raw crossed cells are report-only", "objective": "raw crossed cells also stay in the CP-SAT objective", "exclude": "raw crossed cells are excluded"}.get(cells, cells)
    if source_format == "crossed_cells" and derived:
        return f"{labels.get(source_format, source_format)} — SCOPE derives marginal targets for {', '.join(derived)}; {cells_label}."
    return f"Source target format: {labels.get(source_format, source_format) or 'as provided'}; {cells_label}."


def _rule_number(value: object) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(str(value).strip().rstrip("%"))
    except (TypeError, ValueError):
        return None


def target_row_criteria(row: dict) -> dict[str, str]:
    """Parse canonical dimension criteria from a target row without pandas."""
    raw = row.get("criteria", "")
    if isinstance(raw, dict):
        return {str(key): str(value) for key, value in raw.items() if str(value).strip()}
    text = str(raw or "").strip()
    if text:
        try:
            parsed = json.loads(text)
            if isinstance(parsed, dict):
                return {str(key): str(value) for key, value in parsed.items() if str(value).strip()}
        except (TypeError, ValueError, json.JSONDecodeError):
            pass
    semantic = dimension_semantic(row.get("category", ""))
    group = str(row.get("group", "")).strip()
    return {semantic: group} if semantic and group else {}


def _cell_share(row: dict) -> tuple[str, float] | None:
    """Return ('%', fraction) or ('N', count) for one crossed cell row."""
    raw = row.get("target", "")
    numbers = re.findall(r"-?\d+(?:\.\d+)?", str(raw))
    if not numbers:
        return None
    value = float(numbers[0])
    unit = str(row.get("unit", "")).strip().casefold()
    if unit in {"%", "percent", "percentage"}:
        return ("%", value / 100 if value > 1 else value)
    if unit in {"n", "count"}:
        return ("N", value)
    return None


def derive_marginal_target_rows(
    rows: list[dict],
    dimensions: list[str],
    tolerances: dict | None = None,
    source_label: str = "Derived marginal targets",
) -> list[dict]:
    """Collapse fully crossed cell-by-cell targets into per-dimension marginal rows.

    International adaptations frequently ship a census workbook of fully crossed
    cells (age × sex × region × ethnicity × education).  Those raw cells are the
    *source* target format; the CP-SAT objective should instead use derived
    marginal/group targets with the profile's tolerances (e.g. Region ±10%).
    """
    wanted: list[str] = []
    for value in dimensions or []:
        semantic = dimension_semantic(value)
        if semantic and semantic not in wanted:
            wanted.append(semantic)
    crossed = [(row, target_row_criteria(row)) for row in rows or []]
    crossed = [(row, criteria) for row, criteria in crossed if len(criteria) >= 2]
    if not wanted or not crossed:
        return []
    percent_shares: list[tuple[dict[str, str], float]] = []
    count_shares: list[tuple[dict[str, str], float]] = []
    for row, criteria in crossed:
        share = _cell_share(row)
        if share is None:
            continue
        (percent_shares if share[0] == "%" else count_shares).append((criteria, share[1]))
    shares = percent_shares
    if not shares and count_shares:
        total = sum(value for _criteria, value in count_shares)
        if total > 0:
            shares = [(criteria, value / total) for criteria, value in count_shares]
    if not shares:
        return []
    rules = [normalize_optimization_rule(category, spec) for category, spec in (tolerances or {}).items()]
    derived: list[dict] = []
    for semantic in wanted:
        levels: dict[str, float] = {}
        for criteria, share in shares:
            level = str(criteria.get(semantic, "")).strip()
            if level:
                levels[level] = levels.get(level, 0.0) + share
        spec = next((rule for rule in rules if rule.get("semantic") == semantic and rule.get("applies_to") == "all"), None) or {}
        for level in sorted(levels):
            target_n = _rule_number(spec.get("target_n"))
            tolerance_percent = _rule_number(spec.get("tolerance_percent"))
            rule_name = str(spec.get("rule", "Target"))
            group_spec = next((rule for rule in rules if rule.get("semantic") == semantic and rule.get("applies_to") == "group" and rule.get("group") == level), None)
            effective = group_spec or spec
            if group_spec and _rule_number(group_spec.get("target_n")) is not None:
                target_n = _rule_number(group_spec.get("target_n"))
                tolerance_percent = _rule_number(group_spec.get("tolerance_percent"))
                rule_name = str(group_spec.get("rule", "Exact"))
            if target_n is not None:
                rule_value, target_value, unit = "Exact", f"{target_n:g}", "N"
            elif tolerance_percent:
                pct = levels[level] * 100
                low, high = max(0.0, pct - tolerance_percent), min(100.0, pct + tolerance_percent)
                rule_value, target_value, unit = "Range", f"{low:g}–{high:g}", "%"
            elif rule_name == "Range":
                pct = levels[level] * 100
                rule_value, target_value, unit = "Target", f"{pct:g}", "%"
            else:
                pct = levels[level] * 100
                rule_value, target_value, unit = rule_name, f"{pct:g}", "%"
            derived.append({
                "category": DIMENSION_DISPLAY.get(semantic, semantic.replace("_", " ").title()),
                "group": level,
                "rule": rule_value,
                "target": target_value,
                "unit": unit,
                "tolerance": effective.get("tolerance_n", ""),
                "priority": str(effective.get("priority", "Demographic")),
                "required": True,
                "source": source_label,
                "criteria": json.dumps({semantic: level}, sort_keys=True),
                "origin": str(effective.get("origin", "manual")),
            })
    return derived
