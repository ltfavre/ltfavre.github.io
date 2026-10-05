"""One serializable project configuration shared by GUI and CLI."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path
import re


@dataclass
class ProjectConfig:
    profile_id: str
    project_dir: str
    profile_version: str = ""
    project_file: str = ""
    project_name: str = ""
    workflow: str = "sample_optimization"
    dataset: str = ""
    candidate_datasets: list[dict] = field(default_factory=list)
    master_file: str = ""
    master_sheet: str = "final_all"
    fr_teid_file: str = ""
    scored_dir: str = ""
    case_swap_dir: str = ""
    target_file: str = ""
    recruitment_plan_file: str = ""
    recruitment_actuals_file: str = ""
    swap_sample_file: str = ""
    swap_available_file: str = ""
    swap_case_universe_file: str = ""
    swap_sample_data_file: str = ""
    swap_available_data_file: str = ""
    targeted_swaps: list[dict] = field(default_factory=list)
    output_dir: str = ""
    output_selection: str = "both"
    forms: list[str] = field(default_factory=list)
    language: str = "English"
    study_type: str = "STDZ"
    use_existing_sample: bool = False
    clinical_seeding: bool = False
    filters: dict[str, list[str]] = field(default_factory=dict)
    demographics: list[str] = field(default_factory=list)
    demographic_tables_1way: list[str] = field(default_factory=list)
    demographic_tables_2way: list[str] = field(default_factory=list)
    demographic_tables_3way: list[str] = field(default_factory=list)
    demographic_tables_configured: bool = False
    include_census: bool = True
    include_osep: bool = True
    include_clinical: bool = True
    sample_size: int | None = None
    age_bands: list[dict] = field(default_factory=list)
    target_n_by_agegrp: dict[str, int] = field(default_factory=dict)
    canonical_targets: list[dict] = field(default_factory=list)
    target_overrides: list[dict] = field(default_factory=list)
    target_review_confirmed: bool = False
    locks: list[str] = field(default_factory=lambda: ["vform", "agegrp"])
    exclusions: list[str] = field(default_factory=list)
    target_tolerance: int = 0
    max_solver_seconds: int = 300
    engine_version: str = "SCOPE-1.0"
    run_history: list[dict] = field(default_factory=list)

    def input_value(self, role: str) -> str:
        """Return a canonical input role without exposing legacy field names to clients."""
        attribute = {
            "master_dataset": "dataset", "current_sample": "master_file",
            "scored_data": "scored_dir", "target_specification": "target_file",
            "initial_teid_source": "fr_teid_file", "field_research_report": "fr_teid_file", "case_review_data": "case_swap_dir",
            "recruitment_targets": "recruitment_plan_file", "recruitment_progress": "recruitment_actuals_file",
            "swap_sample": "swap_sample_file", "swap_available_pool": "swap_available_file",
            "swap_case_universe": "swap_case_universe_file",
            "swap_sample_data": "swap_sample_data_file", "swap_available_data": "swap_available_data_file",
        }.get(role)
        return str(getattr(self, attribute, "") or "")

    def input_sources(self) -> dict[str, str]:
        roles = (
            "master_dataset", "current_sample", "scored_data", "target_specification",
            "initial_teid_source", "field_research_report", "case_review_data", "recruitment_targets",
            "recruitment_progress", "swap_sample", "swap_available_pool", "swap_case_universe",
            "swap_sample_data", "swap_available_data",
        )
        return {role: self.input_value(role) for role in roles}

    @property
    def path(self) -> Path:
        name = re.sub(r"[^A-Za-z0-9._-]+", "_", self.project_name.strip()).strip("._") or "SCOPE_Project"
        return Path(self.project_dir) / f"{name}.scope"

    def save(self, path: Path | None = None) -> Path:
        output = path or self.path
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(self.to_storage(), indent=2), encoding="utf-8")
        self.project_file = str(output.resolve())
        return output

    def to_storage(self) -> dict:
        """Stable, readable project file; runtime attributes remain backward-compatible."""
        return {
            "workflow": self.workflow,
            "test_profile": {"id": self.profile_id, "version": self.profile_version},
            "project_configuration": {
                "project_name": self.project_name, "sample_size": self.sample_size, "forms": self.forms, "language": self.language,
                "study_type": self.study_type, "include_census": self.include_census,
                "use_existing_sample": self.use_existing_sample,
                "include_osep": self.include_osep, "include_clinical": self.include_clinical,
                "clinical_seeding": self.clinical_seeding, "demographics": self.demographics,
                "demographic_tables": {
                    "configured": self.demographic_tables_configured,
                    "one_way": self.demographic_tables_1way,
                    "two_way": self.demographic_tables_2way,
                    "three_way": self.demographic_tables_3way,
                },
                "age_bands": self.age_bands, "target_n_by_agegrp": self.target_n_by_agegrp,
                "filters": self.filters, "locks": self.locks, "exclusions": self.exclusions,
                "target_tolerance": self.target_tolerance, "targeted_swaps": self.targeted_swaps,
                "target_overrides": self.target_overrides,
            },
            "inputs": {
                "case_universe_sources": {
                    "master_dataset": self.dataset, "candidate_datasets": self.candidate_datasets, "scored_data": self.scored_dir,
                    "field_research_report": self.fr_teid_file, "case_review_data": self.case_swap_dir,
                },
                "current_sample": {"path": self.master_file, "sheet": self.master_sheet},
                "target_sources": {"target_specification": self.target_file},
                "recruitment": {"targets": self.recruitment_plan_file, "progress": self.recruitment_actuals_file},
                "targeted_swap": {
                    "sample": self.swap_sample_file,
                    "available_pool": self.swap_available_file,
                    "case_universe": self.swap_case_universe_file,
                    "selected_case_data": self.swap_sample_data_file,
                    "available_case_data": self.swap_available_data_file,
                },
            },
            "canonical_targets": {"confirmed": self.target_review_confirmed, "rows": self.canonical_targets},
            "output": {
                "directory": self.output_dir or str(Path(self.project_dir) / "scope_outputs"),
                "project_directory": self.project_dir,
                "selection": self.output_selection,
            },
            "runs": self.run_history,
            "engine_version": self.engine_version,
        }

    @classmethod
    def load(cls, path: str | Path) -> "ProjectConfig":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        if "test_profile" not in data:
            return cls(**data)
        project = data.get("project_configuration") or {}; inputs = data.get("inputs") or {}; output = data.get("output") or {}
        universe = inputs.get("case_universe_sources") or {}
        candidate_datasets = list(universe.get("candidate_datasets") or [])
        legacy_dataset = str(universe.get("master_dataset", inputs.get("dataset", "")))
        if not candidate_datasets and legacy_dataset:
            candidate_datasets = [{"path": legacy_dataset, "source_type": "STDZ", "source_role": "Primary", "priority": 1}]
        current = inputs.get("current_sample") or ""
        current_path = current.get("path", "") if isinstance(current, dict) else current
        current_sheet = current.get("sheet", "final_all") if isinstance(current, dict) else inputs.get("current_sample_sheet", "final_all")
        targets = inputs.get("target_sources") or {}
        recruitment = inputs.get("recruitment") or {}
        targeted_swap = inputs.get("targeted_swap") or {}
        canonical_raw = data.get("canonical_targets") or {}
        canonical = canonical_raw if isinstance(canonical_raw, dict) else {"rows": canonical_raw, "confirmed": False}
        report_tables = project.get("demographic_tables") or {}
        if not isinstance(report_tables, dict):
            report_tables = {}
        return cls(
            profile_id=str((data.get("test_profile") or {}).get("id", "")),
            project_dir=str(output.get("project_directory") or Path(path).resolve().parent),
            profile_version=str((data.get("test_profile") or {}).get("version", "")), project_file=str(Path(path).resolve()),
            project_name=str(project.get("project_name", "")),
            workflow=str(data.get("workflow", "sample_optimization")),
            dataset=legacy_dataset, candidate_datasets=candidate_datasets, master_file=str(current_path),
            master_sheet=str(current_sheet), fr_teid_file=str(universe.get("field_research_report", inputs.get("field_research_teids", ""))),
            scored_dir=str(universe.get("scored_data", inputs.get("scored_data", ""))), case_swap_dir=str(universe.get("case_review_data", inputs.get("case_swap_data", ""))),
            target_file=str(targets.get("target_specification", inputs.get("optimization_targets", ""))),
            recruitment_plan_file=str(recruitment.get("targets", inputs.get("recruitment_plan", ""))),
            recruitment_actuals_file=str(recruitment.get("progress", inputs.get("recruitment_actuals", ""))),
            swap_sample_file=str(targeted_swap.get("sample", inputs.get("swap_sample", ""))),
            swap_available_file=str(targeted_swap.get("available_pool", inputs.get("swap_available_pool", ""))),
            swap_case_universe_file=str(targeted_swap.get("case_universe", inputs.get("swap_case_universe", ""))),
            swap_sample_data_file=str(targeted_swap.get("selected_case_data", inputs.get("swap_sample_data", ""))),
            swap_available_data_file=str(targeted_swap.get("available_case_data", inputs.get("swap_available_data", ""))),
            targeted_swaps=list(project.get("targeted_swaps") or []),
            output_dir=str(output.get("directory", "")),
            output_selection=str(output.get("selection", "both")),
            forms=list(project.get("forms") or []), language=str(project.get("language", "English")),
            study_type=str(project.get("study_type", "STDZ")), use_existing_sample=bool(project.get("use_existing_sample", bool(current_path))), filters=dict(project.get("filters") or {}),
            demographics=list(project.get("demographics") or []), include_census=bool(project.get("include_census", True)),
            demographic_tables_1way=list(report_tables.get("one_way") or project.get("demographic_tables_1way") or []),
            demographic_tables_2way=list(report_tables.get("two_way") or project.get("demographic_tables_2way") or []),
            demographic_tables_3way=list(report_tables.get("three_way") or project.get("demographic_tables_3way") or []),
            demographic_tables_configured=bool(report_tables.get("configured", project.get("demographic_tables_configured", False))),
            include_osep=bool(project.get("include_osep", True)), include_clinical=bool(project.get("include_clinical", True)),
            clinical_seeding=bool(project.get("clinical_seeding", False)), sample_size=project.get("sample_size"),
            age_bands=list(project.get("age_bands") or []), target_n_by_agegrp=dict(project.get("target_n_by_agegrp") or {}),
            canonical_targets=list(canonical.get("rows") or []), target_review_confirmed=bool(canonical.get("confirmed", False)),
            target_overrides=list(project.get("target_overrides") or []),
            locks=list(project.get("locks") or ["vform", "agegrp"]), exclusions=list(project.get("exclusions") or []),
            target_tolerance=int(project.get("target_tolerance", 0)), engine_version=str(data.get("engine_version", "SCOPE-1.0")), run_history=list(data.get("runs") or []),
        )

    def scope_profile(self, profile_name: str, field_map: dict[str, str] | None = None) -> dict:
        return {
            "version": 1, "project_test_name": profile_name, "filters": self.filters,
            "scope_filter": {}, "demographics": self.demographics, "age_bands": self.age_bands,
            "demographic_tables": {
                "configured": self.demographic_tables_configured,
                "one_way": self.demographic_tables_1way,
                "two_way": self.demographic_tables_2way,
                "three_way": self.demographic_tables_3way,
            },
            "target_n_per_stratum": None, "target_n_by_agegrp": self.target_n_by_agegrp,
            "field_map": field_map or {}, "clinical_seeding": self.clinical_seeding,
            "locks": self.locks, "exclusions": self.exclusions,
            "canonical_targets": self.canonical_targets,
        }
