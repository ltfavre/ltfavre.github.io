
from pathlib import Path
from datetime import datetime
import json
import os
import sys
import shutil
import pandas as pd
from openpyxl import load_workbook

# =====================================================
SCRIPT_DIR = Path(__file__).resolve().parent
APP_DIR = SCRIPT_DIR.parent if SCRIPT_DIR.name.lower() == "scripts" else SCRIPT_DIR
APP_SCRIPTS_DIR = APP_DIR / "scripts"

# CONFIG
# =====================================================

PROJECT_DIR = APP_DIR

INPUT_FILE = APP_DIR / "inputs" / "raw_data" / "stratification_summary.xlsx"
TOLERANCE_FILE = APP_DIR / "inputs" / "raw_data" / "stratification_tolerance_rules.xlsx"

FR_TEID_FILE = APP_DIR / "inputs" / "raw_data" / "fr_teid_v1.xlsx"

SCORED_DATA_DIR = APP_DIR / "inputs" / "raw_data"

OUTPUT_DIR = APP_DIR / "outputs" / "qa_tables"
PREP_RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
PACKAGE_APP_DIR = APP_DIR
PACKAGE_CASE_SWAP_PATTERN = "*_case_swap.xlsx"
PACKAGE_TARGET_WORKBOOK = str(APP_DIR / "inputs" / "targets" / "Demo _OSEP_Targets _Template.xlsx")
PACKAGE_CONFIG_WORKBOOK = str(APP_SCRIPTS_DIR / "Stratification_Optimizer_Config.xlsx")
PACKAGE_EXE = str(APP_SCRIPTS_DIR / "StratificationOptimizer.exe")
PACKAGE_CLEAR_EXISTING_POOLS = False
WORKFLOW_SHEET_NAMES = ["START_HERE_WORKFLOW", "START_HERE"]
WORKFLOW_SELECTION_CELL = "G8"
WORKFLOW_RESPONSE_VALUES = {
    "Y",
    "YES",
    "N",
    "NO",
    "BUILD OPTIMIZER PACKAGE FROM RAW DATA",
    "OPTIMIZE EXISTING PREPARED DATA",
}

MASTER_OUTPUT_NAME = "MASTER_SAMPLE.xlsx"
TARGET_OUTPUT_NAME = "Demo _OSEP_Targets _Template.xlsx"
CONFIG_OUTPUT_NAME = "Stratification_Optimizer_Config.xlsx"

POOL_RULES = {
    "TRS_C": {
        "file": "basc4nb_scored_trsc.sas7bdat",
        "testid": [5882, 5797],
        "agegrp": [3, 4],
    },
    "TRS_A": {
        "file": "basc4nb_scored_trsa.sas7bdat",
        "testid": [5886, 5800],
        "agegrp": [5, 6],
    },
    "PRS_C": {
        "file": "basc4nb_scored_prsc.sas7bdat",
        "testid": [5883, 5798],
        "agegrp": [3, 4],
    },
    "PRS_A": {
        "file": "basc4nb_scored_prsa.sas7bdat",
        "testid": [5887, 5801],
        "agegrp": [5, 6],
    },
    "SRP_C": {
        "file": "basc4nb_scored_srpc.sas7bdat",
        "testid": [5885, 5799],
        "agegrp": [4],
    },
    "SRP_A": {
        "file": "basc4nb_scored_srpa.sas7bdat",
        "testid": [5888, 5802],
        "agegrp": [5, 6],
    },
}


# =====================================================
# USER SETTINGS — EDIT THIS SECTION ONLY
# =====================================================
##ADHD##
# PROJECT = "ADHD"
# VFORM = "TRS_C"
# TARGET_FLAG = "tmp_ADHD"
# TARGET_PCT = 5.0

# PROJECT = "ADHD"
# VFORM = "TRS_A"
# TARGET_FLAG = "tmp_ADHD"
# TARGET_PCT = 5.0

# PROJECT = "ADHD"
# VFORM = "PRS_C"
# TARGET_FLAG = "tmp_ADHD"
# TARGET_PCT = 5.0

# PROJECT = "ADHD"
# VFORM = "PRS_A"
# TARGET_FLAG = "tmp_ADHD"
# TARGET_PCT = 5.0

# PROJECT = "ADHD"
# VFORM = "SRP_C"
# TARGET_FLAG = "tmp_ADHD"
# TARGET_PCT = 5.0

PROJECT = "ADHD"
VFORM = "SRP_A"
TARGET_FLAG = "tmp_ADHD"
TARGET_PCT = 5.0

##ASD##
# PROJECT = "ASD"
# VFORM = "TRS_C"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 2.0

# PROJECT = "ASD"
# VFORM = "TRS_A"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 2.0
#
# PROJECT = "ASD"
# VFORM = "PRS_C"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 2.0
#
# PROJECT = "ASD"
# VFORM = "PRS_A"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 2.0

##MOOD##
# PROJECT = "ASD"
# VFORM = "TRS_C"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 0.7
#
# PROJECT = "MOOD"
# VFORM = "TRS_A"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 0.7
#
# PROJECT = "MOOD"
# VFORM = "PRS_C"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 0.7
#
# PROJECT = "MOOD"
# VFORM = "PRS_A"
# TARGET_FLAG = "tmp_ASD"
# TARGET_PCT = 0.7

AGEGRP_TO_RUN = 6

##Expected N is set to 200 for all groups except SRP_C agegrp 4.
##SRP_C agegrp 4---N is set to 400.
EXPECTED_N = 200

MAX_SWAPS = 8
MAX_ADD_CANDIDATES_PER_STEP = 12
FORM_TYPE = VFORM.split("_")[0]


#THIS IS THE LIST OF FORMS TO RUN
# TRS_C agegrp 3, EXPECTED_N = 200
# TRS_C agegrp 4, EXPECTED_N = 200
# TRS_A agegrp 5, EXPECTED_N = 200   with N repair
# TRS_A agegrp 6, EXPECTED_N = 200
#
# PRS_C agegrp 3, EXPECTED_N = 200
# PRS_C agegrp 4, EXPECTED_N = 200
# PRS_A agegrp 5, EXPECTED_N = 200
# PRS_A agegrp 6, EXPECTED_N = 200
#
# SRP_C agegrp 4, EXPECTED_N = 400
# SRP_A agegrp 5, EXPECTED_N = 200
# SRP_A agegrp 6, EXPECTED_N = 200

STRAT_VARS = [
    "sex",
    "eth",
    "edl",
    "region",
    "testlang",
    "edition",
]

FLAG_COLS = [
    "tmp_ADHD",
    "tmp_EBD",
    "tmp_ASD",
    "tmp_DD",
    "tmp_SLD",
    "tmp_SLI",
    "tmp_OTHER",
    "tmp_clin",
]


def to_bool_from_user_value(value, default=True):
    if value is None:
        return default
    text = str(value).strip().upper()
    if text in {
        "Y",
        "YES",
        "TRUE",
        "1",
        "RUN",
        "RAW",
        "RAW_DATA",
        "A",
        "BUILD OPTIMIZER PACKAGE FROM RAW DATA",
    } or text.startswith("BUILD OPTIMIZER PACKAGE"):
        return True
    if text in {
        "N",
        "NO",
        "FALSE",
        "0",
        "SKIP",
        "PREPPED",
        "PREPPED_DATA",
        "B",
        "OPTIMIZE EXISTING PREPARED DATA",
    } or text.startswith("OPTIMIZE EXISTING"):
        return False
    return default


def find_workflow_response(ws):
    preferred_value = ws[WORKFLOW_SELECTION_CELL].value
    preferred_text = str(preferred_value).strip().upper() if preferred_value is not None else ""
    if preferred_text in WORKFLOW_RESPONSE_VALUES:
        return WORKFLOW_SELECTION_CELL, preferred_value

    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            text = str(value).strip().upper() if value is not None else ""
            if text in WORKFLOW_RESPONSE_VALUES:
                return cell.coordinate, value

    return WORKFLOW_SELECTION_CELL, preferred_value


def resolve_config_workbook_path():
    candidate = None

    if len(sys.argv) > 1 and str(sys.argv[1]).strip():
        candidate = Path(sys.argv[1]).expanduser()
    elif os.environ.get("STRAT_CONFIG_WORKBOOK"):
        candidate = Path(os.environ["STRAT_CONFIG_WORKBOOK"]).expanduser()
    elif PACKAGE_CONFIG_WORKBOOK:
        candidate = Path(PACKAGE_CONFIG_WORKBOOK).expanduser()

    if candidate is None:
        return None

    if not candidate.is_absolute():
        candidate = SCRIPT_DIR / candidate

    return candidate.resolve()


def load_workflow_selector(config_workbook_path):
    selector = {
        "sheet_name": None,
        "cell": WORKFLOW_SELECTION_CELL,
        "run_data_prep_first": True,
        "raw_value": None,
    }

    if config_workbook_path is None or not config_workbook_path.exists():
        return selector

    try:
        wb = load_workbook(config_workbook_path, data_only=True, read_only=True)
    except Exception:
        return selector

    sheet_name = next((name for name in WORKFLOW_SHEET_NAMES if name in wb.sheetnames), None)
    if sheet_name is None:
        return selector

    ws = wb[sheet_name]
    selector["sheet_name"] = sheet_name
    selector["cell"], selector["raw_value"] = find_workflow_response(ws)
    selector["run_data_prep_first"] = to_bool_from_user_value(
        selector["raw_value"],
        default=True,
    )
    return selector


CONFIG_WORKBOOK_PATH = resolve_config_workbook_path()
if CONFIG_WORKBOOK_PATH is not None:
    PACKAGE_CONFIG_WORKBOOK = str(CONFIG_WORKBOOK_PATH)

WORKFLOW_SELECTOR = load_workflow_selector(CONFIG_WORKBOOK_PATH)
if not WORKFLOW_SELECTOR.get("run_data_prep_first", True):
    print(
        "START_HERE_WORKFLOW indicates prep is not required "
        f"(cell {WORKFLOW_SELECTOR.get('cell')}={WORKFLOW_SELECTOR.get('raw_value')}). "
        "Skipping Script 01."
    )
    raise SystemExit(0)


def load_pipeline_config():
    config_path = os.environ.get("STRAT_CONFIG")

    if not config_path:
        return None

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_project_path(value):
    path = Path(value)
    return path if path.is_absolute() else PROJECT_DIR / path


pipeline_config = load_pipeline_config()

if pipeline_config is not None:
    PROJECT_DIR = resolve_project_path(
        pipeline_config.get("project_dir", str(PROJECT_DIR))
    )
    INPUT_FILE = resolve_project_path(
        pipeline_config.get("stratification_summary_file", str(INPUT_FILE))
    )
    TOLERANCE_FILE = resolve_project_path(
        pipeline_config.get("tolerance_file", str(TOLERANCE_FILE))
    )
    FR_TEID_FILE = resolve_project_path(
        pipeline_config.get("fr_teid_file", str(FR_TEID_FILE))
    )
    SCORED_DATA_DIR = resolve_project_path(
        pipeline_config.get("scored_data_dir", str(SCORED_DATA_DIR))
    )
    OUTPUT_DIR = resolve_project_path(
        pipeline_config.get("qa_output_dir", str(OUTPUT_DIR))
    )
    PACKAGE_APP_DIR = resolve_project_path(
        pipeline_config.get("package_app_dir", str(PACKAGE_APP_DIR))
    )
    PACKAGE_CASE_SWAP_PATTERN = pipeline_config.get(
        "package_case_swap_pattern",
        PACKAGE_CASE_SWAP_PATTERN
    )
    PACKAGE_TARGET_WORKBOOK = pipeline_config.get(
        "target_workbook",
        PACKAGE_TARGET_WORKBOOK
    )
    PACKAGE_CONFIG_WORKBOOK = pipeline_config.get(
        "optimizer_config_workbook",
        PACKAGE_CONFIG_WORKBOOK
    )
    PACKAGE_EXE = pipeline_config.get("optimizer_exe", PACKAGE_EXE)
    PACKAGE_CLEAR_EXISTING_POOLS = bool(
        pipeline_config.get(
            "package_clear_existing_pools",
            PACKAGE_CLEAR_EXISTING_POOLS
        )
    )

    PROJECT = pipeline_config.get("project", PROJECT)
    TARGET_FLAG = pipeline_config.get("target_flag", TARGET_FLAG)
    TARGET_PCT = pipeline_config.get("target_pct", TARGET_PCT)
    FLAG_COLS = pipeline_config.get("flag_cols", FLAG_COLS)
    POOL_RULES = pipeline_config.get("pool_rules", POOL_RULES)

    group_index = int(os.environ.get("STRAT_GROUP_INDEX", "0"))
    group = pipeline_config["swap_groups"][group_index]

    VFORM = group["vform"]
    AGEGRP_TO_RUN = int(group["agegrp"])
    EXPECTED_N = int(group["expected_n"])
    TARGET_FLAG = group.get("target_flag", TARGET_FLAG)
    TARGET_PCT = float(group.get("target_pct", TARGET_PCT))
    MAX_SWAPS = int(group.get("max_swaps", MAX_SWAPS))
    MAX_ADD_CANDIDATES_PER_STEP = int(
        group.get(
            "max_add_candidates_per_step",
            MAX_ADD_CANDIDATES_PER_STEP
        )
    )
    FORM_TYPE = VFORM.split("_")[0]



OUTPUT_FILE = OUTPUT_DIR / f"{PROJECT}_{VFORM}_AGEGRP{AGEGRP_TO_RUN}_case_swap.xlsx"
OUTPUT_ARCHIVE_FILE = OUTPUT_DIR / f"{PROJECT}_{VFORM}_AGEGRP{AGEGRP_TO_RUN}_{PREP_RUN_TIMESTAMP}_case_swap.xlsx"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def archive_output_file(latest_path, archive_path, label):
    if latest_path == archive_path:
        return
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(latest_path, archive_path)
    print(f"{label} timestamped copy:")
    print(archive_path)

def get_fr_sheet_name(project, form_type):
    project = project.upper()
    form_type = form_type.upper()

    project_map = {
        "ADHD": "adhd",
        "ASD": "asd",
        "MOOD": "md",
    }

    form_map = {
        "TRS": "tch",
        "PRS": "par",
        "SRP": "self",
    }

    if project not in project_map:
        raise ValueError(f"Unknown project: {project}")

    if form_type not in form_map:
        raise ValueError(f"Unknown form type: {form_type}")

    return project_map[project] + form_map[form_type]


# =====================================================
# PACKAGE OPTIMIZER INPUTS
# Integrated from 01b_package_optimizer_inputs.py
# =====================================================

def normalize_teid_for_package(series):
    text = series.astype(str).str.strip()
    numeric = pd.to_numeric(text, errors="coerce")
    mask = numeric.notna()
    text.loc[mask] = numeric.loc[mask].astype("Int64").astype(str)
    return text.replace({"NAN": pd.NA, "NONE": pd.NA, "<NA>": pd.NA})


def ensure_package_dirs(app_dir):
    dirs = {
        "app": app_dir,
        "inputs": app_dir / "inputs",
        "raw_data": app_dir / "inputs" / "raw_data",
        "master": app_dir / "inputs" / "master_sample",
        "pools": app_dir / "inputs" / "replacement_pools",
        "targets": app_dir / "inputs" / "targets",
        "outputs": app_dir / "outputs",
        "logs": app_dir / "logs",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


def read_sheet_if_exists(path, sheet_name):
    try:
        xls = pd.ExcelFile(path)
        if sheet_name not in xls.sheet_names:
            return None
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return None


def collect_case_swap_files(qa_dir, pattern):
    re = __import__("re")
    timestamped_archive = __import__("re").compile(
        r"_\d{8}_\d{6}(?:_\d{6})?_case_swap\.xlsx$",
        re.IGNORECASE,
    )
    agegrp_case_swap = re.compile(r"AGEGRP[_-]?\d+_case_swap\.xlsx$", re.IGNORECASE)
    return sorted(
        p for p in qa_dir.glob(pattern)
        if not p.name.startswith("~$")
        and not timestamped_archive.search(p.name)
        and agegrp_case_swap.search(p.name)
    )


def build_master_from_final_samples(case_swap_files):
    final_parts = []
    log_rows = []

    for path in case_swap_files:
        final_df = read_sheet_if_exists(path, "final_sample")
        pool_df = read_sheet_if_exists(path, "replacement_pool")

        log_rows.append({
            "file_name": path.name,
            "has_final_sample": final_df is not None,
            "final_sample_n": len(final_df) if final_df is not None else 0,
            "has_replacement_pool": pool_df is not None,
            "replacement_pool_n": len(pool_df) if pool_df is not None else 0,
        })

        if final_df is not None and not final_df.empty:
            final_df = final_df.copy()
            final_df["source_case_swap_file"] = path.name
            vform, agegrp = infer_vform_agegrp_from_case_swap_name(path.name)
            if vform is not None:
                if "vform" not in final_df.columns:
                    final_df["vform"] = vform
                else:
                    final_df["vform"] = final_df["vform"].fillna(vform)
            if agegrp is not None:
                if "agegrp" not in final_df.columns:
                    final_df["agegrp"] = agegrp
                else:
                    final_df["agegrp"] = final_df["agegrp"].fillna(agegrp)
            final_parts.append(final_df)

    if not final_parts:
        raise ValueError(
            "No final_sample sheets were found in the case-swap files. "
            "Cannot build MASTER_SAMPLE.xlsx."
        )

    master = pd.concat(final_parts, ignore_index=True, sort=False)

    if "teid" in master.columns:
        master["teid"] = normalize_teid_for_package(master["teid"])
        before = len(master)
        master = master.drop_duplicates(subset=["teid"], keep="first").copy()
        after = len(master)
        deduped = before - after
    else:
        deduped = None

    manifest = pd.DataFrame(log_rows)
    manifest["master_deduped_teids_removed"] = deduped

    return master, manifest


def copy_replacement_pool_files(case_swap_files, dest_dir):
    rows = []
    for src in case_swap_files:
        dest = dest_dir / src.name
        if src.resolve() != dest.resolve():
            shutil.copy2(src, dest)
        pool_df = read_sheet_if_exists(src, "replacement_pool")
        final_df = read_sheet_if_exists(src, "final_sample")
        rows.append({
            "source_file": str(src),
            "copied_to": str(dest),
            "file_name": dest.name,
            "replacement_pool_n": len(pool_df) if pool_df is not None else 0,
            "final_sample_n": len(final_df) if final_df is not None else 0,
            "status": "copied",
        })
    return pd.DataFrame(rows)


def copy_optional_file(src, dest, label):
    if not src:
        return {
            "file_type": label,
            "source": "",
            "destination": str(dest),
            "status": "not supplied",
        }

    src_path = Path(src).expanduser()
    if not src_path.is_absolute():
        src_path = PROJECT_DIR / src_path
    src_path = src_path.resolve()

    if not src_path.exists():
        return {
            "file_type": label,
            "source": str(src_path),
            "destination": str(dest),
            "status": "source not found",
        }

    dest.parent.mkdir(parents=True, exist_ok=True)
    if src_path != dest.resolve():
        shutil.copy2(src_path, dest)
    return {
        "file_type": label,
        "source": str(src_path),
        "destination": str(dest),
        "status": "copied",
    }


def write_master_package(master, master_path):
    master_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="final_all", index=False)
    archive_output_file(
        master_path,
        master_path.with_name(f"{master_path.stem}_{PREP_RUN_TIMESTAMP}{master_path.suffix}"),
        "Master sample",
    )


def write_package_manifest(app_dir, master, case_file_log, pool_copy_log, optional_file_log):
    logs_dir = app_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = logs_dir / "Run_Receipts.xlsx"

    summary_rows = [
        {"item": "app_dir", "value": str(app_dir)},
        {"item": "master_sample_file", "value": str(app_dir / "inputs" / "master_sample" / MASTER_OUTPUT_NAME)},
        {"item": "master_sample_sheet", "value": "final_all"},
        {"item": "master_sample_n", "value": len(master)},
        {"item": "replacement_pool_dir", "value": str(app_dir / "inputs" / "replacement_pools")},
        {"item": "target_file", "value": str(app_dir / "inputs" / "targets" / TARGET_OUTPUT_NAME)},
        {"item": "output_dir", "value": str(app_dir / "outputs")},
        {"item": "logs_dir", "value": str(app_dir / "logs")},
    ]

    with pd.ExcelWriter(manifest_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="summary", index=False)
        case_file_log.to_excel(writer, sheet_name="case_swap_scan", index=False)
        pool_copy_log.to_excel(writer, sheet_name="copied_pools", index=False)
        optional_file_log.to_excel(writer, sheet_name="optional_files", index=False)
        pd.DataFrame({"master_columns": list(master.columns)}).to_excel(
            writer,
            sheet_name="master_columns",
            index=False,
        )
    archive_output_file(
        manifest_path,
        logs_dir / f"Run_Receipts_{PREP_RUN_TIMESTAMP}.xlsx",
        "Run receipts",
    )


def update_optimizer_config_workbook(app_dir, case_swap_files):
    config_path = app_dir / "scripts" / CONFIG_OUTPUT_NAME
    if not config_path.exists():
        print(f"Config workbook not found; skipping config refresh: {config_path}")
        return

    wb = load_workbook(config_path)

    if "1_Project_Files" in wb.sheetnames:
        ws = wb["1_Project_Files"]
        updates = {
            "project_dir": str(app_dir),
            "master_sample_file": r"inputs\master_sample\MASTER_SAMPLE.xlsx",
            "replacement_pool_dir": r"inputs\replacement_pools",
            "target_workbook": rf"inputs\targets\{TARGET_OUTPUT_NAME}",
            "output_dir": "outputs",
        }
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            setting = row[1].value if len(row) > 1 else None
            if setting in updates:
                row[2].value = updates[setting]

    if "3_Pools_Constraints" in wb.sheetnames:
        ws = wb["3_Pools_Constraints"]
        header_row = None
        for r in range(1, ws.max_row + 1):
            values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if "Section" in values and "Item / Constraint" in values:
                header_row = r
                break

        if header_row is not None:
            first_data_row = header_row + 1
            hard_rows = []
            for r in range(first_data_row, ws.max_row + 1):
                if ws.cell(r, 1).value == "Hard Constraint":
                    hard_rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])

            if ws.max_row >= first_data_row:
                ws.delete_rows(first_data_row, ws.max_row - first_data_row + 1)

            output_rows = []
            for path in case_swap_files:
                vform, agegrp = infer_vform_agegrp_from_case_swap_name(path.name)
                if vform is None or agegrp is None:
                    continue
                output_rows.append([
                    "Replacement Pool",
                    path.name,
                    "Y",
                    path.name,
                    vform,
                    agegrp,
                    "replacement_pool sheet",
                    "Reads eligible replacement cases from this workbook",
                ])
            output_rows.extend(hard_rows)

            for row_values in output_rows:
                ws.append(row_values)

    try:
        wb.save(config_path)
        print(f"Config workbook refreshed: {config_path}")
    except PermissionError:
        fallback_dir = app_dir / "outputs" / "config_refresh"
        fallback_dir.mkdir(parents=True, exist_ok=True)
        fallback_path = fallback_dir / f"{config_path.stem}_refreshed_{PREP_RUN_TIMESTAMP}{config_path.suffix}"
        wb.save(fallback_path)
        print("Config workbook is open/locked; refreshed copy saved instead:")
        print(fallback_path)


def infer_vform_agegrp_from_case_swap_name(file_name):
    stem = Path(file_name).stem.upper()
    parts = stem.split("_")
    vform = None
    agegrp = None

    for i in range(len(parts) - 1):
        candidate = f"{parts[i]}_{parts[i + 1]}"
        if candidate in POOL_RULES:
            vform = candidate
            break

    import re
    match = re.search(r"AGEGRP[_-]?(\d+)", stem)
    if match:
        agegrp = int(match.group(1))

    return vform, agegrp


def package_optimizer_inputs():
    qa_dir = OUTPUT_DIR.resolve()
    app_dir = Path(PACKAGE_APP_DIR).expanduser().resolve()

    if not qa_dir.exists():
        raise FileNotFoundError(f"QA/case-swap folder not found: {qa_dir}")

    dirs = ensure_package_dirs(app_dir)

    if PACKAGE_CLEAR_EXISTING_POOLS:
        for f in dirs["pools"].glob("*"):
            if f.is_file():
                f.unlink()

    print("\nPackaging optimizer inputs...")
    case_swap_files = collect_case_swap_files(qa_dir, PACKAGE_CASE_SWAP_PATTERN)
    if not case_swap_files:
        raise ValueError(
            f"No case-swap files found in {qa_dir} using pattern "
            f"{PACKAGE_CASE_SWAP_PATTERN}"
        )

    print(f"Found {len(case_swap_files)} case-swap files.")

    master, case_file_log = build_master_from_final_samples(case_swap_files)
    master_path = dirs["master"] / MASTER_OUTPUT_NAME
    write_master_package(master, master_path)

    pool_copy_log = copy_replacement_pool_files(case_swap_files, dirs["pools"])

    optional_logs = [
        copy_optional_file(PACKAGE_TARGET_WORKBOOK, dirs["targets"] / TARGET_OUTPUT_NAME, "target_workbook"),
        copy_optional_file(PACKAGE_CONFIG_WORKBOOK, app_dir / "scripts" / CONFIG_OUTPUT_NAME, "config_workbook"),
        copy_optional_file(PACKAGE_EXE, app_dir / "scripts" / "StratificationOptimizer.exe", "optimizer_exe"),
    ]
    optional_file_log = pd.DataFrame(optional_logs)

    write_package_manifest(
        app_dir,
        master,
        case_file_log,
        pool_copy_log,
        optional_file_log,
    )
    update_optimizer_config_workbook(app_dir, [Path(row["file_name"]) for _, row in pool_copy_log.iterrows()])

    print("Package created successfully.")
    print(f"App folder: {app_dir}")
    print(f"Master sample: {master_path}")
    print(f"Replacement pools: {dirs['pools']}")
    print(f"Targets folder: {dirs['targets']}")
    print(f"Outputs folder: {dirs['outputs']}")
    print(f"Run receipts: {app_dir / 'logs' / 'Run_Receipts.xlsx'}")

# =====================================================
# LOAD
# =====================================================

case_level = pd.read_excel(INPUT_FILE, sheet_name="case_level")

case_level.columns = [c.strip() for c in case_level.columns]

for col in FLAG_COLS:
    if col in case_level.columns:
        case_level[col] = case_level[col].fillna(0).astype(int)

# Current sample for demographic/clinical flags comes from case_level
sample = case_level[
    (case_level["project"] == PROJECT) &
    (case_level["vform"] == VFORM)
].copy()

if sample.empty:
    raise ValueError("No sample rows found.")

# Bring in official agegrp from scored SAS data
rule = POOL_RULES[VFORM]

agegrp_source = pd.read_sas(
    SCORED_DATA_DIR / rule["file"],
    encoding="latin1"
)

agegrp_source.columns = [
    c.decode() if isinstance(c, bytes) else c
    for c in agegrp_source.columns
]

agegrp_source.columns = [c.strip().lower() for c in agegrp_source.columns]

agegrp_map = agegrp_source[["teid", "agegrp"]].copy()

agegrp_map["teid"] = pd.to_numeric(
    agegrp_map["teid"],
    errors="coerce"
)

agegrp_map = agegrp_map.dropna(subset=["teid"])

agegrp_map["teid"] = agegrp_map["teid"].astype(int)

sample["teid"] = pd.to_numeric(
    sample["teid"],
    errors="coerce"
)

sample = sample.dropna(subset=["teid"])

sample["teid"] = sample["teid"].astype(int)

sample = sample.merge(
    agegrp_map,
    on="teid",
    how="left"
)

print("\nAgegrp merge check:")
print("Sample N before agegrp filter:", len(sample))
print("Matched agegrp N:", sample["agegrp"].notna().sum())

print("\nAgegrp value counts:")
print(sample["agegrp"].value_counts(dropna=False))

sample = sample[
    sample["agegrp"] == AGEGRP_TO_RUN
].copy()

if sample.empty:
    raise ValueError(f"No sample rows found for agegrp {AGEGRP_TO_RUN}.")

print("\nTarget sample:")
print(PROJECT, VFORM, "AGEGRP", AGEGRP_TO_RUN)
print("N =", len(sample))

current_sample = sample.copy()


fr_sheet = get_fr_sheet_name(PROJECT, FORM_TYPE)

fr_norm_sample = pd.read_excel(
    FR_TEID_FILE,
    sheet_name=fr_sheet
)

fr_norm_sample.columns = [c.strip().lower() for c in fr_norm_sample.columns]

if "teid" not in fr_norm_sample.columns:
    raise ValueError(f"'teid' column not found in FR sheet: {fr_sheet}")

fr_norm_teids = set(fr_norm_sample["teid"].dropna().astype(int))

print("\nFR norm sample:")
print("Sheet =", fr_sheet)
print("N TEIDs =", len(fr_norm_teids))

fr_xls = pd.ExcelFile(FR_TEID_FILE)

all_fr_selected_teids = set()

for sheet in fr_xls.sheet_names:
    tmp = pd.read_excel(FR_TEID_FILE, sheet_name=sheet)
    tmp.columns = [c.strip().lower() for c in tmp.columns]

    if "teid" in tmp.columns:
        all_fr_selected_teids.update(
            tmp["teid"].dropna().astype(int).tolist()
        )

print("\nAll FR selected TEIDs:")
print("N =", len(all_fr_selected_teids))


# -----------------------------------------------------
# Load usable universe / candidate replacement pool
# -----------------------------------------------------

rule = POOL_RULES[VFORM]

usable_universe = pd.read_sas(
    SCORED_DATA_DIR / rule["file"],
    encoding="latin1"
)

usable_universe.columns = [
    c.decode() if isinstance(c, bytes) else c
    for c in usable_universe.columns
]

usable_universe.columns = [c.strip().lower() for c in usable_universe.columns]

TARGET_FLAG_POOL = TARGET_FLAG.lower()

# ==========================================
# TEMP DEBUG CHECK
# ==========================================

print("\nClinical flag counts:")

for col in usable_universe.columns:
    if col.lower().startswith("tmp"):
        vals = pd.to_numeric(usable_universe[col], errors="coerce").fillna(0)
        print(col, vals.sum())

# =====================================================
# STANDARDIZE TARGET FLAG IN USABLE UNIVERSE
# =====================================================

if TARGET_FLAG_POOL in usable_universe.columns:
    usable_universe[TARGET_FLAG_POOL] = pd.to_numeric(
        usable_universe[TARGET_FLAG_POOL],
        errors="coerce"
    ).fillna(0).astype(int)
else:
    raise ValueError(f"{TARGET_FLAG_POOL} not found in usable_universe")

# ==========================================
# CONTINUE NORMAL PROCESSING
# ==========================================

usable_universe["testid"] = pd.to_numeric(
    usable_universe["testid"],
    errors="coerce"
)

usable_universe["agegrp"] = pd.to_numeric(
    usable_universe["agegrp"],
    errors="coerce"
)

usable_universe = usable_universe[
    usable_universe["testid"].isin(rule["testid"]) &
    usable_universe["agegrp"].isin(rule["agegrp"])
].copy()

print("\nUsable universe after Anqi pool filter:")
print("VFORM =", VFORM)
print("N =", len(usable_universe))

print("\nUsable universe target flag count after filter:")
print(TARGET_FLAG_POOL, usable_universe[TARGET_FLAG_POOL].sum())

replacement_pool = usable_universe[
    ~usable_universe["teid"].astype(int).isin(all_fr_selected_teids)
].copy()

replacement_pool = replacement_pool[
    replacement_pool["agegrp"] == AGEGRP_TO_RUN
].copy()

print("\nReplacement pool after agegrp filter:")
print("AGEGRP =", AGEGRP_TO_RUN)
print("N =", len(replacement_pool))

if TARGET_FLAG_POOL in replacement_pool.columns:
    replacement_pool[TARGET_FLAG] = replacement_pool[TARGET_FLAG_POOL]
else:
    raise ValueError(f"{TARGET_FLAG_POOL} not found in replacement_pool")

print("\nReplacement pool:")
print("N =", len(replacement_pool))
print("Replacement pool target flag count:")
print(TARGET_FLAG, replacement_pool[TARGET_FLAG].sum())

print("\nPool reduction:")
print("Universe N =", len(usable_universe))
print("Selected FR TEIDs removed =", len(usable_universe) - len(replacement_pool))

# -----------------------------------------------------
# Load tolerance rules
# -----------------------------------------------------

tolerance_rules = pd.read_excel(
    TOLERANCE_FILE,
    sheet_name="tolerance_rules"
)

tolerance_rules = tolerance_rules[
    tolerance_rules["form_type"] == FORM_TYPE
].copy()

print("\nTolerance rules loaded:")
print(tolerance_rules.to_string(index=False))

print("\nProject/VForm counts:")
print(
    case_level
    .groupby(["project", "vform"])
    .size()
    .reset_index(name="n")
    .sort_values("n", ascending=False)
    .head(20)
)

manual_sex_rules = pd.DataFrame({
    "form_type": [FORM_TYPE, FORM_TYPE],
    "variable": ["sex", "sex"],
    "variable_label": ["Sex", "Sex"],
    "group": ["F", "M"],
    "tolerance_pct_points": [0.5, 0.5],
    "source_row": [pd.NA, pd.NA],
    "source_col": [pd.NA, pd.NA],
})

tolerance_rules = pd.concat(
    [tolerance_rules, manual_sex_rules],
    ignore_index=True
)


# =====================================================
# HELPERS
# =====================================================

def load_tolerances():
    tol = pd.read_excel(TOLERANCE_FILE)

    tol.columns = [c.strip().lower() for c in tol.columns]

    # Expected columns from script 55:
    # form_type, demo_var, demo_level, tolerance_pct
    tol["form_type"] = tol["form_type"].astype(str).str.upper().str.strip()
    tol["demo_var"] = tol["demo_var"].astype(str).str.lower().str.strip()
    tol["demo_level"] = tol["demo_level"].astype(str).str.strip()
    tol["tolerance_pct"] = pd.to_numeric(tol["tolerance_pct"], errors="coerce")

    return tol
def clean_group_value(x):
    if pd.isna(x):
        return "MISSING"

    s = str(x).strip()

    # Convert values like 1.0, 2.0, 3.0 into 1, 2, 3
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except ValueError:
        return s.upper()


def get_pct_table(df):
    rows = []
    total = len(df)

    for var in STRAT_VARS:
        if var not in df.columns:
            continue

        s = df[var].apply(clean_group_value)
        counts = s.value_counts(dropna=False)

        for group, n in counts.items():
            rows.append({
                "variable": var,
                "group": group,
                "n": int(n),
                "total": total,
                "pct": round(n / total * 100, 3) if total else 0,
            })

    return pd.DataFrame(rows)

def calc_flag_pct(df, flag_col):
    total = len(df)
    numerator = int(df[flag_col].fillna(0).astype(int).sum())
    pct = numerator / total * 100 if total else 0
    return numerator, total, pct


def demo_impact_vs_baseline(current_df, baseline_demo, baseline_total):
    current_demo = get_pct_table(current_df).rename(columns={
        "n": "current_n",
        "total": "current_total",
        "pct": "current_pct",
    })

    print("\nCURRENT EDL groups:")
    print(
        current_demo.loc[
            current_demo["variable"] == "edl",
            ["group", "current_n", "current_pct"]
        ]
        .drop_duplicates()
        .sort_values("group")
    )

    impact = baseline_demo.merge(
        current_demo,
        on=["variable", "group"],
        how="outer"
    )

    impact["baseline_n"] = impact["baseline_n"].fillna(0)
    impact["baseline_pct"] = impact["baseline_pct"].fillna(0)
    impact["baseline_total"] = impact["baseline_total"].fillna(baseline_total)

    impact["current_n"] = impact["current_n"].fillna(0)
    impact["current_pct"] = impact["current_pct"].fillna(0)
    impact["current_total"] = impact["current_total"].fillna(len(current_df))

    impact["penalty_pct"] = impact["current_pct"] - impact["baseline_pct"]
    impact["abs_penalty_pct"] = impact["penalty_pct"].abs()

    return impact

def tolerance_adjusted_penalty(current_pct, target_pct, tolerance_pct):
    lower = target_pct - tolerance_pct
    upper = target_pct + tolerance_pct

    if lower <= current_pct <= upper:
        return 0.0

    if current_pct < lower:
        return lower - current_pct

    return current_pct - upper
def apply_tolerance_rules(impact, tolerance_rules):
    out = impact.copy()

    out = out.merge(
        tolerance_rules[
            ["variable", "group", "tolerance_pct_points"]
        ],
        on=["variable", "group"],
        how="left"
    )

    out["target_pct"] = out["baseline_pct"]

    out["lower_bound"] = out["target_pct"] - out["tolerance_pct_points"]
    out["upper_bound"] = out["target_pct"] + out["tolerance_pct_points"]

    out["has_tolerance_rule"] = out["tolerance_pct_points"].notna()

    out["within_tolerance"] = pd.NA

    mask = out["has_tolerance_rule"]

    out.loc[mask, "within_tolerance"] = (
        (out.loc[mask, "current_pct"] >= out.loc[mask, "lower_bound"]) &
        (out.loc[mask, "current_pct"] <= out.loc[mask, "upper_bound"])
    )

    out["margin_to_lower"] = out["current_pct"] - out["lower_bound"]
    out["margin_to_upper"] = out["upper_bound"] - out["current_pct"]

    out["margin_remaining"] = out[["margin_to_lower", "margin_to_upper"]].min(axis=1)

    out.loc[~out["has_tolerance_rule"], "margin_remaining"] = pd.NA

    # Tolerance-adjusted penalty:
    out["tolerance_penalty"] = pd.NA

    out.loc[
        mask & (out["current_pct"] < out["lower_bound"]),
        "tolerance_penalty"
    ] = out["lower_bound"] - out["current_pct"]

    out.loc[
        mask & (out["current_pct"] > out["upper_bound"]),
        "tolerance_penalty"
    ] = out["current_pct"] - out["upper_bound"]

    out.loc[
        mask & (out["within_tolerance"] == True),
        "tolerance_penalty"
    ] = 0.0

    return out

def score_single_swap(current_df, remove_teid, add_teid, replacement_pool, baseline_demo, baseline_total):
    remove_rows = current_df[current_df["teid"] == remove_teid]
    add_rows = replacement_pool[replacement_pool["teid"] == add_teid]

    if remove_rows.empty or add_rows.empty:
        return None

    after = current_df[current_df["teid"] != remove_teid].copy()
    after = pd.concat([after, add_rows], ignore_index=True)

    num, total, pct = calc_flag_pct(after, TARGET_FLAG)

    current_demo = get_pct_table(after)

    impact = current_demo.merge(
        baseline_demo,
        on=["variable", "group"],
        how="outer"
    )

    impact["n"] = impact["n"].fillna(0)
    impact["total"] = impact["total"].fillna(total)
    impact["pct"] = impact["pct"].fillna(0)

    impact["baseline_n"] = impact["baseline_n"].fillna(0)
    impact["baseline_total"] = impact["baseline_total"].fillna(baseline_total)
    impact["baseline_pct"] = impact["baseline_pct"].fillna(0)

    impact["delta_pct"] = impact["pct"] - impact["baseline_pct"]
    impact["abs_penalty_pct"] = impact["delta_pct"].abs()

    impact = impact.rename(columns={"pct": "current_pct"})

    impact = apply_tolerance_rules(impact, tolerance_rules)

    tol_penalty = pd.to_numeric(
        impact["tolerance_penalty"],
        errors="coerce"
    )

    return {
        "remove_teid": remove_teid,
        "add_teid": add_teid,
        "after_df": after,
        "after_num": num,
        "after_total": total,
        "after_pct": pct,

        "max_demo_drift": impact["abs_penalty_pct"].max(),
        "mean_demo_drift": impact["abs_penalty_pct"].mean(),
        "sum_demo_drift": impact["abs_penalty_pct"].sum(),

        "max_tolerance_penalty": tol_penalty.max(skipna=True),
        "mean_tolerance_penalty": tol_penalty.mean(skipna=True),
        "sum_tolerance_penalty": tol_penalty.sum(skipna=True),

        "impact": impact,
    }

def score_single_add(current_df, add_teid, replacement_pool, baseline_demo, baseline_total):
    add_case = replacement_pool[replacement_pool["teid"] == add_teid].copy()

    if add_case.empty:
        return None

    after = pd.concat(
        [current_df.copy(), add_case],
        ignore_index=True
    )

    num, total, pct = calc_flag_pct(after, TARGET_FLAG)

    current_demo = get_pct_table(after).rename(columns={
        "n": "current_n",
        "total": "current_total",
        "pct": "current_pct",
    })

    impact = current_demo.merge(
        baseline_demo,
        on=["variable", "group"],
        how="outer"
    )

    impact["current_n"] = impact["current_n"].fillna(0)
    impact["baseline_n"] = impact["baseline_n"].fillna(0)
    impact["current_pct"] = impact["current_pct"].fillna(0)
    impact["baseline_pct"] = impact["baseline_pct"].fillna(0)

    impact["delta_pct"] = impact["current_pct"] - impact["baseline_pct"]
    impact["abs_penalty_pct"] = impact["delta_pct"].abs()

    impact = apply_tolerance_rules(
        impact,
        tolerance_rules
    )

    tol_penalty = pd.to_numeric(
        impact["tolerance_penalty"],
        errors="coerce"
    )

    return {
        "add_teid": add_teid,
        "after_df": after,
        "after_num": num,
        "after_total": total,
        "after_pct": pct,

        "max_demo_drift": impact["abs_penalty_pct"].max(),
        "mean_demo_drift": impact["abs_penalty_pct"].mean(),
        "sum_demo_drift": impact["abs_penalty_pct"].sum(),

        "max_tolerance_penalty": tol_penalty.max(skipna=True),
        "mean_tolerance_penalty": tol_penalty.mean(skipna=True),
        "sum_tolerance_penalty": tol_penalty.sum(skipna=True),

        "impact": impact,
    }

# =====================================================
# BASELINE
# =====================================================

baseline_demo = get_pct_table(sample).rename(columns={
    "n": "baseline_n",
    "total": "baseline_total",
    "pct": "baseline_pct",
})

print("\nBASELINE EDL groups:")
print(
    baseline_demo.loc[
        baseline_demo["variable"] == "edl",
        ["group", "baseline_n", "baseline_pct"]
    ]
    .drop_duplicates()
    .sort_values("group")
)

print("\nSample columns containing tmp:")

print(
    [c for c in sample.columns
     if "tmp" in c.lower()]
)

baseline_num, baseline_total, baseline_pct = calc_flag_pct(sample, TARGET_FLAG)

print("\nBaseline:")
print("Numerator:", baseline_num)
print("Total:", baseline_total)
print("Percent:", round(baseline_pct, 3))
print("Target:", TARGET_PCT)


if baseline_pct > TARGET_PCT:
    SWAP_MODE = "reduce_target"
elif baseline_pct < TARGET_PCT:
    SWAP_MODE = "increase_target"
else:
    SWAP_MODE = "at_target"

print("\nSwap mode:", SWAP_MODE)

if SWAP_MODE == "at_target":
    print("Baseline is already at target. No swaps needed.")

    swap_path = pd.DataFrame([{
        "step": 0,
        "scenario": "Baseline already at target - no swaps needed",
        "removed_teid": None,
        "added_teid": None,
        "n_swaps": 0,
        "num": baseline_num,
        "total": baseline_total,
        "pct": round(baseline_pct, 3),
        "target_pct": TARGET_PCT,
        "distance_to_target": round(abs(baseline_pct - TARGET_PCT), 3),
        "max_tolerance_penalty": 0,
        "n_tolerance_violations": 0,
    }])

    selected_swaps = pd.DataFrame()
    all_step_impacts = pd.DataFrame()
    n_repair = pd.DataFrame()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        swap_path.to_excel(writer, sheet_name="swap_path", index=False)
        selected_swaps.to_excel(writer, sheet_name="selected_swaps", index=False)
        all_step_impacts.to_excel(writer, sheet_name="all_step_impacts", index=False)
        replacement_pool.to_excel(writer, sheet_name="replacement_pool", index=False)
        n_repair.to_excel(writer, sheet_name="n_repair", index=False)

        # Baseline was already at target, so the original sample is the final sample
        sample.to_excel(writer, sheet_name="final_sample", index=False)

    print("\nSaved:")
    print(OUTPUT_FILE)
    archive_output_file(OUTPUT_FILE, OUTPUT_ARCHIVE_FILE, "Case-swap workbook")
    package_optimizer_inputs()

    raise SystemExit

available_replacements = replacement_pool.copy()

# =====================================================
# N REPAIR — ADD CASES IF AGEGRP SAMPLE IS UNDER N
# =====================================================

n_repair_rows = []

if len(sample) < EXPECTED_N:

    n_needed = EXPECTED_N - len(sample)

    print("\nN repair needed:")
    print("Current N =", len(sample))
    print("Expected N =", EXPECTED_N)
    print("Need to add =", n_needed)

    # If sample is over target, add non-target cases.
    # If sample is under target, add target-positive cases.
    if baseline_pct > TARGET_PCT:
        repair_candidates = (
            replacement_pool[replacement_pool[TARGET_FLAG] != 1]
            .head(MAX_ADD_CANDIDATES_PER_STEP)
            .copy()
        )
    else:
        repair_candidates = (
            replacement_pool[replacement_pool[TARGET_FLAG] == 1]
            .head(MAX_ADD_CANDIDATES_PER_STEP)
            .copy()
        )

    for repair_step in range(1, n_needed + 1):

        scored_adds = []

        for _, add_row in repair_candidates.iterrows():
            add_teid = add_row["teid"]

            result = score_single_add(
                current_df=sample,
                add_teid=add_teid,
                replacement_pool=replacement_pool,
                baseline_demo=baseline_demo,
                baseline_total=baseline_total,
            )

            if result is None:
                continue

            scored_adds.append({
                "add_teid": add_teid,
                "after_num": result["after_num"],
                "after_total": result["after_total"],
                "after_pct": round(result["after_pct"], 3),
                "target_pct": TARGET_PCT,
                "distance_to_target": round(abs(result["after_pct"] - TARGET_PCT), 3),
                "max_tolerance_penalty": result["max_tolerance_penalty"],
                "mean_tolerance_penalty": result["mean_tolerance_penalty"],
                "sum_tolerance_penalty": result["sum_tolerance_penalty"],
                "result": result,
            })

        if not scored_adds:
            print("\nNo valid N-repair add candidates found.")
            break

        scored_adds_df = pd.DataFrame([
            {k: v for k, v in x.items() if k != "result"}
            for x in scored_adds
        ])

        scored_adds_df = scored_adds_df.sort_values(
            by=[
                "max_tolerance_penalty",
                "distance_to_target",
                "mean_tolerance_penalty",
                "sum_tolerance_penalty",
            ],
            ascending=[True, True, True, True],
        ).reset_index(drop=True)

        best_add_teid = scored_adds_df.loc[0, "add_teid"]

        best = next(
            x for x in scored_adds
            if x["add_teid"] == best_add_teid
        )

        result = best["result"]

        added_case = replacement_pool[
            replacement_pool["teid"] == best_add_teid
        ].iloc[0].to_dict()

        sample = result["after_df"].copy()

        replacement_pool = replacement_pool[
            replacement_pool["teid"] != best_add_teid
        ].copy()

        n_repair_rows.append({
            "repair_step": repair_step,
            "added_teid": best_add_teid,
            "added_sex": added_case.get("sex"),
            "added_eth": added_case.get("eth"),
            "added_edl": added_case.get("edl"),
            "added_region": added_case.get("region"),
            "after_num": result["after_num"],
            "after_total": result["after_total"],
            "after_pct": round(result["after_pct"], 3),
            "target_pct": TARGET_PCT,
            "distance_to_target": round(abs(result["after_pct"] - TARGET_PCT), 3),
            "max_tolerance_penalty": result["max_tolerance_penalty"],
        })

        print(
            f"N repair {repair_step}: add {best_add_teid} | "
            f"N = {result['after_total']} | "
            f"{TARGET_FLAG} pct = {result['after_pct']:.3f} | "
            f"max tolerance penalty = {result['max_tolerance_penalty']:.3f}"
        )

n_repair = pd.DataFrame(n_repair_rows)

# =====================================================
# GREEDY SWAP PATH
# =====================================================

current = sample.copy()
available_replacements = replacement_pool.copy()

selected_swaps = []
swap_path = []
all_step_impacts = []

# Step 0 baseline
swap_path.append({
    "step": 0,
    "scenario": "Baseline (Original Sample - No Swaps)",
    "removed_teid": None,
    "added_teid": None,
    "n_swaps": 0,
    "num": baseline_num,
    "total": baseline_total,
    "pct": round(baseline_pct, 3),
    "target_pct": TARGET_PCT,
    "distance_to_target": round(abs(baseline_pct - TARGET_PCT), 3),

    "max_demo_drift": 0,
    "mean_demo_drift": 0,
    "sum_demo_drift": 0,

    "max_tolerance_penalty": 0,
    "mean_tolerance_penalty": 0,
    "sum_tolerance_penalty": 0,

    "n_tolerance_violations": 0,
    "closest_margin": pd.NA,
    "closest_tolerance_issue": "",
})

for step in range(1, MAX_SWAPS + 1):

    if SWAP_MODE == "reduce_target":
        remove_candidates = current[current[TARGET_FLAG] == 1].copy()

        add_candidates = (
            available_replacements[available_replacements[TARGET_FLAG] != 1]
            .head(MAX_ADD_CANDIDATES_PER_STEP)
            .copy()
        )

    elif SWAP_MODE == "increase_target":
        remove_candidates = (
            current[current[TARGET_FLAG] != 1]
            .head(50)
            .copy()
        )

        add_candidates = (
            available_replacements[available_replacements[TARGET_FLAG] == 1]
            .head(MAX_ADD_CANDIDATES_PER_STEP)
            .copy()
        )

    elif SWAP_MODE == "at_target":
        print("\nBaseline already saved. Exiting.")
        raise SystemExit
    else:
        raise ValueError(f"Unknown SWAP_MODE: {SWAP_MODE}")


    print(f"\nStep {step}")
    print("Remove candidates:", len(remove_candidates))
    print("Add candidates:", len(add_candidates))

    if remove_candidates.empty or add_candidates.empty:
        print("\nNo more swap candidates available.")
        break

    scored = []

    for _, remove_row in remove_candidates.iterrows():
        remove_teid = remove_row["teid"]

        for _, add_row in add_candidates.iterrows():
            add_teid = add_row["teid"]

            result = score_single_swap(
                current_df=current,
                remove_teid=remove_teid,
                add_teid=add_teid,
                replacement_pool=available_replacements,
                baseline_demo=baseline_demo,
                baseline_total=baseline_total,
            )

            if result is None:
                continue

            scored.append({
                "remove_teid": remove_teid,
                "add_teid": add_teid,
                "after_num": result["after_num"],
                "after_total": result["after_total"],
                "after_pct": round(result["after_pct"], 3),
                "target_pct": TARGET_PCT,
                "distance_to_target": round(abs(result["after_pct"] - TARGET_PCT), 3),

                "max_demo_drift": result["max_demo_drift"],
                "mean_demo_drift": result["mean_demo_drift"],
                "sum_demo_drift": result["sum_demo_drift"],

                "max_tolerance_penalty": result["max_tolerance_penalty"],
                "mean_tolerance_penalty": result["mean_tolerance_penalty"],
                "sum_tolerance_penalty": result["sum_tolerance_penalty"],

                "result": result,
            })

            if len(scored) % 1000 == 0:
                print(f"Scored {len(scored):,} swaps...")

    if not scored:
        print("\nNo valid swaps scored.")
        break

    scored_df = pd.DataFrame([
        {k: v for k, v in x.items() if k != "result"}
        for x in scored
    ])

    # First try swaps that create no tolerance violations
    scored_df = scored_df[
        scored_df["max_tolerance_penalty"] <= 0
    ].copy()

    scored_df = scored_df.sort_values(
        by=["distance_to_target"],
        ascending=[True],
    ).reset_index(drop=True)

    # If no perfect-tolerance swap exists, use least-bad tolerance option
    if scored_df.empty:
        print("\nNo candidate swaps met the current tolerance criteria.")
        print("Relaxing to best available candidate by tolerance penalty + distance to target.")

        scored_df = pd.DataFrame([
            {k: v for k, v in x.items() if k != "result"}
            for x in scored
        ])

        scored_df = scored_df.sort_values(
            by=[
                "max_tolerance_penalty",
                "distance_to_target",
                "mean_tolerance_penalty",
                "sum_tolerance_penalty",
            ],
            ascending=[True, True, True, True],
        ).reset_index(drop=True)

    if scored_df.empty:
        print("\nNo valid swaps available at all.")
        break

    best_remove_teid = scored_df.loc[0, "remove_teid"]
    best_add_teid = scored_df.loc[0, "add_teid"]

    best = next(
        x for x in scored
        if x["remove_teid"] == best_remove_teid and x["add_teid"] == best_add_teid
    )

    result = best["result"]

    removed_case = current[current["teid"] == best_remove_teid].iloc[0].to_dict()
    added_case = available_replacements[
        available_replacements["teid"] == best_add_teid
    ].iloc[0].to_dict()

    current = result["after_df"].copy()

    available_replacements = available_replacements[
        available_replacements["teid"] != best_add_teid
    ].copy()

    selected_swaps.append({
        "step": step,
        "scenario": f"Remove TEID {best_remove_teid}; Add TEID {best_add_teid}",

        "removed_teid": best_remove_teid,
        "removed_sex": removed_case.get("sex"),
        "removed_eth": removed_case.get("eth"),
        "removed_ethdesc": removed_case.get("ethdesc"),
        "removed_edl": removed_case.get("edl"),
        "removed_edcalcdesc": removed_case.get("edcalcdesc"),
        "removed_region": removed_case.get("region"),
        "removed_testlang": removed_case.get("testlang"),
        "removed_edition": removed_case.get("edition"),

        "added_teid": best_add_teid,
        "added_sex": added_case.get("sex"),
        "added_eth": added_case.get("eth"),
        "added_ethdesc": added_case.get("ethdesc"),
        "added_edl": added_case.get("edl"),
        "added_edcalcdesc": added_case.get("edcalcdesc"),
        "added_region": added_case.get("region"),
        "added_testlang": added_case.get("testlang"),
        "added_edition": added_case.get("edition"),
    })

    impact = result["impact"].copy()
    impact["step"] = step
    impact["removed_teid"] = best_remove_teid
    impact["added_teid"] = best_add_teid
    all_step_impacts.append(impact)

    tol_checked = result["impact"][
        result["impact"]["has_tolerance_rule"] == True
    ].copy()

    n_tolerance_violations = int((tol_checked["within_tolerance"] == False).sum())

    closest_margin = (
        tol_checked["margin_remaining"].min()
        if not tol_checked.empty
        else pd.NA
    )

    worst_tolerance_rows = tol_checked.sort_values("margin_remaining").head(3)

    closest_tolerance_issue = (
        "; ".join(
            worst_tolerance_rows.apply(
                lambda r: f"{r['variable']}={r['group']} margin={r['margin_remaining']:.3f}",
                axis=1
            )
        )
        if not worst_tolerance_rows.empty
        else ""
    )

    swap_path.append({
        "step": step,
        "scenario": f"Remove TEID {best_remove_teid}; Add TEID {best_add_teid}",
        "removed_teid": best_remove_teid,
        "added_teid": best_add_teid,
        "n_swaps": step,
        "num": result["after_num"],
        "total": result["after_total"],
        "pct": round(result["after_pct"], 3),
        "target_pct": TARGET_PCT,
        "distance_to_target": round(abs(result["after_pct"] - TARGET_PCT), 3),

        "max_demo_drift": result["max_demo_drift"],
        "mean_demo_drift": result["mean_demo_drift"],
        "sum_demo_drift": result["sum_demo_drift"],

        "max_tolerance_penalty": result["max_tolerance_penalty"],
        "mean_tolerance_penalty": result["mean_tolerance_penalty"],
        "sum_tolerance_penalty": result["sum_tolerance_penalty"],

        "n_tolerance_violations": n_tolerance_violations,
        "closest_margin": closest_margin,
        "closest_tolerance_issue": closest_tolerance_issue,
    })

    print(
        f"Step {step}: remove {best_remove_teid}; add {best_add_teid} | "
        f"{TARGET_FLAG} pct = {result['after_pct']:.3f} | "
        f"max tolerance penalty = {result['max_tolerance_penalty']:.3f}"
    )

    # Mode-specific stopping rule
    if SWAP_MODE == "reduce_target" and result["after_pct"] <= TARGET_PCT:
        print("\nTarget reached.")
        break

    if SWAP_MODE == "increase_target" and result["after_pct"] >= TARGET_PCT:
        print("\nTarget reached.")
        break

swap_path = pd.DataFrame(swap_path)
selected_swaps = pd.DataFrame(selected_swaps)

if all_step_impacts:
    all_step_impacts = pd.concat(all_step_impacts, ignore_index=True)
else:
    all_step_impacts = pd.DataFrame()


# =====================================================
# SAVE
# =====================================================

final_sample = current.copy()

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    swap_path.to_excel(writer, sheet_name="swap_path", index=False)
    selected_swaps.to_excel(writer, sheet_name="selected_swaps", index=False)
    all_step_impacts.to_excel(writer, sheet_name="all_step_impacts", index=False)
    replacement_pool.to_excel(writer, sheet_name="replacement_pool", index=False)
    n_repair.to_excel(writer, sheet_name="n_repair", index=False)

    # Always write final sample, even if no swaps were made
    final_sample.to_excel(writer, sheet_name="final_sample", index=False)

print("\nSaved:")
print(OUTPUT_FILE)
archive_output_file(OUTPUT_FILE, OUTPUT_ARCHIVE_FILE, "Case-swap workbook")
package_optimizer_inputs()
