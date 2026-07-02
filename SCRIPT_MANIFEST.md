# Script Manifest

## Workflow Launchers

| File | Purpose |
| --- | --- |
| `1_RUN_OPTION_A_DATA_PREP.bat` | Top-level user launcher for workflow A data prep. |
| `2_REFRESH_WORKFLOW_STATUS.bat` | Top-level user launcher to refresh workbook status checks. |
| `3_RUN_OPTIMIZER.bat` | Top-level optimizer launcher that keeps the command window open and writes `logs/Optimizer_Run_*.log`. |

## Script Folder

| File | Purpose |
| --- | --- |
| `scripts/01_prepare_stratification_inputs.py` | Builds case-swap outputs, packages optimizer inputs, writes run receipts, and refreshes config rows. |
| `scripts/Run_Option_A_Data_Prep.bat` | Validates required input files and runs `01_prepare_stratification_inputs.py`. |
| `scripts/Refresh_Workflow_Status.ps1` | Checks required files and prep outputs; updates the workbook if open in Excel. |
| `scripts/Refresh_Workflow_Status.bat` | Batch wrapper for the status refresh PowerShell script. |
| `scripts/Stratification_Optimizer_Config.xlsx` | User-facing configuration/control workbook. |

## Current Audit Notes

- Excel hyperlinks cannot launch `.bat`, `.lnk`, or `.exe` files reliably in the current Windows/Excel environment, so workbook actions open the app folder and instruct the user to double-click top-level launchers.
- `01_prepare_stratification_inputs.py` now creates `outputs/qa_tables` before writing case-swap workbooks.
- The at-target data-prep branch now writes the `replacement_pool` sheet, so receipts correctly report replacement pool counts.
- If the config workbook is open/locked during prep, a refreshed copy is saved under `outputs/config_refresh` instead of crashing.

## Archived Project Scripts

The `project_scripts/` folder contains source-only snapshots copied from sibling local projects for version control and audit:

| Folder | Contents |
| --- | --- |
| `project_scripts/Demos/` | Python scripts copied from `C:\Users\UFAVRLA\PythonProjects\Demos`. |
| `project_scripts/KTEA_CELF_LLM_Scoring_Lab/` | Python scripts copied from `C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab`. |

These snapshots intentionally exclude raw data, workbooks, PDFs, outputs, caches, and project metadata.
