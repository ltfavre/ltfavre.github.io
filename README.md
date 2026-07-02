# Stratification Optimizer App

This repository versions the source scripts and workbook control files for the Stratification Optimizer app workflow.

## Main Workflow

1. Open `scripts/Stratification_Optimizer_Config.xlsx`.
2. Select workflow A or B on `START_HERE_WORKFLOW`.
3. For workflow A, use the app-folder launcher:
   - `1_RUN_OPTION_A_DATA_PREP.bat`
4. Refresh/check workflow status when needed:
   - `2_REFRESH_WORKFLOW_STATUS.bat`
5. Run the optimizer through:
   - `3_RUN_OPTIMIZER.bat`

## Versioning Policy

Tracked:
- source scripts
- launcher scripts
- workbook configuration template/control workbook
- documentation and audit notes

Ignored:
- raw input files
- generated master sample/replacement pool workbooks
- logs and run receipts
- timestamped outputs/backups
- packaged runtime internals and executables

Generated artifacts should be attached to releases or stored in controlled data locations, not committed directly.

