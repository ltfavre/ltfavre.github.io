@echo off
setlocal

set "SCRIPTS_DIR=%~dp0"
set "APP_DIR=%SCRIPTS_DIR%.."
cd /d "%APP_DIR%"

echo.
echo Stratification Optimizer - Option A Data Prep
echo =============================================
echo.
echo This will check the required files, refresh the workbook status,
echo and then run scripts\01_prepare_stratification_inputs.py.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS_DIR%Refresh_Workflow_Status.ps1"

set "MISSING=0"

if not exist "inputs\raw_data\stratification_summary.xlsx" (
  echo MISSING: inputs\raw_data\stratification_summary.xlsx
  set "MISSING=1"
)

if not exist "inputs\raw_data\stratification_tolerance_rules.xlsx" (
  echo MISSING: inputs\raw_data\stratification_tolerance_rules.xlsx
  set "MISSING=1"
)

if not exist "inputs\raw_data\fr_teid_v1.xlsx" (
  echo MISSING: inputs\raw_data\fr_teid_v1.xlsx
  set "MISSING=1"
)

for %%F in (
  basc4nb_scored_trsc.sas7bdat
  basc4nb_scored_trsa.sas7bdat
  basc4nb_scored_prsc.sas7bdat
  basc4nb_scored_prsa.sas7bdat
  basc4nb_scored_srpc.sas7bdat
  basc4nb_scored_srpa.sas7bdat
) do (
  if not exist "inputs\raw_data\%%F" (
    echo MISSING: inputs\raw_data\%%F
    set "MISSING=1"
  )
)

dir /b "inputs\targets\*.xlsx" >nul 2>nul
if errorlevel 1 (
  echo MISSING: at least one .xlsx target workbook in inputs\targets
  set "MISSING=1"
)

if "%MISSING%"=="1" (
  echo.
  echo Add the missing files listed above, then run this again.
  echo.
  pause
  exit /b 1
)

echo.
echo Required files were found. Running data prep...
echo.

where py >nul 2>nul
if errorlevel 1 (
  python "scripts\01_prepare_stratification_inputs.py"
) else (
  py -3 "scripts\01_prepare_stratification_inputs.py"
)

if errorlevel 1 (
  echo.
  echo Data prep did not complete. Review the messages above.
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS_DIR%Refresh_Workflow_Status.ps1"

echo.
echo Data prep complete. Return to the workbook and review the next checklist step.
echo.
pause
