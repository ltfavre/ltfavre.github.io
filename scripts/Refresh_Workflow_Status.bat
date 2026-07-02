@echo off
setlocal

set "SCRIPTS_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS_DIR%Refresh_Workflow_Status.ps1"

echo.
echo Status refresh complete. You can return to the workbook.
echo.
pause
