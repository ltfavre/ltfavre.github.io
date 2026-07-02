@echo off
setlocal

cd /d "%~dp0"

if not exist "logs" mkdir "logs"

set "STAMP=%DATE:~-4%%DATE:~4,2%%DATE:~7,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%"
set "STAMP=%STAMP: =0%"
set "LOG=logs\Optimizer_Run_%STAMP%.log"

echo Stratification Optimizer
echo ========================
echo.
echo Started: %DATE% %TIME%
echo Log: %LOG%
echo.
echo Running optimizer. This window will stay open until it finishes.
echo.

(
  echo Stratification Optimizer
  echo Started: %DATE% %TIME%
  echo Working directory: %CD%
  echo.
) > "%LOG%"

"scripts\StratificationOptimizer.exe" >> "%LOG%" 2>&1
set "EXITCODE=%ERRORLEVEL%"

(
  echo.
  echo Finished: %DATE% %TIME%
  echo Exit code: %EXITCODE%
) >> "%LOG%"

echo.
echo Optimizer finished with exit code %EXITCODE%.
echo.
echo Last log lines:
echo ---------------
powershell -NoProfile -Command "Get-Content -LiteralPath '%LOG%' -Tail 40"
echo.
echo Full log saved to:
echo %CD%\%LOG%
echo.
pause
exit /b %EXITCODE%
