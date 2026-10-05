$ErrorActionPreference = 'Stop'
$source = Split-Path -Parent $PSScriptRoot
$distRoot = Join-Path $source 'dist\SCOPE_Profile_Manager_Standalone'
$workRoot = Join-Path $source 'build\SCOPE_Profile_Manager_Standalone'
$arguments = @(
    '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed', '--name', 'SCOPE_Profile_Manager',
    '--distpath', $distRoot, '--workpath', $workRoot, '--specpath', $workRoot,
    '--exclude-module', 'torch', '--exclude-module', 'torchvision', '--exclude-module', 'torchaudio',
    '--exclude-module', 'matplotlib', '--exclude-module', 'scipy', '--exclude-module', 'pandas',
    '--hidden-import', 'openpyxl',
    '--add-data', "$source\profiles;profiles",
    (Join-Path $source 'gui\profile_manager.py')
)
& python @arguments
if ($LASTEXITCODE -ne 0) { throw "PyInstaller exited with code $LASTEXITCODE." }
Write-Host "SCOPE Profile Manager created: $(Join-Path $distRoot 'SCOPE_Profile_Manager')"
