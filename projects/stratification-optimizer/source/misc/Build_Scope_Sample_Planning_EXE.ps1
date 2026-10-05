$ErrorActionPreference = 'Stop'
$source = Split-Path -Parent $PSScriptRoot
$distRoot = Join-Path $source 'dist\SCOPE_Sample_Planning_Standalone'
$workRoot = Join-Path $source 'build\SCOPE_Sample_Planning_Standalone'
$arguments = @(
    '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed', '--name', 'SCOPE_Sample_Planning',
    '--distpath', $distRoot, '--workpath', $workRoot, '--specpath', $workRoot,
    '--collect-all', 'ortools', '--collect-all', 'pyreadstat', '--collect-all', 'scipy',
    '--exclude-module', 'torch', '--exclude-module', 'torchvision', '--exclude-module', 'torchaudio',
    '--exclude-module', 'matplotlib', '--exclude-module', 'pyarrow', '--exclude-module', 'h5py',
    '--add-data', "$source\scope\legacy;scope\legacy",
    '--add-data', "$source\scope\configs;scope\configs",
    '--add-data', "$source\profiles;profiles",
    '--add-data', "$source\gui\assets;gui\assets",
    (Join-Path $source 'gui\sample_planning_app.py')
)
& python @arguments
if ($LASTEXITCODE -ne 0) { throw "PyInstaller exited with code $LASTEXITCODE." }
Write-Host "SCOPE Sample Planning app created: $(Join-Path $distRoot 'SCOPE_Sample_Planning')"
