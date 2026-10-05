param([string]$OutputRoot = '')
$ErrorActionPreference = 'Stop'
$source = Split-Path -Parent $PSScriptRoot
$distRoot = if ($OutputRoot) { $OutputRoot } else { Join-Path $source 'dist\SCOPE_Standalone' }
$workRoot = Join-Path $source 'build\SCOPE_Standalone'
$arguments = @(
    '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed', '--name', 'SCOPE',
    '--distpath', $distRoot, '--workpath', $workRoot, '--specpath', $workRoot,
    '--hidden-import', 'ortools.sat.python.cp_model', '--collect-binaries', 'ortools',
    '--hidden-import', 'pyreadstat',
    '--exclude-module', 'torch', '--exclude-module', 'torchvision', '--exclude-module', 'torchaudio',
    '--exclude-module', 'matplotlib', '--exclude-module', 'pyarrow', '--exclude-module', 'h5py',
    '--exclude-module', 'scipy',
    '--exclude-module', 'IPython', '--exclude-module', 'pytest', '--exclude-module', 'jedi',
    '--exclude-module', 'parso', '--exclude-module', 'nbformat', '--exclude-module', 'jupyter',
    '--exclude-module', 'notebook', '--exclude-module', 'traitlets', '--exclude-module', 'pygments',
    '--exclude-module', 'zmq', '--exclude-module', 'fsspec',
    '--add-data', "$source\scope\legacy;scope\legacy",
    '--add-data', "$source\scope\configs;scope\configs",
    '--add-data', "$source\profiles;profiles",
    '--add-data', "$source\gui\assets;gui\assets",
    (Join-Path $source 'gui\app.py')
)
& python @arguments
if ($LASTEXITCODE -ne 0) { throw "PyInstaller exited with code $LASTEXITCODE." }
Write-Host "SCOPE standalone app created: $(Join-Path $distRoot 'SCOPE')"
