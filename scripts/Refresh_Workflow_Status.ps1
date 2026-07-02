$ErrorActionPreference = "Stop"

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appDir = Split-Path -Parent $scriptsDir
$workbookPath = Join-Path $scriptsDir "Stratification_Optimizer_Config.xlsx"

$rawDir = Join-Path $appDir "inputs\raw_data"
$targetsDir = Join-Path $appDir "inputs\targets"
$masterPath = Join-Path $appDir "inputs\master_sample\MASTER_SAMPLE.xlsx"
$poolsDir = Join-Path $appDir "inputs\replacement_pools"

$requiredRawFiles = @(
    "stratification_summary.xlsx",
    "stratification_tolerance_rules.xlsx",
    "fr_teid_v1.xlsx",
    "basc4nb_scored_trsc.sas7bdat",
    "basc4nb_scored_trsa.sas7bdat",
    "basc4nb_scored_prsc.sas7bdat",
    "basc4nb_scored_prsa.sas7bdat",
    "basc4nb_scored_srpc.sas7bdat",
    "basc4nb_scored_srpa.sas7bdat"
)

$missingRawFiles = @()
foreach ($name in $requiredRawFiles) {
    if (-not (Test-Path (Join-Path $rawDir $name))) {
        $missingRawFiles += $name
    }
}

$targetFiles = @()
if (Test-Path $targetsDir) {
    $targetFiles = Get-ChildItem -LiteralPath $targetsDir -File -Filter "*.xlsx" |
        Where-Object { -not $_.Name.StartsWith("~$") }
}

$poolFiles = @()
if (Test-Path $poolsDir) {
    $poolFiles = Get-ChildItem -LiteralPath $poolsDir -File -Filter "*_case_swap.xlsx" |
        Where-Object { -not $_.Name.StartsWith("~$") }
}

$inputFilesAdded = ($missingRawFiles.Count -eq 0 -and $targetFiles.Count -gt 0)
$prepOutputsExist = ((Test-Path $masterPath) -and $poolFiles.Count -gt 0)

$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Host "Excel is not currently open, so live workbook status was not updated."
    Write-Host "Input files added: $inputFilesAdded"
    Write-Host "Prep outputs exist: $prepOutputsExist"
    if ($missingRawFiles.Count -gt 0) {
        Write-Host "Missing raw/prep files:"
        foreach ($name in $missingRawFiles) {
            Write-Host " - $name"
        }
    }
    exit 0
}

$targetWorkbook = $null
foreach ($wb in $excel.Workbooks) {
    if ($wb.FullName -eq $workbookPath) {
        $targetWorkbook = $wb
        break
    }
}

if ($null -eq $targetWorkbook) {
    Write-Host "Config workbook is not open in Excel, so live status was not updated."
    Write-Host "Input files added: $inputFilesAdded"
    Write-Host "Prep outputs exist: $prepOutputsExist"
    if ($missingRawFiles.Count -gt 0) {
        Write-Host "Missing raw/prep files:"
        foreach ($name in $missingRawFiles) {
            Write-Host " - $name"
        }
    }
    exit 0
}

$ws = $targetWorkbook.Worksheets.Item("START_HERE_WORKFLOW")
$ws.Range("G4").Value2 = $(if ($inputFilesAdded) { "[x]" } else { "[ ]" })
$ws.Range("G5").Value2 = $(if ($prepOutputsExist) { "[x]" } else { "[ ]" })

if ($inputFilesAdded) {
    $ws.Range("H4").Value2 = "Input files added"
} else {
    $ws.Range("H4").Value2 = "Input files needed"
}

if ($prepOutputsExist) {
    $ws.Range("H5").Value2 = "Data prep complete"
} else {
    $ws.Range("H5").Value2 = "Data prep not run"
}

$targetWorkbook.Save()

Write-Host "Workflow status refreshed in Excel."
Write-Host "Input files added: $inputFilesAdded"
Write-Host "Prep outputs exist: $prepOutputsExist"
if ($missingRawFiles.Count -gt 0) {
    Write-Host "Missing raw/prep files:"
    foreach ($name in $missingRawFiles) {
        Write-Host " - $name"
    }
}
