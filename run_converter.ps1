param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Template,
    [Parameter(Mandatory = $true)] [string] $PriceDate,
    [Parameter(Mandatory = $true)] [string] $TreeCategory,
    [string] $ReferenceIndex = ".\references\suzu_reference_index.json",
    [string] $Rules = ".\mapping_rules.csv",
    [string[]] $CodeHistory = @(),
    [string] $District = "",
    [string] $Location = "",
    [string] $WorkCategory = "",
    [string] $ExtractionWorkDirectory = ".\tmp\quantity-extraction",
    [string] $Output = ".\outputs\GAIA_candidate.xlsx"
)

$ErrorActionPreference = "Stop"
$python = if (Test-Path ".\.venv\Scripts\python.exe") {
    ".\.venv\Scripts\python.exe"
} else {
    "python"
}

if ($CodeHistory.Count -eq 0) {
    $CodeHistory = @($Template)
}

foreach ($path in @($Source, $Template, $ReferenceIndex, $Rules) + $CodeHistory) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required file not found: $path"
    }
}

if ($PriceDate -notmatch "^\d{4}-\d{2}-\d{2}$") {
    throw "PriceDate must use YYYY-MM-DD, for example 2025-10-24."
}

$converterArgs = @(
    ".\gaia_converter.py",
    "--source", $Source,
    "--template", $Template,
    "--rules", $Rules,
    "--reference-index", $ReferenceIndex,
    "--tree-category", $TreeCategory,
    "--output", $Output,
    "--extraction-work-dir", $ExtractionWorkDirectory,
    "--price-date", $PriceDate
)

foreach ($historyPath in $CodeHistory) {
    $converterArgs += @("--code-history", $historyPath)
}

if ($District) {
    $converterArgs += @("--district", $District)
}
if ($Location) {
    $converterArgs += @("--location", $Location)
}
if ($WorkCategory) {
    $converterArgs += @("--work-category", $WorkCategory)
}

& $python @converterArgs
