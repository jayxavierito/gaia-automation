param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [string] $Template = ".\assets\gaia_suzu_7sheet_template.xlsx",
    [string] $PriceDate = "",
    [string] $TreeCategory = "",
    [string] $ReferenceIndex = ".\references\suzu_reference_index.json",
    [string] $Rules = ".\mapping_rules.csv",
    [string[]] $CodeHistory = @(),
    [string] $District = "珠洲",
    [string] $Location = "",
    [string] $WorkCategory = "橋梁工事",
    [string] $ExtractionWorkDirectory = ".\tmp\quantity-extraction",
    [string] $Output = ".\outputs\GAIA_candidate.xlsx"
)

$ErrorActionPreference = "Stop"
$python = if (Test-Path ".\.venv\Scripts\python.exe") {
    ".\.venv\Scripts\python.exe"
} else {
    "python"
}

if (-not (Test-Path -LiteralPath $ReferenceIndex)) {
    $ReferenceIndex = ".\assets\suzu_level_tree_index.json"
}

foreach ($path in @($Source, $Template, $ReferenceIndex, $Rules) + $CodeHistory) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required file not found: $path"
    }
}

if ($PriceDate -and $PriceDate -notmatch "^\d{4}-\d{2}-\d{2}$") {
    throw "PriceDate must use YYYY-MM-DD, for example 2025-10-24."
}

$converterArgs = @(
    ".\gaia_converter.py",
    "--source", $Source,
    "--template", $Template,
    "--rules", $Rules,
    "--reference-index", $ReferenceIndex,
    "--output", $Output,
    "--extraction-work-dir", $ExtractionWorkDirectory
)

if ($PriceDate) {
    $converterArgs += @("--price-date", $PriceDate)
}
if ($TreeCategory) {
    $converterArgs += @("--tree-category", $TreeCategory)
}

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
