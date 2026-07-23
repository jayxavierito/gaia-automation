param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [string] $Template = ".\assets\gaia_suzu_7sheet_template.xlsx",
    [string] $Catalog = ".\assets\gaia_import_catalog.json",
    [string] $PriceDate = "",
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

foreach ($path in @($Source, $Template, $Catalog)) {
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
    "--catalog", $Catalog,
    "--output", $Output,
    "--extraction-work-dir", $ExtractionWorkDirectory
)

if ($PriceDate) {
    $converterArgs += @("--price-date", $PriceDate)
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
