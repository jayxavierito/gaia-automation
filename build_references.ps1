param(
    [Parameter(Mandatory = $true)] [string] $LevelTree,
    [Parameter(Mandatory = $true)] [string] $QuantityGuideline,
    [Parameter(Mandatory = $true)] [string] $NationalReferenceXlsx,
    [Parameter(Mandatory = $true)] [string] $NationalPackagePdf,
    [Parameter(Mandatory = $true)] [string] $IshikawaPackageXlsx,
    [Parameter(Mandatory = $true)] [string] $IshikawaPackagePdf,
    [string] $Output = ".\references\suzu_reference_index.json"
)

$ErrorActionPreference = "Stop"
$python = if (Test-Path ".\.venv\Scripts\python.exe") {
    ".\.venv\Scripts\python.exe"
} else {
    "python"
}

foreach ($path in @(
    $LevelTree,
    $QuantityGuideline,
    $NationalReferenceXlsx,
    $NationalPackagePdf,
    $IshikawaPackageXlsx,
    $IshikawaPackagePdf
)) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Reference file not found: $path"
    }
}

& $python .\build_reference_index.py `
    --level-tree $LevelTree `
    --quantity-guideline $QuantityGuideline `
    --national-reference-xlsx $NationalReferenceXlsx `
    --national-package-pdf $NationalPackagePdf `
    --ishikawa-package-xlsx $IshikawaPackageXlsx `
    --ishikawa-package-pdf $IshikawaPackagePdf `
    --output $Output
