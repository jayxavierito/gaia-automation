param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [string] $Output = "",
    [string] $WorkDirectory = ".\tmp\quantity-extraction"
)

$ErrorActionPreference = "Stop"
$python = if (Test-Path ".\.venv\Scripts\python.exe") {
    ".\.venv\Scripts\python.exe"
} else {
    "python"
}

if (-not (Test-Path -LiteralPath $Source)) {
    throw "Quantity source not found: $Source"
}

$arguments = @(
    ".\quantity_extract.py",
    "--source", $Source,
    "--work-dir", $WorkDirectory
)
if ($Output) {
    $arguments += @("--output", $Output)
}

& $python @arguments
