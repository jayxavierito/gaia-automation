param(
    [string] $Name = "GaiaQuantityConverter"
)

$ErrorActionPreference = "Stop"
$python = ".\.venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Python environment not found. Run .\setup.ps1 first."
}

& $python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is not installed. Run: .\.venv\Scripts\python.exe -m pip install -r requirements-build.txt"
}

$arguments = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--noupx",
    "--name", $Name,
    "--add-data", "windows_ocr.ps1;.",
    "--add-data", "mapping_rules.csv;.",
    "--add-data", "assets;assets",
    "--collect-all", "pypdfium2",
    "--collect-all", "pdfplumber"
)
if (Test-Path -LiteralPath ".\references\suzu_reference_index.json") {
    $arguments += @(
        "--add-data",
        "references\suzu_reference_index.json;references"
    )
}
$arguments += ".\gaia_quantity_gui.py"

& $python @arguments
if ($LASTEXITCODE -ne 0) {
    throw "EXE build failed."
}

$releaseDirectory = Join-Path ".\dist" $Name
Copy-Item -LiteralPath ".\PORTABLE_README_JA.txt" -Destination $releaseDirectory -Force
Copy-Item -LiteralPath ".\README.md" -Destination (Join-Path $releaseDirectory "README_日本語.md") -Force

$zipPath = Join-Path ".\dist" "$Name-win64.zip"
Compress-Archive -Path (Join-Path $releaseDirectory "*") -DestinationPath $zipPath -Force

Write-Host "Portable EXE: $(Resolve-Path (Join-Path $releaseDirectory "$Name.exe"))"
Write-Host "Distribution ZIP: $(Resolve-Path $zipPath)"
