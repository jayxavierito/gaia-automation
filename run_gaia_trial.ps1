param(
    [Parameter(Mandatory = $true)] [string] $Candidate,
    [Parameter(Mandatory = $true)] [string] $SourceReview,
    [Parameter(Mandatory = $true)] [string] $ProjectName,
    [string] $FolderName = "",
    [string] $OutputDir = ".\outputs\ui_trial",
    [switch] $ExportReview,
    [switch] $DryRun
)

$ErrorActionPreference = "Stop"
$python = if (Test-Path ".\.venv\Scripts\python.exe") {
    ".\.venv\Scripts\python.exe"
} else {
    "python"
}

foreach ($path in @($Candidate, $SourceReview)) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required file not found: $path"
    }
}
$trialArgs = @(
    ".\gaia_ui_automation.py",
    "--candidate", $Candidate,
    "--source-review", $SourceReview,
    "--project-name", $ProjectName,
    "--output-dir", $OutputDir
)
if ($FolderName) {
    $trialArgs += @("--folder-name", $FolderName)
}
if ($ExportReview) {
    $trialArgs += "--export-review"
}
if ($DryRun) {
    $trialArgs += "--dry-run"
} else {
    $trialArgs += "--confirm-start"
}

& $python @trialArgs
