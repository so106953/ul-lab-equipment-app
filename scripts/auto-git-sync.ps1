$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$git = "C:\Users\10695\.codex\tools\PortableGit\cmd\git.exe"
$logDirectory = Join-Path $env:LOCALAPPDATA "UL-Lab-Equipment"
$logFile = Join-Path $logDirectory "git-auto-sync.log"

New-Item -ItemType Directory -Force -Path $logDirectory | Out-Null

function Write-SyncLog([string]$message) {
  "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $message" | Add-Content -LiteralPath $logFile -Encoding utf8
}

try {
  if (-not (Test-Path $git)) { throw "Portable Git was not found." }

  $changes = & $git -C $projectRoot status --porcelain
  if ([string]::IsNullOrWhiteSpace($changes)) {
    Write-SyncLog "No changes to back up."
    exit 0
  }

  & $git -C $projectRoot add --all
  & $git -C $projectRoot commit -m ("Auto backup " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
  & $git -C $projectRoot push origin main
  Write-SyncLog "Backup pushed successfully."
  exit 0
}
catch {
  Write-SyncLog ("Backup failed: " + $_.Exception.Message)
  exit 1
}
