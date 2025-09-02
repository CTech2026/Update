<#
.SYNOPSIS
  Upgrade the system to W11 by first Disabling system sleep on both AC power (plugged in) and DC power (battery) to send the necessary commands to upgrade the PC to W11.

.NOTES
  - Requires Administrator.
  - Affects the active power plan only.
  - Does not disable display sleep/turn off unless you also set those explicitly.
  - No UI, no prompts
  - No automatic reboot (reboot later yourself)
  - Dynamic Update DISABLED by default to avoid 46% stalls
#>

[CmdletBinding()]
param(
  [string]$Source = "C:\ProgramData\W11\Win11_24H2.iso",
  [ValidateSet("Enable","Disable")]
  [string]$DynamicUpdate = "Disable"
)

function Set-NeverSleep {
    # Get the active power scheme GUID
    $schemeGuid = (powercfg /GetActiveScheme) -replace '.*:\s*',''
    Write-Host "Active Power Scheme: $schemeGuid"

    # 0 = Never
    Write-Host "Disabling sleep on AC (plugged in)..."
    powercfg /Change standby-timeout-ac 0

    Write-Host "Disabling sleep on DC (battery)..."
    powercfg /Change standby-timeout-dc 0

    # Optional: also disable hibernate timeout if desired
    Write-Host "Disabling hibernate timeout on AC/DC..."
    powercfg /Change hibernate-timeout-ac 0
    powercfg /Change hibernate-timeout-dc 0

    Write-Host "Sleep and hibernate timeouts set to Never."
}
 
function Resolve-SetupSource {
  param([string]$Source)
 
  $mount = $null
  $root  = $null
 
  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Source path not found: $Source"
  }
 
  if ($Source -match '\.iso$') {
    $img = Mount-DiskImage -ImagePath $Source -PassThru -ErrorAction Stop
    $vol = Get-Volume -DiskImage $img | Where-Object DriveLetter | Select-Object -First 1
    if (-not $vol) { throw "Could not resolve mounted ISO volume/drive letter." }
    $root = ($vol.DriveLetter + ":\")
    $mount = $img
  } else {
    $root = (Resolve-Path -LiteralPath $Source).Path
  }
 
  $setupExe = Join-Path -Path $root -ChildPath "setup.exe"
  if (-not (Test-Path -LiteralPath $setupExe)) {
    throw "setup.exe not found in '$root'. Provide an ISO or a folder that contains setup.exe."
  }
 
  [PSCustomObject]@{ Root = $root; Mount = $mount; Setup = $setupExe }
}
 
function Start-Upgrade {
  param(
    [string]$SetupExe,
    [string]$DynamicUpdate
  )
 
  $logDir = "$env:SystemDrive\Win11SetupLogs"
  New-Item -Path $logDir -ItemType Directory -Force | Out-Null
 
  # Fully silent, no UI, no auto-reboot; Dynamic Update default = Disable
  $args = @(
    "/auto","upgrade",
    "/quiet",
    "/eula","accept",
    "/noreboot",
    "/DynamicUpdate",$DynamicUpdate,
    "/Telemetry","Disable",
    "/copylogs",$logDir
  )
 
  $p = Start-Process -FilePath $SetupExe -ArgumentList $args -PassThru -Wait
  return $p.ExitCode
}
 
# ---- main ----

try {
    Set-NeverSleep
} catch {
    Write-Error $_
}

$src = $null
try {
  $src = Resolve-SetupSource -Source $Source
  $exit = Start-Upgrade -SetupExe $src.Setup -DynamicUpdate $DynamicUpdate
 
  # Record status; do NOT close the console or reboot
  $statusPath = Join-Path -Path $env:SystemDrive -ChildPath "Win11SetupLogs\status.txt"
  "[{0}] ExitCode={1} (DynamicUpdate={2})" -f (Get-Date -Format s), $exit, $DynamicUpdate |
    Out-File -FilePath $statusPath -Encoding UTF8 -Append
 
  $global:LASTEXITCODE = $exit
  Write-Output "Setup ExitCode=$exit (DynamicUpdate=$DynamicUpdate). Reboot later to continue upgrade."
}
finally {
  try {
    if ($src -and $src.Mount) { Dismount-DiskImage -ImagePath $Source | Out-Null }
  } catch { }

}
