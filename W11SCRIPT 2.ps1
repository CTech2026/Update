<#
.SYNOPSIS
  Upgrade the system to W11 by first disabling sleep/hibernate timeouts to avoid mid-upgrade pauses,
  then restoring the original power settings at the end. On success, deletes C:\ProgramData\W11.
  NOW ALSO: Optionally runs Windows Update first (software + optional drivers).

.NOTES
  - Requires Administrator.
  - Affects the active power plan only.
  - Does not change display sleep unless you set those explicitly.
  - No UI, no prompts. No automatic reboot (reboot later yourself).
  - Dynamic Update DISABLED by default to avoid 46% stalls.
#>

[CmdletBinding()]
param(
  [string]$Source = "C:\ProgramData\W11\Win11_24H2.iso",
  [ValidateSet("Enable","Disable")]
  [string]$DynamicUpdate = "Disable",

  # New: control Windows Update behavior
  [switch]$RunWindowsUpdate,          # If present, run Windows Update before OS setup
  [switch]$IncludeDriverUpdates       # If present, include drivers in Windows Update
)

# --- Power settings helpers ---
function Get-PowerTimeouts {
  $outSleep = powercfg /q SCHEME_CURRENT SUB_SLEEP STANDBYIDLE
  $outHib   = powercfg /q SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE

  $acHexSleep = ($outSleep | Select-String 'Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value
  $dcHexSleep = ($outSleep | Select-String 'Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value
  $acHexHib   = ($outHib   | Select-String 'Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value
  $dcHexHib   = ($outHib   | Select-String 'Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value

  [pscustomobject]@{
    StandbyTimeoutAC   = if ($acHexSleep) { [Convert]::ToInt32($acHexSleep,16) } else { $null }
    StandbyTimeoutDC   = if ($dcHexSleep) { [Convert]::ToInt32($dcHexSleep,16) } else { $null }
    HibernateTimeoutAC = if ($acHexHib)   { [Convert]::ToInt32($acHexHib,16)   } else { $null }
    HibernateTimeoutDC = if ($dcHexHib)   { [Convert]::ToInt32($dcHexHib,16)   } else { $null }
  }
}

function Set-PowerTimeouts {
  param(
    [Parameter(Mandatory)] [int]$StandbyTimeoutAC,
    [Parameter(Mandatory)] [int]$StandbyTimeoutDC,
    [Parameter(Mandatory)] [int]$HibernateTimeoutAC,
    [Parameter(Mandatory)] [int]$HibernateTimeoutDC
  )
  Write-Host "Setting sleep/hibernate timeouts (AC=$StandbyTimeoutAC/DC=$StandbyTimeoutDC; HibAC=$HibernateTimeoutAC/HibDC=$HibernateTimeoutDC)..."
  powercfg /Change standby-timeout-ac   $StandbyTimeoutAC   | Out-Null
  powercfg /Change standby-timeout-dc   $StandbyTimeoutDC   | Out-Null
  powercfg /Change hibernate-timeout-ac $HibernateTimeoutAC | Out-Null
  powercfg /Change hibernate-timeout-dc $HibernateTimeoutDC | Out-Null
}

function Set-NeverSleep {
  Write-Host "Disabling sleep/hibernate timeouts for AC/DC (setting to Never = 0)..."
  powercfg /Change standby-timeout-ac   0 | Out-Null
  powercfg /Change standby-timeout-dc   0 | Out-Null
  powercfg /Change hibernate-timeout-ac 0 | Out-Null
  powercfg /Change hibernate-timeout-dc 0 | Out-Null
  Write-Host "Sleep and hibernate timeouts set to Never."
}

# --- Windows Update (built-in COM API) ---
function Invoke-WindowsUpdate {
  param(
    [switch]$IncludeDrivers,
    [string]$LogFile
  )

  Write-Host "Starting Windows Update search (IncludeDrivers=$($IncludeDrivers.IsPresent))..."
  Add-Type -AssemblyName 'System.Runtime.InteropServices' | Out-Null

  $session   = New-Object -ComObject Microsoft.Update.Session
  $searcher  = $session.CreateUpdateSearcher()

  $criteria = "IsInstalled=0 and IsHidden=0"
  if (-not $IncludeDrivers) { $criteria += " and Type='Software'" }

  $sr = $searcher.Search($criteria)
  if ($sr.Updates.Count -eq 0) {
    "[{0}] WU: No applicable updates found." -f (Get-Date -Format s) | Out-File -FilePath $LogFile -Encoding UTF8 -Append
    return [pscustomobject]@{ Installed=0; RebootRequired=$false; HResult=$sr.ResultCode }
  }

  # Build collection
  $updates = New-Object -ComObject Microsoft.Update.UpdateColl
  $names = @()
  for ($i=0; $i -lt $sr.Updates.Count; $i++) {
    [void]$updates.Add($sr.Updates.Item($i))
    $names += $sr.Updates.Item($i).Title
  }

  "[{0}] WU: Found {1} update(s): {2}" -f (Get-Date -Format s), $updates.Count, ($names -join '; ') |
    Out-File -FilePath $LogFile -Encoding UTF8 -Append

  # Download
  $downloader = $session.CreateUpdateDownloader()
  $downloader.Updates = $updates
  $dr = $downloader.Download()
  "[{0}] WU: Download result = {1}" -f (Get-Date -Format s), $dr.ResultCode |
    Out-File -FilePath $LogFile -Encoding UTF8 -Append

  # Filter to downloaded
  $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
  for ($i=0; $i -lt $updates.Count; $i++) {
    if ($updates.Item($i).IsDownloaded) { [void]$toInstall.Add($updates.Item($i)) }
  }

  if ($toInstall.Count -eq 0) {
    "[{0}] WU: Nothing downloaded; skipping install." -f (Get-Date -Format s) | Out-File -FilePath $LogFile -Encoding UTF8 -Append
    return [pscustomobject]@{ Installed=0; RebootRequired=$false; HResult=$dr.ResultCode }
  }

  $installer = $session.CreateUpdateInstaller()
  $installer.ForceQuiet = $true
  $installer.Updates = $toInstall

  $ir = $installer.Install()
  $installedCount = ($ir.ResultCode -eq 2) ? $ir.Updates.Count : ($ir.Updates.Count)  # 2 = Succeeded
  "[{0}] WU: Install result = {1}; RebootRequired={2}; InstalledCount={3}" -f (Get-Date -Format s), $ir.ResultCode, $ir.RebootRequired, $installedCount |
    Out-File -FilePath $LogFile -Encoding UTF8 -Append

  return [pscustomobject]@{
    Installed      = $installedCount
    RebootRequired = [bool]$ir.RebootRequired
    HResult        = $ir.ResultCode
  }
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

  Write-Host "Launching: $SetupExe $($args -join ' ')"
  $p = Start-Process -FilePath $SetupExe -ArgumentList $args -PassThru -Wait
  return $p.ExitCode
}

# ---- main ----
$prevTimeouts = $null
$src = $null
$exit = $null
$w11Folder = "C:\ProgramData\W11"
$statusPath = Join-Path -Path $env:SystemDrive -ChildPath "Win11SetupLogs\status.txt"
New-Item -Path (Split-Path $statusPath) -ItemType Directory -Force | Out-Null

try {
  # Capture current power timeouts, then set to Never
  try {
    $prevTimeouts = Get-PowerTimeouts
    Write-Host ("Captured current timeouts: Sleep AC/DC = {0}/{1} min, Hibernate AC/DC = {2}/{3} min" -f `
      $prevTimeouts.StandbyTimeoutAC, $prevTimeouts.StandbyTimeoutDC, $prevTimeouts.HibernateTimeoutAC, $prevTimeouts.HibernateTimeoutDC)
  } catch {
    Write-Warning "Could not capture existing power timeouts: $($_.Exception.Message)"
  }

  try { Set-NeverSleep } catch { Write-Warning "Failed to set NeverSleep: $($_.Exception.Message)" }

  # (NEW) Run Windows Update first, if requested
  $wuRebootNeeded = $false
  if ($RunWindowsUpdate.IsPresent) {
    "[{0}] WU: Starting pre-upgrade updates (IncludeDrivers={1})" -f (Get-Date -Format s), $IncludeDriverUpdates.IsPresent |
      Out-File -FilePath $statusPath -Encoding UTF8 -Append

    try {
      $wu = Invoke-WindowsUpdate -IncludeDrivers:$IncludeDriverUpdates.IsPresent -LogFile $statusPath
      if ($wu.RebootRequired) {
        $wuRebootNeeded = $true
        "[{0}] WU: Reboot required before OS upgrade. Aborting upgrade step." -f (Get-Date -Format s) |
          Out-File -FilePath $statusPath -Encoding UTF8 -Append
      }
    } catch {
      "[{0}] WU: ERROR {1}" -f (Get-Date -Format s), $_.Exception.Message |
        Out-File -FilePath $statusPath -Encoding UTF8 -Append
    }
  }

  if ($wuRebootNeeded) {
    $exit = 3010  # Standard "success with reboot required"
    $global:LASTEXITCODE = $exit
    Write-Output "Windows Update requires a reboot (ExitCode 3010). Reboot, then run this script again to continue the OS upgrade."
  }
  else {
    # Resolve source & run upgrade
    $src  = Resolve-SetupSource -Source $Source
    $exit = Start-Upgrade -SetupExe $src.Setup -DynamicUpdate $DynamicUpdate

    # Record status; do NOT close the console or reboot
    "[{0}] ExitCode={1} (DynamicUpdate={2})" -f (Get-Date -Format s), $exit, $DynamicUpdate |
      Out-File -FilePath $statusPath -Encoding UTF8 -Append

    $global:LASTEXITCODE = $exit
    Write-Output "Setup ExitCode=$exit (DynamicUpdate=$DynamicUpdate). Reboot later to continue upgrade."
  }
}
finally {
  # Always dismount the ISO if we mounted it
  try {
    if ($src -and $src.Mount) { Dismount-DiskImage -ImagePath $Source | Out-Null }
  } catch {
    Write-Warning "Failed to dismount ISO: $($_.Exception.Message)"
  }

  # Restore prior power timeouts if we captured them
  try {
    if ($prevTimeouts -and $prevTimeouts.StandbyTimeoutAC -ne $null) {
      Set-PowerTimeouts `
        -StandbyTimeoutAC   $prevTimeouts.StandbyTimeoutAC `
        -StandbyTimeoutDC   $prevTimeouts.StandbyTimeoutDC `
        -HibernateTimeoutAC $prevTimeouts.HibernateTimeoutAC `
        -HibernateTimeoutDC $prevTimeouts.HibernateTimeoutDC
      Write-Host "Restored previous power timeouts."
    } else {
      Write-Warning "Previous power timeouts unknown; leaving current settings in place."
    }
  } catch {
    Write-Warning "Failed to restore power timeouts: $($_.Exception.Message)"
  }

  # If setup succeeded (ExitCode 0), remove the W11 folder
  try {
    if ($exit -eq 0 -and (Test-Path -LiteralPath $w11Folder)) {
      Write-Host "Upgrade reported success (ExitCode 0). Removing $w11Folder ..."
      Remove-Item -LiteralPath $w11Folder -Recurse -Force -ErrorAction Stop
      Write-Host "Removed $w11Folder."
    }
  } catch {
    Write-Warning "Failed to remove w11 Folder: $($_.Exception.Message)"
  }
}
