<#
.SYNOPSIS
  Upgrade the system to W11 by first disabling sleep/hibernate timeouts to avoid mid-upgrade pauses,
  ALWAYS running Windows Update first, then restoring original power settings at the end.
  On success, deletes C:\ProgramData\W11.

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
  [string]$DynamicUpdate = "Disable"
)

# -------- hard-coded behavior (no switches) ----------
# Always run Windows Update first; include driver updates as well.
$RUN_WINDOWS_UPDATE       = $true
$INCLUDE_DRIVER_UPDATES   = $true   # set to $false if you want software-only updates

# ------------------ helpers --------------------------

function Try-Get {
  param([scriptblock]$Do)
  try { & $Do } catch { $null }
}

function Get-PowerTimeouts {
  $outSleep = Try-Get { powercfg /q SCHEME_CURRENT SUB_SLEEP STANDBYIDLE }
  $outHib   = Try-Get { powercfg /q SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE }

  $acHexSleep = if ($outSleep) { ($outSleep | Select-String 'Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value } else { $null }
  $dcHexSleep = if ($outSleep) { ($outSleep | Select-String 'Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value } else { $null }
  $acHexHib   = if ($outHib)   { ($outHib   | Select-String 'Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value } else { $null }
  $dcHexHib   = if ($outHib)   { ($outHib   | Select-String 'Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)').Matches.Groups[1].Value } else { $null }

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
  Write-Host "Setting sleep/hibernate timeouts AC/DC=$StandbyTimeoutAC/$StandbyTimeoutDC; Hib AC/DC=$HibernateTimeoutAC/$HibernateTimeoutDC"
  Try-Get { powercfg /Change standby-timeout-ac   $StandbyTimeoutAC   | Out-Null } | Out-Null
  Try-Get { powercfg /Change standby-timeout-dc   $StandbyTimeoutDC   | Out-Null } | Out-Null
  Try-Get { powercfg /Change hibernate-timeout-ac $HibernateTimeoutAC | Out-Null } | Out-Null
  Try-Get { powercfg /Change hibernate-timeout-dc $HibernateTimeoutDC | Out-Null } | Out-Null
}

function Set-NeverSleep {
  Write-Host "Disabling sleep/hibernate timeouts (setting to Never = 0)..."
  Try-Get { powercfg /Change standby-timeout-ac   0 | Out-Null } | Out-Null
  Try-Get { powercfg /Change standby-timeout-dc   0 | Out-Null } | Out-Null
  Try-Get { powercfg /Change hibernate-timeout-ac 0 | Out-Null } | Out-Null
  Try-Get { powercfg /Change hibernate-timeout-dc 0 | Out-Null } | Out-Null
  Write-Host "Sleep and hibernate timeouts set to Never."
}

# Windows Update using built-in COM API (PowerShell 5.1 compatible)
function Invoke-WindowsUpdate {
  param(
    [switch]$IncludeDrivers,
    [string]$LogFile
  )

  Write-Host "Starting Windows Update search (IncludeDrivers=$($IncludeDrivers.IsPresent))..."

  $session  = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()

  $criteria = "IsInstalled=0 and IsHidden=0"
  if (-not $IncludeDrivers) { $criteria += " and Type='Software'" }

  $sr = $searcher.Search($criteria)
  if ($sr.Updates.Count -eq 0) {
    ("[{0}] WU: No applicable updates found." -f (Get-Date -Format s)) |
      Out-File -FilePath $LogFile -Encoding UTF8 -Append
    return [pscustomobject]@{ Installed=0; RebootRequired=$false; HResult=$sr.ResultCode }
  }

  # Build collection
  $updates = New-Object -ComObject Microsoft.Update.UpdateColl
  $names = @()
  for ($i=0; $i -lt $sr.Updates.Count; $i++) {
    $u = $sr.Updates.Item($i)
    [void]$updates.Add($u)
    $names += $u.Title
  }

  ("[{0}] WU: Found {1} update(s): {2}" -f (Get-Date -Format s), $updates.Count, ($names -join '; ')) |
    Out-File -FilePath $LogFile -Encoding UTF8 -Append

  # Download
  $downloader = $session.CreateUpdateDownloader()
  $downloader.Updates = $updates
  $dr = $downloader.Download()

  ("[{0}] WU: Download result = {1}" -f (Get-Date -Format s), $dr.ResultCode) |
    Out-File -FilePath $LogFile -Encoding UTF8 -Append

  # Filter to downloaded
  $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
  for ($i=0; $i -lt $updates.Count; $i++) {
    if ($updates.Item($i).IsDownloaded) { [void]$toInstall.Add($updates.Item($i)) }
  }

  if ($toInstall.Count -eq 0) {
    ("[{0}] WU: Nothing downloaded; skipping install." -f (Get-Date -Format s)) |
      Out-File -FilePath $LogFile -Encoding UTF8 -Append
    return [pscustomobject]@{ Installed=0; RebootRequired=$false; HResult=$dr.ResultCode }
  }

  # Install silently
  $installer = $session.CreateUpdateInstaller()
  $installer.ForceQuiet = $true
  $installer.Updates = $toInstall
  $ir = $installer.Install()

  # Count succeeded (2) or succeeded with errors (3)
  $installedCount = 0
  for ($i = 0; $i -lt $toInstall.Count; $i++) {
    $ur = $ir.GetUpdateResult($i)
    if ($ur.ResultCode -eq 2 -or $ur.ResultCode -eq 3) { $installedCount++ }
  }

  ("[{0}] WU: Install result = {1}; RebootRequired={2}; InstalledCount={3}" -f `
    (Get-Date -Format s), $ir.ResultCode, $ir.RebootRequired, $installedCount) |
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

function Get-ExitMessage {
  param([int]$Code)
  $map = @{
    0          = 'Success'
    1          = 'Restart required to continue'
    3          = 'General error'
    4          = 'Compatibility blocks (hard)'
    5          = 'Download/DU failure'
    302        = 'Cancelled'
    3010       = 'Success with reboot required (WU)'
    0xC1900101 = 'Driver rollback (C1900101)'
    0xC1900208 = 'Incompatible app detected'
    0xC1900204 = 'Unsupported OS edition/target'
    0x8007001F = 'Device/driver error during upgrade'
  }
  if ($map.ContainsKey($Code)) { $map[$Code] } else { 'Unknown' }
}

# --------------------- main ---------------------------

$prevTimeouts = $null
$src = $null
$exit = $null
$w11Folder = "C:\ProgramData\W11"
$statusPath = Join-Path -Path $env:SystemDrive -ChildPath "Win11SetupLogs\status.txt"
New-Item -Path (Split-Path $statusPath) -ItemType Directory -Force | Out-Null

try {
  # Capture and disable power timeouts
  try {
    $prevTimeouts = Get-PowerTimeouts
    Write-Host ("Captured timeouts: Sleep AC/DC={0}/{1}  Hib AC/DC={2}/{3}" -f `
      $prevTimeouts.StandbyTimeoutAC, $prevTimeouts.StandbyTimeoutDC, $prevTimeouts.HibernateTimeoutAC, $prevTimeouts.HibernateTimeoutDC)
  } catch {
    Write-Warning "Could not capture existing power timeouts: $($_.Exception.Message)"
  }

  Try-Get { Set-NeverSleep } | Out-Null

  # ALWAYS: Windows Update first
  $wuRebootNeeded = $false
  if ($RUN_WINDOWS_UPDATE) {
    ("[{0}] WU: Starting pre-upgrade updates (IncludeDrivers={1})" -f (Get-Date -Format s), $INCLUDE_DRIVER_UPDATES) |
      Out-File -FilePath $statusPath -Encoding UTF8 -Append

    try {
      $wu = Invoke-WindowsUpdate -IncludeDrivers:($INCLUDE_DRIVER_UPDATES) -LogFile $statusPath
      if ($wu -and $wu.RebootRequired) {
        $wuRebootNeeded = $true
        ("[{0}] WU: Reboot required before OS upgrade. Aborting upgrade step." -f (Get-Date -Format s)) |
          Out-File -FilePath $statusPath -Encoding UTF8 -Append
      }
    } catch {
      ("[{0}] WU: ERROR {1}" -f (Get-Date -Format s), $_.Exception.Message) |
        Out-File -FilePath $statusPath -Encoding UTF8 -Append
    }
  }

  if ($wuRebootNeeded) {
    $exit = 3010
    $msg  = Get-ExitMessage -Code $exit
    ("[{0}] ExitCode={1} ({2}) DynamicUpdate={3}" -f (Get-Date -Format s), $exit, $msg, $DynamicUpdate) |
      Out-File -FilePath $statusPath -Encoding UTF8 -Append
    $global:LASTEXITCODE = $exit
    Write-Output "Windows Update requires a reboot (ExitCode $exit - $msg). Reboot, then run this script again to proceed with the OS upgrade."
    return
  }

  # Resolve source & run upgrade
  $src  = Resolve-SetupSource -Source $Source
  $exit = Start-Upgrade -SetupExe $src.Setup -DynamicUpdate $DynamicUpdate

  $msg = Get-ExitMessage -Code $exit
  ("[{0}] ExitCode={1} ({2}) DynamicUpdate={3}" -f (Get-Date -Format s), $exit, $msg, $DynamicUpdate) |
    Out-File -FilePath $statusPath -Encoding UTF8 -Append

  $global:LASTEXITCODE = $exit
  Write-Output "Setup ExitCode=$exit ($msg). DynamicUpdate=$DynamicUpdate. Reboot later to continue."
}
finally {
  # Dismount ISO if mounted
  Try-Get { if ($src -and $src.Mount) { Dismount-DiskImage -ImagePath $Source | Out-Null } } | Out-Null

  # Restore prior power timeouts if captured (fallbacks to 0 if any nulls)
  Try-Get {
    if ($prevTimeouts -and $prevTimeouts.StandbyTimeoutAC -ne $null) {
      $sbAC  = if ($prevTimeouts.StandbyTimeoutAC   -ne $null) { [int]$prevTimeouts.StandbyTimeoutAC   } else { 0 }
      $sbDC  = if ($prevTimeouts.StandbyTimeoutDC   -ne $null) { [int]$prevTimeouts.StandbyTimeoutDC   } else { 0 }
      $hibAC = if ($prevTimeouts.HibernateTimeoutAC -ne $null) { [int]$prevTimeouts.HibernateTimeoutAC } else { 0 }
      $hibDC = if ($prevTimeouts.HibernateTimeoutDC -ne $null) { [int]$prevTimeouts.HibernateTimeoutDC } else { 0 }
      Set-PowerTimeouts -StandbyTimeoutAC $sbAC -StandbyTimeoutDC $sbDC -HibernateTimeoutAC $hibAC -HibernateTimeoutDC $hibDC
      Write-Host "Restored previous power timeouts."
    } else {
      Write-Warning "Previous power timeouts unknown; leaving current settings."
    }
  } | Out-Null

  # If setup succeeded (ExitCode 0), remove the W11 folder
  Try-Get {
    if ($exit -eq 0 -and (Test-Path -LiteralPath $w11Folder)) {
      Write-Host "Upgrade reported success. Removing $w11Folder ..."
      Remove-Item -LiteralPath $w11Folder -Recurse -Force
      Write-Host "Removed $w11Folder."
    }
  } | Out-Null
}
