# Path to the most recent CompatData.xml
$CompatFile = Get-ChildItem "C:\$WINDOWS.~BT\Sources\Panther\" -Filter "CompatData*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($CompatFile) {
    Write-Host "Checking $($CompatFile.FullName)`n"
    [xml]$xml = Get-Content $CompatFile.FullName

    # Look for blocking apps
    $blocks = $xml.CompatReport.CompatibilityInfo.BlockingApplications.BlockingApplication
    if ($blocks) {
        $blocks | ForEach-Object {
            Write-Host ("Blocking Application: {0} -> {1}" -f $_.Name, $_.Resolution)
        }
    }

    # Look for blocking drivers
    $drivers = $xml.CompatReport.CompatibilityInfo.BlockingDrivers.BlockingDriver
    if ($drivers) {
        $drivers | ForEach-Object {
            Write-Host ("Blocking Driver: {0} -> {1}" -f $_.Name, $_.Resolution)
        }
    }

    if (-not $blocks -and -not $drivers) {
        Write-Host "No blocking apps or drivers found in report."
    }
} else {
    Write-Host "No CompatData.xml found."
}
