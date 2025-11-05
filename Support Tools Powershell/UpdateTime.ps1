<# 
Run from an elevated PowerShell (Run as Administrator).
Sets California time zone and syncs time.
#>

# --- Elevate if needed ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  if ($PSCommandPath) {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
  } else {
    Write-Error "Please run this script from an elevated (Run as administrator) PowerShell."
    exit 1
  }
}

# --- Set California (Pacific) time zone ---
$tz = "Pacific Standard Time"   # Windows TZ ID covering Pacific Time (handles DST automatically)
try {
  Set-TimeZone -Name $tz
  Write-Host "Time zone set to: $tz"
} catch {
  Write-Error "Failed to set time zone to '$tz': $($_.Exception.Message)"
}

# --- Ensure Windows Time service is automatic & running ---
try {
  Set-Service -Name W32Time -StartupType Automatic
  if ((Get-Service -Name W32Time).Status -ne 'Running') {
    Start-Service -Name W32Time
  } else {
    Restart-Service -Name W32Time -Force
  }
  Write-Host "Windows Time (W32Time) service is running and set to Automatic."
} catch {
  Write-Error "Could not configure W32Time: $($_.Exception.Message)"
}

# --- Configure NTP on non-domain machines ---
$partOfDomain = $false
try { $partOfDomain = (Get-CimInstance Win32_ComputerSystem).PartOfDomain } catch {}
if (-not $partOfDomain) {
  $ntp = "time.windows.com,0x9"
  try {
    w32tm /config /manualpeerlist:$ntp /syncfromflags:manual /update | Out-Null
    Write-Host "NTP server configured to: $ntp"
  } catch {
    Write-Warning "Failed to configure NTP server: $($_.Exception.Message)"
  }
} else {
  Write-Host "Domain-joined machine detected; using domain time source."
}

# --- Trigger built-in sync task if present (non-fatal if missing) ---
try { Start-ScheduledTask -TaskPath '\Microsoft\Windows\Time Synchronization\' -TaskName 'SynchronizeTime' -ErrorAction SilentlyContinue } catch {}

# --- Force an immediate time sync ---
try {
  w32tm /resync /nowait
  Write-Host "Time sync triggered."
} catch {
  Write-Warning "Resync failed; rediscovering sources then retrying..."
  try {
    w32tm /config /update /rediscover | Out-Null
    Start-Sleep -Seconds 2
    w32tm /resync /nowait
  } catch {
    Write-Error "Resync failed: $($_.Exception.Message)"
  }
}

# --- Show status and configuration summary ---
Write-Host "`n=== Current Time Status ==="
w32tm /query /status
Write-Host "`n=== Time Configuration ==="
w32tm /query /configuration
