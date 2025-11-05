<#
    Printer Client – search / add / remove (INF-based driver install, folder-aware)
    - SEARCH: Lists queues on a print server (fallback to NET VIEW if RPC is blocked)
    - ADD   : Prompts for DRIVER .INF or a FOLDER; if a folder is given it finds printer INFs,
              lets you pick, stages driver (with model selection + fallbacks), then adds \\server\share
    - REMOVE: Removes local printer(s); optionally uninstalls driver if unused (with Force option)
    - EXIT  : exit/quit/q works reliably via labeled loop

    PowerShell 5.1 compatible
#>

# ====== RUN FIRST: copy INF to local temp ======
$src = '\\BLANK_\BLANK_\BLANK_\BLANK_.inf'
$dst = 'C:\BLANK_\BLANK_.inf'
Copy-Item -Path $src -Destination $dst -Force
# ===============================================

# ====== SETTINGS ======
$Global:PrintServer = "BLANK_"   # e.g., "YOUR_SERVER_NAME"
# ======================

function Write-Header { param([string]$Text) Write-Host "`n=== $Text ===" -ForegroundColor Cyan }

function Ensure-LocalSpooler {
    $svc = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -ne 'Running') {
        try {
            Write-Host "Starting local Print Spooler..." -ForegroundColor Yellow
            Start-Service -Name Spooler -ErrorAction Stop
        } catch { Write-Warning "Could not start local Spooler: $_" }
    }
}

function Get-SharedPrintersFromNetView {
    param([Parameter(Mandatory)][string]$Server)
    $output = & cmd.exe /c "net view \\$Server" 2>$null
    if (-not $output) { return @() }
    $printers = @()
    foreach ($line in $output) {
        if ($line -match '^\s*([^\s$][^\s]*)\s{2,}') {
            $name = $Matches[1].Trim()
            if ($name -in @('IPC$', 'ADMIN$', 'C$','D$','Print$')) { continue }
            $printers += [pscustomobject]@{ Name = $name; ShareName = $name; DriverName = $null }
        }
    }
    return $printers
}

function List-ServerQueues {
    param([Parameter(Mandatory)][string]$Server)
    Ensure-LocalSpooler
    try {
        Write-Header "Queues on \\$Server (remote spooler)"
        $q = Get-Printer -ComputerName $Server -ErrorAction Stop | Sort-Object Name
        if ($q -and $q.Count -gt 0) {
            $i=1; $q | ForEach-Object {
                $share = if ($_.ShareName) { $_.ShareName } else { $_.Name }
                Write-Host ("[{0}] \\{1}\{2}  (Driver: {3})" -f $i, $Server, $share, $_.DriverName); $i++
            }
            return $q
        }
    } catch {
        Write-Warning "Remote spooler query failed. Falling back to NET VIEW…"
    }

    $fallback = Get-SharedPrintersFromNetView -Server $Server
    if ($fallback.Count -gt 0) {
        Write-Header "Queues on \\$Server (net view fallback)"
        $i=1; $fallback | ForEach-Object {
            Write-Host ("[{0}] \\{1}\{2}" -f $i, $Server, $_.ShareName); $i++
        }
        return $fallback
    }

    Write-Error "Failed to query \\$Server via both methods. Verify server name, connectivity, and permissions."
    return @()
}

function Choose-ServerAndQueue {
    if ([string]::IsNullOrWhiteSpace($Global:PrintServer) -or $Global:PrintServer -eq 'YOUR_PRINT_SERVER_NAME') {
        $Global:PrintServer = Read-Host "Enter print server (e.g., BLANK_)"
    }
    if ([string]::IsNullOrWhiteSpace($Global:PrintServer) -or $Global:PrintServer -eq 'YOUR_PRINT_SERVER_NAME') {
        Write-Warning "No valid print server set."; return $null
    }

    $queues = List-ServerQueues -Server $Global:PrintServer
    if ($queues.Count -gt 0) {
        $i=1; $queues | ForEach-Object {
            $share = if ($_.ShareName) { $_.ShareName } else { $_.Name }
            Write-Host ("[{0}] \\{1}\{2}" -f $i, $Global:PrintServer, $share); $i++
        }
    } else { return $null }

    $choice = Read-Host "Enter number to ADD that queue (or press Enter to cancel)"
    if ($choice -notmatch '^\d+$') { return $null }

    $sel = $queues[[int]$choice-1]
    $share = if ($sel.ShareName) { $sel.ShareName } else { $_.Name }
    $serverDriver = $null
    if ($sel.PSObject.Properties.Match('DriverName').Count -gt 0) { $serverDriver = $sel.DriverName }

    return @{ Server = $Global:PrintServer; Share = $share; ServerDriverName = $serverDriver; QueueName = $sel.Name }
}

# ---------- INF-aware driver staging (with model discovery + fallbacks) ----------
function Get-InfModelNames {
    param([Parameter(Mandatory)][string]$InfPath)

    if (-not (Test-Path $InfPath -PathType Leaf)) { throw "INF file not found: $InfPath" }

    $lines = Get-Content -LiteralPath $InfPath -ErrorAction Stop
    $models = @()

    foreach ($ln in $lines) {
        if ($ln -match '^\s*"([^"]+)"\s*=\s*') {
            $models += $Matches[1].Trim()
        }
    }
    foreach ($ln in $lines) {
        if ($ln -match '=\s*"([^"]+)"\s*$') {
            $candidate = $Matches[1].Trim()
            if ($candidate -match '\b(PCL|UFR|PS|XPS|LaserJet|ImageRUNNER|Universal|PostScript|KX)\b') {
                $models += $candidate
            }
        }
    }
    $inStrings = $false
    foreach ($ln in $lines) {
        if ($ln -match '^\s*\[Strings\]\s*$') { $inStrings = $true; continue }
        if ($inStrings -and $ln -match '^\s*\[') { $inStrings = $false } # next section
        if ($inStrings -and $ln -match '=\s*"([^"]+)"') {
            $candidate = $Matches[1].Trim()
            if ($candidate -match '\b(PCL|UFR|PS|XPS|LaserJet|ImageRUNNER|Universal|PostScript|KX)\b') {
                $models += $candidate
            }
        }
    }

    $models = $models | Where-Object { $_ -and $_.Trim().Length -gt 0 } | Sort-Object -Unique
    return $models
}

function Install-DriverViaPrintUI {
    param(
        [Parameter(Mandatory)][string]$DriverModelName,
        [Parameter(Mandatory)][string]$InfPath,
        [string]$Architecture = "x64",
        [string]$Version = "Type 3 - User Mode"
    )
    $args = @("printui.dll,PrintUIEntry","/ia","/m",$DriverModelName,"/h",$Architecture,"/v",$Version,"/f",$InfPath)
    try {
        Write-Host "PrintUI install: $($args -join ' ')" -ForegroundColor Yellow
        $p = Start-Process -FilePath rundll32.exe -ArgumentList $args -NoNewWindow -PassThru -Wait
        return ($p.ExitCode -eq 0)
    } catch {
        Write-Warning "PrintUI install failed: $_"; return $false
    }
}

function Install-DriverViaPnPUtil {
    param([Parameter(Mandatory)][string]$InfPath)
    try {
        $p1 = Start-Process -FilePath pnputil.exe -ArgumentList @("/add-driver","`"$InfPath`"","/subdirs") -NoNewWindow -PassThru -Wait
        if ($p1.ExitCode -ne 0) { Write-Warning "pnputil add-driver exit $($p1.ExitCode)"; return $false }
        $p2 = Start-Process -FilePath pnputil.exe -ArgumentList @("/install","`"$InfPath`"") -NoNewWindow -PassThru -Wait
        if ($p2.ExitCode -ne 0) { Write-Host "pnputil install exit $($p2.ExitCode) (continuing)" -ForegroundColor DarkYellow }
        return $true
    } catch { Write-Warning "pnputil failed: $_"; return $false }
}

# Resolve user input (file or folder) to a definite printer INF path
function Resolve-InfPath {
    param(
        [string]$InputPath,
        [string]$Hint
    )

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $InputPath = Read-Host "Enter FULL path to the driver's INF file OR a FOLDER that contains it"
        if ([string]::IsNullOrWhiteSpace($InputPath)) { return $null }
    }
    if (-not (Test-Path -LiteralPath $InputPath)) { Write-Error "Path not found: $InputPath"; return $null }

    # If it's already an INF file, return it
    if ( (Test-Path -LiteralPath $InputPath -PathType Leaf) -and (([IO.Path]::GetExtension($InputPath)).ToLowerInvariant() -eq '.inf') ) {
        return (Resolve-Path -LiteralPath $InputPath).Path
    }

    # Treat as directory: search for printer INFs
    $root = if (Test-Path -LiteralPath $InputPath -PathType Container) { $InputPath } else { Split-Path -LiteralPath $InputPath -Parent }
    if (-not $root) { Write-Error "Could not determine a folder to search."; return $null }

    Write-Host "Scanning '$root' for printer INFs..." -ForegroundColor DarkGray
    $infFiles = @()
    try {
        $infFiles = Get-ChildItem -LiteralPath $root -Recurse -Filter *.inf -File -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Some subfolders were not accessible while scanning."
    }
    if (-not $infFiles -or $infFiles.Count -eq 0) { Write-Error "No .inf files found under $root"; return $null }

    # Keep likely printer-class INFs using case-insensitive checks for Class=Printer OR Printer class GUID
    $printerPattern = '(?im)^\s*Class\s*=\s*Printer\b|^\s*ClassGuid\s*=\s*\{4d36e979-e325-11ce-bfc1-08002be10318\}'
    $printerInfs = $infFiles | ForEach-Object {
        try {
            if (Select-String -Path $_.FullName -Pattern $printerPattern -Quiet -ErrorAction SilentlyContinue) { $_ }
        } catch { }
    } | Where-Object { $_ }

    # If still none, fall back to showing ALL INFs
    if (-not $printerInfs -or $printerInfs.Count -eq 0) {
        Write-Warning "No INFs matched printer heuristics. Showing all INFs so you can pick."
        $printerInfs = $infFiles
    }

    # Score results using optional hint (e.g., server driver name) and common tokens
    $printerInfs = $printerInfs | Select-Object FullName, Name, Length, @{
        Name='Score'; Expression={
            $n = $_.Name.ToLower()
            $s = 0
            if ($n -match 'hp')  { $s += 5 }
            if ($n -match 'pcl') { $s += 3 }
            if ($n -match 'ps')  { $s += 1 }
            if ($Hint) {
                $h = $Hint.ToLower()
                if ($n -match [regex]::Escape($h)) { $s += 6 }
            }
            $s
        }
    }

    $sorted = $printerInfs | Sort-Object Score, Length -Descending

    if ($sorted.Count -eq 1) {
        Write-Host "Found one candidate INF: $($sorted[0].FullName)" -ForegroundColor DarkGray
        return $sorted[0].FullName
    }

    Write-Host "Select the INF to use:" -ForegroundColor Cyan
    $max = [Math]::Min(20, $sorted.Count)
    for ($i=0; $i -lt $max; $i++) {
        Write-Host ("[{0}] {1}" -f ($i+1), $sorted[$i].FullName)
    }
    if ($sorted.Count -gt $max) { Write-Host ("...and {0} more not shown" -f ($sorted.Count - $max)) -ForegroundColor DarkGray }

    $sel = Read-Host "Enter number (1-$max) or press Enter to cancel"
    if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $max) {
        return $sorted[[int]$sel-1].FullName
    }
    return $null
}

function Prestage-DriverFromInf {
    param(
        [string]$ExpectedDriverName,
        [string]$InfPathOrFolder
    )
    $resolvedInf = Resolve-InfPath -InputPath $InfPathOrFolder -Hint $ExpectedDriverName
    if (-not $resolvedInf) { Write-Warning "No INF selected."; return $null }

    try { Unblock-File -LiteralPath $resolvedInf -ErrorAction SilentlyContinue } catch {}

    $models = Get-InfModelNames -InfPath $resolvedInf
    if (-not $models -or $models.Count -eq 0) { Write-Error "No printer model names discovered in the INF."; return $null }

    $chosen = $ExpectedDriverName
    if (-not $chosen -or -not ($models -contains $chosen)) {
        if ($chosen) { Write-Warning "Expected driver name '$chosen' not in INF. Choose one from the list." }
        Write-Host "Select a model from the INF:" -ForegroundColor Cyan
        $i=1; $models | ForEach-Object { Write-Host ("[{0}] {1}" -f $i, $_); $i++ }
        $sel = Read-Host "Enter number (or press Enter to cancel)"
        if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $models.Count) { $chosen = $models[[int]$sel-1] } else { return $null }
    }

    # 1) Preferred
    try {
        Write-Host "Staging driver via Add-PrinterDriver: '$chosen' from INF: $resolvedInf" -ForegroundColor Yellow
        Add-PrinterDriver -Name $chosen -InfPath $resolvedInf -ErrorAction Stop
        Write-Host "Driver staged: $chosen" -ForegroundColor Green
        return $chosen
    } catch {
        Write-Warning "Add-PrinterDriver failed: $($_.Exception.Message)"
    }

    # 2) Fallback: PrintUI
    if (Install-DriverViaPrintUI -DriverModelName $chosen -InfPath $resolvedInf) {
        $drv = Get-PrinterDriver -Name $chosen -ErrorAction SilentlyContinue
        if ($drv) { Write-Host "Driver staged via PrintUI: $chosen" -ForegroundColor Green; return $chosen }
        Write-Warning "PrintUI reported success but driver not visible yet. Restarting Spooler and rechecking..."
        try { Stop-Service Spooler -Force -ErrorAction Stop; Start-Sleep 2; Start-Service Spooler -ErrorAction Stop } catch {}
        $drv = Get-PrinterDriver -Name $chosen -ErrorAction SilentlyContinue
        if ($drv) { return $chosen }
    }

    # 3) Last resort: pnputil + register
    if (Install-DriverViaPnPUtil -InfPath $resolvedInf) {
        try {
            Add-PrinterDriver -Name $chosen -ErrorAction Stop
            Write-Host "Driver staged via pnputil + Add-PrinterDriver: $chosen" -ForegroundColor Green
            return $chosen
        } catch {
            $installed = Get-PrinterDriver | Where-Object { $_.Name -like "*$chosen*" }
            if ($installed) {
                Write-Host "Driver package installed. Found installed driver(s):" -ForegroundColor Yellow
                $installed | Select-Object Name, Manufacturer, InfPath | Format-Table -Auto
                return ($installed | Select-Object -First 1).Name
            }
        }
    }

    Write-Error "Failed to stage driver from INF. Confirm 64-bit INF, chosen model, and policy."
    return $null
}

# ---------- Driver removal helpers ----------
function Restart-SpoolerToReleaseHandles {
    Write-Host "Restarting Print Spooler to release driver files..." -ForegroundColor Yellow
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Host "Spooler restarted." -ForegroundColor Green
        return $true
    } catch { Write-Warning "Failed to restart Spooler: $_"; return $false }
}

function Get-OemInfForDriver {
    param([Parameter(Mandatory)][string]$DriverName)
    $out = & pnputil.exe /enum-drivers 2>$null
    if (-not $out) { return @() }
    $pubName = $null; $class = $null; $drv = $null; $found = @()
    foreach ($line in $out) {
        if ($line -match '^\s*$') {
            if ($class -like '*Printer*' -and $drv -eq $DriverName -and $pubName) { $found += $pubName }
            $pubName = $null; $class = $null; $drv = $null; continue
        }
        if ($line -match 'Published Name\s*:\s*(.+)$') { $pubName = $Matches[1].Trim(); continue }
        if ($line -match 'Class Name\s*:\s*(.+)$')     { $class   = $Matches[1].Trim(); continue }
        if ($line -match 'Driver Name\s*:\s*(.+)$')    { $drv     = $Matches[1].Trim(); continue }
    }
    if ($class -like '*Printer*' -and $drv -eq $DriverName -and $pubName) { $found += $pubName }
    return $found | Select-Object -Unique
}

function Uninstall-DriverIfUnused {
    param([Parameter(Mandatory)][string]$DriverName,[switch]$Force)
    $stillUsed = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.DriverName -eq $DriverName }
    if ($stillUsed) { Write-Host "Driver '$DriverName' is still used by other printers. Skipping driver removal." -ForegroundColor Yellow; return }
    $removedViaCmdlet = $false
    try {
        Write-Host "Removing printer driver '$DriverName'..." -ForegroundColor Yellow
        Remove-PrinterDriver -Name $DriverName -ErrorAction Stop
        Write-Host "Removed printer driver '$DriverName'." -ForegroundColor Green
        $removedViaCmdlet = $true
    } catch {
        Write-Warning "Remove-PrinterDriver failed: $($_.Exception.Message)"
        if (Restart-SpoolerToReleaseHandles) {
            try {
                Remove-PrinterDriver -Name $DriverName -ErrorAction Stop
                Write-Host "Removed printer driver '$DriverName' after spooler restart." -ForegroundColor Green
                $removedViaCmdlet = $true
            } catch { Write-Warning "Still could not remove via Remove-PrinterDriver. Will try pnputil…" }
        }
    }
    $oems = Get-OemInfForDriver -DriverName $DriverName
    foreach ($o in $oems) {
        $args = @('/delete-driver', $o)
        if ($Force) { $args += '/uninstall'; $args += '/force' }
        try {
            Write-Host "pnputil $($args -join ' ')" -ForegroundColor Yellow
            $p = Start-Process -FilePath pnputil.exe -ArgumentList $args -NoNewWindow -PassThru -Wait
            if ($p.ExitCode -eq 0) { Write-Host "Removed driver package $o." -ForegroundColor Green }
            else { Write-Warning "pnputil failed for $o (exit $($p.ExitCode)). It may already be gone or in use." }
        } catch { Write-Warning "Failed to run pnputil for $o. $_" }
    }
    if (-not $removedViaCmdlet -and -not $oems) {
        Write-Host "No removable package found and cmdlet removal failed; driver may still be in use or blocked by policy." -ForegroundColor Yellow
    }
}

# ---------- Add & Remove ----------
function Add-SharedPrinter {
    param([Parameter(Mandatory)][string]$Server,[Parameter(Mandatory)][string]$Share,[string]$ServerDriverName,[string]$QueueName)
    $finalDriver = Prestage-DriverFromInf -ExpectedDriverName $ServerDriverName
    if (-not $finalDriver) { Write-Error "Driver staging failed or was cancelled. Aborting add."; return }
    $conn = "\\$Server\$Share"
    try {
        Write-Host "Adding $conn ..." -ForegroundColor Yellow
        Add-Printer -ConnectionName $conn -ErrorAction Stop
        Write-Host "Added: $conn (Driver: $finalDriver)" -ForegroundColor Green
    } catch { Write-Error "Failed to add $conn. $_" }
}

function List-LocalPrinters {
    Write-Header "Printers on THIS PC"
    Get-Printer | Sort-Object Name | Format-Table -Auto Name, DriverName, PortName, Default
}

function Choose-LocalPrinter {
    $list = Get-Printer | Sort-Object Name
    if (-not $list) { Write-Warning "No local printers."; return $null }
    $i=1; $list | ForEach-Object {
        $mark = if ($_.Default) { "*" } else { " " }
        Write-Host ("[{0}] {1} {2}" -f $i, $_.Name, $mark); $i++
    }
    $idx = Read-Host "Enter number (or press Enter to cancel)"
    if ($idx -match '^\d+$' -and [int]$idx -ge 1 -and [int]$idx -le $list.Count) { return $list[[int]$idx-1].Name }
    return $null
}

function Remove-LocalPrinter {
    $name = Choose-LocalPrinter
    if (-not $name) { return $false }
    $driverName = (Get-Printer -Name $name -ErrorAction SilentlyContinue).DriverName
    try {
        Write-Host "Removing '$name'..." -ForegroundColor Yellow
        Remove-Printer -Name $name -ErrorAction Stop
        Write-Host "Removed." -ForegroundColor Green
    } catch { Write-Error "Failed to remove '$name'. $_"; return $false }

    $ans = Read-Host "Also uninstall this printer's driver if unused? (y/N)"
    if ($ans -match '^[Yy]' -and $driverName) {
        $force = (Read-Host "If locked, force driver package removal with pnputil? (y/N)") -match '^[Yy]'
        Uninstall-DriverIfUnused -DriverName $driverName -Force:$force
    }
    return $true
}

# ------------------- MAIN MENU (labeled loop so 'exit' truly exits) -------------------
Write-Host "Printer Client ready. Type 'help' for commands; 'exit' to quit." -ForegroundColor Cyan

:Main while ($true) {
    $cmd = (Read-Host "> (search/add/remove/server/help/exit)").Trim().ToLower()
    switch ($cmd) {
        'search' {
            if ([string]::IsNullOrWhiteSpace($Global:PrintServer) -or $Global:PrintServer -eq 'YOUR_PRINT_SERVER_NAME') {
                $Global:PrintServer = Read-Host "Enter print server (e.g., print.city.local)"
            }
            if ($Global:PrintServer -and $Global:PrintServer -ne 'YOUR_PRINT_SERVER_NAME') {
                [void](List-ServerQueues -Server $Global:PrintServer)
            } else {
                Write-Warning "Please set a valid print server name first (use 'server' command)."
            }
        }
        'add' {
            while ($true) {
                $pick = Choose-ServerAndQueue
                if ($pick) { Add-SharedPrinter -Server $pick.Server -Share $pick.Share -ServerDriverName $pick.ServerDriverName -QueueName $pick.QueueName }
                else { break }
                $again = Read-Host "Add another printer? (y/N)"
                if ($again -notmatch '^[Yy]') { break }
            }
        }
        'remove' {
            while ($true) {
                $ok = Remove-LocalPrinter
                if (-not $ok) { break }
                $again = Read-Host "Remove another printer? (y/N)"
                if ($again -notmatch '^[Yy]') { break }
            }
        }
        'server' {
            $Global:PrintServer = Read-Host "Set/Change print server (current: $Global:PrintServer)"
            if (-not $Global:PrintServer) { Write-Warning "Server cleared." }
        }
        'help' {
@"
Commands:
  search  - List queues on \\$Global:PrintServer (uses remote spooler, falls back to NET VIEW)
  add     - Prompt for DRIVER .INF or a FOLDER (auto-scan) then add one or more \\server\share
  remove  - Remove one or more local printer connections; optionally uninstall the driver if unused
  server  - Set/change the print server name
  exit    - Quit the tool (also accepts 'quit' or 'q')

Tips:
  - You can paste a folder like 'C:\Users\BLANK_\Downloads' and the script will scan for printer INFs.
  - If none are detected as 'printer', it will show ALL INFs so you can pick the right one.
  - Run as Administrator.

Note:
  - At start, the script copies: $src -> $dst
    When prompted during 'add', you can point to 'C:\Temp\hpcu270u.inf'.
"@ | Write-Host
        }
        'exit' { break Main }
        'quit' { break Main }
        'q'    { break Main }
        default { if ($cmd) { Write-Warning "Unknown command '$cmd'. Type 'help'." } }
    }
}

Write-Host "Goodbye." -ForegroundColor Cyan
