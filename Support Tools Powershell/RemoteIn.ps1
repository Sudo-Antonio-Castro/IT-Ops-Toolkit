<#
.SYNOPSIS
  Connect to a remote PC by name or by sAMAccountName, then stay connected with a persistent PSSession
  so you can run multiple PowerShell commands. Includes a local status GUI and a remote user notification.

.NOTES
  - Requires WinRM on the target (Enable-PSRemoting -Force).
  - AD module only needed if you use the username scan path.
#>

[CmdletBinding()]
param()

function Prompt-YN($msg) {
    while ($true) {
        $r = Read-Host "$msg [Y/N]"
        switch ($r.ToUpper()) {
            'Y' { return $true }
            'N' { return $false }
            default { Write-Host "Please enter Y or N." -ForegroundColor Yellow }
        }
    }
}

function Pick-FromList($items, $title = "Pick one") {
    for ($i = 0; $i -lt $items.Count; $i++) {
        Write-Host ("[{0}] {1}" -f $i, $items[$i])
    }
    while ($true) {
        $choice = Read-Host "$title (enter index)"
        if ($choice -match '^\d+$' -and [int]$choice -ge 0 -and [int]$choice -lt $items.Count) {
            return $items[[int]$choice]
        }
        Write-Host "Invalid selection." -ForegroundColor Yellow
    }
}

# --- Local status GUI (on your machine) ---
function Show-ConnectionStatus {
    param(
        [Parameter(Mandatory)] [string] $Target,
        [Parameter(Mandatory)] [ValidateSet('Starting','Testing','Connecting','Notifying','Connected','Failed')] [string] $State,
        [string] $Message = ''
    )
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing       | Out-Null

    if (-not $script:StatusForm) {
        $form               = New-Object System.Windows.Forms.Form
        $form.Text          = "Remote session status"
        $form.Size          = New-Object System.Drawing.Size(460,140)
        $form.StartPosition = "CenterScreen"
        $form.TopMost       = $true

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.AutoSize = $false
        $lbl.Size     = New-Object System.Drawing.Size(430,30)
        $lbl.Location = New-Object System.Drawing.Point(10,10)
        $lbl.Font     = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
        $form.Controls.Add($lbl)

        $pb = New-Object System.Windows.Forms.ProgressBar
        $pb.Style    = 'Marquee'
        $pb.Size     = New-Object System.Drawing.Size(430,20)
        $pb.Location = New-Object System.Drawing.Point(10,50)
        $form.Controls.Add($pb)

        $sub = New-Object System.Windows.Forms.Label
        $sub.AutoSize = $false
        $sub.Size     = New-Object System.Drawing.Size(430,20)
        $sub.Location = New-Object System.Drawing.Point(10,80)
        $form.Controls.Add($sub)

        $script:StatusForm  = $form
        $script:StatusLabel = $lbl
        $script:StatusPB    = $pb
        $script:StatusSub   = $sub
        $form.Show()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $script:StatusLabel.Text = "Target: $Target"
    switch ($State) {
        'Starting'   { $script:StatusPB.Style = 'Marquee'; $script:StatusSub.Text = "Starting... $Message" }
        'Testing'    { $script:StatusPB.Style = 'Marquee'; $script:StatusSub.Text = "Testing connectivity... $Message" }
        'Connecting' { $script:StatusPB.Style = 'Marquee'; $script:StatusSub.Text = "Creating remote session... $Message" }
        'Notifying'  { $script:StatusPB.Style = 'Marquee'; $script:StatusSub.Text = "Notifying remote user... $Message" }
        'Connected'  {
            $script:StatusPB.Style = 'Blocks'
            $script:StatusPB.Value = 100
            $script:StatusSub.Text = "Connected! Type commands in the console."
            $script:StatusSub.ForeColor = [System.Drawing.Color]::Green
        }
        'Failed'     {
            $script:StatusPB.Style = 'Blocks'
            $script:StatusPB.Value = 0
            $script:StatusSub.Text = "Connection failed. $Message"
            $script:StatusSub.ForeColor = [System.Drawing.Color]::Red
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Close-ConnectionStatus {
    if ($script:StatusForm) { $script:StatusForm.Close(); $script:StatusForm.Dispose(); $script:StatusForm = $null }
}

# --- Remote popup (user sees a terminal message) ---
function Notify-RemoteUser {
    param(
        [Parameter(Mandatory)] [string] $ComputerName,
        [Parameter(Mandatory)] [pscredential] $Credential,
        [string] $Message = "IT has remotely connected to your PC for support. If you did not expect this, please contact the Service Desk."
    )
    Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock {
        param($msg)
        try { cmd.exe /c "msg * $msg" } catch { Write-Host $msg }
    } -ArgumentList $Message -ErrorAction SilentlyContinue | Out-Null
}

# --- Helper for username normalization ---
function Normalize-UserName($raw) {
    if (-not $raw) { return $null }
    $s = $raw.Trim()
    if ($s -match '\\') { return ($s.Split('\')[-1]).ToLower() }
    elseif ($s -match '@') { return ($s.Split('@')[0]).ToLower() }
    else { return $s.ToLower() }
}

# --- MAIN ---
Clear-Host
Write-Host "Remote (Persistent) by PC or User helper" -ForegroundColor Cyan
Write-Host "----------------------------------------`n"

$pcInput   = Read-Host "Enter PC name (FQDN or NetBIOS). Leave blank if unknown"
$userInput = Read-Host "Enter sAMAccountName (leave blank if you provided PC name)"

if ([string]::IsNullOrWhiteSpace($pcInput) -and [string]::IsNullOrWhiteSpace($userInput)) {
    Write-Error "You must supply either a PC name or a sAMAccountName."
    exit 1
}

$cred = Get-Credential -Message "Enter credentials for remote operations"

# Resolve target
$target = $null

if (-not [string]::IsNullOrWhiteSpace($pcInput)) {
    $target = $pcInput.Trim()
} else {
    # Scan by sAMAccountName
    $sam = $userInput.Trim()
    Write-Host "`nLocating computers where '$sam' is logged on..." -ForegroundColor Cyan

    $canUseAD = $false
    try { Import-Module ActiveDirectory -ErrorAction Stop; $canUseAD = $true } catch {
        Write-Warning "ActiveDirectory module not available. Provide a PC name instead."
        exit 1
    }

    if (-not (Prompt-YN "Scan domain computers now? (may take time)")) { exit }

    try {
        $computers = Get-ADComputer -Filter 'Enabled -eq $true' -Properties DNSHostName | Select-Object -ExpandProperty DNSHostName
    } catch {
        Write-Error "Failed to enumerate domain computers: $($_.Exception.Message)"
        exit 1
    }

    $matches = New-Object System.Collections.Generic.List[string]
    $i = 0; $total = $computers.Count
    foreach ($c in $computers) {
        $i++; Write-Host ("[{0}/{1}] {2}" -f $i, $total, $c) -NoNewline
        $up = $false
        try { $up = Test-Connection -ComputerName $c -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { $up = $false }
        if (-not $up) { Write-Host " - offline/skipped"; continue }
        try {
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $c -ErrorAction Stop -OperationTimeoutSec 4
            if ((Normalize-UserName $cs.UserName) -eq $sam.ToLower()) {
                Write-Host " - MATCH" -ForegroundColor Green
                $matches.Add($c) | Out-Null
            } else {
                Write-Host " - no"
            }
        } catch { Write-Host " - query failed" }
    }

    if ($matches.Count -eq 0) { Write-Warning "No active logon found for '$sam'."; exit 0 }
    if ($matches.Count -eq 1) {
        if (Prompt-YN "Connect to $($matches[0]) now?") { $target = $matches[0] } else { Write-Host "Canceled."; exit }
    } else {
        $target = Pick-FromList -items $matches -title "Select computer to connect to"
    }
}

# Connect & hold session
Show-ConnectionStatus -Target $target -State Starting
Show-ConnectionStatus -Target $target -State Testing

$pingOk = $false
try { $pingOk = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue } catch {}
if (-not $pingOk) { Write-Warning "Host '$target' did not respond to ping. Continuing (ICMP may be blocked)." }

try {
    Test-WSMan -ComputerName $target -ErrorAction Stop | Out-Null
} catch {
    Show-ConnectionStatus -Target $target -State Failed -Message $_.Exception.Message
    Write-Error "WinRM test failed for $target : $($_.Exception.Message)"
    Close-ConnectionStatus
    exit 1
}

Show-ConnectionStatus -Target $target -State Notifying
Notify-RemoteUser -ComputerName $target -Credential $cred

# Create persistent session
Show-ConnectionStatus -Target $target -State Connecting
try {
    $s = New-PSSession -ComputerName $target -Credential $cred -ErrorAction Stop
} catch {
    Show-ConnectionStatus -Target $target -State Failed -Message $_.Exception.Message
    Write-Error "Failed to create PSSession to $target : $($_.Exception.Message)"
    Close-ConnectionStatus
    exit 1
}

Show-ConnectionStatus -Target $target -State Connected
Start-Sleep 1
Close-ConnectionStatus

Write-Host ""
Write-Host "Connected to $target. Type PowerShell commands to run **on the remote machine**." -ForegroundColor Green
Write-Host "Type 'exit' or 'quit' to disconnect." -ForegroundColor Yellow
Write-Host ""

# Simple interactive loop
while ($true) {
    $prompt = "[$target] PS> "
    $cmd = Read-Host -Prompt $prompt

    if ($null -eq $cmd) { continue }
    if ($cmd.Trim().ToLower() -in @('exit','quit')) { break }
    if ([string]::IsNullOrWhiteSpace($cmd)) { continue }

    try {
        # Run the exact command text remotely
        Invoke-Command -Session $s -ScriptBlock {
            param($commandText)
            try {
                Invoke-Expression $commandText
            } catch {
                Write-Error $_
            }
        } -ArgumentList $cmd
    } catch {
        Write-Host "Remote execution error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Cleanup
if ($s) {
    try { Remove-PSSession $s -ErrorAction SilentlyContinue } catch {}
}
Write-Host "Disconnected from $target." -ForegroundColor Cyan
