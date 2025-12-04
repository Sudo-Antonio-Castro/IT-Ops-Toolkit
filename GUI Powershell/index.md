---
layout: default
title: GUI PowerShell Toolkit
---

# IT Ops Toolkit

## PowerShell GUI 

```powershell
# ADUserLookup.GUI.ps1
# GUI wrapper for AD user lookups (email or name). Requires RSAT/ActiveDirectory module.

# --- UI Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Prereqs ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "ActiveDirectory module not found. Install RSAT / AD module and try again.",
        "Missing Module",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

# --- Helpers ---
function Show-Status($msg) { $StatusLabel.Text = $msg }
function Clear-Results {
    $List.Items.Clear()
    Show-Status "Ready"
}

function Add-Row($u){
    $name = $u.Name
    $sam  = $u.SamAccountName
    $mail = $u.mail
    $titl = $u.Title
    $dept = $u.Department
    $ou   = $u.DistinguishedName
    $item = New-Object System.Windows.Forms.ListViewItem($name)
    $item.SubItems.Add($sam)  | Out-Null
    $item.SubItems.Add($mail) | Out-Null
    $item.SubItems.Add($titl) | Out-Null
    $item.SubItems.Add($dept) | Out-Null
    $item.SubItems.Add($ou)   | Out-Null
    $List.Items.Add($item)    | Out-Null
}

function Search-ByEmail {
    Clear-Results
    $email = $EmailBox.Text.Trim()
    if (-not $email) { Show-Status "Enter an email."; return }
    Show-Status "Searching by email..."
    try {
        # Using 'mail' attribute is most reliable across ADs
        $u = Get-ADUser -LDAPFilter "(mail=$email)" -Properties mail,Name,SamAccountName,Title,Department,DistinguishedName
        if ($u) {
            Add-Row $u
            Show-Status "Found 1 result."
        } else {
            Show-Status "No user found with that email."
        }
    } catch {
        Show-Status "Error: $($_.Exception.Message)"
    }
}

function Search-ByName {
    Clear-Results
    $name = $NameBox.Text.Trim()
    if (-not $name) { Show-Status "Enter a name (partial allowed)."; return }
    Show-Status "Searching by name..."
    try {
        # Name is display name; wildcards allowed
        $users = Get-ADUser -Filter "Name -like '*$name*'" -Properties mail,Name,SamAccountName,Title,Department,DistinguishedName | Sort-Object Name
        if ($users) {
            foreach ($u in $users) { Add-Row $u }
            Show-Status "Found $($users.Count) result(s)."
        } else {
            Show-Status "No users found with that name."
        }
    } catch {
        Show-Status "Error: $($_.Exception.Message)"
    }
}

# --- UI ---
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "AD User Lookup"
$form.StartPosition = "CenterScreen"
$form.Size          = New-Object System.Drawing.Size(980, 560)
$form.MaximizeBox   = $true

# Email search row
$EmailLbl           = New-Object System.Windows.Forms.Label
$EmailLbl.Text      = "Email:"
$EmailLbl.Location  = New-Object System.Drawing.Point(16,16)
$EmailLbl.AutoSize  = $true

$EmailBox           = New-Object System.Windows.Forms.TextBox
$EmailBox.Location  = New-Object System.Drawing.Point(70, 12)
$EmailBox.Width     = 420

$EmailBtn           = New-Object System.Windows.Forms.Button
$EmailBtn.Text      = "Search by Email"
$EmailBtn.Location  = New-Object System.Drawing.Point(500, 10)
$EmailBtn.Add_Click({ Search-ByEmail })

# Name search row
$NameLbl            = New-Object System.Windows.Forms.Label
$NameLbl.Text       = "Name:"
$NameLbl.Location   = New-Object System.Drawing.Point(16,50)
$NameLbl.AutoSize   = $true

$NameBox            = New-Object System.Windows.Forms.TextBox
$NameBox.Location   = New-Object System.Drawing.Point(70, 46)
$NameBox.Width      = 420

$NameBtn            = New-Object System.Windows.Forms.Button
$NameBtn.Text       = "Search by Name"
$NameBtn.Location   = New-Object System.Drawing.Point(500, 44)
$NameBtn.Add_Click({ Search-ByName })

$ClearBtn           = New-Object System.Windows.Forms.Button
$ClearBtn.Text      = "Clear"
$ClearBtn.Location  = New-Object System.Drawing.Point(640, 44)
$ClearBtn.Add_Click({
    Clear-Results
    $EmailBox.Clear()
    $NameBox.Clear()
})

$CloseBtn           = New-Object System.Windows.Forms.Button
$CloseBtn.Text      = "Close"
$CloseBtn.Location  = New-Object System.Drawing.Point(720, 44)
$CloseBtn.Add_Click({ $form.Close() })

# Results list
$List               = New-Object System.Windows.Forms.ListView
$List.Location      = New-Object System.Drawing.Point(16, 90)
$List.Size          = New-Object System.Drawing.Size(940, 380)
$List.View          = 'Details'
$List.FullRowSelect = $true
$List.GridLines     = $true

# Columns
$List.Columns.Add("Name",           180) | Out-Null
$List.Columns.Add("Username",       120) | Out-Null
$List.Columns.Add("Email",          180) | Out-Null
$List.Columns.Add("Title",          140) | Out-Null
$List.Columns.Add("Department",     140) | Out-Null
$List.Columns.Add("OU Path",        160) | Out-Null

# Status bar
$StatusLabel            = New-Object System.Windows.Forms.Label
$StatusLabel.AutoSize   = $false
$StatusLabel.Text       = "Ready"
$StatusLabel.TextAlign  = "MiddleLeft"
$StatusLabel.BorderStyle= 'Fixed3D'
$StatusLabel.Location   = New-Object System.Drawing.Point(16, 480)
$StatusLabel.Size       = New-Object System.Drawing.Size(940, 28)

# Add controls
$form.Controls.AddRange(@(
    $EmailLbl, $EmailBox, $EmailBtn,
    $NameLbl,  $NameBox,  $NameBtn,
    $ClearBtn, $CloseBtn,
    $List, $StatusLabel
))

# Keyboard shortcuts
$EmailBox.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { Search-ByEmail; $_.SuppressKeyPress = $true }
})
$NameBox.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { Search-ByName; $_.SuppressKeyPress = $true }
})

# Show
[void]$form.ShowDialog()
```
