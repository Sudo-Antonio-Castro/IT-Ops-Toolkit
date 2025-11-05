# --- Ensure script runs as Administrator ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Elevating to Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-ExecutionPolicy Bypass -File "' + $MyInvocation.MyCommand.Definition + '"')
    exit
}
# --- End elevation check ---


Import-Module ActiveDirectory

# ---------- Helpers ----------

function Write-Success($msg) { Write-Host "✔ $msg" -ForegroundColor Green }
function Write-Fail($msg)    { Write-Host "✖ $msg" -ForegroundColor Red }
function Write-Info($msg)    { Write-Host "• $msg" -ForegroundColor Cyan }
function Confirm-Yes($prompt){
    (Read-Host "$prompt (Y/N)") -match '^[Yy]$'
}

function Test-ModuleAvailable([string]$Name){
    try { Get-Module -ListAvailable -Name $Name | Out-Null; return $true } catch { return $false }
}

function Enable-ADObjectInheritance([string]$DistinguishedName){
    try {
        $entry = New-Object DirectoryServices.DirectoryEntry("LDAP://$DistinguishedName")
        $sec = $entry.ObjectSecurity
        if ($sec.AreAccessRulesProtected) {
            $sec.SetAccessRuleProtection($false, $true)   # enable inheritance, preserve existing explicit ACEs
            $entry.ObjectSecurity = $sec
            $entry.CommitChanges()
            Write-Success "Enabled permission inheritance on object."
        } else {
            Write-Info "Inheritance already enabled."
        }
        return $true
    } catch {
        Write-Fail "Could not enable inheritance: $($_.Exception.Message)"
        return $false
    }
}

function Set-ProtectedFromAccidentalDeletion([string]$Identity, [bool]$Protected){
    try {
        Set-ADObject -Identity $Identity -ProtectedFromAccidentalDeletion:$Protected -ErrorAction Stop
        $state = if($Protected){"enabled"}else{"disabled"}
        Write-Info "Accidental deletion protection $state on $Identity"
        return $true
    } catch {
        Write-Fail "Failed to set accidental deletion protection on $Identity : $($_.Exception.Message)"
        return $false
    }
}

function Hide-FromAddressLists($Sam){
    # Prefer Exchange cmdlets if available; otherwise set the raw AD attribute.
    $didSomething = $false
    if (Test-ModuleAvailable "ExchangeOnlineManagement" -or (Get-Command Set-Mailbox -ErrorAction SilentlyContinue)) {
        try {
            Set-Mailbox -Identity $Sam -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            Write-Success "Hidden from address lists via Set-Mailbox."
            $didSomething = $true
        } catch {
            try {
                Set-RemoteMailbox -Identity $Sam -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                Write-Success "Hidden from address lists via Set-RemoteMailbox."
                $didSomething = $true
            } catch {
                Write-Info "Exchange cmdlets unavailable or failed; will try writing AD attribute directly."
            }
        }
    }
    if (-not $didSomething) {
        try {
            Set-ADUser -Identity $Sam -Replace @{msExchHideFromAddressLists=$true} -ErrorAction Stop
            Write-Success "Hidden from address lists by setting msExchHideFromAddressLists."
        } catch {
            Write-Fail "Failed to hide from address lists: $($_.Exception.Message)"
        }
    }
}

function Remove-NonDefaultGroups($User){
    # Domain Users is usually the primary group and may NOT appear in MemberOf.
    # We'll exclude any group whose CN equals "Domain Users".
    try {
        $memberOf = @()
        if ($User.MemberOf) { $memberOf = $User.MemberOf }
        $toRemove = $memberOf | Where-Object { $_ -notmatch '^CN=Domain Users,' }
        if (-not $toRemove -or $toRemove.Count -eq 0) {
            Write-Info "No removable group memberships found."
            return
        }
        foreach ($groupDN in $toRemove) {
            try {
                Remove-ADGroupMember -Identity $groupDN -Members $User -Confirm:$false -ErrorAction Stop
                Write-Info "Removed from group: $groupDN"
            } catch {
                Write-Fail "Failed to remove from $groupDN : $($_.Exception.Message)"
            }
        }
        Write-Success "Group memberships removed (except Domain Users)."
    } catch {
        Write-Fail "Group removal failed: $($_.Exception.Message)"
    }
}

function Show-UserSummary($user) {
    Write-Host "`nUser selected:"
    Write-Host "Name: $($user.Name)"
    Write-Host "Username: $($user.SamAccountName)"
    Write-Host "Email: $($user.EmailAddress)"
    Write-Host "Title: $($user.Title)"
    Write-Host "Department: $($user.Department)"
    Write-Host "Manager: $($user.Manager)"
    Write-Host "OU Path: $($user.DistinguishedName)"
}

# ---------- OU Resolver for Offboarding (smart + interactive, robust arrays) ----------

function Resolve-OffboardingOU {
    param(
        [string]$PreferredParentName = "Azure AD Sync",
        [string]$OffboardingName = "Offboarding"
    )

    try {
        $domainDN = (Get-ADDomain).DistinguishedName
    } catch {
        Write-Fail "Unable to determine domain DN: $($_.Exception.Message)"
        return $null
    }

    # 1) Try the most likely exact path first
    $expectedDN = "OU=$OffboardingName,OU=$PreferredParentName,$domainDN"
    try {
        $ou = Get-ADOrganizationalUnit -Identity $expectedDN -ErrorAction Stop
        return $ou.DistinguishedName
    } catch { }

    # 2) Locate parent OU(s) similar to PreferredParentName anywhere in the domain
    $parentCandidates = @()
    try {
        $parentCandidates = @(Get-ADOrganizationalUnit -LDAPFilter "(name=*$PreferredParentName*)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop)
    } catch {
        Write-Info "Could not enumerate parent OU '$PreferredParentName' candidates: $($_.Exception.Message)"
    }

    $candidates = @()

    foreach ($p in @($parentCandidates)) {
        try {
            # Prefer one-level child called exactly Offboarding
            $child = @(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName)" -SearchBase $p.DistinguishedName -SearchScope OneLevel -ErrorAction Stop)
            if ($child) { $candidates += @($child) }
        } catch { }

        try {
            # Else, any Offboard* anywhere under that parent
            $childAny = @(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName*)" -SearchBase $p.DistinguishedName -SearchScope Subtree -ErrorAction Stop)
            if ($childAny) { $candidates += @($childAny) }
        } catch { }
    }

    # 3) If still nothing, search the entire domain for Offboard*
    if (-not $candidates -or @($candidates).Length -eq 0) {
        try {
            $candidates = @(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName*)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop)
        } catch {
            Write-Fail "Failed searching for '$OffboardingName*' OUs in the domain: $($_.Exception.Message)"
            return $null
        }
    }

    # Deduplicate and force a real array
    $candidates = @($candidates | Sort-Object DistinguishedName -Unique)

    if ($candidates.Length -eq 0) {
        Write-Fail "Offboarding OU not found anywhere under $domainDN."
        return $null
    }

    if ($candidates.Length -eq 1) {
        return $candidates[0].DistinguishedName
    }

    Write-Host "`nMultiple 'Offboard*' OUs found:"
    for ($i = 0; $i -lt $candidates.Length; $i++) {
        Write-Host ("{0}. {1}" -f ($i+1), $candidates[$i].DistinguishedName)
    }

    $sel = Read-Host "Enter the number to use, or paste an OU distinguishedName, or press Enter to cancel"

    if ($sel -match '^\d+$') {
        $idx = [int]$sel
        if ($idx -ge 1 -and $idx -le $candidates.Length) {
            return $candidates[$idx - 1].DistinguishedName
        } else {
            Write-Fail "Selection out of range."
            return $null
        }
    }

    if ($sel -and $sel -match '^OU=.*') {
        return $sel
    }

    Write-Fail "No selection made."
    return $null
}

# ---------- Main workflow functions ----------

function RunCleanupPrompts($user) {
    Show-UserSummary $user

    # Pre-checks for common permission issues
    $u = Get-ADUser $user -Properties AdminCount, ntSecurityDescriptor, ProtectedFromAccidentalDeletion, MemberOf
    if ($u.AdminCount -eq 1) {
        Write-Info "This account is admin-protected (AdminCount=1). Delegated permissions may NOT apply."
        if (Confirm-Yes "Re-enable inheritance on the user object now") {
            Enable-ADObjectInheritance -DistinguishedName $u.DistinguishedName | Out-Null
        }
    } else {
        # Even for non-admin, inheritance may be off
        try {
            $entry = New-Object DirectoryServices.DirectoryEntry("LDAP://$($u.DistinguishedName)")
            if ($entry.ObjectSecurity.AreAccessRulesProtected) {
                if (Confirm-Yes "Inheritance is disabled on this user. Enable it now") {
                    Enable-ADObjectInheritance -DistinguishedName $u.DistinguishedName | Out-Null
                }
            }
        } catch { }
    }

    if (Confirm-Yes "`nDo you want to clear Job Title, Department, and Manager?") {
        try {
            Set-ADUser -Identity $u.SamAccountName -Clear Title,Department,Manager -ErrorAction Stop
            Write-Success "Properties cleared."
        } catch {
            Write-Fail "Failed to clear properties: $($_.Exception.Message)"
        }
    }

    if (Confirm-Yes "`nDo you wish to remove all 'Member of' groups except 'Domain Users'?") {
        Remove-NonDefaultGroups -User $u
    }

    if (Confirm-Yes "`nDo you want to hide user from address lists (GAL/SharePoint People Picker)?") {
        Hide-FromAddressLists -Sam $u.SamAccountName
    }

    if (Confirm-Yes "`nWould you like to disable the account?") {
        try {
            Disable-ADAccount -Identity $u.SamAccountName -ErrorAction Stop
            Write-Success "Account disabled."
        } catch {
            Write-Fail "Failed to disable account: $($_.Exception.Message)"
        }
    }

    # ---- MOVE: to Offboarding OU; toggle accidental deletion ONLY on the USER ----
    if (Confirm-Yes "`nWould you like to move this user to the Offboarding OU?") {
        $offboardingOU = Resolve-OffboardingOU
        if (-not $offboardingOU) {
            Write-Fail "Move aborted: Offboarding OU not found."
        } else {
            $userProtChanged = $false

            # Ensure the user object isn't protected from accidental deletion (needed for move delete on source)
            try {
                $uRef = Get-ADUser -Identity $u.DistinguishedName -Properties ProtectedFromAccidentalDeletion
                if ($uRef.ProtectedFromAccidentalDeletion) {
                    if (Set-ProtectedFromAccidentalDeletion -Identity $uRef.DistinguishedName -Protected:$false) {
                        Write-Info "Accidental deletion protection disabled on user for the move."
                        $userProtChanged = $true
                    }
                }
            } catch {
                Write-Info "Couldn't read user's protection state (continuing): $($_.Exception.Message)"
            }

            try {
                Move-ADObject -Identity $u.DistinguishedName -TargetPath $offboardingOU -ErrorAction Stop
                Write-Success "User moved to Offboarding OU."
            } catch {
                Write-Fail "Failed to move user: $($_.Exception.Message)"
            } finally {
                if ($userProtChanged) {
                    try {
                        Set-ProtectedFromAccidentalDeletion -Identity $u.DistinguishedName -Protected:$true | Out-Null
                        Write-Info "Accidental deletion protection re-enabled on user."
                    } catch {
                        Write-Info "Couldn't re-enable user protection (manual follow-up may be needed): $($_.Exception.Message)"
                    }
                }
            }
        }
    }

    if (Confirm-Yes "`nIs there another user you wish to search for?") {
        Search-ADUserWorkflow
    } else {
        Read-Host "`nPress Enter to close" | Out-Null
    }
}

# ---------- Search workflow (email exact, name partial + retry loop) ----------

function Search-ADUserWorkflow {
    $email = Read-Host "Enter the email address to search for"
    # Exact email match on EmailAddress
    $emailEscaped = $email -replace "'", "''"
    $user = Get-ADUser -Filter "EmailAddress -eq '$emailEscaped'" -Properties EmailAddress, Name, SamAccountName, Title, Department, Manager, DistinguishedName, MemberOf

    if ($user) {
        RunCleanupPrompts $user
        return
    }

    Write-Host "`nNo user found with that email."

    if (-not (Confirm-Yes "Would you like to search by name instead")) {
        Read-Host "`nPress Enter to close" | Out-Null
        return
    }

    # Name search loop (partial match on several attributes)
    while ($true) {
        $nameInput = Read-Host "Enter the name or partial name (e.g., 'james', 'john j', 'jjames')"
        if (-not $nameInput) {
            Write-Info "No input provided. Exiting."
            Read-Host "`nPress Enter to close" | Out-Null
            return
        }

        # Build LDAP OR filter for partial matches; escape special chars: \ * ( )
        $escaped = $nameInput -replace '\\','\5c' -replace '\*','\2a' -replace '\(','\28' -replace '\)','\29'
        $ldapFilter = "(|(name=*$escaped*)(displayName=*$escaped*)(sAMAccountName=*$escaped*))"

        try {
            $matches = Get-ADUser -LDAPFilter $ldapFilter -Properties EmailAddress, Name, SamAccountName, Title, Department, Manager, DistinguishedName, MemberOf | Sort-Object Name
            $arr = @($matches)
        } catch {
            Write-Fail "Lookup failed: $($_.Exception.Message)"
            if (Confirm-Yes "Try another name") { continue } else { Read-Host "`nPress Enter to close" | Out-Null; return }
        }

        if ($arr.Length -eq 0) {
            Write-Host "`nNo users found with that name."
            if (Confirm-Yes "Try another name") { continue } else { Read-Host "`nPress Enter to close" | Out-Null; return }
        }

        if ($arr.Length -eq 1) {
            RunCleanupPrompts $arr[0]
            return
        }

        # Multiple results: present a numbered list and allow refine
        Write-Host "`nMultiple users found:"
        for ($i=0; $i -lt $arr.Length; $i++) {
            $item = $arr[$i]
            Write-Host ("{0}. {1}  -  {2}  -  {3}" -f ($i+1), $item.Name, $item.SamAccountName, $item.EmailAddress)
        }

        $selection = Read-Host "`nEnter the number to select, or press Enter to refine your search"
        if ([string]::IsNullOrWhiteSpace($selection)) {
            continue
        }

        if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $arr.Length) {
            $selectedUser = $arr[[int]$selection - 1]
            RunCleanupPrompts $selectedUser
            return
        } else {
            Write-Fail "Invalid selection."
            if (Confirm-Yes "Try another name") { continue } else { Read-Host "`nPress Enter to close" | Out-Null; return }
        }
    }
}

# ---------- Start ----------
Search-ADUserWorkflow
