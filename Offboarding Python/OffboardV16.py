# OffboardV16.py
# Tkinter AD Offboarding GUI (no external GUI dependencies)
# - Windows only (calls PowerShell + RSAT ActiveDirectory)
# - Build:
#   python -m pip install --upgrade pip
#   python -m pip install pyinstaller
#   python -m PyInstaller --noconsole --onefile OffboardV16.py

__author__ = "Created by: Antonio C."  # hidden credit in source

import os, sys, json, ctypes, subprocess, re, tkinter as tk
from tkinter import ttk, messagebox
import tkinter.simpledialog as sd  # for email/UPN prompt

# -------- Elevation (Admin) --------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def relaunch_as_admin():
    exe = sys.executable
    params = " ".join([f'"{os.path.abspath(sys.argv[0])}"'] + [f'"{a}"' for a in sys.argv[1:]])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, None, 1)

if not is_admin():
    relaunch_as_admin()
    sys.exit(0)

# -------- PowerShell helpers (hidden windows) --------
POWERSHELL = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'),
                          'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')

# Hide popup console windows for subprocess on Windows
SI = None
CREATE_NO_WINDOW = 0x08000000  # subprocess.CREATE_NO_WINDOW
if os.name == "nt":
    SI = subprocess.STARTUPINFO()
    SI.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    SI.wShowWindow = 0  # SW_HIDE

def ps_run(ps_cmd, timeout=300):
    """
    Run the given PowerShell command string and return (ok, stdout_or_error).
    Uses -WindowStyle Hidden and CREATE_NO_WINDOW/STARTUPINFO to suppress popups.
    """
    try:
        proc = subprocess.run(
            [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-Command", ps_cmd],
            text=True,
            capture_output=True,
            timeout=timeout,
            startupinfo=SI,
            creationflags=CREATE_NO_WINDOW
        )
        if proc.returncode != 0:
            return False, (proc.stderr or proc.stdout).strip()
        return True, proc.stdout.strip()
    except subprocess.TimeoutExpired:
        return False, "PowerShell command timed out."
    except Exception as ex:
        return False, str(ex)

def ps_json(ps_body, timeout=300):
    """
    Convert PowerShell objects to JSON *inside* PowerShell to avoid parsing issues.
    """
    wrapped = "$__r = @(\n" + ps_body + "\n)\n$__r | ConvertTo-Json -Depth 8"
    ok, out = ps_run(wrapped, timeout=timeout)
    if not ok:
        return False, out
    if not out:
        return True, None
    try:
        data = json.loads(out)
        return True, data
    except json.JSONDecodeError:
        return False, f"JSON parse failed: {out[:4000]}"

# -------- AD queries & operations --------
BASE_USER_SELECT = "Name,SamAccountName,EmailAddress,Title,Department,Manager,DistinguishedName"

def ad_check_module():
    ps = r"""
        try {
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                throw "ActiveDirectory module not available. Install RSAT."
            }
            Import-Module ActiveDirectory -ErrorAction Stop
            "OK"
        } catch { $_.Exception.Message }
    """
    ok, out = ps_run(ps)
    return ok and out.strip()=="OK", out

def ad_search_by_email(email):
    email = email.replace("'", "''")
    ps = f"""
        Import-Module ActiveDirectory
        $u = Get-ADUser -Filter "EmailAddress -eq '{email}'" -Properties {BASE_USER_SELECT.replace(',',' , ')}
        if ($u) {{ $u | Select-Object {BASE_USER_SELECT} }}
    """
    return ps_json(ps)

def ad_search_by_name(name_part):
    esc = name_part
    esc = re.sub(r'\\', r'\\5c', esc)
    esc = re.sub(r'\*', r'\\2a', esc)
    esc = re.sub(r'\(', r'\\28', esc)
    esc = re.sub(r'\)', r'\\29', esc)
    ps = f"""
        Import-Module ActiveDirectory
        $ldap = "(|(name=*{esc}*)(displayName=*{esc}*)(sAMAccountName=*{esc}*))"
        Get-ADUser -LDAPFilter $ldap -Properties {BASE_USER_SELECT.replace(',',' , ')} |
        Sort-Object Name |
        Select-Object {BASE_USER_SELECT}
    """
    return ps_json(ps)

def ad_user_details(dn):
    ps = f"""
        Import-Module ActiveDirectory
        Get-ADUser -Identity '{dn}' -Properties {BASE_USER_SELECT.replace(',',' , ')} |
        Select-Object {BASE_USER_SELECT}
    """
    return ps_json(ps)

def ad_clear_properties(sam):
    ps = f"""
        Import-Module ActiveDirectory
        Set-ADUser -Identity '{sam}' -Clear Title,Department,Manager -ErrorAction Stop
        "Cleared Title/Department/Manager"
    """
    return ps_run(ps)

def ad_remove_non_default_groups(dn):
    ps = f"""
        Import-Module ActiveDirectory
        $u = Get-ADUser -Identity '{dn}' -Properties MemberOf
        $memberOf = @(); if ($u.MemberOf) {{ $memberOf = $u.MemberOf }}
        $toRemove = $memberOf | Where-Object {{ $_ -notmatch '^CN=Domain Users,' }}
        if (-not $toRemove -or $toRemove.Count -eq 0) {{ "No removable groups" }}
        else {{
            foreach($g in $toRemove) {{
                try {{ Remove-ADGroupMember -Identity $g -Members $u -Confirm:$false -ErrorAction Stop; "Removed: $g" }}
                catch {{ "ERROR removing $g : $($_.Exception.Message)" }}
            }}
            "Group cleanup done"
        }}
    """
    return ps_run(ps)

def ad_hide_from_address_lists(sam):
    ps = f"""
        $did=$false
        if (Get-Command Set-Mailbox -ErrorAction SilentlyContinue) {{
            try {{ Set-Mailbox -Identity '{sam}' -HiddenFromAddressListsEnabled $true -ErrorAction Stop; "Hidden via Set-Mailbox"; $did=$true }} catch {{}}
        }}
        if (-not $did -and (Get-Command Set-RemoteMailbox -ErrorAction SilentlyContinue)) {{
            try {{ Set-RemoteMailbox -Identity '{sam}' -HiddenFromAddressListsEnabled $true -ErrorAction Stop; "Hidden via Set-RemoteMailbox"; $did=$true }} catch {{}}
        }}
        if (-not $did) {{
            Import-Module ActiveDirectory
            Set-ADUser -Identity '{sam}' -Replace @{{msExchHideFromAddressLists=$true}} -ErrorAction Stop
            "Hidden by msExchHideFromAddressLists"
        }}
    """
    return ps_run(ps)

def ad_disable_account(sam):
    ps = f"""
        Import-Module ActiveDirectory
        Disable-ADAccount -Identity '{sam}' -ErrorAction Stop
        "Account disabled"
    """
    return ps_run(ps)

def ad_resolve_offboarding_ou():
    ps = r"""
        Import-Module ActiveDirectory
        try { $domainDN = (Get-ADDomain).DistinguishedName } catch { throw $_.Exception.Message }
        $PreferredParentName = "Azure AD Sync"; $OffboardingName = "Offboarding"
        $expectedDN = "OU=$OffboardingName,OU=$PreferredParentName,$domainDN"
        try { $ou = Get-ADOrganizationalUnit -Identity $expectedDN -ErrorAction Stop; $ou.DistinguishedName; return } catch {}
        $parentCandidates=@(); try { $parentCandidates = @(Get-ADOrganizationalUnit -LDAPFilter "(name=*$PreferredParentName*)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop) } catch {}
        $cands=@()
        foreach($p in @($parentCandidates)){
            try{ $c=@(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName)" -SearchBase $p.DistinguishedName -SearchScope OneLevel -ErrorAction Stop); if($c){$cands+=$c} }catch{}
            try{ $c=@(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName*)" -SearchBase $p.DistinguishedName -SearchScope Subtree -ErrorAction Stop); if($c){$cands+=$c} }catch{}
        }
        if (-not $cands -or @($cands).Length -eq 0) {
            try { $cands=@(Get-ADOrganizationalUnit -LDAPFilter "(name=$OffboardingName*)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop) } catch { throw $_.Exception.Message }
        }
        $cands | Sort-Object DistinguishedName -Unique | Select-Object -ExpandProperty DistinguishedName
    """
    return ps_json(ps)

def ad_move_user(dn, target_ou):
    ps = f"""
        Import-Module ActiveDirectory
        $u = Get-ADUser -Identity '{dn}' -Properties ProtectedFromAccidentalDeletion
        $changed=$false
        if ($u.ProtectedFromAccidentalDeletion) {{ Set-ADObject -Identity '{dn}' -ProtectedFromAccidentalDeletion:$false -ErrorAction Stop; $changed=$true }}
        try {{ Move-ADObject -Identity '{dn}' -TargetPath '{target_ou}' -ErrorAction Stop; "Moved to: {target_ou}" }}
        catch {{ throw $_.Exception.Message }}
        finally {{ if ($changed) {{ Set-ADObject -Identity '{dn}' -ProtectedFromAccidentalDeletion:$true -ErrorAction SilentlyContinue }} }}
    """
    return ps_run(ps)

# -------- Microsoft Entra group removal (Graph with auto-install, fallback to AzureAD) --------
def entra_remove_groups_except_all_users(upn_or_email):
    ps = f"""
        $ErrorActionPreference = "Stop"
        $target = "{upn_or_email}"

        function Ensure-TlsAndGallery {{
            try {{
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
                if ($repo.InstallationPolicy -ne 'Trusted') {{
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                }}
            }} catch {{ }}
        }}

        function Ensure-Module($name, $minVersion) {{
            try {{
                if (-not (Get-Module -ListAvailable -Name $name)) {{
                    Install-Module $name -Scope AllUsers -Force -AllowClobber -MinimumVersion $minVersion -ErrorAction Stop
                }}
                Import-Module $name -ErrorAction Stop
                return $true
            }} catch {{
                "Ensure-Module failed for ${{name}}: $($_.Exception.Message)"
                return $false
            }}
        }}

        function RemoveViaGraph {{
            try {{
                if (-not (Ensure-Module -name 'Microsoft.Graph' -minVersion '2.0.0')) {{ return $false }}
                if (-not (Get-MgContext)) {{
                    Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All" | Out-Null
                }}
                $u = Get-MgUser -UserId $target
                $uid = $u.Id
                $groups = Get-MgUserMemberOf -UserId $uid -All | Where-Object {{ $_.'@odata.type' -eq '#microsoft.graph.group' }}
                $toRemove = @()
                foreach($g in $groups) {{
                    $name = $g.AdditionalProperties["displayName"]
                    if ($name -and $name -ne "All Users") {{ $toRemove += [PSCustomObject]@{{ Id=$g.Id; Name=$name }} }}
                }}
                if (-not $toRemove -or $toRemove.Count -eq 0) {{ "No Entra groups to remove."; return $true }}

                foreach($g in $toRemove) {{
                    try {{
                        Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $uid -ErrorAction Stop
                        "Removed Entra group: {{" + "$($g.Name)" + "}}"
                    }} catch {{
                        "ERROR removing Entra group {{" + "$($g.Name)" + "}}: {{" + "$($_.Exception.Message)" + "}}"
                    }}
                }}
                "Entra group cleanup done (Graph)."
                return $true
            }} catch {{
                "Graph path failed: {{" + "$($_.Exception.Message)" + "}}"
                return $false
            }}
        }}

        function RemoveViaAzureAD {{
            try {{
                if (-not (Ensure-Module -name 'AzureAD' -minVersion '2.0.2.140')) {{ return $false }}
                try {{ $null = Get-AzureADCurrentSessionInfo -ErrorAction Stop }} catch {{ Connect-AzureAD | Out-Null }}
                $u = Get-AzureADUser -ObjectId $target
                if (-not $u) {{ $u = Get-AzureADUser -Filter "userPrincipalName eq '{upn_or_email}'" }}
                if (-not $u) {{ "User not found in AzureAD."; return $false }}
                $groups = Get-AzureADUserMembership -ObjectId $u.ObjectId | Where-Object {{$_.ObjectType -eq "Group"}}
                $toRemove = $groups | Where-Object {{$_.DisplayName -ne "All Users"}}
                if (-not $toRemove -or $toRemove.Count -eq 0) {{ "No Entra groups to remove."; return $true }}
                foreach($g in $toRemove) {{
                    try {{
                        Remove-AzureADGroupMember -ObjectId $g.ObjectId -MemberId $u.ObjectId -ErrorAction Stop
                        "Removed Entra group: {{" + "$($g.DisplayName)" + "}}"
                    }} catch {{
                        "ERROR removing Entra group {{" + "$($g.DisplayName)" + "}}: {{" + "$($_.Exception.Message)" + "}}"
                    }}
                }}
                "Entra group cleanup done (AzureAD)."
                return $true
            }} catch {{
                "AzureAD path failed: {{" + "$($_.Exception.Message)" + "}}"
                return $false
            }}
        }}

        Ensure-TlsAndGallery

        if (-not (RemoveViaGraph)) {{
            "Falling back to AzureAD..."
            if (-not (RemoveViaAzureAD)) {{
                "Could not remove Entra groups. Graph/AzureAD not available or insufficient permissions."
            }}
        }}
    """
    return ps_run(ps)

# -------- UI Helpers --------
def build_details(u):
    return "\n".join([
        f"Name: {u.get('Name','')}",
        f"Username (sAMAccountName): {u.get('SamAccountName','')}",
        f"Email: {u.get('EmailAddress','')}",
        f"Title: {u.get('Title','')}",
        f"Department: {u.get('Department','')}",
        f"Manager: {u.get('Manager','')}",
        f"OU Path: {u.get('DistinguishedName','')}",
    ])

def split_csv(s):
    return [p.strip() for p in s.split(",") if p and p.strip()]

# -------- Tkinter GUI --------
LARGE = ("Segoe UI", 10)
MONO11 = ("Consolas", 11)
MONO10 = ("Consolas", 10)

root = tk.Tk()
root.title("AD Offboarding - IT Tool")
root.geometry("1400x880")
root.minsize(1200, 780)

# Top search frame
top = ttk.Frame(root, padding=(8,8,8,4))
top.pack(side=tk.TOP, fill=tk.X)

ttk.Style().configure("TLabel", font=LARGE)
ttk.Style().configure("TButton", font=LARGE)
ttk.Style().configure("Treeview", rowheight=24, font=LARGE)
ttk.Style().configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

lblEmail = ttk.Label(top, text="Search by Email (exact, comma-separated):")
lblEmail.grid(row=0, column=0, sticky="w", padx=(0,8), pady=4)

emailVar = tk.StringVar()
txtEmail = ttk.Entry(top, textvariable=emailVar, width=60, font=LARGE)
txtEmail.grid(row=0, column=1, sticky="ew", padx=(0,8), pady=4)

btnEmail = ttk.Button(top, text="Find by Email(s)")
btnEmail.grid(row=0, column=2, sticky="e", pady=4)

lblName = ttk.Label(top, text="Search by Name (partial, comma-separated):")
lblName.grid(row=1, column=0, sticky="w", padx=(0,8), pady=4)

nameVar = tk.StringVar()
txtName = ttk.Entry(top, textvariable=nameVar, width=60, font=LARGE)
txtName.grid(row=1, column=1, sticky="ew", padx=(0,8), pady=4)

btnName = ttk.Button(top, text="Find by Name(s)")
btnName.grid(row=1, column=2, sticky="e", pady=4)

appendVar = tk.BooleanVar(value=True)
chkAppend = ttk.Checkbutton(top, text="Append results", variable=appendVar)
chkAppend.grid(row=0, column=3, padx=8, sticky="w")

btnClearResults = ttk.Button(top, text="Clear Results")
btnClearResults.grid(row=1, column=3, padx=8, sticky="w")

top.columnconfigure(1, weight=1)

# Main paned window: left (search results) | right (details + selected users)
paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))

left = ttk.Frame(paned)
right = ttk.Frame(paned)
paned.add(left, weight=2)
paned.add(right, weight=3)

# Left: results tree with checkbox column
columns = ("Selected","Name","SamAccountName","Email","DistinguishedName")
tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse")
for col, w, anchor in zip(columns, (90,220,160,280,520), ("center","w","w","w","w")):
    tree.heading(col, text=col)
    tree.column(col, width=w, anchor=anchor, stretch=True if col=="DistinguishedName" else False)
ys = ttk.Scrollbar(left, orient="vertical", command=tree.yview)
tree.configure(yscroll=ys.set)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
ys.pack(side=tk.RIGHT, fill=tk.Y)

# Right top: details + action buttons
right_top = ttk.LabelFrame(right, text="Selected User Details")
right_top.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=(0,0), pady=(0,8))
txtDetails = tk.Text(right_top, height=10, wrap="none", font=MONO11)
txtDetails.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

btns_frame = ttk.Frame(right)
btns_frame.pack(side=tk.TOP, fill=tk.X, pady=(0,8))

btnSelectAll = ttk.Button(btns_frame, text="Select All in Results")
btnSelectAll.pack(side=tk.LEFT, padx=(0,8))

btnClearSelection = ttk.Button(btns_frame, text="Clear Selection")
btnClearSelection.pack(side=tk.LEFT, padx=(0,8))

btnOffboardSelected = ttk.Button(btns_frame, text="Offboard Selected (0)")
btnOffboardSelected.pack(side=tk.LEFT, padx=(0,8))

# Right middle: DOCKED "Selected Users" panel (persistent bucket view)
dock_frame = ttk.LabelFrame(right, text="Selected Users (Persistent)")
dock_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0,0), pady=(0,8))

sel_cols = ("Name","SamAccountName","Email","DistinguishedName")
selected_tree_docked = ttk.Treeview(dock_frame, columns=sel_cols, show="headings", selectmode="extended")
for col, w in zip(sel_cols, (220,160,260,520)):
    selected_tree_docked.heading(col, text=col)
    selected_tree_docked.column(col, width=w, anchor="w", stretch=True if col=="DistinguishedName" else False)
ys2 = ttk.Scrollbar(dock_frame, orient="vertical", command=selected_tree_docked.yview)
selected_tree_docked.configure(yscroll=ys2.set)
selected_tree_docked.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8,0), pady=8)
ys2.pack(side=tk.LEFT, fill=tk.Y, pady=8)

dock_btns = ttk.Frame(dock_frame)
dock_btns.pack(side=tk.RIGHT, fill=tk.Y, padx=8, pady=8)
def dock_remove_selected():
    rows = selected_tree_docked.selection()
    removed = 0
    for iid in rows:
        dn = selected_tree_docked.item(iid, "values")[3]
        if dn in selection_bucket:
            del selection_bucket[dn]; removed += 1
        if dn in selected_dns:
            selected_dns.remove(dn)
    refresh_selected_views()
    update_selected_count()
    if removed:
        log(f"Removed {removed} from Selected bucket.", "OK")
def dock_clear_all():
    if not selection_bucket: return
    if messagebox.askyesno("Clear Selected Users", "Remove all users from the Selected bucket?") != True:
        return
    selection_bucket.clear()
    selected_dns.clear()
    refresh_selected_views()
    update_selected_count()
    log("Cleared Selected bucket.", "OK")
def dock_offboard_all():
    bucket_users = list(selection_bucket.values())
    if not bucket_users:
        messagebox.showinfo("Offboarding", "No users in Selected bucket.")
        return
    names = ", ".join([u.get("Name","") for u in bucket_users][:5])
    more = "" if len(bucket_users) <= 5 else f" (+{len(bucket_users)-5} more)"
    if messagebox.askyesno("Confirm", f'Proceed with offboarding for {len(bucket_users)} user(s): {names}{more}?') != True:
        return
    choice = offboarding_options_dialog(root)
    if choice is None:
        log("Offboarding cancelled.", "WARN"); return
    set_status("Running offboarding...")
    for u in bucket_users:
        offboard_single_user(u, choice)
    set_status("Offboarding run completed.")

ttk.Button(dock_btns, text="Remove Selected", command=dock_remove_selected).pack(fill=tk.X, pady=(0,8))
ttk.Button(dock_btns, text="Clear All", command=dock_clear_all).pack(fill=tk.X, pady=(0,8))
ttk.Button(dock_btns, text="Offboard All", command=dock_offboard_all).pack(fill=tk.X)

# Right bottom: log
right_log = ttk.LabelFrame(right, text="Log")
right_log.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
txtLog = tk.Text(right_log, height=10, wrap="word", font=MONO10, state="disabled")
txtLog.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

# Bottom status bar
status_bar = ttk.Frame(root)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)
status_left = ttk.Label(status_bar, text="Ready.", anchor="w")
status_left.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8,4), pady=(2,4))
credit_right = ttk.Label(status_bar, text="Created by: ME", anchor="e")
credit_right.pack(side=tk.RIGHT, padx=(4,8), pady=(2,4))

# -------- App State --------
results = []                 # list[dict] of users in the tree (order = rows)
selected_dns = set()         # distinguishedNames checked in results (UI assist)
iid_to_index = {}            # tree iid -> index in results
selection_bucket = {}        # dn -> user dict (PERSISTENT across clears)

def set_status(t):
    status_left.config(text=t)
    status_left.update_idletasks()

def log(line, level="INFO"):
    prefix = {"OK":"✔ ", "ERR":"✖ ", "WARN":"⚠ "}.get(level, "• ")
    txtLog.configure(state="normal")
    txtLog.insert("end", prefix + line + "\n")
    txtLog.configure(state="disabled")
    txtLog.see("end")

def refresh_details_from_selection():
    sel = tree.selection()
    if not sel:
        txtDetails.delete("1.0","end"); return
    idx = iid_to_index.get(sel[0])
    if idx is None or idx >= len(results):
        txtDetails.delete("1.0","end"); return
    u = results[idx]
    txtDetails.delete("1.0","end")
    txtDetails.insert("1.0", build_details(u))

def update_selected_count():
    btnOffboardSelected.config(text=f"Offboard Selected ({len(selection_bucket)})")

def user_key(u):
    return u.get("DistinguishedName") or (u.get("SamAccountName"), u.get("EmailAddress"))

def populate_tree(new_users, append=False):
    global results, iid_to_index
    if not append:
        results = []
        tree.delete(*tree.get_children())
        iid_to_index.clear()

    existing_keys = {user_key(u) for u in results}
    to_add = []
    for u in (new_users or []):
        k = user_key(u)
        if k and k not in existing_keys:
            existing_keys.add(k)
            to_add.append(u)

    start_idx = len(results)
    results.extend(to_add)

    for i, u in enumerate(to_add, start=start_idx):
        dn = u.get("DistinguishedName","")
        checked = "☑" if dn in selection_bucket else "☐"
        if dn in selection_bucket:
            selected_dns.add(dn)
        iid = tree.insert("", "end", values=(
            checked,
            u.get("Name",""),
            u.get("SamAccountName",""),
            u.get("EmailAddress",""),
            dn,
        ))
        iid_to_index[iid] = i

    refresh_details_from_selection()
    update_selected_count()
    refresh_selected_views()

def set_checkbox_for_row(iid, checked_bool):
    idx = iid_to_index.get(iid)
    if idx is None: return
    u = results[idx]
    dn = u.get("DistinguishedName","")
    if checked_bool:
        if dn:
            selected_dns.add(dn)
            selection_bucket[dn] = u
    else:
        if dn in selected_dns:
            selected_dns.remove(dn)
        if dn in selection_bucket:
            del selection_bucket[dn]
    vals = list(tree.item(iid, "values"))
    vals[0] = "☑" if checked_bool else "☐"
    tree.item(iid, values=vals)
    update_selected_count()
    refresh_selected_views()

def toggle_checkbox_click(event):
    region = tree.identify("region", event.x, event.y)
    if region != "cell": return
    col = tree.identify_column(event.x)  # '#1' is first
    if col != "#1": return
    row = tree.identify_row(event.y)
    if not row: return
    current = tree.item(row, "values")[0]
    set_checkbox_for_row(row, current == "☐")

def select_all_results():
    for iid in tree.get_children():
        set_checkbox_for_row(iid, True)

def clear_selection():
    for iid in tree.get_children():
        set_checkbox_for_row(iid, False)

def clear_results():
    global results, iid_to_index
    results = []
    iid_to_index = {}
    tree.delete(*tree.get_children())
    txtDetails.delete("1.0","end")
    # keep selection_bucket intact
    update_selected_count()
    refresh_selected_views()
    set_status("Results cleared. Selected Users bucket preserved.")

# ----- Offboarding options dialog -----
def offboarding_options_dialog(parent):
    top = tk.Toplevel(parent)
    top.title("Choose offboarding actions")
    top.geometry("520x360")
    top.transient(parent)
    top.grab_set()

    frm = ttk.Frame(top, padding=12)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text="Select the actions to perform for selected user(s):").pack(anchor="w", pady=(0,8))

    opts = {
        "clear_props": tk.BooleanVar(value=True),
        "remove_ad_groups": tk.BooleanVar(value=True),
        "remove_entra_groups": tk.BooleanVar(value=True),
        "hide_from_gal": tk.BooleanVar(value=True),
        "disable_acct": tk.BooleanVar(value=True),
        "move_to_ou": tk.BooleanVar(value=True),
    }

    rows = [
        ("Clear Title / Department / Manager", "clear_props"),
        ("Remove AD groups (except Domain Users)", "remove_ad_groups"),
        ("Remove Entra groups (except \"All Users\")", "remove_entra_groups"),
        ("Hide from address lists (GAL/People Picker)", "hide_from_gal"),
        ("Disable the account", "disable_acct"),
        ("Move to Offboarding OU", "move_to_ou"),
    ]

    for text, key in rows:
        ttk.Checkbutton(frm, text=text, variable=opts[key]).pack(anchor="w", pady=3)

    sel_all = tk.BooleanVar(value=True)
    def on_toggle_all():
        v = sel_all.get()
        for _, key in rows:
            opts[key].set(v)

    ttk.Separator(frm).pack(fill="x", pady=8)
    all_frame = ttk.Frame(frm); all_frame.pack(fill="x")
    ttk.Checkbutton(all_frame, text="Select all", variable=sel_all, command=on_toggle_all).pack(side="left")

    btns = ttk.Frame(frm); btns.pack(fill="x", pady=(16,0))
    sel = {"value": None}
    def ok_cmd():
        sel["value"] = {k: v.get() for k, v in opts.items()}
        top.destroy()
    def cancel_cmd():
        sel["value"] = None
        top.destroy()
    ttk.Button(btns, text="Run", command=ok_cmd).pack(side="right", padx=(8,0))
    ttk.Button(btns, text="Cancel", command=cancel_cmd).pack(side="right")

    parent.update_idletasks()
    top.wait_window()
    return sel["value"]

def refresh_selected_docked():
    selected_tree_docked.delete(*selected_tree_docked.get_children())
    for dn, u in selection_bucket.items():
        selected_tree_docked.insert("", "end", values=(
            u.get("Name",""),
            u.get("SamAccountName",""),
            u.get("EmailAddress",""),
            u.get("DistinguishedName",""),
        ))

def refresh_selected_views():
    refresh_selected_docked()

# -------- Event handlers --------
def run_bulk_email_search(emails, append):
    found = []; errors = 0
    for e in emails:
        ok, data = ad_search_by_email(e)
        if not ok:
            log(f"Email search failed for {e}: {data}", "ERR"); errors += 1; continue
        if not data:
            log(f"No user found for email: {e}", "WARN"); continue
        users = [data] if isinstance(data, dict) else data
        found.extend(users)
    populate_tree(found, append=append)
    set_status(f"Found {len(found)} user(s). {'Appended' if append else 'Replaced'} results.")
    if errors: set_status(f"Finished with {errors} error(s).")

def run_bulk_name_search(names, append):
    found = []; errors = 0
    for n in names:
        ok, data = ad_search_by_name(n)
        if not ok:
            log(f"Name search failed for {n}: {data}", "ERR"); errors += 1; continue
        if data:
            users = data if isinstance(data, list) else [data]
            found.extend(users)
        else:
            log(f"No users matched: {n}", "WARN")
    populate_tree(found, append=append)
    set_status(f"Found {len(found)} user(s). {'Appended' if append else 'Replaced'} results.")
    if errors: set_status(f"Finished with {errors} error(s).")

def on_find_email():
    raw = emailVar.get().strip()
    if not raw:
        messagebox.showinfo("Search", "Enter one or more email addresses (comma-separated)."); return
    emails = split_csv(raw)
    set_status("Searching by email(s)...")
    run_bulk_email_search(emails, append=appendVar.get())

def on_find_name():
    raw = nameVar.get().strip()
    if not raw:
        messagebox.showinfo("Search", "Enter one or more names (comma-separated)."); return
    names = split_csv(raw)
    set_status("Searching by name(s)...")
    run_bulk_name_search(names, append=appendVar.get())
    txtDetails.delete("1.0","end")

def offboard_single_user(u, choice):
    name = u.get("Name",""); sam = u.get("SamAccountName",""); dn = u.get("DistinguishedName",""); email = u.get("EmailAddress","") or ""
    log(f"— Offboarding: {name} ({sam})", "INFO")

    if choice.get("clear_props"):
        ok, out = ad_clear_properties(sam)
        log(out if ok else f"Clear properties failed: {out}", "OK" if ok else "ERR")

    if choice.get("remove_ad_groups"):
        ok, out = ad_remove_non_default_groups(dn)
        log(out if ok else f"AD group cleanup failed: {out}", "OK" if ok else "ERR")

    if choice.get("remove_entra_groups"):
        upn_for_entra = email or sd.askstring("Entra Group Cleanup", f"Enter UPN/email for {name} ({sam}):", initialvalue="")
        if upn_for_entra and upn_for_entra.strip():
            ok, out = entra_remove_groups_except_all_users(upn_for_entra.strip())
            log(out if ok else f"Entra group cleanup failed: {out}", "OK" if ok else "ERR")
        else:
            log("Entra group cleanup skipped (no email/UPN provided).", "WARN")

    if choice.get("hide_from_gal"):
        ok, out = ad_hide_from_address_lists(sam)
        log(out if ok else f"Hide from address lists failed: {out}", "OK" if ok else "ERR")

    if choice.get("disable_acct"):
        ok, out = ad_disable_account(sam)
        log(out if ok else f"Disable account failed: {out}", "OK" if ok else "ERR")

    if choice.get("move_to_ou"):
        ok, data = ad_resolve_offboarding_ou()
        if not ok or not data:
            log("Could not locate Offboarding OU(s)." + ("" if ok else f" {data}"), "ERR")
        else:
            ou_list = data if isinstance(data, list) else [data]
            selected_ou = None
            if len(ou_list) == 1:
                selected_ou = ou_list[0]
            else:
                picker = tk.Toplevel(root)
                picker.title(f"Select Offboarding OU for {name}")
                picker.geometry("900x300")
                lb = tk.Listbox(picker, selectmode="single")
                lb.pack(fill=tk.BOTH, expand=True)
                for ou in ou_list:
                    lb.insert("end", ou)
                sel_val = {"value": None}
                def ok_cmd():
                    try:
                        i = lb.curselection()[0]
                        sel_val["value"] = lb.get(i)
                    except Exception:
                        pass
                    picker.destroy()
                ttk.Button(picker, text="OK", command=ok_cmd).pack(pady=8)
                picker.transient(root); picker.grab_set(); picker.wait_window()
                selected_ou = sel_val["value"]
            if selected_ou:
                ok, out = ad_move_user(dn, selected_ou)
                log(out if ok else f"Move failed: {out}", "OK" if ok else "ERR")
            else:
                log("Move aborted by user.", "WARN")

def on_offboard_selected():
    bucket_users = list(selection_bucket.values())
    if not bucket_users:
        sel = tree.selection()
        if sel:
            idx = iid_to_index.get(sel[0])
            if idx is not None:
                u = results[idx]
                if u.get("DistinguishedName",""):
                    bucket_users = [u]
    if not bucket_users:
        messagebox.showinfo("Offboarding", "Select (check) users in the results or from the docked Selected Users panel."); return

    names = ", ".join([u.get("Name","") for u in bucket_users][:5])
    more = "" if len(bucket_users) <= 5 else f" (+{len(bucket_users)-5} more)"
    if messagebox.askyesno("Confirm", f'Proceed with offboarding for {len(bucket_users)} user(s): {names}{more}?') != True:
        return

    choice = offboarding_options_dialog(root)
    if choice is None:
        log("Offboarding cancelled.", "WARN"); return

    set_status("Running offboarding...")
    for u in bucket_users:
        offboard_single_user(u, choice)
        dn = u.get("DistinguishedName","")
        ok, data = ad_user_details(dn)
        if ok and data:
            txtDetails.delete("1.0","end"); txtDetails.insert("1.0", build_details(data))
    set_status("Offboarding run completed.")

# Wire events
btnEmail.configure(command=on_find_email)
btnName.configure(command=on_find_name)
btnOffboardSelected.configure(command=on_offboard_selected)
btnSelectAll.configure(command=select_all_results)
btnClearSelection.configure(command=clear_selection)
btnClearResults.configure(command=clear_results)

tree.bind("<<TreeviewSelect>>", lambda e: refresh_details_from_selection())
tree.bind("<Button-1>", toggle_checkbox_click)  # click to toggle checkbox in column 1
root.bind("<Return>", lambda e: on_find_name() if txtName.focus_get()==txtName else None)

# Module check log
ok_mod, mod_msg = ad_check_module()
def init_log(msg, level="OK"):
    prefix = {"OK":"✔ ", "ERR":"✖ ", "WARN":"⚠ "}.get(level, "• ")
    txtLog.configure(state="normal")
    txtLog.insert("end", prefix + msg + "\n")
    txtLog.configure(state="disabled")

if not ok_mod:
    init_log(f"ActiveDirectory module load warning: {mod_msg}", "ERR")
else:
    init_log("ActiveDirectory module loaded.", "OK")

set_status("Ready. Search by email or name (comma-separated). Check users to build your batch in the docked Selected Users panel, then Offboard Selected.")
root.mainloop()
