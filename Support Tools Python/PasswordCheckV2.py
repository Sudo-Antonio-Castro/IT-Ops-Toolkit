#!/usr/bin/env python3
# AD Password Expiry (Fast, Patched)
# - Caches MaxPasswordAge once
# - Uses EncodedCommand; avoids f-string braces conflicts with PowerShell
#
# Build to EXE (Windows):
#   py -m pip install pyinstaller
#   py -m PyInstaller --noconsole --onefile --name "AD Password Expiry" ADPasswordExpiryGUI_fast_fixed.py

import base64
import json
import platform
import subprocess
from datetime import datetime
from tkinter import Tk, StringVar, N, S, E, W, messagebox
from tkinter.ttk import Frame, Label, Entry, Button, Style

APP_TITLE = "AD Password Expiry"
DATE_FMT = "%Y-%m-%d %H:%M:%S"

def is_windows():
    return platform.system().lower().startswith("win")

def ps_encoded(cmd: str) -> list:
    # PowerShell -EncodedCommand expects UTF-16LE
    data = cmd.encode("utf-16le")
    b64 = base64.b64encode(data).decode("ascii")
    return ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", b64]

def run_ps_command_get_json(cmd: str):
    if not is_windows():
        return {"Success": False, "Error": "Windows required (PowerShell + AD module)."}
    # Hide window on Windows
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    try:
        proc = subprocess.run(
            ps_encoded(cmd),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except Exception as ex:
        return {"Success": False, "Error": f"PowerShell launch failed: {ex}"}
    if proc.returncode != 0:
        out = proc.stdout.strip()
        try:
            return json.loads(out) if out else {"Success": False, "Error": proc.stderr.strip() or "Unknown PowerShell error."}
        except Exception:
            return {"Success": False, "Error": proc.stderr.strip() or out or "Unknown PowerShell error."}
    out = proc.stdout.strip()
    try:
        return json.loads(out)
    except Exception as ex:
        return {"Success": False, "Error": f"Failed to parse JSON. Raw: {out[:400]} ..."}

def get_domain_max_age_seconds():
    # Cache MaxPasswordAge (in seconds) once at startup
    ps = r"""
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        $o = @{ Success=$false; Error="ActiveDirectory module not found. Install RSAT." } | ConvertTo-Json -Compress
        [Console]::Out.Write($o); exit 1
    }
    try {
        $max = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
        $secs = [int64]$max.TotalSeconds
    } catch {
        $secs = 0
    }
    @{ Success=$true; MaxAgeSeconds=$secs } | ConvertTo-Json -Compress
    """
    res = run_ps_command_get_json(ps)
    if not res.get("Success"):
        return 0, res.get("Error", "Unknown error.")
    return int(res.get("MaxAgeSeconds", 0)), None

def check_user(sam: str, max_age_seconds: int):
    # PowerShell payload as a raw (non-f) string so Python doesn't interpret {}
    ps = r"""
    param([string]$Sam,[Int64]$MaxAgeSecs)
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch {
        $o = @{ Success=$false; Error="ActiveDirectory module not found. Install RSAT." } | ConvertTo-Json -Compress
        [Console]::Out.Write($o); exit 1
    }
    try {
        $u = Get-ADUser -Identity $Sam -Properties DisplayName,pwdLastSet,msDS-UserPasswordExpiryTimeComputed,Enabled,PasswordNeverExpires,UserPrincipalName
    } catch {
        $o = @{ Success=$false; Error="User '$Sam' not found or lookup failed: $($_.Exception.Message)" } | ConvertTo-Json -Compress
        [Console]::Out.Write($o); exit 1
    }
    if (-not $u) {
        $o = @{ Success=$false; Error="User '$Sam' not found." } | ConvertTo-Json -Compress
        [Console]::Out.Write($o); exit 1
    }

    $pls = $null
    if ($u.pwdLastSet) { $pls = [datetime]::FromFileTime($u.pwdLastSet) }

    $expiry = $null
    if ($u.msDS_UserPasswordExpiryTimeComputed) {
        $expiry = [datetime]::FromFileTime($u.msDS_UserPasswordExpiryTimeComputed)
    } elseif ($pls -and $MaxAgeSecs -gt 0 -and -not $u.PasswordNeverExpires) {
        $expiry = $pls + [TimeSpan]::FromSeconds($MaxAgeSecs)
    }

    $now = Get-Date
    $expired = $false
    if ($expiry) { $expired = ($expiry -lt $now) }

    $o = [ordered]@{
        Success=$true
        SamAccountName=$u.SamAccountName
        DisplayName=$u.DisplayName
        UserPrincipalName=$u.UserPrincipalName
        PasswordLastSet = $(if ($pls) { $pls.ToString("s") } else { $null })
        PasswordExpires  = $(if ($expiry) { $expiry.ToString("s") } else { $null })
        Expired = $expired
        Enabled = [bool]$u.Enabled
        PasswordNeverExpires = [bool]$u.PasswordNeverExpires
    }
    $o | ConvertTo-Json -Compress
    """
    # Build the invocation without f-strings to avoid brace conflicts
    sam_escaped = sam.replace("'", "''")
    param_prefix = "$argsSam='" + sam_escaped + "'; $argsMax=" + str(int(max_age_seconds)) + "; & "
    param_body = "{ param([string]$Sam,[Int64]$MaxAgeSecs) " + ps + " } -Sam $argsSam -MaxAgeSecs $argsMax"
    param_invoke = param_prefix + param_body
    return run_ps_command_get_json(param_invoke)

class App:
    def __init__(self, master: Tk):
        self.master = master
        master.title(APP_TITLE)
        master.geometry("520x260")
        master.minsize(500, 250)

        s = Style()
        try:
            s.theme_use("clam")
        except Exception:
            pass

        self.max_age_seconds = 0
        self.sam_var = StringVar()

        main = Frame(master, padding=12)
        main.grid(column=0, row=0, sticky=(N,S,E,W))
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        # Input controls
        Label(main, text="sAMAccountName").grid(column=0, row=0, sticky=E, padx=(0,8), pady=(0,6))
        self.sam_entry = Entry(main, textvariable=self.sam_var, width=28)
        self.sam_entry.grid(column=1, row=0, sticky=W, pady=(0,6))
        self.sam_entry.focus_set()

        self.check_btn = Button(main, text="Check", command=self.do_check)
        self.check_btn.grid(column=2, row=0, sticky=W, padx=(8,0), pady=(0,6))

        # Result labels
        self.v_name = StringVar(); self.v_upn = StringVar()
        self.v_pls = StringVar(); self.v_pexp = StringVar()
        self.v_expired = StringVar(); self.v_enabled = StringVar(); self.v_pne = StringVar()
        Label(main, text="Display Name:").grid(column=0, row=1, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_name).grid(column=1, row=1, columnspan=2, sticky=W)
        Label(main, text="UPN:").grid(column=0, row=2, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_upn).grid(column=1, row=2, columnspan=2, sticky=W)
        Label(main, text="Password Last Set:").grid(column=0, row=3, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_pls).grid(column=1, row=3, columnspan=2, sticky=W)
        Label(main, text="Password Expires:").grid(column=0, row=4, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_pexp).grid(column=1, row=4, columnspan=2, sticky=W)
        Label(main, text="Expired?").grid(column=0, row=5, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_expired).grid(column=1, row=5, sticky=W)
        Label(main, text="Enabled?").grid(column=0, row=6, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_enabled).grid(column=1, row=6, sticky=W)
        Label(main, text="Pwd Never Expires?").grid(column=0, row=7, sticky=E, padx=(0,8), pady=2); Label(main, textvariable=self.v_pne).grid(column=1, row=7, sticky=W)

        self.status = StringVar(value="Loading domain password policy…")
        Label(main, textvariable=self.status).grid(column=0, row=8, columnspan=3, sticky=W, pady=(8,0))

        # Layout weights
        for i in range(3):
            main.columnconfigure(i, weight=1)
        for r in range(1,8):
            main.rowconfigure(r, weight=0)

        # Load policy once
        self.master.after(100, self.load_policy)

    def fmt_dt(self, iso):
        if not iso: return ""
        try:
            dt = datetime.fromisoformat(iso)
            return dt.strftime(DATE_FMT)
        except Exception:
            return iso

    def load_policy(self):
        secs, err = get_domain_max_age_seconds()
        if err:
            self.status.set(f"Policy load warning: {err}")
        else:
            self.status.set(f"Domain MaxPasswordAge: {secs//86400} day(s)")
        self.max_age_seconds = secs

    def do_check(self):
        sam = self.sam_var.get().strip()
        if not sam:
            messagebox.showwarning(APP_TITLE, "Please enter a sAMAccountName.")
            return
        self.status.set(f"Checking '{sam}'…")
        self.master.update_idletasks()
        res = check_user(sam, self.max_age_seconds)
        if not res.get("Success"):
            messagebox.showerror(APP_TITLE, res.get("Error","Unknown error."))
            self.status.set("Error.")
            return
        self.v_name.set(res.get("DisplayName",""))
        self.v_upn.set(res.get("UserPrincipalName",""))
        self.v_pls.set(self.fmt_dt(res.get("PasswordLastSet")))
        self.v_pexp.set(self.fmt_dt(res.get("PasswordExpires")))
        self.v_expired.set(str(res.get("Expired","")))
        self.v_enabled.set(str(res.get("Enabled","")))
        self.v_pne.set(str(res.get("PasswordNeverExpires","")))
        self.status.set(f"Checked '{sam}'.")

def main():
    root = Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
