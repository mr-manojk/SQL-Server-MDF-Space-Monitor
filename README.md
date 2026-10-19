# SQL Server MDF Space Monitor

A lightweight PowerShell script that automatically monitors your SQL Server MDF (data file) free space and sends a styled HTML email alert when it drops below a threshold you define. Runs silently via Windows Task Scheduler — no SQL Server Agent required.

---

## What it does

- Queries `sys.database_files` to calculate free space on your MDF file
- Compares it against your configured threshold (MB)
- Sends a styled HTML warning email if free space is low
- Logs `OK` or `ALERT sent` to the console on every run
- Works with any SMTP provider — Office 365, Gmail, internal relay

---

## Email preview

The alert email uses a fully inline-CSS, Outlook-safe HTML template with an orange warning design:

- Header bar with ⚠ warning icon
- Detail table showing server, database, file name, path, total size and free space
- Free space highlighted in an orange badge when below threshold
- Timestamp of the check at the bottom

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows PowerShell 5.1+ | Built into Windows 10 / Server 2016 and later |
| `SqlServer` PowerShell module | Run `Install-Module -Name SqlServer` once |
| SQL Server login | Needs `VIEW DATABASE STATE` permission |
| SMTP access | Office 365, Gmail (app password), or internal relay |

---

## Setup

### 1. Clone or download

```bash
git clone https://github.com/mr-manojk/SQL-Server-MDF-Space-Monitor.git
```

### 2. Fill in your credentials

Copy `config.ps1` and fill in your real values — this file is excluded from git via `.gitignore` so credentials never get committed:

```powershell
$SqlServer    = "YOUR_SERVER\SQLEXPRESS"
$SqlDatabase  = "YOUR_DATABASE"
$ThresholdMB  = 900                         # alert below this many MB free
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$SmtpUser     = "sender@yourdomain.com"
$SmtpPassword = "YOUR_APP_PASSWORD"
$ToEmail      = "recipient@yourdomain.com"
$FromEmail    = "sender@yourdomain.com"
```

> `config.ps1` is listed in `.gitignore` — it will never be staged or pushed. `MonitorMDFSpace.ps1` contains no credentials and is safe to commit as-is.

### 3. Test manually

Open PowerShell and run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\MonitorMDFSpace.ps1"
```

Temporarily raise `$ThresholdMB` above your current free space to force a test alert email.

### 4. Schedule with Windows Task Scheduler

Open **Task Scheduler** → **Create Task** and configure:

| Setting | Value |
|---|---|
| **General → Run whether user is logged on or not** | Checked |
| **General → Run with highest privileges** | Checked |
| **Trigger → Schedule** | Daily, repeat every 30 minutes, indefinitely |
| **Action → Program** | `powershell.exe` |
| **Action → Arguments** | `-NonInteractive -ExecutionPolicy Bypass -File "C:\path\to\MonitorMDFSpace.ps1"` |
| **Settings → Stop task if runs longer than** | 5 minutes |

---

## SMTP provider notes

**Office 365**
```powershell
$SmtpServer = "smtp.office365.com"
$SmtpPort   = 587
```
Ensure SMTP AUTH is enabled for the mailbox in the Microsoft 365 admin centre.

**Gmail**
```powershell
$SmtpServer = "smtp.gmail.com"
$SmtpPort   = 587
```
Use an [App Password](https://support.google.com/accounts/answer/185833) — not your regular Google password. Requires 2-Step Verification to be on.

**Internal SMTP relay (no auth)**

Remove the `$smtpClient.Credentials` line and set your relay server:
```powershell
$SmtpServer = "relay.internal.company.com"
$SmtpPort   = 25
$smtpClient.EnableSsl = $false
```

---

## Keeping credentials safe

For production, store credentials outside the script using one of these approaches:

```powershell
# Option A — environment variable
$SmtpPassword = $env:SMTP_PASSWORD

# Option B — Windows Credential Manager (most secure)
$cred = Get-StoredCredential -Target "SQLMonitorSMTP"
$SmtpPassword = $cred.GetNetworkCredential().Password
```

Add a `config.ps1` (with real values) to `.gitignore` and dot-source it:
```powershell
. ".\config.ps1"
```

---

## File structure

```
sql-mdf-monitor/
├── MonitorMDFSpace.ps1   # Main script (safe to commit — no credentials)
├── config.ps1            # Your private credentials (git-ignored, never committed)
├── .gitignore            # Excludes config.ps1
└── README.md
```

---

## License

MIT — free to use, modify and distribute.
