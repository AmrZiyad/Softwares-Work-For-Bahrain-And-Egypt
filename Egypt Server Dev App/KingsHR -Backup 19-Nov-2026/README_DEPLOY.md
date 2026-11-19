# Kings School Egypt — HR Web App Deployment Guide

## What This Is

A web-based version of the Kings School Egypt HR & Payroll desktop app.  
HR staff (Amr, HSE, ARN) open a browser and go to **http://10.4.172.15/** — no software installation needed on their laptops.

---

## Server: App01-KSCR (10.4.172.15)

| Item | Value |
|------|-------|
| Server IP | 10.4.172.15 |
| Windows login | Administrator / Mad@123456 |
| SQL login | businessapp / Kings@123 |
| DB | KingsSchoolEgypt |
| App port | 80 (HTTP) |

---

## First-Time Installation

### Step 1 — Copy files to server

Copy the entire `kings_web` folder to the server, e.g. via:
- USB drive / shared folder → paste to `C:\KingsHR`
- Or RDP to the server (`mstsc /v:10.4.172.15`) and copy from a network share

### Step 2 — Run the installer

On the server, open **PowerShell as Administrator** and run:

```powershell
cd C:\KingsHR
Set-ExecutionPolicy Bypass -Scope Process
.\install.ps1
```

The installer will:
1. Download and install **Python 3.11** (if not present)
2. Copy all app files to `C:\KingsHR`
3. Install Python packages (`Flask`, `pyodbc`, `gunicorn`)
4. Install **ODBC Driver 18 for SQL Server** (if not present)
5. Open **port 80** in the Windows firewall
6. Download **NSSM** (service manager)
7. Register and start the `KingsHRApp` Windows service (auto-starts on boot)

### Step 3 — Test

Open a browser on a laptop and go to: **http://10.4.172.15/**

You should see the Kings School Egypt login page.

---

## Default Login

The app reads users from the `Users` table in the database.  
If no users exist yet, add the first Admin user directly in SQL:

```sql
INSERT INTO Users (Username, Password, FullName, Role, IsActive)
VALUES ('admin', 'Kings@123', 'System Admin', 'Admin', 1);
```

Then log in with `admin` / `Kings@123` and add the rest through the User Manager screen.

### Roles

| Role | Access |
|------|--------|
| Admin | Everything: users, payroll, employees, settings |
| HR | Employees + Payroll + Reports |
| IT | View only (extend as needed) |
| Ops | View only (extend as needed) |
| Emp | Employee self-service (extend as needed) |

---

## Managing the Service

RDP to the server and open PowerShell as Admin:

```powershell
# Check status
Get-Service KingsHRApp

# Stop / start / restart
nssm stop    KingsHRApp
nssm start   KingsHRApp
nssm restart KingsHRApp

# View logs
Get-Content C:\KingsHR\logs\service.log -Tail 50
```

---

## Files

```
C:\KingsHR\
├── app.py              # Flask routes (main app logic)
├── db.py               # Database connection (SQL Server)
├── payroll.py          # Payroll calculation engine
├── requirements.txt    # Python packages
├── templates/          # HTML pages
│   ├── base.html
│   ├── login.html
│   ├── dashboard.html
│   ├── employees.html
│   ├── employee_form.html
│   ├── payroll_menu.html
│   ├── payroll_manual.html
│   ├── payroll_run_all.html
│   ├── payroll_history.html
│   ├── payroll_history_detail.html
│   ├── payroll_rollback.html
│   ├── payroll_settings.html
│   ├── reports_menu.html
│   ├── reports_run.html
│   └── user_manager.html
├── static/
│   ├── css/
│   └── img/
└── logs/
    └── service.log
```

---

## Changing the Database Password

Edit `C:\KingsHR\db.py` and update the `CONN_STR` line:

```python
CONN_STR = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    "SERVER=localhost;DATABASE=KingsSchoolEgypt;"
    "UID=businessapp;PWD=Kings@123;TrustServerCertificate=yes;"
    "Connection Timeout=5;"
)
```

After editing, restart the service:
```powershell
nssm restart KingsHRApp
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Can't reach http://10.4.172.15/ | Check firewall: `Get-NetFirewallRule -DisplayName KingsHRApp-Port80` |
| Login page loads but login fails | Check `Users` table in DB; add a user via SQL (see above) |
| "Database error" message | Check SQL Server is running; verify credentials in `db.py` |
| Service not starting | Check `C:\KingsHR\logs\service.log` |
| Python not found after install | Reopen PowerShell as Admin (PATH refresh needed) |
