# Cloud Server Scheduling Guide

## Goal
Process all new Excel files uploaded to the shared input folder every night, generate enriched outputs, and archive processed inputs.

## Shared Folder Paths
- Input: \\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\input
- Output: \\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\output
- Archive (created automatically): \\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\archive

## Files to Deploy on Server
- Script runner: automation/run_nightly_contact_company_enrichment.ps1
- Python app: run_contact_company_enrichment.py
- Dependencies file: requirements.txt

## One-Time Server Setup
1. Install Python 3.11+ on the server.
2. Open PowerShell as admin and create/activate virtual environment:
   - python -m venv C:\Automation\venv
   - C:\Automation\venv\Scripts\Activate.ps1
3. Install dependencies from project folder:
   - pip install -r requirements.txt
4. Ensure the service account has read/write access to:
   - Top_Opportunity\input
   - Top_Opportunity\output
   - Top_Opportunity\archive
5. Ensure firewall/network policy allows access to DIVE endpoint used by the Python script.
6. For email notifications from `R-IFX-DEMMA@infineon.com`, grant the task service account `Send As` permission on that mailbox.
7. Confirm SMTP relay host/port that allows sending from the server.

## Test Run Before Scheduling
Run this once manually from PowerShell:

C:\Automation\venv\Scripts\python.exe C:\path\to\run_contact_company_enrichment.py --input-dir "\\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\input" --output-dir "\\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\output" --workbook "*.xlsx"

If credentials are not in environment variables, set them before running:
- $env:DIVE_UID = "uid"
- $env:DIVE_PWD = "password"

## Schedule as Nightly Task (Task Scheduler UI)
1. Open Task Scheduler.
2. Create Task (not Basic Task).
3. General tab:
   - Name: TopOpportunity Contact Company Enrichment Nightly
   - Run whether user is logged on or not
   - Run with highest privileges
   - Configure for: Windows Server version in use
4. Triggers tab:
   - New trigger: Daily
   - Start time: 01:00 AM
5. Actions tab:
   - Program/script: powershell.exe
   - Add arguments:
       -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\automation\run_nightly_contact_company_enrichment.ps1" -PythonExe "C:\Automation\venv\Scripts\python.exe"
   - Start in:
     C:\path\to\contact_company_enrichment
6. Conditions tab:
   - Uncheck options that block task in server contexts, as needed.
7. Settings tab:
   - Allow task to be run on demand
   - If task fails, restart every 15 minutes (recommended)
   - Stop task if it runs longer than 8 hours

## Optional schtasks Command
Use this if you prefer command-line registration:

schtasks /Create /TN "TopOpportunity_ContactCompany_Nightly" /SC DAILY /ST 01:00 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\path\to\automation\run_nightly_contact_company_enrichment.ps1\" -PythonExe \"C:\Automation\venv\Scripts\python.exe\"" /RU "DOMAIN\\service-account" /RP "<password>"

## Enable Email Notifications (Per Output File)

The runner supports optional email notification as soon as each file output is created.

### Required parameters
- `-EnableEmailNotifications`
- `-SmtpServer` (SMTP relay hostname)
- `-MailTo` (comma-separated recipients)

### Optional parameters
- `-MailFrom` default: `R-IFX-DEMMA@infineon.com`
- `-MailCc`
- `-SmtpPort` default: `25`
- `-UseSsl`

### Task action arguments example (with email)

`-NoProfile -ExecutionPolicy Bypass -File "C:\path\to\automation\run_nightly_contact_company_enrichment.ps1" -PythonExe "C:\Automation\venv\Scripts\python.exe" -EnableEmailNotifications -SmtpServer "smtp-relay.infineon.com" -SmtpPort 25 -MailFrom "R-IFX-DEMMA@infineon.com" -MailTo "user1@infineon.com,user2@infineon.com"`

### Notes
1. Sender spoofing controls may block mail if `Send As` is not granted.
2. If SMTP requires authentication, configure relay exception or update script to use credentials.
3. Email send failures are logged as warnings and do not stop file processing.

## Runtime Behavior
1. Script scans input folder for .xlsx files.
2. Each file is processed individually.
3. Output is written to output folder.
4. Successfully processed input files are moved to archive\YYYYMMDD_HHMMSS.
5. Failed files remain in input for retry next run.
6. Logs are written to project logs folder as nightly_run_YYYYMMDD_HHMMSS.log.
7. If email notification is enabled, a mail is sent per successfully generated output file.

## Operational Checks
1. Every morning verify:
   - New enriched files in output
   - Previous night inputs moved to archive
2. On failure:
   - Check latest log in logs folder
   - Confirm share permissions and DIVE credentials
   - Re-run task manually from Task Scheduler

## Security Notes
1. Prefer running task under a managed service account.
2. Avoid plain-text passwords in scripts.
3. If required, pass DIVE credentials via secure environment configuration managed by server admins.
