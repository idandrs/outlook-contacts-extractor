# Windows 11 Guide

This guide is for a Windows 11 user who wants to:

1. download the project from GitHub
2. run it
3. open the results in Excel

Direct ZIP download:

- `https://github.com/idandrs/outlook-contacts-extractor-public/archive/refs/heads/main.zip`

## What This Tool Does

This project can export two kinds of data from classic Outlook for Windows:

- saved contacts from Outlook Contacts folders
- email addresses discovered from incoming and outgoing mail, even if they are not saved as contacts

If your goal is to build a spreadsheet of everyone seen in mail, use the mail-based export.

## Before You Start

You need all of these:

- Windows 11
- classic Outlook for Windows installed
- Outlook signed into the mailbox/profile you want to scan
- Python installed on Windows

Important:

- this does not work with the new Outlook app by itself
- the default runner already scans the last `365` days from both Inbox and Sent Mail

## Step 1: Install Python

If Python is not already installed:

1. Go to `https://www.python.org/downloads/windows/`
2. Install Python 3.
3. During install, enable `Add Python to PATH` if the installer offers it.

Check that Python works:

```powershell
py -3 --version
```

If `py` does not work, try:

```powershell
python --version
```

## Step 2: Make Sure Outlook Is Ready

1. Open classic Outlook for Windows.
2. Sign into the mailbox/profile you want to scan.
3. Wait until Outlook finishes loading mail folders.
4. Close Outlook, or leave it open if you prefer.

If you are using the new Outlook app, switch back to classic Outlook first.

## Step 3: Download the Project ZIP

Open this ZIP link in your browser:

`https://github.com/idandrs/outlook-contacts-extractor-public/archive/refs/heads/main.zip`

Then:

1. Download the file
2. Save it somewhere convenient, for example `Downloads`
3. Right-click the ZIP file and choose `Extract All...`
4. Extract it to a folder such as:

```text
C:\Users\<your-user>\Desktop\outlook-contacts-extractor
```

5. Open the extracted folder
6. Right-click inside the folder and choose `Open in Terminal` or `Open PowerShell window here`

## Step 4: Install the Project Dependencies

Still in PowerShell, run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\bootstrap.ps1
```

This creates a local Python environment in `.venv` and installs everything the exporter needs.

## Step 5: Run the Mail-Based Export

This is the mode to use if you want addresses from real emails, including people who are not saved in Contacts.

### Default no-parameter run

```powershell
.\scripts\run_default_mail_export.ps1
```

What this does:

- scans Inbox and Sent Mail
- includes subfolders
- looks at the last 365 days of mail
- creates Excel, CSV, JSON, and summary files

### Same default using the Python command directly

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py
```

### If you want every sender and recipient in both directions

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py --days-back 365 --address-scope all-participants
```

### If the mailbox is very large

Start with a smaller scan:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py --days-back 90 --max-messages 5000
```

## Step 6: Find the Results

After the command finishes, PowerShell prints the output file paths.

The main Excel result is:

```text
output\mail-addresses\latest\mail_addresses.xlsx
```

Other output files:

- `output\mail-addresses\latest\mail_addresses.csv`
- `output\mail-addresses\latest\mail_addresses.json`
- `output\mail-addresses\latest\manifest.json`
- `output\mail-addresses\latest\summary.md`

## Step 7: Open the Results in Excel

From File Explorer:

1. Open the project folder.
2. Open `output`.
3. Open `mail-addresses`.
4. Open `latest`.
5. Double-click `mail_addresses.xlsx`.

Or from PowerShell:

```powershell
ii .\output\mail-addresses\latest\mail_addresses.xlsx
```

## What You Will See in the Spreadsheet

Each row is one deduplicated email address discovered from mail.

Useful columns include:

- `email_address`
- `primary_display_name`
- `display_names`
- `incoming_sender_count`
- `incoming_recipient_count`
- `outgoing_sender_count`
- `outgoing_recipient_count`
- `message_count_total`
- `first_seen_at`
- `last_seen_at`
- `stores`
- `folders`
- `sample_subjects`

## Optional: Export Only Saved Contacts

If you want Outlook Contacts folder data instead of mail-derived addresses, run:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_contacts.py
```

The Excel output for saved contacts is:

```text
output\latest\contacts.xlsx
```

## Common Problems

### Python is not recognized

Install Python, close PowerShell, then open it again.

### Outlook opens but the script cannot read mailbox data

Make sure:

- you are using classic Outlook, not new Outlook
- Outlook has already been opened at least once
- the mailbox/profile is signed in and available

### The scan is too slow

Reduce the scope:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py --days-back 30 --max-messages 2000
```

### I only want Inbox or only Sent Mail

Use:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py --roots inbox
```

or:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_addresses_from_mail.py --roots sent
```

## Fastest Copy-Paste Version

If Python and classic Outlook are already installed, this is the shortest path after you extract the ZIP and open PowerShell in that folder:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\bootstrap.ps1
.\scripts\run_default_mail_export.ps1
ii .\output\mail-addresses\latest\mail_addresses.xlsx
```
