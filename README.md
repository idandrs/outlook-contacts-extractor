# Outlook Contacts Extractor

Local Outlook contacts export and analysis tools for Windows 11.

This project exports your Outlook contacts into:

- JSON for bulk analysis
- CSV for import into Excel or other tools
- XLSX for direct use in Excel
- A local MCP server for Codex CLI sessions

## What This Uses

The exporter talks to Outlook through the Windows COM automation interface exposed by classic Outlook for Windows.

It does not scrape mailboxes directly and it does not need Microsoft Graph.

## Important Windows Requirement

This project is for Windows 11 with classic Outlook for Windows installed and signed into a profile.

It will not work with the new Outlook for Windows by itself because the new Outlook client does not expose the same COM automation model.

Before running the exporter:

1. Install classic Outlook for Windows.
2. Open Outlook at least once and sign into the profile whose contacts you want to export.
3. Close Outlook if you want, but keep the profile configured.
4. Make sure Python for Windows is installed.

## Files

- `scripts/bootstrap.ps1`: create `.venv` and install dependencies on Windows
- `scripts/export_contacts.py`: export Outlook contacts to JSON, CSV, XLSX, and summary files
- `outlook_contacts.py`: shared Outlook COM access and export logic
- `mcp_server.py`: local stdio MCP server for Codex CLI

## Setup

In PowerShell on Windows:

```powershell
cd "$HOME\Personal\outlook contacts extractor"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\bootstrap.ps1
```

## Export All Contacts

```powershell
cd "$HOME\Personal\outlook contacts extractor"
.\.venv\Scripts\python.exe .\scripts\export_contacts.py
```

Default output goes to:

- `output\latest\contacts.json`
- `output\latest\contacts.csv`
- `output\latest\contacts.xlsx`
- `output\latest\manifest.json`
- `output\latest\summary.md`

## MCP For Codex CLI

Register it with Codex from PowerShell on Windows:

```powershell
cd "$HOME\Personal\outlook contacts extractor"
codex mcp add outlook_contacts -- "$PWD\.venv\Scripts\python.exe" "$PWD\mcp_server.py"
```

Verify:

```powershell
codex mcp list
codex mcp get outlook_contacts
```

Then restart Codex CLI.

The server exposes tools such as:

- `outlook_status`
- `list_contact_stores`
- `list_contact_folders`
- `list_contacts`
- `search_contacts`
- `get_contact`
- `export_contacts_snapshot`

## Useful Options

Export only one Outlook store:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_contacts.py --store-name "your@email.com"
```

Export only one contact folder:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_contacts.py --folder-path "Contacts/Customers"
```

Skip distribution lists:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_contacts.py --skip-distribution-lists
```

Recreate the export directory first:

```powershell
.\.venv\Scripts\python.exe .\scripts\export_contacts.py --clean
```

## Output Schema

The main export includes fields such as:

- `full_name`
- `company_name`
- `job_title`
- `email_1_address`
- `email_2_address`
- `email_3_address`
- `business_phone`
- `mobile_phone`
- `home_phone`
- `business_*`, `home_*`, `other_*` address fields
- `categories`
- `notes`
- `store_name`
- `folder_path`
- `item_type`

## Limitations

- This project requires classic Outlook for Windows. If you only use the new Outlook client, use Microsoft Graph instead.
- Outlook COM access depends on the local Windows profile and whatever stores that Outlook profile can open.
- Shared or delegated contacts may or may not appear depending on how they are mounted in your Outlook profile.
- Distribution lists are exported as rows with `item_type = distribution_list` and best-effort member expansion.
