#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from outlook_contacts import (
    DEFAULT_OUTPUT_DIR,
    ContactFilters,
    OutlookAccessError,
    collect_contact_folders,
    collect_store_summaries,
    collect_status,
    export_contacts_snapshot as run_export_contacts_snapshot,
    get_contact_by_identity,
    preview_contacts,
    search_contact_rows,
)
from outlook_mail_addresses import (
    DEFAULT_MAIL_OUTPUT_DIR,
    MailAddressFilters,
    collect_mail_folders,
    export_mail_addresses_snapshot as run_export_mail_addresses_snapshot,
    get_mail_address as run_get_mail_address,
    preview_mail_addresses,
    search_mail_address_rows,
)


mcp = FastMCP("Outlook Contacts Extractor")


def _output_dir(raw: str | None) -> Path:
    return Path(raw).expanduser().resolve() if raw else DEFAULT_OUTPUT_DIR


def _mail_output_dir(raw: str | None) -> Path:
    return Path(raw).expanduser().resolve() if raw else DEFAULT_MAIL_OUTPUT_DIR


@mcp.tool()
def outlook_status() -> dict[str, Any]:
    """Report Outlook COM access status for the local Windows machine."""
    return collect_status()


@mcp.tool()
def list_contact_stores() -> dict[str, Any]:
    """List Outlook stores that are visible to the current profile."""
    stores = collect_store_summaries()
    return {"count": len(stores), "stores": stores}


@mcp.tool()
def list_contact_folders(store_name: str | None = None) -> dict[str, Any]:
    """List contact folders, optionally filtered to one Outlook store."""
    folders = collect_contact_folders(store_name=store_name)
    return {"count": len(folders), "folders": folders}


@mcp.tool()
def list_contacts(
    store_name: str | None = None,
    folder_path: str | None = None,
    limit: int = 25,
    offset: int = 0,
    include_notes_preview: bool = False,
    preview_chars: int = 300,
    include_distribution_lists: bool = True,
) -> dict[str, Any]:
    """List contacts with optional store and folder filters."""
    filters = ContactFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_distribution_lists=include_distribution_lists,
    )
    return preview_contacts(
        filters=filters,
        limit=limit,
        offset=offset,
        include_notes_preview=include_notes_preview,
        preview_chars=preview_chars,
    )


@mcp.tool()
def search_contacts(
    query: str,
    store_name: str | None = None,
    folder_path: str | None = None,
    limit: int = 25,
    include_notes_preview: bool = False,
    preview_chars: int = 300,
    include_distribution_lists: bool = True,
) -> dict[str, Any]:
    """Search contacts by name, company, email, phone, categories, notes, or members."""
    filters = ContactFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_distribution_lists=include_distribution_lists,
    )
    rows = search_contact_rows(query, filters, limit=limit)
    if not include_notes_preview:
        rows = [{key: value for key, value in row.items() if key != "notes"} for row in rows]
    else:
        materialized: list[dict[str, str]] = []
        for row in rows:
            item = dict(row)
            notes = item.pop("notes", "")
            item["notes_preview"] = notes[:preview_chars] if len(notes) <= preview_chars else notes[:preview_chars] + "\n...[truncated]"
            materialized.append(item)
        rows = materialized
    return {"count": len(rows), "results": rows}


@mcp.tool()
def get_contact(
    entry_id: str | None = None,
    full_name: str | None = None,
    store_name: str | None = None,
    folder_path: str | None = None,
    include_distribution_lists: bool = True,
) -> dict[str, Any]:
    """Fetch one contact by entry_id or full_name."""
    filters = ContactFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_distribution_lists=include_distribution_lists,
    )
    return get_contact_by_identity(
        entry_id=entry_id,
        full_name=full_name,
        filters=filters,
    )


@mcp.tool()
def export_contacts_snapshot(
    output_dir: str | None = None,
    clean: bool = False,
    store_name: str | None = None,
    folder_path: str | None = None,
    include_distribution_lists: bool = True,
) -> dict[str, Any]:
    """Export contacts to JSON, CSV, XLSX, manifest, and summary files."""
    filters = ContactFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_distribution_lists=include_distribution_lists,
    )
    return run_export_contacts_snapshot(
        output_dir=_output_dir(output_dir),
        filters=filters,
        clean=clean,
    )


@mcp.tool()
def list_mail_folders(
    store_name: str | None = None,
    folder_path: str | None = None,
    include_subfolders: bool = True,
    scan_inbox: bool = True,
    scan_sent_mail: bool = True,
) -> dict[str, Any]:
    """List Inbox and Sent Mail folders that can be scanned for discovered addresses."""
    filters = MailAddressFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_subfolders=include_subfolders,
        scan_inbox=scan_inbox,
        scan_sent_mail=scan_sent_mail,
    )
    folders = collect_mail_folders(filters)
    return {"count": len(folders), "folders": folders}


@mcp.tool()
def list_mail_addresses(
    store_name: str | None = None,
    folder_path: str | None = None,
    include_subfolders: bool = True,
    scan_inbox: bool = True,
    scan_sent_mail: bool = True,
    days_back: int | None = None,
    max_messages: int | None = None,
    address_scope: str = "correspondents",
    limit: int = 25,
    offset: int = 0,
    include_subject_preview: bool = True,
    preview_chars: int = 200,
) -> dict[str, Any]:
    """List unique addresses discovered from Outlook mail items."""
    filters = MailAddressFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_subfolders=include_subfolders,
        scan_inbox=scan_inbox,
        scan_sent_mail=scan_sent_mail,
        days_back=days_back,
        max_messages=max_messages,
        address_scope=address_scope,
    )
    return preview_mail_addresses(
        filters=filters,
        limit=limit,
        offset=offset,
        include_subject_preview=include_subject_preview,
        preview_chars=preview_chars,
    )


@mcp.tool()
def search_mail_addresses(
    query: str,
    store_name: str | None = None,
    folder_path: str | None = None,
    include_subfolders: bool = True,
    scan_inbox: bool = True,
    scan_sent_mail: bool = True,
    days_back: int | None = None,
    max_messages: int | None = None,
    address_scope: str = "correspondents",
    limit: int = 25,
) -> dict[str, Any]:
    """Search unique addresses discovered from Outlook mail items."""
    filters = MailAddressFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_subfolders=include_subfolders,
        scan_inbox=scan_inbox,
        scan_sent_mail=scan_sent_mail,
        days_back=days_back,
        max_messages=max_messages,
        address_scope=address_scope,
    )
    return search_mail_address_rows(query, filters, limit=limit)


@mcp.tool()
def get_mail_address(
    email_address: str,
    store_name: str | None = None,
    folder_path: str | None = None,
    include_subfolders: bool = True,
    scan_inbox: bool = True,
    scan_sent_mail: bool = True,
    days_back: int | None = None,
    max_messages: int | None = None,
    address_scope: str = "correspondents",
) -> dict[str, Any]:
    """Fetch one discovered mail-derived address by email address."""
    filters = MailAddressFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_subfolders=include_subfolders,
        scan_inbox=scan_inbox,
        scan_sent_mail=scan_sent_mail,
        days_back=days_back,
        max_messages=max_messages,
        address_scope=address_scope,
    )
    return run_get_mail_address(email_address=email_address, filters=filters)


@mcp.tool()
def export_mail_addresses_snapshot(
    output_dir: str | None = None,
    clean: bool = False,
    store_name: str | None = None,
    folder_path: str | None = None,
    include_subfolders: bool = True,
    scan_inbox: bool = True,
    scan_sent_mail: bool = True,
    days_back: int | None = None,
    max_messages: int | None = None,
    address_scope: str = "correspondents",
) -> dict[str, Any]:
    """Export unique addresses discovered from Inbox and Sent Mail folders."""
    filters = MailAddressFilters(
        store_name=store_name,
        folder_path=folder_path,
        include_subfolders=include_subfolders,
        scan_inbox=scan_inbox,
        scan_sent_mail=scan_sent_mail,
        days_back=days_back,
        max_messages=max_messages,
        address_scope=address_scope,
    )
    return run_export_mail_addresses_snapshot(
        output_dir=_mail_output_dir(output_dir),
        filters=filters,
        clean=clean,
    )


if __name__ == "__main__":
    try:
        mcp.run()
    except OutlookAccessError as exc:
        raise SystemExit(str(exc))
