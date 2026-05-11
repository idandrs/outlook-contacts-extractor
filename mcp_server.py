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


mcp = FastMCP("Outlook Contacts Extractor")


def _output_dir(raw: str | None) -> Path:
    return Path(raw).expanduser().resolve() if raw else DEFAULT_OUTPUT_DIR


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


if __name__ == "__main__":
    try:
        mcp.run()
    except OutlookAccessError as exc:
        raise SystemExit(str(exc))
