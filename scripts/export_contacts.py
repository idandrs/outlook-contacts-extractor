#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from outlook_contacts import (  # noqa: E402
    DEFAULT_OUTPUT_DIR,
    ContactFilters,
    OutlookAccessError,
    export_contacts_snapshot,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Outlook contacts to JSON, CSV, XLSX, and summary files."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for export output",
    )
    parser.add_argument(
        "--store-name",
        type=str,
        default=None,
        help="Limit export to one Outlook store display name",
    )
    parser.add_argument(
        "--folder-path",
        type=str,
        default=None,
        help="Limit export to one contact folder path, for example Contacts/Customers",
    )
    parser.add_argument(
        "--skip-distribution-lists",
        action="store_true",
        help="Do not export distribution lists",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Delete the output directory before exporting",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    filters = ContactFilters(
        store_name=args.store_name,
        folder_path=args.folder_path,
        include_distribution_lists=not args.skip_distribution_lists,
    )
    manifest = export_contacts_snapshot(
        output_dir=args.output_dir.expanduser().resolve(),
        filters=filters,
        clean=args.clean,
    )

    print("Export complete.")
    print(f"Rows: {manifest['row_count']}")
    print(f"Contacts: {manifest['contact_count']}")
    print(f"Distribution lists: {manifest['distribution_list_count']}")
    print(f"XLSX: {manifest['files']['contacts_xlsx']}")
    print(f"CSV: {manifest['files']['contacts_csv']}")
    print(f"JSON: {manifest['files']['contacts_json']}")
    print(f"Summary: {manifest['files']['summary_md']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except OutlookAccessError as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
