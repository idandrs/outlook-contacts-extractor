#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from outlook_contacts import OutlookAccessError  # noqa: E402
from outlook_mail_addresses import (  # noqa: E402
    DEFAULT_MAIL_OUTPUT_DIR,
    MailAddressFilters,
    export_mail_addresses_snapshot,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export email addresses mined from Outlook Inbox and Sent Mail folders."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_MAIL_OUTPUT_DIR,
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
        help="Limit export to one mail folder path, for example Inbox/Customers",
    )
    parser.add_argument(
        "--days-back",
        type=int,
        default=None,
        help="Only include mail from the last N days",
    )
    parser.add_argument(
        "--max-messages",
        type=int,
        default=None,
        help="Stop after scanning this many mail items",
    )
    parser.add_argument(
        "--address-scope",
        choices=["correspondents", "all-participants"],
        default="correspondents",
        help=(
            "correspondents scans senders of incoming mail and recipients of outgoing mail. "
            "all-participants scans both senders and recipients in both directions."
        ),
    )
    parser.add_argument(
        "--roots",
        choices=["both", "inbox", "sent"],
        default="both",
        help="Choose which default mail trees to scan",
    )
    parser.add_argument(
        "--no-subfolders",
        action="store_true",
        help="Only scan the root Inbox and/or Sent Mail folders",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Delete the output directory before exporting",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    filters = MailAddressFilters(
        store_name=args.store_name,
        folder_path=args.folder_path,
        include_subfolders=not args.no_subfolders,
        scan_inbox=args.roots in {"both", "inbox"},
        scan_sent_mail=args.roots in {"both", "sent"},
        days_back=args.days_back,
        max_messages=args.max_messages,
        address_scope=args.address_scope,
    )
    manifest = export_mail_addresses_snapshot(
        output_dir=args.output_dir.expanduser().resolve(),
        filters=filters,
        clean=args.clean,
    )

    print("Mail address export complete.")
    print(f"Unique addresses: {manifest['unique_address_count']}")
    print(f"Messages scanned: {manifest['scanned_message_count']}")
    print(f"XLSX: {manifest['files']['mail_addresses_xlsx']}")
    print(f"CSV: {manifest['files']['mail_addresses_csv']}")
    print(f"JSON: {manifest['files']['mail_addresses_json']}")
    print(f"Summary: {manifest['files']['summary_md']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except OutlookAccessError as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
