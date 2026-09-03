from __future__ import annotations

import argparse
import json

from .assign import run_assign
from .config import ORDER_BATCHES_LIST, ORDER_FORM_LIST, REQUISITIONS_LIST, VENDORS_LIST, load_settings
from .data import load_lists_context
from .generate import run_generate
from .graph import GraphClient, resolve_site
from .sharepoint import get_item, get_list_by_display_name, print_schema


def build_parser() -> argparse.ArgumentParser:
    """Define the command-line interface.

    `assign` and `generate` can run independently. `run-all` performs both
    actions after a single device-code login.
    """
    parser = argparse.ArgumentParser(prog="python -m ece_orders")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("inspect-schema", help="Print internal field names for the SharePoint lists")

    read_item = subparsers.add_parser("read-item", help="Print raw Graph fields for one SharePoint list item")
    read_item.add_argument("--list", required=True, help="SharePoint list display name, like 'ECE Requisitions'")
    read_item.add_argument("--item-id", required=True, help="SharePoint list item ID")

    assign = subparsers.add_parser("assign", help="Create requisitions and assign approved orders to them")
    assign.add_argument("--dry-run", action="store_true", help="Preview changes without writing")

    generate = subparsers.add_parser("generate", help="Generate PDFs/docs for pending requisitions")
    generate.add_argument("--dry-run", action="store_true", help="Preview changes without writing")

    run_all = subparsers.add_parser("run-all", help="Assign orders, then generate pending requisitions in one login")
    run_all.add_argument("--dry-run", action="store_true", help="Preview both steps without writing")
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    settings = load_settings()
    client = GraphClient(settings)
    client.login_device_code()
    site = resolve_site(client, settings.site_hostname, settings.site_path)
    site_id = site["id"]

    if args.command == "inspect-schema":
        print_schema(client, site_id, [ORDER_FORM_LIST, REQUISITIONS_LIST, VENDORS_LIST, ORDER_BATCHES_LIST])
        return
    if args.command == "read-item":
        sp_list = get_list_by_display_name(client, site_id, args.list)
        item = get_item(client, site_id, sp_list.id, args.item_id)
        print(json.dumps(item.get("fields", item), indent=2, sort_keys=True))
        return

    dry_run = args.dry_run
    ctx = load_lists_context(client, site_id)
    if args.command == "assign":
        run_assign(client, site_id, ctx, dry_run=dry_run)
    elif args.command == "generate":
        run_generate(client, site_id, ctx, settings, dry_run=dry_run)
    elif args.command == "run-all":
        print("\n=== Step 1: assign ===")
        run_assign(client, site_id, ctx, dry_run=dry_run)
        print("\n=== Step 2: generate ===")
        run_generate(client, site_id, ctx, settings, dry_run=dry_run)
