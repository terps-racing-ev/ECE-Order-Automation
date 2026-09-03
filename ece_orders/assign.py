from __future__ import annotations

import datetime as dt
from collections import defaultdict
from typing import Any

from . import sharepoint as sp
from .config import MAX_ROWS_PER_PDF, PENDING_CREATION
from .data import ListsContext, field, load_vendors, load_vendors_by_id, lookup_id_value, normalize_name
from .graph import GraphClient


def chunk(items: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def is_truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"true", "yes", "1"}


def collect_assignable_orders(client: GraphClient, site_id: str, ctx: ListsContext) -> tuple[list[dict[str, Any]], list[str]]:
    """Read approved order-list items that do not yet point at a requisition."""
    item_col = ctx.order_cols.title()
    vendor_col = ctx.order_cols.internal("Vendor")
    chief_col = ctx.order_cols.internal("Chief Approved")
    req_form_col = ctx.order_cols.internal("Req Form")
    part_col = ctx.order_cols.internal("Part Number")
    cost_col = ctx.order_cols.internal("Unit Cost")
    qty_col = ctx.order_cols.internal("Quantity")

    orders: list[dict[str, Any]] = []
    skipped: list[str] = []
    for item in sp.iter_items(client, site_id, ctx.order_form.id):
        fields = item.get("fields", {})
        if not is_truthy(field(fields, chief_col, False)):
            continue
        if lookup_id_value(fields, req_form_col):
            continue
        vendor_lookup_id = lookup_id_value(fields, vendor_col)
        vendor = str(field(fields, vendor_col, "")).strip()
        if not vendor_lookup_id and not vendor:
            skipped.append(f"Order {item['id']} has no Vendor")
            continue
        orders.append(
            {
                "id": item["id"],
                "item": field(fields, item_col, ""),
                "vendor": vendor,
                "vendor_lookup_id": vendor_lookup_id,
                "part_number": field(fields, part_col, ""),
                "unit_cost": field(fields, cost_col, ""),
                "quantity": field(fields, qty_col, ""),
            }
        )
    return orders, skipped


def get_or_create_batch(client: GraphClient, site_id: str, ctx: ListsContext, batch_id: str) -> dict[str, Any]:
    """Return the ECE Order Batches item for this run, creating it if needed."""
    title_col = ctx.batch_cols.title()
    matches = []
    for item in sp.iter_items(client, site_id, ctx.order_batches.id, select_fields=[title_col]):
        fields = item.get("fields", {})
        if str(field(fields, title_col, "")).strip() == batch_id:
            matches.append(item)

    if len(matches) > 1:
        ids = ", ".join(str(item["id"]) for item in matches)
        raise RuntimeError(f"Batch ID '{batch_id}' matched multiple ECE Order Batches items: {ids}")
    if matches:
        return matches[0]

    return sp.create_item(client, site_id, ctx.order_batches.id, {title_col: batch_id})


def run_assign(client: GraphClient, site_id: str, ctx: ListsContext, dry_run: bool) -> None:
    """Create requisitions and write their lookup IDs back to order items."""
    vendors = load_vendors(client, site_id, ctx)
    vendors_by_id = load_vendors_by_id(client, site_id, ctx)
    orders, skipped = collect_assignable_orders(client, site_id, ctx)
    valid_by_vendor: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for order in orders:
        # Vendor is a lookup column. The lookup ID is the authoritative value;
        # name matching is only a fallback for older/manual data.
        vendor_match = vendors_by_id.get(str(order["vendor_lookup_id"])) if order["vendor_lookup_id"] else None
        if not vendor_match:
            key = normalize_name(order["vendor"])
            matches = vendors.get(key, {}).get("matches", [])
            if len(matches) != 1:
                label = order["vendor"] or f"lookup id {order['vendor_lookup_id']}"
                reason = "missing" if not matches else f"ambiguous ({len(matches)} matches)"
                skipped.append(f"Order {order['id']} vendor '{label}' is {reason} in ECE Approved Vendors")
                continue
            vendor_match = matches[0]
        order["vendor_lookup_id"] = vendor_match["id"]
        order["vendor_name"] = vendor_match["name"]
        valid_by_vendor[vendor_match["name"]].append(order)

    today = dt.datetime.now()
    today_id = today.strftime("%m%d%y")
    today_date = today.strftime("%Y-%m-%d")
    letters = [chr(code) for code in range(ord("A"), ord("Z") + 1)]

    print(f"Found {len(orders)} chief-approved unassigned order(s).")
    if skipped:
        print("\nSkipped:")
        for message in skipped:
            print(f" - {message}")

    planned: list[tuple[str, str, list[dict[str, Any]]]] = []
    for vendor_name, vendor_orders in sorted(valid_by_vendor.items()):
        for idx, group in enumerate(chunk(vendor_orders, MAX_ROWS_PER_PDF)):
            suffix = letters[idx] if idx < len(letters) else str(idx + 1)
            req_id = f"{vendor_name} {today_id} {suffix}"
            planned.append((vendor_name, req_id, group))

    print(f"\nPlanned requisition(s): {len(planned)}")
    for vendor_name, req_id, group in planned:
        print(f" - {req_id}: {len(group)} order(s) for {vendor_name}")

    if dry_run:
        print(f"\nDry run: batch {today_id} was not created or linked, no requisitions created, and no orders patched.")
        return

    title_col = ctx.req_cols.title()
    req_vendor_lookup = ctx.req_cols.lookup_id("Vendor")
    req_batch_lookup = ctx.req_cols.lookup_id("Batch")
    status_col = ctx.req_cols.internal("Requisition Status")
    date_col = ctx.req_cols.internal("Date Created")
    order_req_lookup = ctx.order_cols.lookup_id("Req Form")
    batch_id = get_or_create_batch(client, site_id, ctx, today_id)["id"] if planned else ""

    created = 0
    patched_orders = 0
    for _vendor_name, req_id, group in planned:
        # Create one requisition item per vendor chunk, then point each order's
        # Req Form lookup at the newly-created requisition item.
        req = sp.create_item(
            client,
            site_id,
            ctx.requisitions.id,
            {
                title_col: req_id,
                req_vendor_lookup: group[0]["vendor_lookup_id"],
                req_batch_lookup: batch_id,
                status_col: PENDING_CREATION,
                date_col: today_date,
            },
        )
        created += 1
        for order in group:
            sp.patch_item_fields(client, site_id, ctx.order_form.id, order["id"], {order_req_lookup: req["id"]})
            patched_orders += 1

    print(f"\nCreated {created} requisition(s) and assigned {patched_orders} order(s).")
