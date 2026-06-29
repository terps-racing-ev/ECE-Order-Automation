from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from . import sharepoint as sp
from .config import PENDING_ADVISOR_APPROVAL, PENDING_CREATION, Settings
from .data import ListsContext, field, load_vendors, load_vendors_by_id, lookup_id_value, normalize_name
from .documents import build_links_docx
from .graph import GraphClient, get_default_drive_id
from .pdf import fill_order_pdf
from .storage import ensure_date_folder, upload_small


def collect_pending_requisitions(client: GraphClient, site_id: str, ctx: ListsContext) -> list[dict[str, Any]]:
    """Read requisitions whose status says their PDF/docx still need creation."""
    status_col = ctx.req_cols.internal("Requisition Status")
    title_col = ctx.req_cols.title()
    vendor_col = ctx.req_cols.internal("Vendor")
    pending = []
    for item in sp.iter_items(client, site_id, ctx.requisitions.id):
        fields = item.get("fields", {})
        if str(field(fields, status_col, "")).strip() != PENDING_CREATION:
            continue
        pending.append(
            {
                "id": item["id"],
                "internal_req_id": field(fields, title_col, f"Req {item['id']}"),
                "vendor_lookup_id": lookup_id_value(fields, vendor_col),
            }
        )
    return pending


def collect_orders_for_req(client: GraphClient, site_id: str, ctx: ListsContext, req_id: str) -> list[dict[str, Any]]:
    """Find order-list items assigned to one requisition lookup ID."""
    order_cols = ctx.order_cols
    title_col = order_cols.title()
    req_form_col = order_cols.internal("Req Form")
    fields_needed = {
        "part_number": order_cols.internal("Part Number"),
        "vendor": order_cols.internal("Vendor"),
        "link": order_cols.internal("Link"),
        "unit_cost": order_cols.internal("Unit Cost"),
        "quantity": order_cols.internal("Quantity"),
    }
    optional_special = order_cols.columns.get("Special Instructions", {}).get("name")

    orders: list[dict[str, Any]] = []
    for item in sp.iter_items(client, site_id, ctx.order_form.id):
        fields = item.get("fields", {})
        if lookup_id_value(fields, req_form_col) != str(req_id):
            continue
        order = {
            "id": item["id"],
            "item": field(fields, title_col, ""),
            "part_number": field(fields, fields_needed["part_number"], ""),
            "vendor": field(fields, fields_needed["vendor"], ""),
            "vendor_lookup_id": lookup_id_value(fields, fields_needed["vendor"]),
            "link": field(fields, fields_needed["link"], ""),
            "unit_cost": field(fields, fields_needed["unit_cost"], 0),
            "quantity": field(fields, fields_needed["quantity"], 0),
            "special_instructions": "",
        }
        if optional_special:
            order["special_instructions"] = field(fields, optional_special, "")
        orders.append(order)
    return orders


def vendor_by_lookup_id(vendors: dict[str, dict[str, Any]], lookup_id: str) -> dict[str, Any] | None:
    for vendor_bucket in vendors.values():
        for vendor in vendor_bucket["matches"]:
            if str(vendor["id"]) == str(lookup_id):
                return vendor
    return None


def safe_filename(value: str) -> str:
    forbidden = '<>:"/\\|?*'
    cleaned = "".join("_" if char in forbidden else char for char in value).strip()
    return cleaned or "Requisition"


def output_base_name(vendor_name: str, internal_req_id: str) -> str:
    parts = str(internal_req_id or "").split()
    if len(parts) >= 2:
        date_suffix = " ".join(parts[-2:])
        return safe_filename(f"{vendor_name} Order - {date_suffix}")
    return safe_filename(f"{vendor_name} Order - {internal_req_id}")


def run_generate(
    client: GraphClient,
    site_id: str,
    ctx: ListsContext,
    settings: Settings,
    dry_run: bool,
) -> None:
    """Generate PDFs/docs for pending requisitions and update requisition rows."""
    if not settings.blank_order_pdf.exists():
        raise FileNotFoundError(f"Missing template PDF: {settings.blank_order_pdf}")

    requisitions = collect_pending_requisitions(client, site_id, ctx)
    vendors = load_vendors(client, site_id, ctx)
    vendors_by_id = load_vendors_by_id(client, site_id, ctx)
    print(f"Found {len(requisitions)} requisition(s) with status '{PENDING_CREATION}'.")

    plans: list[dict[str, Any]] = []
    skipped: list[str] = []
    for req in requisitions:
        orders = collect_orders_for_req(client, site_id, ctx, req["id"])
        if not orders:
            skipped.append(f"Requisition {req['internal_req_id']} has no linked ECE Order Form items")
            continue
        vendor = vendors_by_id.get(str(req["vendor_lookup_id"])) or vendor_by_lookup_id(vendors, req["vendor_lookup_id"])
        if not vendor:
            # Requisitions should have their own Vendor lookup, but falling back
            # to the first order's vendor keeps older partially-created data usable.
            order_vendor_lookup_id = orders[0].get("vendor_lookup_id", "")
            vendor = vendors_by_id.get(str(order_vendor_lookup_id)) if order_vendor_lookup_id else None
            if not vendor:
                fallback_key = normalize_name(orders[0].get("vendor", ""))
                matches = vendors.get(fallback_key, {}).get("matches", [])
                vendor = matches[0] if len(matches) == 1 else None
        if not vendor:
            skipped.append(f"Requisition {req['internal_req_id']} could not resolve its approved vendor lookup")
            continue
        base = output_base_name(vendor["name"], req["internal_req_id"])
        plans.append(
            {
                "req": req,
                "vendor": vendor,
                "orders": orders,
                "pdf_name": f"{base}.pdf",
                "docx_name": f"{base} Links.docx",
            }
        )

    for message in skipped:
        print(f" - Skipped: {message}")
    print(f"\nPlanned generation(s): {len(plans)}")
    for plan in plans:
        print(f" - {plan['req']['internal_req_id']}: {plan['pdf_name']} and {plan['docx_name']}")

    if dry_run:
        print(f"\nDry run: no files uploaded and no requisitions patched. Destination: {settings.output_parent_sp_path}/{{MMDDYY}}")
        return

    drive_id = get_default_drive_id(client, site_id)
    date_folder = ensure_date_folder(client, drive_id, settings.output_parent_sp_path)
    form_col = ctx.req_cols.internal("Requisition Form")
    links_col = ctx.req_cols.internal("Links Document")
    status_col = ctx.req_cols.internal("Requisition Status")

    completed = 0
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for plan in plans:
            # Generate locally, upload to the date folder, then store the uploaded
            # SharePoint URLs in plain text columns on the requisition item.
            pdf_path = tmpdir / plan["pdf_name"]
            fill_order_pdf(
                settings.blank_order_pdf,
                pdf_path,
                account="",
                vendor=plan["vendor"]["name"],
                vendor_info=plan["vendor"],
                items=plan["orders"],
            )
            pdf_upload = upload_small(client, drive_id, date_folder["id"], plan["pdf_name"], pdf_path.read_bytes())

            docx_bytes = build_links_docx(plan["vendor"]["name"], plan["orders"])
            docx_upload = None
            if docx_bytes:
                docx_upload = upload_small(client, drive_id, date_folder["id"], plan["docx_name"], docx_bytes)

            links = {form_col: (pdf_upload.get("webUrl", ""), plan["pdf_name"])}
            if docx_upload:
                links[links_col] = (docx_upload.get("webUrl", ""), plan["docx_name"])
            sp.patch_url_text_fields(
                client,
                site_id,
                ctx.requisitions.id,
                plan["req"]["id"],
                links,
            )
            sp.patch_item_fields(
                client,
                site_id,
                ctx.requisitions.id,
                plan["req"]["id"],
                {status_col: PENDING_ADVISOR_APPROVAL},
            )
            completed += 1

    print(f"\nGenerated and linked {completed} requisition(s).")
