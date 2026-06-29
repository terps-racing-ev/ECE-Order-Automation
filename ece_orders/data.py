from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from . import sharepoint as sp
from .config import ORDER_FORM_LIST, REQUISITIONS_LIST, VENDORS_LIST
from .graph import GraphClient


@dataclass(frozen=True)
class ListsContext:
    """Resolved list IDs plus display-name-to-internal-name column maps."""

    order_form: sp.SharePointList
    requisitions: sp.SharePointList
    vendors: sp.SharePointList
    order_cols: sp.ColumnMap
    req_cols: sp.ColumnMap
    vendor_cols: sp.ColumnMap


def load_lists_context(client: GraphClient, site_id: str) -> ListsContext:
    """Resolve the three business lists and cache their column definitions."""
    order_form = sp.get_list_by_display_name(client, site_id, ORDER_FORM_LIST)
    requisitions = sp.get_list_by_display_name(client, site_id, REQUISITIONS_LIST)
    vendors = sp.get_list_by_display_name(client, site_id, VENDORS_LIST)
    return ListsContext(
        order_form=order_form,
        requisitions=requisitions,
        vendors=vendors,
        order_cols=sp.get_column_map(client, site_id, order_form),
        req_cols=sp.get_column_map(client, site_id, requisitions),
        vendor_cols=sp.get_column_map(client, site_id, vendors),
    )


def normalize_name(value: Any) -> str:
    return str(value or "").strip().casefold()


def field(fields: dict[str, Any], internal_name: str, default: Any = "") -> Any:
    return fields.get(internal_name, default)


def lookup_id_value(fields: dict[str, Any], internal_name: str) -> str:
    """Return the lookup ID SharePoint exposes beside a lookup display field."""
    return str(fields.get(f"{internal_name}LookupId") or fields.get(internal_name) or "").strip()


def load_vendors(client: GraphClient, site_id: str, ctx: ListsContext) -> dict[str, dict[str, Any]]:
    """Load approved vendors keyed by normalized title/name for fallback matching."""
    title_col = ctx.vendor_cols.title()
    address1 = ctx.vendor_cols.internal("Address 1")
    address2 = ctx.vendor_cols.internal("Address 2")
    phone = ctx.vendor_cols.internal("Phone")
    website = ctx.vendor_cols.internal("Website")
    vendors: dict[str, dict[str, Any]] = {}
    for item in sp.iter_items(client, site_id, ctx.vendors.id):
        fields = item.get("fields", {})
        name = str(field(fields, title_col, "")).strip()
        if not name:
            continue
        vendors.setdefault(normalize_name(name), {"matches": []})["matches"].append(
            {
                "id": item["id"],
                "name": name,
                "address1": field(fields, address1, ""),
                "address2": field(fields, address2, ""),
                "phone": field(fields, phone, ""),
                "website": field(fields, website, ""),
            }
        )
    return vendors


def load_vendors_by_id(client: GraphClient, site_id: str, ctx: ListsContext) -> dict[str, dict[str, Any]]:
    """Load approved vendors keyed by SharePoint item ID for lookup-column joins."""
    vendors_by_id: dict[str, dict[str, Any]] = {}
    for vendor_bucket in load_vendors(client, site_id, ctx).values():
        for vendor in vendor_bucket["matches"]:
            vendors_by_id[str(vendor["id"])] = vendor
    return vendors_by_id
