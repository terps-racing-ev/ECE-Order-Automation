from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable
from urllib.parse import quote, urlsplit, urlunsplit

from .config import GRAPH_BASE
from .graph import GraphClient


@dataclass(frozen=True)
class SharePointList:
    id: str
    name: str
    display_name: str


@dataclass(frozen=True)
class ColumnMap:
    """Resolve friendly display names to the internal field names Graph writes."""

    columns: dict[str, dict[str, Any]]

    @classmethod
    def from_columns(cls, columns: Iterable[dict[str, Any]]) -> "ColumnMap":
        by_display = {}
        for column in columns:
            display = column.get("displayName") or column.get("name")
            if display:
                by_display[display] = column
            internal = column.get("name")
            if internal:
                by_display.setdefault(internal, column)
        return cls(by_display)

    def internal(self, display_name: str) -> str:
        column = self.columns.get(display_name)
        if not column:
            available = ", ".join(sorted(self.columns))
            raise KeyError(f"Could not find column display name '{display_name}'. Available: {available}")
        return column["name"]

    def lookup_id(self, display_name: str) -> str:
        return f"{self.internal(display_name)}LookupId"

    def title(self) -> str:
        # SharePoint renamed Title columns often appear alongside read-only
        # computed columns like LinkTitleNoMenu. Writes must target Title.
        return "Title"


def list_site_lists(client: GraphClient, site_id: str) -> list[SharePointList]:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,name,displayName"
    data = client.get(url)
    return [
        SharePointList(id=item["id"], name=item.get("name", ""), display_name=item.get("displayName", item.get("name", "")))
        for item in data.get("value", [])
    ]


def get_list_by_display_name(client: GraphClient, site_id: str, display_name: str) -> SharePointList:
    matches = [item for item in list_site_lists(client, site_id) if item.display_name == display_name or item.name == display_name]
    if not matches:
        known = ", ".join(item.display_name for item in list_site_lists(client, site_id))
        raise RuntimeError(f"List '{display_name}' not found. Known lists: {known}")
    if len(matches) > 1:
        ids = ", ".join(item.id for item in matches)
        raise RuntimeError(f"List '{display_name}' matched multiple lists: {ids}")
    return matches[0]


def get_columns(client: GraphClient, site_id: str, list_id: str) -> list[dict[str, Any]]:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/columns?$top=200"
    return client.get(url).get("value", [])


def get_column_map(client: GraphClient, site_id: str, sp_list: SharePointList) -> ColumnMap:
    return ColumnMap.from_columns(get_columns(client, site_id, sp_list.id))


def print_schema(client: GraphClient, site_id: str, list_names: Iterable[str]) -> None:
    """Print enough column metadata to debug renamed fields and lookup columns."""
    for list_name in list_names:
        sp_list = get_list_by_display_name(client, site_id, list_name)
        print(f"\n=== {sp_list.display_name} ===")
        print(f"id={sp_list.id} name={sp_list.name}")
        for col in get_columns(client, site_id, sp_list.id):
            column_type = next(
                (key for key in ("text", "choice", "boolean", "dateTime", "number", "currency", "lookup", "hyperlinkOrPicture") if key in col),
                "unknown",
            )
            extra = ""
            if "lookup" in col:
                lookup = col["lookup"]
                extra = f" lookupList={lookup.get('listId', '')} lookupColumn={lookup.get('columnName', '')}"
            elif "choice" in col:
                extra = f" choices={col['choice'].get('choices', [])}"
            print(
                f"{col.get('displayName', ''):<35} "
                f"internal={col.get('name', ''):<35} "
                f"type={column_type:<18} "
                f"hidden={col.get('hidden', False)} "
                f"readOnly={col.get('readOnly', False)}"
                f"{extra}"
            )


def iter_items(client: GraphClient, site_id: str, list_id: str, select_fields: list[str] | None = None) -> list[dict[str, Any]]:
    """Return all list items with expanded `fields`, following Graph pagination."""
    expand = "fields"
    if select_fields:
        encoded_fields = ",".join(select_fields)
        expand = f"fields(select={encoded_fields})"
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?$expand={quote(expand, safe='=(),')}&$top=200"
    items: list[dict[str, Any]] = []
    while url:
        data = client.get(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def get_item(client: GraphClient, site_id: str, list_id: str, item_id: str) -> dict[str, Any]:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}?$expand=fields"
    return client.get(url)


def create_item(client: GraphClient, site_id: str, list_id: str, fields: dict[str, Any]) -> dict[str, Any]:
    return client.post(f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items", {"fields": fields})


def patch_item_fields(client: GraphClient, site_id: str, list_id: str, item_id: str, fields: dict[str, Any]) -> dict[str, Any]:
    return client.patch(f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields", fields)


def patch_url_text_fields(
    client: GraphClient,
    site_id: str,
    list_id: str,
    item_id: str,
    links: dict[str, tuple[str, str]],
) -> None:
    """Write generated file URLs to text columns on a requisition item."""
    payload = {field: normalize_hyperlink_url(url) for field, (url, _description) in links.items() if url}
    if payload:
        patch_item_fields(client, site_id, list_id, item_id, payload)


def normalize_hyperlink_url(url: str) -> str:
    parts = urlsplit(url)
    if not parts.scheme or not parts.netloc:
        return url
    path = quote(parts.path, safe="/%")
    query = quote(parts.query, safe="=&%?/:;+,$")
    fragment = quote(parts.fragment, safe="")
    return urlunsplit((parts.scheme, parts.netloc, path, query, fragment))
