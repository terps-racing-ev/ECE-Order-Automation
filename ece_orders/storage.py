from __future__ import annotations

import datetime as dt
from typing import Any

from .config import GRAPH_BASE
from .graph import GraphClient


def get_item_by_path(client: GraphClient, drive_id: str, path: str) -> dict[str, Any]:
    return client.get(f"{GRAPH_BASE}/drives/{drive_id}/root:{path}")


def ensure_folder(client: GraphClient, drive_id: str, parent_path: str, name: str) -> dict[str, Any]:
    parent = get_item_by_path(client, drive_id, parent_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent['id']}/children"
    payload = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    response = client.request("POST", url, headers={"Content-Type": "application/json"}, json=payload)
    if response.status_code == 409:
        return get_item_by_path(client, drive_id, f"{parent_path}/{name}")
    if not response.ok:
        raise RuntimeError(f"Create folder failed: {response.status_code} {response.text}")
    return response.json()


def date_folder_name(value: Any) -> str:
    """Return the MMDDYY SharePoint folder name for a date-like value."""
    if isinstance(value, dt.datetime):
        date_value = value.date()
    elif isinstance(value, dt.date):
        date_value = value
    else:
        text = str(value or "").strip()
        if not text:
            raise ValueError("Date Created is empty")
        try:
            date_value = dt.date.fromisoformat(text[:10])
        except ValueError as exc:
            raise ValueError(f"Date Created is not an ISO date: {value!r}") from exc
    return date_value.strftime("%m%d%y")


def ensure_date_folder(client: GraphClient, drive_id: str, parent_path: str, date_created: Any) -> dict[str, Any]:
    """Create or reuse the output folder for a requisition's Date Created value."""
    base = date_folder_name(date_created)
    for attempt in range(10):
        name = base if attempt == 0 else f"{base} {attempt}"
        try:
            return ensure_folder(client, drive_id, parent_path, name)
        except RuntimeError:
            continue
    raise RuntimeError("Unable to create/find date folder after 10 attempts")


def upload_small(client: GraphClient, drive_id: str, dest_folder_id: str, filename: str, data: bytes) -> dict[str, Any]:
    """Upload small generated artifacts with Graph's simple upload endpoint."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{dest_folder_id}:/{filename}:/content"
    return client.put_bytes(url, data)
