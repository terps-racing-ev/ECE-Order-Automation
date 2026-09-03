from __future__ import annotations

import json
import time
from typing import Any

import msal
import requests

from .config import GRAPH_BASE, SCOPES, Settings


class GraphClient:
    """Small Microsoft Graph wrapper with device-code auth and retry handling."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self.token_cache = msal.SerializableTokenCache()
        if self.settings.sharepoint_token_cache_path.exists():
            self.token_cache.deserialize(self.settings.sharepoint_token_cache_path.read_text())
        self.app = msal.PublicClientApplication(
            client_id=self.settings.client_id,
            authority=f"https://login.microsoftonline.com/{self.settings.tenant_id}",
            token_cache=self.token_cache,
        )
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})

    def login_device_code(self) -> None:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        if not result:
            flow = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Failed to create device flow: {flow}")
            print("\n=== Microsoft sign-in ===")
            print(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Token acquisition failed: {result}")
        self._save_token_cache()
        self.session.headers["Authorization"] = f"Bearer {result['access_token']}"

    def _save_token_cache(self) -> None:
        if not self.token_cache.has_state_changed:
            return
        cache_path = self.settings.sharepoint_token_cache_path
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(self.token_cache.serialize())

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        # Graph occasionally throttles or flakes during workbook/list/file calls.
        # Keep retry policy here so the business logic stays readable.
        last_response: requests.Response | None = None
        for attempt in range(6):
            response = self.session.request(method, url, timeout=60, **kwargs)
            last_response = response
            if response.status_code in (429, 500, 502, 503, 504):
                delay = float(response.headers.get("Retry-After", 0)) or min(0.5 * (2**attempt), 8.0)
                time.sleep(delay)
                continue
            return response
        if last_response is None:
            raise RuntimeError(f"{method} {url} did not return a response")
        return last_response

    def get(self, url: str, **kwargs: Any) -> Any:
        response = self.request("GET", url, **kwargs)
        if not response.ok:
            raise RuntimeError(f"GET {url} -> {response.status_code}: {response.text}")
        return response.json()

    def post(self, url: str, payload: Any, headers: dict[str, str] | None = None) -> Any:
        request_headers = {"Content-Type": "application/json"}
        if headers:
            request_headers.update(headers)
        response = self.request("POST", url, data=json.dumps(payload), headers=request_headers)
        if not response.ok:
            raise RuntimeError(f"POST {url} -> {response.status_code}: {response.text}")
        return response.json()

    def patch(self, url: str, payload: Any, headers: dict[str, str] | None = None) -> Any:
        request_headers = {"Content-Type": "application/json"}
        if headers:
            request_headers.update(headers)
        response = self.request("PATCH", url, data=json.dumps(payload), headers=request_headers)
        if not response.ok:
            raise RuntimeError(f"PATCH {url} -> {response.status_code}: {response.text}")
        return response.json() if response.text else {}

    def put_bytes(self, url: str, content: bytes, content_type: str = "application/octet-stream") -> dict[str, Any]:
        response = self.request("PUT", url, data=content, headers={"Content-Type": content_type})
        if not response.ok:
            raise RuntimeError(f"PUT {url} -> {response.status_code}: {response.text}")
        return response.json()


def resolve_site(client: GraphClient, hostname: str, site_path: str) -> dict[str, Any]:
    """Return the Graph site object for the configured SharePoint team site."""
    return client.get(f"{GRAPH_BASE}/sites/{hostname}:/sites{site_path}")


def get_default_drive_id(client: GraphClient, site_id: str) -> str:
    """Return the default document-library drive used for generated uploads."""
    drive = client.get(f"{GRAPH_BASE}/sites/{site_id}/drive")
    return drive["id"]
