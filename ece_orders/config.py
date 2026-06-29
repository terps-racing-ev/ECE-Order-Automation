from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Selected"]

ORDER_FORM_LIST = "ECE Order Form"
REQUISITIONS_LIST = "ECE Requisitions"
VENDORS_LIST = "ECE Approved Vendors"

PENDING_CREATION = "Pending Creation"
PENDING_ADVISOR_APPROVAL = "Pending Advisor Approval"

MAX_ROWS_PER_PDF = 10


@dataclass(frozen=True)
class Settings:
    client_id: str
    tenant_id: str
    site_hostname: str
    site_path: str
    output_parent_sp_path: str
    blank_order_pdf: Path


def load_settings() -> Settings:
    load_dotenv()
    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID")
    if not client_id or not tenant_id:
        raise RuntimeError("Missing CLIENT_ID or TENANT_ID in environment/.env")

    return Settings(
        client_id=client_id,
        tenant_id=tenant_id,
        site_hostname=os.getenv("SITE_HOSTNAME", "umd0.sharepoint.com"),
        site_path=os.getenv("SITE_PATH", "/TeamsTerpsRacingEV"),
        output_parent_sp_path=os.getenv("OUTPUT_PARENT_SP_PATH", "/General/_EV26/Finance/ECE/Order Forms"),
        blank_order_pdf=Path(__file__).resolve().parent.parent / "Blank Order.pdf",
    )
