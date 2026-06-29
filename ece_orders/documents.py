from __future__ import annotations

import io
from typing import Any

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def extract_url(value: Any) -> str:
    """Accept either a plain text URL or Graph's object shape for URL columns."""
    if isinstance(value, dict):
        return str(value.get("Url") or value.get("url") or value.get("Description") or value.get("description") or "").strip()
    return str(value or "").strip()


def add_hyperlink(paragraph, url: str, text: str) -> None:
    """python-docx has no public hyperlink helper, so build the Word XML node."""
    part = paragraph.part
    relationship_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)

    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    props.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.append(underline)

    run.append(props)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def build_links_docx(vendor: str, items: list[dict]) -> bytes | None:
    """Build the vendor links document in memory for upload."""
    links = [(item.get("item", "Link"), extract_url(item.get("link", ""))) for item in items if extract_url(item.get("link", ""))]
    if not links:
        return None

    doc = Document()
    doc.add_paragraph(vendor, style="Title")
    for label, link in links:
        paragraph = doc.add_paragraph(style="List Bullet")
        add_hyperlink(paragraph, link, str(label or link))

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()
