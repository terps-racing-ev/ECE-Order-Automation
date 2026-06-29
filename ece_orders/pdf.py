from __future__ import annotations

import random
from pathlib import Path


HEADER_FIELDS = {
    "account": "KFS Acct",
    "vendor": "Vendor",
    "street": "Street",
    "city_state_zip": "City State Zip",
    "phone": "Phone_2",
    "web": "Web Address",
    "date": "Date",
    "justification": "How is purchase related to University business project event unit 1",
    "subtotal": "Subtotal",
    "total": "Total Cost",
    "special1": "Special Instructions 1",
    "special2": "Special Instructions 2",
    "special3": "Special Instructions 3",
}

ROW_FIELDS = {
    "part": "Part Row{n}",
    "desc": "Detailed description including type of product eg monitor chemical name book microscope etcRow{n}",
    "qty": "QtyRow{n}",
    "unit": "Unit PriceRow{n}",
    "cost": "CostRow{n}",
}

JUSTIFICATIONS = [
    "Supports electric vehicle development, testing, and competition readiness under FSAE requirements",
    "Provides materials needed for safe and reliable progress on team fabrication and validation work",
    "Helps the team maintain project momentum while meeting competition, safety, and documentation needs",
    "Supports shop operations and subsystem development for the electric vehicle program",
    "Ensures required parts and supplies are available for FSAE design, build, and test activities",
]


def wrap_special(instr: str) -> list[str]:
    if not instr:
        return ["", "", ""]
    words = instr.split()
    lines: list[str] = []
    current = ""
    for word in words:
        if len(current) + (1 if current else 0) + len(word) > 80:
            lines.append(current)
            current = word
        else:
            current = word if not current else f"{current} {word}"
    if current:
        lines.append(current)
    return (lines + ["", "", ""])[:3]


def fill_order_pdf(
    template_path: Path,
    out_path: Path,
    account: str,
    vendor: str,
    vendor_info: dict[str, str],
    items: list[dict],
) -> None:
    """Fill the university order-form PDF template from one requisition group."""
    from pdfrw import PdfDict, PdfName, PdfObject, PdfReader, PdfWriter

    pdf = PdfReader(str(template_path))
    if getattr(pdf, "Root", None) and getattr(pdf.Root, "AcroForm", None):
        pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))

    subtotal = sum(float(item["unit_cost"]) * int(item["quantity"]) for item in items)

    def set_val(annot, value: str) -> None:
        # Clearing AP lets Acrobat/Preview regenerate visible field appearances.
        annot.update(PdfDict(V=str(value)))
        if getattr(annot, "AP", None) is not None:
            annot.AP = None

    for page in pdf.pages:
        if not getattr(page, "Annots", None):
            continue
        for annot in page.Annots:
            if getattr(annot, "Subtype", None) != PdfName.Widget:
                continue
            name = (annot.T or "").strip("()") if getattr(annot, "T", None) else ""
            if name == HEADER_FIELDS["account"]:
                set_val(annot, account)
            elif name == HEADER_FIELDS["vendor"]:
                set_val(annot, vendor)
            elif name == HEADER_FIELDS["street"]:
                set_val(annot, vendor_info.get("address1", ""))
            elif name == HEADER_FIELDS["city_state_zip"]:
                set_val(annot, vendor_info.get("address2", ""))
            elif name == HEADER_FIELDS["phone"]:
                set_val(annot, vendor_info.get("phone", ""))
            elif name == HEADER_FIELDS["web"]:
                set_val(annot, vendor_info.get("website", ""))
            elif name == HEADER_FIELDS["justification"]:
                set_val(annot, random.choice(JUSTIFICATIONS))
            elif name == HEADER_FIELDS["subtotal"]:
                set_val(annot, f"{subtotal:.2f}")
            elif name == HEADER_FIELDS["total"]:
                set_val(annot, f"{subtotal:.2f}")
            else:
                for row_num, item in enumerate(items, start=1):
                    mapping = {
                        ROW_FIELDS["part"].format(n=row_num): item.get("part_number", ""),
                        ROW_FIELDS["desc"].format(n=row_num): item.get("item", ""),
                        ROW_FIELDS["qty"].format(n=row_num): str(item.get("quantity", "")),
                        ROW_FIELDS["unit"].format(n=row_num): f"{float(item.get('unit_cost', 0.0)):.2f}",
                        ROW_FIELDS["cost"].format(n=row_num): f"{float(item.get('unit_cost', 0.0)) * int(item.get('quantity', 0)):.2f}",
                    }
                    if name in mapping:
                        set_val(annot, mapping[name])
                        break

    special = "; ".join(item.get("special_instructions", "").strip() for item in items if item.get("special_instructions"))
    special1, special2, special3 = wrap_special(special)
    for page in pdf.pages:
        if not getattr(page, "Annots", None):
            continue
        for annot in page.Annots:
            name = (annot.T or "").strip("()") if getattr(annot, "T", None) else ""
            if name == HEADER_FIELDS["special1"]:
                set_val(annot, special1)
            elif name == HEADER_FIELDS["special2"]:
                set_val(annot, special2)
            elif name == HEADER_FIELDS["special3"]:
                set_val(annot, special3)

    if getattr(pdf, "Root", None):
        if getattr(pdf.Root, "AcroForm", None):
            pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))
        pdf.Root.OpenAction = PdfDict(S=PdfName.JavaScript, JS=PdfObject("(this.calculateNow();)"))

    PdfWriter().write(str(out_path), pdf)
