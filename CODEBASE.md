# Codebase Guide

This project is a mostly sequential pipeline. The normal command is:

```powershell
python -m ece_orders run-all
```

That runs assignment first, then generation, using the same Microsoft Graph login.

## Start Of The Program

Execution starts in `ece_orders/__main__.py`, which immediately calls `ece_orders.cli.main()`.

`cli.main()` does the shared setup for every command:

1. `config.load_settings()` reads `.env`, checks `CLIENT_ID` and `TENANT_ID`, and loads defaults like the SharePoint site and output folder.
2. `GraphClient(settings)` creates one reusable HTTP session for Microsoft Graph.
3. `client.login_device_code()` shows the Microsoft sign-in prompt and stores the Graph token on the session.
4. `graph.resolve_site()` resolves the configured SharePoint site and returns the site ID.
5. For write commands, `data.load_lists_context()` resolves the three SharePoint lists and builds column maps.

The column maps are important: the code asks for columns by display name, then uses the resolved internal names for Graph reads and writes.

## `run-all`

`run-all` is just the two business passes back to back:

```python
run_assign(client, site_id, ctx, dry_run=dry_run)
run_generate(client, site_id, ctx, settings, dry_run=dry_run)
```

Because both calls receive the same `GraphClient`, the user only signs in once.

## Assignment Pass

The assignment pass starts in `assign.run_assign()`.

First, it loads vendor reference data:

1. `data.load_vendors()` reads `ECE Approved Vendors` and builds a name-keyed vendor map.
2. `data.load_vendors_by_id()` builds an ID-keyed vendor map, which is the normal path because order vendors are lookup fields.

Then it calls `assign.collect_assignable_orders()`.

`collect_assignable_orders()` loops through `ECE Order Form` items using `sharepoint.iter_items()`. It keeps only rows where:

- `Chief Approved` is true.
- `Req Form` is empty.
- `Vendor` exists as either a lookup ID or fallback display value.

Back in `run_assign()`, each order is matched to an approved vendor. Lookup ID matching is preferred; name matching only exists as a fallback for messy data.

The valid orders are grouped by vendor and chunked by `MAX_ROWS_PER_PDF`, currently 10. For each chunk, `run_assign()` creates one `ECE Requisitions` item with `sharepoint.create_item()`. It writes:

- `Title`: internal requisition ID, like `Digikey 062826 A`
- `Vendor`: lookup ID to `ECE Approved Vendors`
- `Requisition Status`: `Pending Creation`
- `Date Created`: today

After creating the requisition, it patches each order's `Req Form` lookup with `sharepoint.patch_item_fields()`.

## Generation Pass

The generation pass starts in `generate.run_generate()`.

First, it checks that `Blank Order.pdf` exists. Then it calls `generate.collect_pending_requisitions()`, which reads `ECE Requisitions` and keeps only items whose `Requisition Status` is `Pending Creation`.

For each pending requisition, `run_generate()` calls `generate.collect_orders_for_req()`.

`collect_orders_for_req()` scans `ECE Order Form` and keeps order items whose `Req FormLookupId` matches the requisition item ID. It extracts the values needed for the PDF and links document:

- `Item`
- `Part Number`
- `Vendor`
- `Link`
- `Unit Cost`
- `Quantity`
- `Special Instructions`

Then `run_generate()` resolves vendor address/phone/website data from `ECE Approved Vendors`.

For each valid requisition plan, the actual file work happens in a temporary local folder:

1. `pdf.fill_order_pdf()` fills `Blank Order.pdf`.
2. `documents.build_links_docx()` creates the Word links document.
3. `graph.get_default_drive_id()` finds the SharePoint document library.
4. `storage.ensure_date_folder()` creates or reuses the output folder named from the requisition's `Date Created` value.
5. `storage.upload_small()` uploads the generated PDF and docx.
6. `sharepoint.patch_url_text_fields()` writes the uploaded file URLs into the requisition text fields.
7. `sharepoint.patch_item_fields()` writes `/Shared Documents/...` paths and changes `Requisition Status` to `Pending Advisor Approval`.

## Diagnostics

`inspect-schema` calls `sharepoint.print_schema()`. It prints each list's display names, internal names, types, lookup targets, hidden flags, and read-only flags.

`read-item` calls `sharepoint.get_item()` and prints the raw expanded Graph `fields` JSON for one list item. This is useful when SharePoint returns a surprising field shape.

## Field Rules

The code is configured with display names, then immediately resolves those to internal names at runtime. Runtime Graph writes use internal names.

Important expectations:

- `ECE Order Form.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Order Form.Req Form` is a lookup to `ECE Requisitions`.
- `ECE Order Form.Link` is a long text column. The links document extracts the first URL from it.
- `ECE Order Form.Special Instructions` is a long text column. The PDF joins item instructions with `; ` and wraps them across the three special-instruction fields.
- `ECE Requisitions.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Requisitions.Requisition Form` is a text column containing the uploaded PDF URL.
- `ECE Requisitions.Links Document` is a text column containing the uploaded docx URL.
- `ECE Requisitions.Requisition Form Path` is a text column containing the relative PDF path.
- `ECE Requisitions.Links Document Path` is a text column containing the relative docx path.

Use `inspect-schema` after list changes. If a column display name changes, update the matching string in code or restore the SharePoint display name.

## Safety

`assign`, `generate`, and `run-all` write by default. Pass `--dry-run` to preview without mutating SharePoint.
