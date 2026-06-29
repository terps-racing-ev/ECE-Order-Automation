# Codebase Guide

This package automates the ECE order workflow against SharePoint Lists through Microsoft Graph. It has two business steps: assign approved order items to requisition list items, then generate/upload the PDF and links document for pending requisitions.

## Runtime Flow

`python -m ece_orders` enters through `ece_orders/__main__.py`, which calls `ece_orders.cli.main`.

The CLI does the shared setup once:

1. Load `.env` and defaults with `config.load_settings`.
2. Authenticate with `GraphClient.login_device_code`.
3. Resolve the SharePoint site with `graph.resolve_site`.
4. Resolve list IDs and column maps with `data.load_lists_context`.

Commands:

- `inspect-schema`: prints display names, internal names, types, lookup targets, and read-only flags for the three business lists.
- `read-item`: prints raw expanded Graph `fields` for one list item.
- `assign`: creates requisition list items from chief-approved unassigned order items.
- `generate`: creates PDFs/docs for requisitions marked `Pending Creation`.
- `run-all`: runs `assign` and then `generate` in one authenticated session.

## Important Modules

`ece_orders.config`

- Defines list display names, statuses, Graph scope, output folder defaults, and `Settings`.
- `load_settings()` is the only place that reads `.env`.

`ece_orders.graph`

- `GraphClient` wraps `requests.Session`, MSAL device-code auth, retry handling, and basic `get/post/patch/put_bytes` helpers.
- `resolve_site()` returns the configured SharePoint site object.
- `get_default_drive_id()` finds the document library used for generated uploads.

`ece_orders.sharepoint`

- `SharePointList` and `ColumnMap` hold list metadata.
- `ColumnMap.internal("Display Name")` resolves a display name to the internal field name Graph needs.
- `ColumnMap.lookup_id("Vendor")` returns lookup ID fields like `VendorLookupId`.
- `iter_items()`, `create_item()`, and `patch_item_fields()` are the generic list item helpers.
- `patch_url_text_fields()` writes generated file URLs to text columns on `ECE Requisitions`.

`ece_orders.data`

- `ListsContext` bundles the three resolved lists and their column maps.
- `load_vendors()` loads approved vendors by normalized name for fallback matching.
- `load_vendors_by_id()` loads approved vendors by SharePoint item ID, which is the normal lookup-column path.

`ece_orders.assign`

- `collect_assignable_orders()` reads `ECE Order Form` items where `Chief Approved` is true and `Req Form` is empty.
- `run_assign()` groups valid orders by approved-vendor lookup, chunks them by 10, creates `ECE Requisitions` items, and writes each order's `Req Form` lookup.

`ece_orders.generate`

- `collect_pending_requisitions()` finds requisitions with status `Pending Creation`.
- `collect_orders_for_req()` finds order items whose `Req Form` lookup points at a requisition.
- `run_generate()` creates the filled PDF, builds the links docx, uploads both files, writes their URLs to text fields, and moves the requisition to `Pending Advisor Approval`.

`ece_orders.pdf`

- `fill_order_pdf()` fills `Blank Order.pdf` using `pdfrw`.
- The PDF description field uses `ECE Order Form.Item`, not `Notes`.

`ece_orders.documents`

- `build_links_docx()` creates a Word document containing clickable purchase links.
- `extract_url()` accepts either text URLs or SharePoint/Graph URL objects.

`ece_orders.storage`

- `ensure_date_folder()` creates/reuses the SharePoint output folder for today.
- `upload_small()` uploads generated PDFs/docx files with Graph's simple upload endpoint.

## Field Rules

The code is configured with display names, then immediately resolves those to internal names at runtime. Runtime Graph writes use internal names.

Important expectations:

- `ECE Order Form.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Order Form.Req Form` is a lookup to `ECE Requisitions`.
- `ECE Requisitions.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Requisitions.Requisition Form` is a text column containing the uploaded PDF URL.
- `ECE Requisitions.Links Document` is a text column containing the uploaded docx URL.

Use `inspect-schema` after list changes. If a column display name changes, update the matching string in code or restore the SharePoint display name.

## Safety

`assign`, `generate`, and `run-all` are dry-run by default. Pass `--write` to mutate SharePoint.
