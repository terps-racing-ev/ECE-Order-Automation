# ECE Order Automation

Automates the ECE purchasing flow using SharePoint Lists.

For normal use, run the full workflow:

```powershell
python -m ece_orders run-all --write
```

This signs in once, assigns approved orders to requisitions, generates the order PDFs and links documents, uploads them to SharePoint, writes the file URLs back to the requisition list, and moves generated requisitions to `Pending Advisor Approval`.

## Setup

1. Install Python 3.10+.
2. Install the package dependencies from this repo:

   ```powershell
   python -m pip install -e .
   ```

3. Create `.env` from `.env.example` and fill in:

   ```text
   CLIENT_ID=
   TENANT_ID=
   ```

4. Make sure `Blank Order.pdf` is present in the repo root.

The default SharePoint site and output folder are already configured in `.env.example`.

## Normal Commands

Run everything:

```powershell
python -m ece_orders run-all --write
```

Only assign approved orders to requisitions:

```powershell
python -m ece_orders assign --write
```

Only generate PDFs/docs for pending requisitions:

```powershell
python -m ece_orders generate --write
```

## SharePoint Requirements

The script expects these lists:

- `ECE Order Form`
- `ECE Requisitions`
- `ECE Approved Vendors`

Important column assumptions:

- `ECE Order Form.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Order Form.Req Form` is a lookup to `ECE Requisitions`.
- `ECE Requisitions.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Requisitions.Requisition Form` is a text column for the generated PDF URL.
- `ECE Requisitions.Links Document` is a text column for the generated links document URL.

## What It Does

Step 1, assignment:

- Finds order items where `Chief Approved` is true and `Req Form` is empty.
- Groups them by approved vendor, 10 items per requisition.
- Creates `ECE Requisitions` items with status `Pending Creation`.
- Writes each order item's `Req Form` lookup.

Step 2, generation:

- Finds requisitions with status `Pending Creation`.
- Generates the PDF order form and links `.docx`.
- Uploads both files to the configured SharePoint order forms folder.
- Writes the uploaded file URLs to the requisition item.
- Changes status to `Pending Advisor Approval`.

## Admin And Debug Commands

Preview commands without writing:

```powershell
python -m ece_orders run-all
python -m ece_orders assign
python -m ece_orders generate
```

Inspect SharePoint column internal names:

```powershell
python -m ece_orders inspect-schema
```

Read one raw SharePoint list item:

```powershell
python -m ece_orders read-item --list "ECE Requisitions" --item-id 2
```

For implementation details, see [CODEBASE.md](CODEBASE.md).
