# ECE Order Automation

Automates the ECE purchasing flow using SharePoint Lists.

For normal use, run the full workflow:

```powershell
python -m ece_orders run-all
```

This signs in once, assigns approved orders to requisitions, generates the order PDFs and links documents, uploads them to SharePoint, writes the file URLs back to the requisition list, and moves generated requisitions to `Req Form Created`.

Installing the package also adds an `ece-orders` command, so `ece-orders run-all` works the same as `python -m ece_orders run-all`.

## Setup

1. Install Python 3.10+.
2. Install the package dependencies from this repo:

   ```powershell
   python -m pip install -e .
   ```

3. Copy `.env.example` to `.env` and fill in:

   ```text
   CLIENT_ID=
   TENANT_ID=
   ```

   The SharePoint site and output folder defaults are built into `ece_orders/config.py`. To override them, set any of these in `.env`:

   ```text
   SITE_HOSTNAME=umd0.sharepoint.com
   SITE_PATH=/TeamsTerpsRacingEV
   OUTPUT_PARENT_SP_PATH=/General/_EV26/Finance/ECE/Order Forms
   ```

   The Microsoft Graph token cache is stored at `env/sharepoint_token_cache.json` by default after the first successful sign-in. To use a different location, set:

   ```text
   SHAREPOINT_TOKEN_CACHE_PATH=
   ```

4. Create `env/jlc_pcb.env` for the JLC PCB links-document password:

   ```text
   JLC_PCB_PASSWORD=
   ```

   This is only used when a requisition's vendor is named exactly `JLC PCB` (case and extra spaces ignored). The links document for that vendor gets the team login email and this password at the top. If the file or variable is missing, the document says `Missing JLC_PCB_PASSWORD` instead.

5. Make sure `Blank Order.pdf` is present in the repo root.

## Normal Commands

Run everything:

```powershell
python -m ece_orders run-all
```

Only assign approved orders to requisitions:

```powershell
python -m ece_orders assign
```

Only generate PDFs/docs for pending requisitions:

```powershell
python -m ece_orders generate
```

## SharePoint Requirements

The script expects these lists:

- `ECE Order Form`
- `ECE Requisitions`
- `ECE Approved Vendors`
- `ECE Order Batches`

Important column assumptions:

- `ECE Order Form.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Order Form.Req Form` is a lookup to `ECE Requisitions`.
- `ECE Order Form.Link` is a long text column containing the item URL.
- `ECE Order Form.Special Instructions` is a long text column.
- `ECE Requisitions.Vendor` is a lookup to `ECE Approved Vendors`.
- `ECE Requisitions.Batch` is a lookup to `ECE Order Batches`.
- `ECE Requisitions.Requisition Form` is a text column for the generated PDF URL.
- `ECE Requisitions.Links Document` is a text column for the generated links document URL.
- `ECE Order Batches.Title` is the batch ID, displayed in SharePoint as `Batch ID`, like `071626`.

## What It Does

Step 1, assignment:

- Finds order items where `Chief Approved` is true and `Req Form` is empty.
- Matches each order to `ECE Approved Vendors` by its Vendor lookup ID, falling back to the vendor name.
- Groups them by approved vendor, 10 items per requisition.
- Creates or reuses today's `ECE Order Batches` item, with Batch ID `MMDDYY`.
- Creates `ECE Requisitions` items linked to that batch with status `Pending Creation` and `Date Created` set to today.
- Names each requisition `{Vendor} MMDDYY {letter}`, like `Digikey 071626 A`. A vendor with more than 10 orders gets `B`, `C`, and so on.
- Writes each order item's `Req Form` lookup.

Orders are skipped and left unassigned if they have no vendor, or if the vendor is missing from `ECE Approved Vendors` or matches more than one entry there. Skipped orders are listed in the output and picked up on a later run once fixed.

Step 2, generation:

- Finds requisitions with status `Pending Creation`.
- Collects the order items whose `Req Form` points at that requisition.
- Generates the PDF order form, filled with the vendor's address, phone and website from `ECE Approved Vendors`, the order rows, the totals, the special instructions and a randomly chosen justification line.
- Generates the links `.docx` from the orders' `Link` values. If none of the orders has a usable URL, no links document is created and `Links Document` is left as it was.
- Uploads the files to `{OUTPUT_PARENT_SP_PATH}/MMDDYY`, where `MMDDYY` comes from the requisition's `Date Created`. The folder is created if needed. If it can't be created, the script tries `MMDDYY 1`, `MMDDYY 2`, and so on, up to `MMDDYY 9`.
- Names the files `{requisition title}.pdf` and `{requisition title} Links.docx`. Characters that aren't allowed in file names are replaced with `_`.
- Writes the uploaded file URLs to the requisition item.
- Changes status to `Req Form Created`.

Requisitions are skipped and left at `Pending Creation` if they have no linked order items, if their vendor can't be resolved, or if `Date Created` is empty or invalid.

## Regenerating A Requisition

To regenerate a requisition, set its `Requisition Status` back to `Pending Creation` and run `generate` or `run-all`. Regeneration only redoes step 2, so this works the same whether it's the same day or a few days later:

- **Same identity.** The requisition item, its title, `Batch`, `Vendor` and `Date Created` are not changed. It stays in its original batch and does not move to today's.
- **Same location and file names.** The files go to the folder for the original `Date Created`, under the same names. The new upload replaces the old file there (SharePoint keeps the old one in version history), so the URLs in `Requisition Form` and `Links Document` stay the same.
- **Current contents.** The PDF and links document are rebuilt from the order items that are linked to the requisition now, and the vendor details as they are now. Edits to quantities, costs, links or special instructions are picked up, and so are orders you manually linked or unlinked. The justification line is picked at random again, so it may change.
- **New approved orders are not added.** `assign` only puts unassigned orders on new requisitions. To add an order to an existing requisition, set that order's `Req Form` yourself before regenerating, and keep the total at 10 or fewer items, since the PDF has 10 rows.

Edge cases:

- If you rename the requisition, the new files get the new name and the old files stay in the folder.
- If you change `Date Created`, the files go to the folder for the new date and the old files stay in the old folder.
- If the orders no longer have any links, no new links document is made. The old one stays in the folder and `Links Document` still points at it.

## Admin And Debug Commands

Preview commands without writing:

```powershell
python -m ece_orders run-all --dry-run
python -m ece_orders assign --dry-run
python -m ece_orders generate --dry-run
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
