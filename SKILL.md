---
name: bl-check
description: >
  Checks a shipper/client Bill of Lading PDF against the Heung-A Line (HAL) DRAFT B/L PDF
  field by field before the real B/L is issued, and writes a colour-coded Excel report
  "BL CHECK <B/L NO>.xlsx". Compares all 27 items of the B/L check list (B/L no., shipper,
  consignee, notify party, pre-carriage, place of receipt, vessel/voy, POL, POD, place of
  delivery, final destination, container/seal, no. of pkgs, marks, description, gross
  weight, measurement, total in words, freight, ex rate, prepaid at, payable at, place/date
  of issue, total prepaid, no. of originals, date/laden on board, CY/CY) plus the RIDER /
  attached sheet — including fields written "AS PER ATTACHED RIDER / SEE ATTACHED SHEET /
  AS PER ATTACHED DESCRIPTIONS", which are looked up in the rider. Extra warnings: container
  check digit (ISO 6346), gross weight below N.W., container count vs total in words, B/L
  no. in file name vs document, on-board date after issue date.
  Use whenever the user asks to "check B/L", "เช็ค B/L", "ตรวจ B/L", "เทียบ B/L ลูกค้ากับ
  DRAFT", "CHECK DETAILS CONTENT ON B/L", or points to PDFs named like CLIENT <BLNO>.pdf and
  DRAFT HAL <BLNO>.pdf — even if only one of the two files is attached yet.
---

# CLIENT B/L × DRAFT HAL B/L → BL CHECK.xlsx

This packages a working, already-validated program (`scripts/bl_checker.py`, field boxes
in `scripts/layouts.json`) and an identical browser version (`docs/core.js`, live at
<https://rattanao.github.io/B-L-CHECK-REPORT-Export/>). Run the script; don't re-derive the
comparison by hand.

## When the user gives you files

1. Find the pair: the HAL file has `DRAFT` in its name, the client file does not; both names
   carry the B/L no. A separate rider PDF has `RIDER` (or `ATTACH`) in its name. Pages 2+
   of either B/L PDF are read as that side's rider automatically.
2. Run (needs `pip install pdfplumber openpyxl`):
   ```bash
   python scripts/bl_checker.py --cli "<client.pdf>" "<draft.pdf>"
   python scripts/bl_checker.py --cli "<client.pdf>" "<draft.pdf>" "<client rider.pdf>" "<draft rider.pdf>"
   python scripts/bl_checker.py --cli <folder> --outdir <folder>\output     # every pair in a folder
   ```
   It prints every field that does not match with the exact difference, the rider
   differences, and other warnings, and writes `BL CHECK <B/L NO>.xlsx` (sheet `B-L CHECK`,
   plus `RIDER CHECK` when a rider exists or "AS PER ATTACHED" is written).
3. Report to the user as a table of the unmatched fields (CLIENT value vs DRAFT value and
   what differs), then the warnings — lead with weight / container / DG items. Mention
   "MATCH (format differs)" items only briefly (e.g. `3 / THREE` vs `THREE(3)`).
4. If the report file is open in Excel, saving fails with PermissionError — ask the user to
   close it, or write to another name.

## Result meanings

| RESULT | Colour | Meaning |
|---|---|---|
| NOT MATCH | red | values differ — DETAILS says which words / how much |
| MISSING IN DRAFT / MISSING IN CLIENT | yellow | filled on one B/L, blank on the other |
| MATCH (via rider) | blue | one side says AS PER ATTACHED…, everything the other side prints was found in that rider |
| MATCH (format differs) | light green | same value, different spelling/spacing |
| MATCH | green | identical |
| BOTH EMPTY | grey | blank on both |

## How fields are read (when something looks wrong)

- Page 1 is read **by position**: each layout in `scripts/layouts.json` gives a box
  `[x0, top, x1, bottom]` in PDF points (A4 595 × 842, origin top-left) per field, and the
  cargo table is split into columns (`left` = container/seal + marks, `no_pkgs`,
  `description`, `gross_weight`, `measurement`). Text that overflows from DESCRIPTION into
  the weight column stays in DESCRIPTION; right-aligned package counts starting just left
  of a column line move right.
- Two layouts exist: `HAL DRAFT (system print)` (has tiny printed labels such as
  "Shipper", size < 6 pt, which are ignored) and `CLIENT (shipper print)`. The best-scoring
  layout is picked automatically; a warning "Form layout not recognised well" means a new
  form. To add one: dump words with
  `pdfplumber` (`page.extract_words(extra_attrs=["size"])`), copy a layout block in
  `layouts.json`, adjust the boxes, and **mirror the same layout in `docs/core.js`
  (`LAYOUTS`)** so the web version stays identical.
- Scanned PDFs (no text layer) are not supported — say so instead of guessing.

## Rider rules

- Riders are compared line by line (order-insensitive, line wrapping ignored), container
  by container, and by the sum of all KGS figures.
- A rider line also counts as found if the other B/L prints it on page 1.
- "AS PER ATTACHED" with no rider page → NOT MATCH + warning.

## Keeping the two versions in sync

`scripts/bl_checker.py` and `docs/core.js` implement the same rules. Any change to
matching logic or layouts must be made in both; verify by running both on the same PDFs
and comparing row statuses.
