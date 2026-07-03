# salvage-excel

Salvage data from messy Excel files to Parquet.

## Purpose

Standard Excel readers (pandas, R, etc.) can struggle with real-world `.xlsx` files:

- **Auto-converted input**: Excel silently turns some user-entered strings like `"+2"` into formulas, and readers may return the cached computed value (often an error code) instead of the intended string.
- **Large files**: Loading entire files into memory can make Excel and similar programs like LibreOffice Calc unresponsive, so correcting errors in Excel becomes hard.
- **Data corruption**: Complex formatting or unusual structures sometimes cause readers to fail or silently lose data.

`salvage-excel` bypasses standard readers by reading the internal XML directly. It converts via an intermediate CSV (data) + JSON (schema) representation, enabling manual inspection and correction.

## Installation

Requires Python ≥ 3.10.

**Option 1: Install with pipx**

```bash
pipx install git+https://github.com/allefeld/salvage-excel
```

**Option 2: Install with uv tool**

```bash
uv tool install git+https://github.com/allefeld/salvage-excel
```

**Option 3: Clone and sync**

```bash
git clone <repo>
cd salvage-excel
uv sync
```

Then use `uv run salvage-excel` from within the repository directory (this version is not added to `PATH`).

## Usage

```bash
# List sheets in an Excel file
salvage-excel data.xlsx

# Convert a sheet to Parquet
salvage-excel data.xlsx "Sheet1"

# Use cached cell values instead of formula text
salvage-excel data.xlsx "Sheet1" --prefer-cached

# Show raw XML parse events (for debugging)
salvage-excel data.xlsx "Sheet1" --debug
```

By default, when a cell contains a formula, the tool records the formula text (e.g. `=A1+B1`) rather than the cached computed value stored in the file. This avoids the auto-conversion problem described above. Use `--prefer-cached` when you actually want the computed values — for example, when formulas reference external workbooks that are no longer available.

The tool first creates two output files for each sheet

- `data-Sheet1.csv` — raw data extracted from Excel
- `data-Sheet1.json` — inferred schema and PyArrow options

and then converts these into

- `data-Sheet1.parquet` — final Parquet output

Output file names follow the pattern `<stem>-<sheet>.<ext>`, where `<stem>` is the input file name without extension and `<sheet>` is the sheet name as given on the command line.

## Workflow

1. Convert an Excel file's sheet.
2. Inspect output and intermediate CSV and schema for issues.
3. If necessary, correct intermediate CSV and schema.
4. Re-run the tool to regenerate the Parquet with your corrections.

The pipeline is idempotent: if the CSV already exists, the tool skips Excel parsing and goes straight to Parquet conversion, making it fast to iterate on schema fixes.

---

This software is copyrighted © 2026 by Carsten Allefeld and released under the terms of the GNU General Public License, version 3 or later.
