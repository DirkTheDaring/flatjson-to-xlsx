# xlsx_from_json(1) — write/update XLSX from JSON with ordering, filters, PK merge, and hyperlinks

## NAME
**xlsx_from_json** — convert JSON (array, object, or NDJSON) into an Excel `.xlsx` sheet, preserving existing formatting, merging rows by primary keys, filtering/ordering columns, and adding per‑column hyperlinks.

> Binary version: **0.7.0**  
> (This README documents the CLI implemented in the provided Rust source.)

## SYNOPSIS
```sh
xlsx_from_json --out OUT.xlsx [--sheet Sheet1] [--pk col1,col2,...] \
               [--array | --ndjson] \
               [--include name1,name2,...] [--include-regex r1,r2,...] [--include-substr s1,s2,...] \
               [--order n1,n2,...] [--order-regex r1,r2,...] [--order-substr s1,s2,...] [--order-rest existing|alpha|none] \
               [--pk-first | --no-pk-first] [--link col=BASE[,col2=BASE2,...]] [--config file.toml] < input.json
```

## DESCRIPTION
This tool reads JSON from **stdin** and writes it into an Excel workbook (`.xlsx`). If the output file already exists, it **preserves existing formatting/styles** and **updates the data in place**.

Input can be:
- A **JSON array** of objects (preferred)
- A single **JSON object** (treated as one row)
- **NDJSON** (one JSON object per line) when `--ndjson` is supplied

Each input object is expected to be **already flattened** (key → scalar), e.g. `{ "a.b": 1 }`. The program does not flatten nested JSON; it maps keys to columns and values to cells.

Major features:
- **Column inclusion** (exact/regex/substring)
- **Column ordering** (exact/regex/substring groups + remainder policy)
- **Primary key (PK) merge** into an existing workbook
- **Per‑column hyperlinks** via Excel `HYPERLINK()` formulas
- **Formatting preservation** (headers/data) when updating existing files
- **Natural sort** for leftover columns (e.g. `c.2` < `c.10`)

## INPUT MODES
- `--array`  
  Force parsing as a single JSON array from stdin.
- `--ndjson`  
  Parse stdin as NDJSON (one JSON object per line).  
  **Auto‑override:** if `--ndjson` is set but the payload **starts with `[`**, the program assumes it’s a JSON array and switches to array mode (emits a note to stderr).
- Default (no flag): the tool tries to parse stdin as a JSON array; if it’s an object, it becomes a single row.

## CONFIG FILE
You can provide a TOML config with `--config file.toml` (or `-c file.toml`). CLI flags override the config. Example:

```toml
# file: export.toml
out = "report.xlsx"
sheet = "Data"
ndjson = false
pk = ["id", "subid"]

# Filters (columns to include)
include = ["id", "name"]
include_regex = ["^meta\\..+$"]
include_substr = ["_score"]

# Ordering
order = ["id","subid","name"]
order_regex = ["^meta\\..+$"]
order_substr = ["_score"]
order_rest = "alpha" # "existing" | "alpha" | "none"

# PK positioning
pk_first = true

# Per-column hyperlink bases
[hyperlink]
ticket = "https://tracker.local/browse/"
doc_id = "https://docs.local/view?id="
```

## OPTIONS
```
--out, -o <FILE.xlsx>
    Output workbook path. Required. Must end with .xlsx (exit code 2 if not).

--sheet, -s <NAME>
    Sheet name to write/update (default: "Sheet1").
    If the workbook exists and the sheet exists, it is updated in place.
    Otherwise the sheet is created (or "Sheet1" is renamed).

--array
    Treat input as a single JSON array.

--ndjson
    Treat input as NDJSON. If input begins with '[', switches to array mode.

--pk, -k col1,col2,...
    Primary key columns. When OUT.xlsx already has rows, rows are merged by
    composite key (concat of PK values). Rows with missing PK values are appended.

--pk-first / --no-pk-first
    Whether PK columns are forced to the front (default: true; config: pk_first).

--include, -i name1,name2,...
--include-regex r1,r2,...
--include-substr s1,s2,...
    Column inclusion filters. If any inclusion list is present, inclusion turns
    ACTIVE and only matching columns (plus all PKs) are kept.

--order n1,n2,...
--order-regex r1,r2,...
--order-substr s1,s2,...
--order-rest existing|alpha|none
    Column ordering controls. Final order is:
      1) PKs (if pk_first)
      2) Exact names in --order
      3) Names matched by --order-regex (in discovery order)
      4) Names containing any --order-substr entries (in discovery order)
      5) Remainder: "existing" (keep workbook/header order), "alpha"
         (natural sort), or "none" (omit remainder).

--link col=BASE[,col2=BASE2,...]
    Per-column hyperlink base URLs. When a cell has a non-empty value `v` in
    one of these columns, the cell is set to:
        HYPERLINK("<BASE><v>", "<v>")
    so the display shows just `v` but is clickable.

--config, -c file.toml
    Read defaults from a TOML config (fields mirror this README).

-h, --help
    Show usage help.

-V, --version
    Show program version and exit.
```

## MERGE BY PRIMARY KEY (PK)
If `--pk` is provided, the tool builds an index of existing rows in the target sheet using the **composite PK** (all PK column values joined — internal delimiter, not visible in Excel).  
For every input row:
- If the composite PK is **present** and **found** in existing data, that row is **updated** (replaced).
- If the composite PK is **present** but **not found**, the row is **appended**.
- If any PK value is **missing**, the row is **appended** (no merge).

## COLUMN UNIVERSE & ORDERING
1. Start from **existing headers** (non-empty), excluding PKs.
2. Add **all keys** discovered in input rows.
3. Apply **inclusion** (if active).
4. Build final order:
   - PKs first (when `pk_first=true`)
   - `--order` exact names (deduped)
   - `--order-regex` matches (in the order keys are discovered)
   - `--order-substr` matches (in the order keys are discovered)
   - Remainder via `--order-rest`:
     - `existing` — keep workbook/header + discovery order
     - `alpha` — natural sort (`c.2` < `c.10`)
     - `none` — drop leftover columns
5. If `pk_first=false`, ensure PK columns appear somewhere (append if missing).

## FORMATTING & WRITING
- Existing workbook is opened with **umya-spreadsheet** and **styles are preserved**.
- Headers are written to row 1; data begin at row 2.
- Hyperlink columns are written as `HYPERLINK()` formulas; non-link values are written with their native types when possible (bool, number, string).

## EXAMPLES

### Write a fresh workbook from a JSON array
```sh
cat data.json | xlsx_from_json --out report.xlsx --sheet Data
```

### Update existing workbook, merging by PK, keeping PKs first
```sh
cat rows.ndjson | xlsx_from_json --out report.xlsx --sheet Data --ndjson \
  --pk id,subid --pk-first
```

### Include only selected columns and meta.* (regex), sort the rest alphabetically
```sh
cat data.json | xlsx_from_json --out report.xlsx --sheet Data \
  --include id,name --include-regex '^meta\\..+$' \
  --order id,name --order-rest alpha
```

### Add clickable links for ticket and doc_id columns
```sh
cat data.json | xlsx_from_json --out report.xlsx --sheet Data \
  --link ticket=https://tracker.local/browse/,doc_id=https://docs.local/view?id=
```

### Use a TOML config and override sheet on CLI
```sh
cat data.json | xlsx_from_json -c export.toml --sheet Latest
```

## EXIT STATUS
- `0` on success
- `2` when `--out` does not end with `.xlsx`
- Non-zero on IO/parse/config errors (propagated from libraries)

## BUILDING
**Dependencies (Cargo):** `calamine`, `umya-spreadsheet`, `regex`, `serde`, `serde_json`, `toml`.

`Cargo.toml` snippet:
```toml
[package]
name = "xlsx_from_json"
version = "0.7.0"
edition = "2021"

[dependencies]
calamine = "0.25"
umya-spreadsheet = "1"
regex = "1"
serde = { version = "1", features = ["derive"] }
serde_json = "1"
toml = "0.8"
```

**Build & Run**
```sh
cargo build --release
cat input.json | target/release/xlsx_from_json --out OUT.xlsx [options]
```

## NOTES
- If any inclusion list is specified, **inclusion mode** activates and only matching columns (plus PKs) are kept.
- When `--ndjson` is used but the input begins with `[` (array), the tool switches to array mode and logs a **note** on stderr.
- Empty header cells in an existing workbook are ignored.
- Rows that are completely empty (all values null/empty) are skipped on readback.
- Numbers are written as Excel numbers when representable; otherwise as strings.

## VERSION
`xlsx_from_json` **0.7.0**
