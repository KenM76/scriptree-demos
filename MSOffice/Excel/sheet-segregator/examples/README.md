# Example — Column-Based Sheet Segregator

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the sample. |
| `sample_sales.xlsx` | One flat table on a sheet named **Sales**: `Date, Region, Product, Units, Revenue` with a header row and 12 data rows. |

The **Region** column has four distinct values, interleaved (not grouped):
East ×4, West ×3, North ×3, South ×2.

## How to try it

1. Open `sample_sales.xlsx` in Excel (the **Sales** sheet is active).
2. Run **Column-Based Sheet Segregator** from ScripTree.

Form values:

| Field | Value |
|---|---|
| Key column | `Region` |
| First row is a header | *(checked)* |
| Where to put the sheets | `New workbook (recommended)` |
| Sheet name prefix | *(leave blank)* |
| Backup workbook first | *(ignored in New-workbook mode)* |
| Max distinct values | `50` |
| Output format | `markdown` |

**Expected result** — a brand-new workbook (your `sample_sales.xlsx` is
left untouched) containing **4** sheets:

```
East   — header + 4 rows
West   — header + 3 rows
North  — header + 3 rows
South  — header + 2 rows
```

The result workbook is left **open and unsaved** so you can review and
save (or discard) it yourself.

### Variations to try

* **Key column** = `B` (column letter) or `2` (1-based position) gives the
  identical result — three ways to name the same column.
* **Sheet name prefix** = `R-` names the sheets `R-East`, `R-West`, …
* **Where to put the sheets** = `Add sheets to current workbook` appends
  the four sheets after Sales in the open workbook. In that mode tick
  **Backup workbook first** to get a `sample_sales_backup` copy on disk
  before any sheet is added (so save the file once first).
* The **Max distinct values** cap: if you point the key column at
  something near-unique (e.g. `Date`, 12 distinct values is fine, but
  imagine an ID column), the tool refuses rather than spawning a sheet per
  row. Lower the cap to `3` and run on Region to see the refusal.

## What this demonstrates

* Gathering scattered rows by a key column into one sheet per value.
* Header replication onto every output sheet.
* The safe New-workbook default (original never touched) vs. the
  Add-sheets mode with its backup guard.
* The distinct-value safety cap.
