# Example — Multi-CSV Sheet Aggregator

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the folder. |
| `monthly_exports/` | The folder you point the tool at. Contains five files (below). |

Inside `monthly_exports/`:

| File | Delimiter | Why it's here |
|---|---|---|
| `jan.csv` | comma | ordinary monthly export |
| `feb.csv` | comma | ordinary monthly export |
| `mar.csv` | comma | ordinary monthly export |
| `europe_q1.csv` | **semicolon** | European CSV — proves per-file auto-detect (a comma split would mangle it) |
| `notes.txt` | **tab** | shows `.txt` files import too, and tab is detected |

## How to try it

1. Open **any** workbook in Excel (even a blank one) and **save it once**
   — the default **Backup workbook first** needs a folder for the backup.
   The new sheets are added to whatever workbook is active.
2. Run **Multi-CSV Sheet Aggregator** from ScripTree.

Form values:

| Field | Value |
|---|---|
| Source folder | *(browse to this `monthly_exports` folder)* |
| Delimiter | `Auto-detect (recommended)` |
| Backup workbook first | *(checked)* |
| Output format | `markdown` |

**Expected result** — **5** new sheets added to the open workbook, one
per file, each named after its source file (sanitised, ≤31 chars,
de-duplicated):

```
jan        — Product/Units/Revenue, 3 rows
feb        — 3 rows
mar        — 3 rows
europe_q1  — Produkt/Einheiten/Umsatz, 3 rows  (parsed on ';', NOT one mashed column)
notes      — Region/Manager/Target, 3 rows     (parsed on tab)
```

A `<name>_Backup` copy is written to disk first; the workbook is then left
**open and unsaved** for you to review and save.

### See auto-detect matter

Re-run with **Delimiter** forced to `Comma ( , )`. Now `europe_q1.csv`
imports as a **single** column (`Produkt;Einheiten;Umsatz` all in column
A) because there are no commas to split on — exactly the mangling that
per-file auto-detect avoids.

## What this demonstrates

* Merging a folder of files into one workbook, a sheet per file (the
  inverse of the Segregator).
* Per-file delimiter auto-detection (comma / semicolon / tab).
* `.csv` **and** `.txt` are both imported; subfolders are not scanned.
* Source files are only read; the backup + leave-open-unsaved safety model.
