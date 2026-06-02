# Column-Based Sheet Segregator (Excel)

Split the flat table on the **active Excel worksheet** into separate sheets —
**one sheet per distinct value** in a chosen "key" column. Classic example: a
sales export with a `Region` column → one sheet per region, each holding only
that region's rows (plus a copy of the header row).

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

You have a table on a worksheet — rows of data, usually with a header row in
row 1. You pick a column. The tool produces one new sheet for every distinct
value found in that column, each sheet containing the header (if you have one)
plus exactly the rows whose key-column cell holds that value.

### The form fields

The form groups its fields into two collapsible sections — **Key column** and **Output**.

| Section | Field | Meaning |
|---|---|---|
| Key column | **Key column** (required) | Which column to split on. Accepts a **header name** (e.g. `Region` — only when "First row is a header" is on), a **column letter** (e.g. `C`), or a **1-based position within the used range** (e.g. `3` = third column of the table). |
| Key column | **First row is a header** (default ON) | When on, row 1 is treated as column headings: it is copied to the top of every output sheet, and you may name the key column by its heading. Turn off if the data starts on row 1 with no header. |
| Output | **Where to put the sheets** | `New workbook` (default, safest) writes the per-value sheets into a brand-new workbook; your original is never touched. `Add sheets to current workbook` appends the new sheets after the active sheet. |
| Output | **Sheet name prefix** (optional) | Text prepended to every created sheet name (prefix `R-` turns value `East` into sheet `R-East`). Illegal characters and over-length names are sanitised automatically. |
| Output | **Backup workbook first** (default ON) | Only used in "Add sheets to current workbook" mode. Saves a `<name>_backup` copy to disk before any sheet is added. Ignored in "New workbook" mode. If the workbook has never been saved, the run is refused (nothing to back up). |
| Output | **Max distinct values** (default 50, range 1–1000) | Safety cap. If the key column has more distinct values than this, the tool refuses to run — a guard against accidentally picking a near-unique column (an ID, a timestamp) and spawning thousands of sheets. |
| Output | **Output format** | `Markdown` or `Plain text` for the result summary in the output pane. |

### What it never does

* It **never modifies your source data.** The source sheet is only read.
* In **New workbook** mode the original workbook is not touched at all.
* In **Add sheets** mode only *new* sheets are added; the source sheet is
  read-only, and the optional backup is written to a **new** path
  (`<name>_backup.<ext>`) — it never overwrites your working file.
* The result is always left **OPEN and UNSAVED** in Excel so you review and
  save (or discard) it yourself.

### Prerequisites

* The target workbook is **already open** in a running Excel instance (the tool
  attaches to it via combridge; it does not open files).
* ScripTree's bundled `lib/combridge/combridge.exe` is present (it is, in a
  normal install).

---

## 2. The logic (reconstruct-the-tool spec)

The tool is a **Strategy-A shim**: a ScripTree form runs a Python shim
(`sheet_segregator.py`); the shim bakes the form values into a C# Roslyn script
rendered from `sheet_segregator.csx.template`, then runs it through combridge,
which owns the live COM connection to Excel. Read the module docstring of
`sheet_segregator.py` for the full rationale; the essentials:

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`xlApp` / `xlBook` / `xlSheet`) and environment variables. So form
  values are baked into the script text by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — a clean run always exits
  0 (only compile=3 / unhandled-throw=4 / host=5 differ). So the script prints a
  machine-readable **sentinel** as its first stdout line and the shim translates
  it into the process exit code ScripTree actually reads.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__XLSEG__ STATUS=<code> [key=value ...]
```

The shim strips that line, prints the rest as the human report, and maps:

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Segregation succeeded; report follows. |
| `NOWB` | 2 | No workbook / no active sheet open. |
| `NODATA` | 2 | Active sheet has no usable table. |
| `BADCOL` | 2 | Key column could not be resolved. |
| `TOOMANY` | 2 | More distinct key values than the safety cap. |
| `UNSAVED` | 2 | Backup requested but the workbook was never saved. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes, surfaced verbatim. |

Exit 2 = a **guarded refusal / nothing done** (distinct from a crash at 3/4/5
and from success at 0).

### The eight steps the `.csx` performs

1. **Guard + bulk read.** If `xlBook`/`xlSheet` are null → `NOWB`. Read
   `xlSheet.UsedRange.Value` as a 1-based `object[,]`; record `nRows`/`nCols`
   and the used range's absolute top-left (`firstRow`/`firstCol`). A single-cell
   used range returns a scalar (not an array) → treat as `NODATA`. Require at
   least `hasHeader ? 2 : 1` rows, else `NODATA`.
2. **Resolve the key column** to a 1-based index *within the used range*, trying
   in order: (a) if `hasHeader`, case-insensitive trimmed match against the
   header cells; (b) a column **letter** (`A`=1 … `AB`=28) converted to absolute
   then offset into the range (`abs - firstCol + 1`); (c) a 1-based **position**
   within the range. No match → `BADCOL` (the report lists the available
   headers).
3. **Group rows** in **first-seen order**: an ordered list of distinct keys plus
   a `Dictionary<string, List<int>>` of row indices. Empty/blank key cells go in
   a `(blank)` bucket. If the distinct count exceeds `maxGroups` → `TOOMANY`.
4. **Capture per-column number formats** from the first data row (skip
   `General`; wrap in try/catch) so output cells keep dates/currency formatting.
5. **Prepare the output target.** In `new_sheets` mode with `backup_first`: the
   workbook must have been saved (`xlBook.Path` non-empty) else `UNSAVED`; with
   `DisplayAlerts` off, `xlBook.SaveCopyAs(<name>_backup.<ext>)` (copy-without-
   changing-open-workbook). The target book is `xlBook` in `new_sheets` mode, or
   a fresh `xlApp.Workbooks.Add()` in `new_workbook` mode (record its default
   sheets so they can be removed at the end).
6. **Build unique sheet names.** Seed a `HashSet` with existing sheet names; a
   local `MakeSheetName(raw)` applies the prefix, replaces the illegal characters
   `: \ / ? * [ ]`, truncates to 31 chars, and de-dupes with a ` (n)` suffix.
7. **Create + bulk-write** one sheet per key: build a 0-based `object[,]` block
   (header row if any, then the key's data rows), `Worksheets.Add(After: anchor)`,
   set `ws.Name`, write the block to `ws.Range[Cells[1,1], Cells[outRows,nCols]]`
   in one assignment, and re-apply the captured number formats per column. After
   the loop, in `new_workbook` mode (and only if at least one sheet was written),
   delete the recorded default sheets with `DisplayAlerts` off.
8. **Emit** the sentinel `__XLSEG__ STATUS=OK groups={N} mode={outputMode}`
   followed by the markdown/plain-text report (source sheet, key column and how
   it was resolved, data-row count, sheets created, mode, backup path if any,
   whether the header was copied, and a per-sheet row-count table), ending with a
   reminder that the result is open but unsaved.

### Excel COM facts this relies on (see the office-com RAG)

* `UsedRange.Value` is a **1-based** `[r,c]` array; a single cell returns a
  scalar. Bulk read/write is dramatically faster than cell-by-cell.
* Assigning a **0-based** `object[,]` to a `Range.Value` maps `[i,j]` → the
  cell at `(i+1, j+1)` of that range.
* `Worksheet.Delete` and `SaveAs`/`SaveCopyAs` raise **modal** dialogs that hang
  a hidden COM-launched Excel — always set `xlApp.DisplayAlerts = false` around
  them.
* `xlBook.Path` is `""` until the workbook has been saved to disk.

---

## 3. Files in this app

| File | Role |
|---|---|
| `sheet-segregator.scriptree` | The form definition (7 params; `filter` icon embedded as PNG). Validate with `python -m scriptree.cli.validate <path>` from `D:\Dev\ScripTree`. |
| `sheet-segregator.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (hides the IDE chrome, pops up on success/error). |
| `sheet_segregator.py` | The Strategy-A shim: parses argv, renders the `.csx`, runs combridge, parses the sentinel, owns the exit code. |
| `sheet_segregator.csx.template` | The Roslyn C# script template with `__TOKEN__` placeholders the shim fills. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
sheet_segregator.py
  --key-column <str>            (required)
  [--has-header]                (flag; first row is a header)
  --output-mode new_workbook|new_sheets
  --sheet-name-prefix <str>
  [--backup-first]              (flag; new_sheets mode only)
  --max-groups <int>            (clamped to >= 1 in the shim)
  --output-format markdown|text
```

The form's `argument_template` emits exactly this. Boolean flags use the
`{id?--flag}` conditional form; string passthroughs use a `["--flag","{id}"]`
token group with `no_split: true` on the param so values with spaces stay
single tokens.

---

## 4. Editing / maintenance notes

* **Re-embedding the icon strips hand-edited `min`/`max`/`step`.** `embed_icon`
  does a load→save round-trip through the Param model, which does not persist
  those numeric bounds. Workflow: embed the icon **first**, then add
  `max_groups`'s `min`/`max` as the **last** hand-edit, then validate. (The shim
  also clamps `max_groups` to `>= 1` defensively.)
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree.cli.validate <path>`.
* **Offline render-check:** render the template with sample values and grep the
  output for `__[A-Z_]+__` — the only legitimate matches are the `__XLSEG__`
  sentinel and the `__PLACEHOLDER__` mention inside a template comment; any other
  `__TOKEN__` is an unfilled placeholder (a compile error waiting to happen).
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path; this repo does not
  bundle combridge.
