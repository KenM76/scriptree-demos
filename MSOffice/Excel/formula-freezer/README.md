# Formula-to-Value Freezer (Excel)

Convert **every formula** in the open Excel workbook into its current **static
value** — locking down the calculated data state so the numbers can never
silently change again. On a **copy by default**, so your live formulas stay
intact in the workbook you are looking at.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

You have a workbook full of formulas — `=SUM(...)`, `=A2*B2` columns, cross-sheet
references like `=Sales!B9`. You want to "lock down" the current calculated
state: replace each formula with the number it currently produces, so the file
is a permanent snapshot. Typical reasons: a month-end figure handed to an
auditor; a file sent outside the company where the source data behind the
references won't be available; or simply freezing a result before someone edits
an input by mistake.

### The form fields

| Field | Meaning |
|---|---|
| **Scope** | `All worksheets` (default) freezes every formula in the workbook. `Active sheet only` freezes just the one sheet currently active in Excel. |
| **Work on a copy** (default ON) | **Safety guard.** When on, the workbook is copied to a sibling `<name>_Frozen.xlsx`, the freeze happens in that copy, and the copy is saved and closed — your original stays open with its live formulas intact. When off, the open workbook is frozen in memory and left **open and UNSAVED** for you to review. |
| **Output format** | `Markdown` or `Plain text` for the result summary in the output pane. |

### What it does to your data

* Each formula cell is overwritten with the **constant value it currently
  shows**. The displayed numbers are identical; the formulas are gone.
* **Number formats are preserved** — dates stay dates, currency stays currency.
  Only the underlying formula is replaced, not the cell's display format.

### What it never does

* In **copy mode** (default) the original workbook is **not touched at all** —
  the freeze happens entirely in the `_Frozen.xlsx` copy, which is the only file
  written.
* In **in-place mode** the open workbook is changed **in memory only** and left
  **UNSAVED** — the file on disk is never overwritten by this tool; you decide
  whether to save (or Ctrl+Z to undo).
* It refuses (changing nothing) if a target worksheet is **protected**, or if
  copy mode is requested for a workbook that has **never been saved**.

### Prerequisites

* The target workbook is **already open** in a running Excel instance (the tool
  attaches to it via combridge; it does not open files).
* ScripTree's bundled `lib/combridge/combridge.exe` is present (it is, in a
  normal install).

---

## 2. The logic (reconstruct-the-tool spec)

The tool is a **Strategy-A shim**: a ScripTree form runs a Python shim
(`freeze_formulas.py`); the shim bakes the form values into a C# Roslyn script
rendered from `freeze_formulas.csx.template`, then runs it through combridge,
which owns the live COM connection to Excel. Read the module docstring of
`freeze_formulas.py` for the full rationale; the essentials:

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
__XLFREEZE__ STATUS=<code> [frozen=N] [sheets=N] [copy=0/1]
```

The shim strips that line, prints the rest as the human report, and maps:

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Freeze succeeded; report follows. |
| `NOBOOK` | 2 | No workbook is open. |
| `PROTECTED` | 2 | A target worksheet is protected; its cells can't be overwritten. Aborted before any change. |
| `UNSAVED` | 2 | Copy mode requested but the workbook was never saved (no folder for the copy). |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes, surfaced verbatim. |

Exit 2 = a **guarded refusal / nothing done** (distinct from a crash at 3/4/5
and from success at 0). The shim's `PRECONDITION_FAILS` tuple is
`("NOBOOK", "PROTECTED", "UNSAVED")`.

### The six steps the `.csx` performs

1. **Guard.** If `xlBook` is null → `NOBOOK` (exit 2). (`xlApp` is never null;
   `xlBook` is the active workbook.)
2. **Choose target sheets** per `scope`: `active_sheet` → just the active
   sheet's name (`xlSheet`, falling back to `xlBook.ActiveSheet`); `all_sheets`
   → every `xlBook.Worksheets` name. The names are resolved against the **open**
   workbook so the protection report uses names the user recognises. (No
   worksheet at all → `OK` with `frozen=0`.)
3. **Protection guard.** Enumerate the target sheets; if any has
   `ws.ProtectContents == true`, print `PROTECTED` naming that sheet and abort
   **before any mutation** (a protected sheet rejects the value write and would
   otherwise throw mid-run). `ProtectContents` is read inside try/catch.
4. **Choose the working workbook (copy-mode safety).**
   * **Copy mode** (`work_on_copy` true): if `xlBook.Path == ""` (never saved) →
     `UNSAVED`. Otherwise build `frozenPath = <dir>/<base>_Frozen.xlsx` (the
     `.xlsx` extension is forced regardless of the original's `.xls`/`.xlsm`),
     set `xlApp.DisplayAlerts = false`, `xlBook.SaveCopyAs(frozenPath)` (writes a
     byte-identical copy **without** changing the open book's path or dirty
     state), then `work = xlApp.Workbooks.Open(frozenPath)` and operate on
     `work`. The user's `xlBook` is never modified.
   * **In-place mode** (`work_on_copy` false): `work = xlBook`. We never `Save`
     it; it is left open and unsaved.
5. **Freeze each target sheet of `work`.** For each target sheet: take
   `Range ur = ws.UsedRange` (skip if null). Count formula cells **for the
   report** via `ur.SpecialCells(Excel.XlCellType.xlCellTypeFormulas).Count`,
   wrapped in try/catch because `SpecialCells` **throws "No cells were found"**
   when the sheet has no formulas (catch → 0). Then **strip the formulas** with
   the value round-trip:

   ```csharp
   object vals = ur.Value2;   // current constants — NO formulas attached
   ur.Value2  = vals;         // write them back -> formulas replaced by constants
   ```

   Reading `Value2` returns the cells' calculated constants; writing them back
   overwrites the formula cells with those constants. `NumberFormat` is untouched
   (Value2 carries values only), so dates/currency still display correctly. The
   **single-cell** used range returns a scalar from `Value2` (not an array);
   assigning a scalar back is valid and freezes it, so no special-casing is
   needed. `Value2` (not `Value`) is used so currency/date cells round-trip as
   raw doubles rather than Currency/DateTime COM variants.
6. **Copy mode finish + report.** In copy mode, `work.Save()` then
   `work.Close(false)` — the frozen deliverable is saved to disk and closed; the
   user's original re-activates. Restore `DisplayAlerts` in a `finally`. Finally
   print the sentinel `__XLFREEZE__ STATUS=OK frozen={count} sheets={n}
   copy={0|1}` followed by the markdown/plain-text report (workbook name, scope,
   sheets processed, formula cells converted, copy path or in-place note, and a
   per-sheet table of formula-cell counts), ending with the appropriate
   "copy saved / left unsaved" reminder.

### Excel COM facts this relies on (see the office-com RAG)

* **`Range.Value2 = Range.Value2` freezes formulas.** Reading `Value2` yields
  the current constants with no formula; assigning them back replaces the
  formulas. `NumberFormat` is preserved (Value2 carries values, not formatting).
  Use **`Value2`**, not `Value`, to avoid Currency/DateTime variant round-trips.
* **`SpecialCells(xlCellTypeFormulas)` throws when there are no formulas** ("No
  cells were found") — always wrap in try/catch and treat the throw as zero.
* **`ProtectContents`** is `true` on a protected sheet; freezing such a sheet
  would throw, so guard up front and abort.
* **`SaveCopyAs` ≠ `SaveAs`** — `SaveCopyAs` writes a copy to disk **without**
  changing the open workbook's `Path` or dirty/clean state, so the original is
  left exactly as the user had it. `Workbooks.Open(path)` then loads that copy to
  edit it headlessly.
* **`DisplayAlerts` is a `bool`** on Excel (`xlApp.DisplayAlerts = false`).
  `SaveCopyAs` / `Save` / `Open` can raise **modal** dialogs that hang a COM-
  driven Excel — set `DisplayAlerts = false` around them and restore after.
* **`xlBook.Path` is `""`** until the workbook has been saved to disk.

---

## 3. Files in this app

| File | Role |
|---|---|
| `formula-freezer.scriptree` | The form definition (3 params; snowflake-over-grid icon embedded as PNG). Validate with `python -m scriptree validate <path>` from `D:\Dev\ScripTree`. |
| `formula-freezer.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (hides the IDE chrome, pops up on success/error). |
| `freeze_formulas.py` | The Strategy-A shim: parses argv, renders the `.csx`, runs combridge, parses the sentinel, owns the exit code. |
| `freeze_formulas.csx.template` | The Roslyn C# script template with `__TOKEN__` placeholders the shim fills. |
| `examples/` | A generator (`make_example.py`) + `sample_formulas.xlsx` (two sheets of real formulas) + a scenario README. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
freeze_formulas.py
  --scope all_sheets|active_sheet
  [--work-on-copy]              (flag; freeze a saved <name>_Frozen.xlsx copy)
  --output-format markdown|text
```

The form's `argument_template` emits exactly this. The boolean uses the
`{work_on_copy?--work-on-copy}` conditional form; the enums use a
`["--flag","{id}"]` token group. There are no free-text params, so no
`no_split` is needed (the `csharp_literal` escaper in the shim is retained for
robustness and catalog consistency).

---

## 4. Editing / maintenance notes

* **Re-embedding the icon strips hand-edited `min`/`max`/`step`.** `embed_icon`
  (any load→save round-trip through the Param model) does not persist those
  numeric bounds. This app has no bounded numeric param, so it is unaffected —
  but if one is added, embed the icon **first**, add the bounds as the **last**
  hand-edit, then validate (and clamp in the shim). The `.scriptree` here was
  generated by a throwaway `_build_scriptree.py` (deleted after use).
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with sample values and grep the
  output for `__[A-Z_]+__` — the only legitimate matches are the `__XLFREEZE__`
  sentinel and the `__PLACEHOLDER__` mention inside a template comment; any other
  `__TOKEN__` is an unfilled placeholder (a compile error waiting to happen).
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path; this repo does not
  bundle combridge, so running the shim here returns the "could not locate" error
  with exit 1 (correct — the app is deploy-only).
