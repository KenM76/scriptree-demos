# Multi-CSV Sheet Aggregator (Excel)

Import **every `.csv` and `.txt` file in a folder** as a **separate new sheet**
in the Excel workbook you already have open — one sheet per file, each named
after its source file. Classic example: a folder of monthly exports
(`jan.csv`, `feb.csv`, …) becomes one workbook with a sheet per month, ready to
consolidate with formulas or a PivotTable.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. What the user sees (end-user guide)

You have a folder full of delimited text files. You point the tool at it. The
tool adds one new worksheet per file to your open workbook, each sheet named
after its file (`sales_q1.csv` → sheet `sales_q1`), and imports that file's rows
into it.

### The form fields

| Field | Meaning |
|---|---|
| **Source folder** (required) | The folder to scan. Every `.csv` and `.txt` file **directly inside** it is imported (subfolders are not searched). |
| **Delimiter** (default Auto-detect) | How each file's columns are separated. **Auto-detect** inspects each file's first non-empty line and picks comma, tab, or semicolon — this correctly handles European semicolon-delimited CSVs that a plain comma split would mangle. Force a specific delimiter only if auto-detect guesses wrong. |
| **Backup workbook first** (default ON) | Saves a `<name>_Backup` copy of the open workbook to disk **before** any sheet is added. If the workbook has never been saved there is no folder for the backup, so the run is refused. Untick to import without a backup. |
| **Output format** | `Markdown` or `Plain text` for the result summary in the output pane. |

### What it does and never does

* The source files on disk are **only read, never changed.**
* The tool **only adds sheets** to your open workbook. It never overwrites your
  existing sheets and never saves over your workbook file.
* The optional backup is written to a **new** path (`<name>_Backup.<ext>`) via
  `SaveCopyAs`, which does **not** change the open workbook's own path or its
  dirty/clean state.
* After the import the workbook is left **OPEN and UNSAVED** so you review the
  new sheets and save (or discard) the workbook yourself.
* Files that fail to import are **skipped and listed** in the report; the rest
  still import (one bad file never aborts the batch).

### Prerequisites

* A workbook is **already open** in a running Excel instance (the tool attaches
  to it via combridge; it does not open the workbook for you).
* ScripTree's bundled `lib/combridge/combridge.exe` is present (it is, in a
  normal install).

---

## 2. The logic (reconstruct-the-tool spec)

The tool is a **Strategy-A shim**: a ScripTree form runs a Python shim
(`csv_aggregator.py`); the shim bakes the form values into a C# Roslyn script
rendered from `csv_aggregator.csx.template`, then runs it through combridge,
which owns the live COM connection to Excel. The essentials (full rationale in
the shim's module docstring):

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`xlApp` / `xlBook` / `xlSheet`) and environment variables. So form
  values are baked into the script text by replacing `__TOKEN__` placeholders,
  each value C#-escaped (`csharp_literal`) or rendered as a C# bool literal.
* **combridge swallows the script's `return` value** — a clean run always exits
  0 (only compile=3 / unhandled-throw=4 / host=5 differ). So the script prints a
  machine-readable **sentinel** as its first stdout line and the shim translates
  it into the process exit code ScripTree actually reads.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__CSVAGG__ STATUS=<code> [files=N] [failed=N]
```

The shim strips that line, prints the rest as the human report, and maps:

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Import ran; `files=`/`failed=` counts and report follow. |
| `NOWB` | 2 | No workbook open to import into. |
| `NOFILES` | 2 | No `.csv`/`.txt` files in the chosen folder (or folder unreadable). |
| `UNSAVED` | 2 | Backup requested but the workbook was never saved. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes, surfaced verbatim. |

Exit 2 = a **guarded refusal / nothing done** (distinct from a crash at 3/4/5
and from success at 0). Note `OK` with `failed>0` is still exit 0 — partial
success is a success; the skipped files are named in the report.

### The five steps the `.csx` performs

1. **Guard.** If `xlBook` is null → `NOWB`. (`xlSheet` is never read or modified;
   new sheets are appended after the current last sheet.)
2. **Gather source files.** `Directory.GetFiles(sourceFolder)` filtered to
   extensions `.csv`/`.txt` (case-insensitive), sorted by name
   (`StringComparer.OrdinalIgnoreCase`) for stable, deterministic sheet order. A
   missing/unreadable folder is caught and treated as empty. Empty list →
   `NOFILES`.
3. **Optional backup.** If `backup_first`: the workbook must have been saved
   (`xlBook.Path` non-empty) else `UNSAVED`; with `DisplayAlerts` off,
   `xlBook.SaveCopyAs(<dir>/<baseName>_Backup<ext>)` (ext defaults to `.xlsx` if
   the open book somehow has none).
4. **Import each file.** Seed a `HashSet` (case-insensitive) with existing sheet
   names. With `DisplayAlerts` off for the whole batch, for each file (in its own
   try/catch so one failure is recorded and skipped, not fatal):
   * Resolve the delimiter — if `delimiterMode == "auto"`, **sniff per file**
     (see below); else use the forced choice.
   * `Worksheets.Add(After: anchor)` (anchor starts at the last sheet, advances
     each iteration so imports stay in folder order at the end), name it via
     `MakeSheetName(fileNameWithoutExtension)` (sanitize `: \ / ? * [ ]`,
     truncate to 31, de-dupe with a ` (n)` suffix that itself stays ≤ 31).
   * `QueryTables.Add("TEXT;" + path, ws.Range["A1"])` → configure as a
     **delimited** import: `TextFileParseType = xlDelimited`,
     `TextFileConsecutiveDelimiter = false` (don't merge empty cells), set
     exactly one of `TextFileCommaDelimiter` / `TextFileTabDelimiter` /
     `TextFileSemicolonDelimiter` true, others (`Space`, `Other`) off,
     `AdjustColumnWidth = true`, `RefreshStyle = xlOverwriteCells`.
   * `qt.Refresh(false)` (synchronous; `BackgroundQuery:false`), then
     `qt.Delete()` — this **keeps the imported cells but drops the live query**,
     leaving plain values rather than a connected external-data table.
   * Record `ws.UsedRange.Rows.Count` as the imported row count.
5. **Emit** the sentinel `__CSVAGG__ STATUS=OK files={imported} failed={F}`
   followed by the markdown/plain-text report (source folder, files imported,
   files skipped if any, the delimiter mode, the backup path if one was made, a
   per-sheet table of *sheet / source file / rows / delimiter*, a "skipped files"
   section with the error for each), ending with a reminder that the workbook now
   has the new sheets but is not yet saved.

### The auto-detect delimiter sniffer

For `delimiter = auto`, `SniffDelimiter(path)` reads the **first non-empty
line** and counts `,`, `\t`, `;`:

* `tab` wins if `tabs > commas && tabs >= semis`;
* else `semicolon` wins if `semis > commas && semis > tabs`;
* else `comma` (the default, and the tie-break).

An empty or unreadable file falls back to `comma` (harmless — the import just
produces a single-column sheet). The sniff is **per file**, so a folder mixing
comma and semicolon files imports each correctly.

### Excel COM facts this relies on (see the office-com RAG)

* `QueryTables.Add(connection, destination)` with a `"TEXT;<path>"` connection
  string is the scriptable text-import engine; `Refresh(false)` pulls
  synchronously and `Delete()` drops the query while keeping the data.
* `SaveCopyAs` backs up to disk **without** changing the open workbook's `Path`
  or dirty state; `xlBook.Path` is `""` until the workbook has been saved once.
* `Worksheets.Add`, `SaveCopyAs`, and the QueryTable refresh can raise **modal**
  dialogs that hang a hidden COM-attached Excel — always set
  `xlApp.DisplayAlerts = false` around them (restored in a `finally`).
* Sheet names: max 31 chars; may not contain `: \ / ? * [ ]`; must be unique
  (case-insensitive) within the workbook.
* `System.IO` **is** in the Excel plugin's default ScriptUsings (contrast the
  Outlook plugin, where it must be added) — `Directory`/`Path`/`File` are
  available.

---

## 3. Files in this app

| File | Role |
|---|---|
| `csv-aggregator.scriptree` | The form definition (4 params; `link` icon embedded as PNG). Validate with `python -m scriptree validate <path>` from `D:\Dev\ScripTree`. |
| `csv-aggregator.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (hides the IDE chrome, pops up on success/error). |
| `csv_aggregator.py` | The Strategy-A shim: parses argv, renders the `.csx`, runs combridge, parses the sentinel, owns the exit code. |
| `csv_aggregator.csx.template` | The Roslyn C# script template with `__TOKEN__` placeholders the shim fills. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
csv_aggregator.py
  --source-folder <path>                  (required)
  --delimiter auto|comma|tab|semicolon    (default auto)
  [--backup-first]                        (flag)
  --output-format markdown|text
```

The form's `argument_template` emits exactly this. The boolean flag uses the
`{id?--flag}` conditional form; `--source-folder`/`--delimiter`/
`--output-format` use `["--flag","{id}"]` token groups.

---

## 4. Editing / maintenance notes

* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template with sample values and grep the
  output for `__[A-Z_]+__` — the only legitimate matches are the `__CSVAGG__`
  sentinel and any `__PLACEHOLDER__` mention inside a template comment; any other
  `__TOKEN__` is an unfilled placeholder (a compile error waiting to happen).
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path; this repo does not
  bundle combridge.
* **Why a mutator, not read-only:** importing necessarily adds sheets to the open
  workbook. The safety model is therefore (a) source files are read-only, (b) the
  optional pre-import backup, and (c) leaving the result open-and-unsaved so the
  user has the final commit decision — not an in-place file overwrite.
