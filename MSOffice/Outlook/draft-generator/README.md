# Draft Generator from Template (Outlook)

Turn an Excel mailing list into **one DRAFT Outlook email per row** — a mail
merge that **stops at drafts** for human review. Give it an `.xlsx` whose first
row is column headers (`Email`, `Name`, `Company`, …) plus a subject and body
template containing `{ColumnName}` placeholders; it creates a personalised draft
in your Outlook **Drafts** folder for each data row.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax that
> enacts it. If you change behaviour, change this file in the same commit.

---

## 1. The safety contract (read this first)

**This tool creates DRAFTS only. It NEVER sends.** That is the headline property
and the entire reason the tool is shaped the way it is. The guarantee is enforced
in three independent places, and every one of them must survive any future edit:

1. **No send call exists.** The generated `.csx` does only `MailItem.Save()` on
   each new item — which lands an *unsent* mail item in the Drafts folder. There
   is **no `.Send()` anywhere** in `generate_drafts.csx.template`. Grepping the
   rendered script for `.Send(` must return nothing. Sending is forbidden;
   creating drafts is allowed.
2. **Hard cap of 50 drafts per run.** `max_drafts` is clamped to `[1, 50]` in the
   Python shim (`clamp_max`) *before* any rows are collected, so the `.csx` can
   never be handed more than 50 rows. A bad merge cannot blast a list even as
   drafts. **50 is the absolute maximum**, not just the default.
3. **The shim only reads and writes a temp file.** It reads the `.xlsx` with
   openpyxl and writes a temporary `.csx`; it has no code path that could send
   mail.

After a run you have up to 50 **unsent drafts** to review in Outlook. Nothing
leaves the mailbox until *you* click Send. If a future change adds any send
operation, it has broken the tool's defining promise — don't.

---

## 2. What the user sees (end-user guide)

### The form fields

The form groups its fields into three collapsible sections — **Mailing list**, **Message**, and **Options**.

| Section | Field | Meaning |
|---|---|---|
| Mailing list | **Mailing list (.xlsx)** (required) | The Excel workbook of recipients. **Row 1 = column headers**; each row below is one recipient. The **active sheet** is used. Read with openpyxl — Excel need not be installed or open. |
| Mailing list | **Recipient column header** | The exact header text of the column holding email addresses. Matched **case-insensitively**. Default `Email`. If no header matches, the run is refused (exit 2). |
| Message | **Subject template** | The subject line for every draft. May contain `{Column}` placeholders, e.g. `Your Q3 statement, {Name}`. |
| Message | **Body template** (required) | The message body (multi-line). Use `{Column}` placeholders to merge per-row values. |
| Options | **Maximum drafts (cap 50)** | The most drafts to make this run. Hard-capped at 50, floored at 1. If the list has more usable rows, only the first N are drafted. |
| Options | **Output format** | `Markdown` or `Plain text` for the result summary. |

### How `{Column}` tokens map to the spreadsheet

A token is the **brace-wrapped exact header text**, matched **case-sensitively**
against the column headers in row 1 of the workbook:

* If your sheet has a header `Name`, then `{Name}` in the subject/body is replaced
  by that row's value in the `Name` column. `{name}` (lower-case) would **not**
  match — the token must match the header's exact casing.
* A header with a space works too: header `First Name` → token `{First Name}`.
* A blank cell renders as nothing (empty string), so `{MiddleName}` for a row
  with no middle name just disappears.
* Braces that don't correspond to any real header are **left untouched** — a body
  containing `{not a column}` or a JSON snippet is not mangled.

(The recipient-column match for **Recipient column header** is the one place that
is case-*insensitive*, so `Email` / `email` / `EMAIL` all find the column. The
`{Column}` token match in the templates is case-*sensitive*.)

### Prerequisites

* **Outlook is running** (a profile loaded). No window or explorer needs to be
  open — `CreateItem` works headless.
* Excel does **not** need to be installed or running (openpyxl reads the file
  directly).
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 3. The logic (reconstruct-the-tool spec)

The tool is a **Strategy-A shim**, but with an important twist from the rest of
the catalog: **the Python shim does the heavy lifting (reading Excel + rendering
templates), and the `.csx` only talks to Outlook.**

```
ScripTree form ──argv──► generate_drafts.py (shim)
                              │  1. openpyxl reads the .xlsx, renders {Column} tokens per row
                              │  2. bakes (to, subject, body) triples into a .csx from the template
                              ▼
                         combridge.exe outlook run-script <temp.csx> -
                              │  3. .csx loops the baked rows, CreateItem + Save() (DRAFTS ONLY)
                              ▼
                         first-line sentinel ──► shim parses ──► process exit code
```

### Why the shim reads the Excel (not the `.csx`)

* `openpyxl` is installed in ScripTree's Python and is a robust, Excel-free way
  to read `.xlsx`. We open with `load_workbook(path, read_only=True,
  data_only=True)` — `read_only` for low memory on big lists, `data_only` so a
  formula cell yields its *computed value*, not the formula text.
* Token rendering (`{Column}` → cell value) is trivial string work that belongs
  in Python.
* This means the only live application the tool needs is **Outlook**.

### combridge constraints this works around

* **`run-script` has no argv channel** — a `.csx` sees only the plugin globals
  (`olApp` / `olNs` / `olExplorer`) and environment variables. So the data is
  **baked into the script text** by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return`** — a clean run always exits 0 (only
  compile=3 / unhandled-throw=4 / host=5 differ). So the `.csx` prints a
  machine-readable **sentinel** first line and the shim translates it to the exit
  code ScripTree reads.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__OLDRAFT__ STATUS=<code> [key=value ...]
```

The shim strips that line, prints the rest as the report, and maps:

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Merge ran; `created=N failed=F` follow. A real run is always exit 0. |
| `NOITEMS` | 2 | Nothing to draft (zero usable rows handed to the `.csx`). In normal operation the **shim short-circuits this first** (see below) and never launches combridge. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes, surfaced verbatim. |

The shim also returns **exit 2** for its own guarded refusals (mailing list
missing, recipient column absent, zero usable rows) with a clear stderr message,
and **exit 1** if it cannot find the template or combridge.

### The shim's algorithm (`generate_drafts.py`)

1. **Validate inputs.** Empty/missing `--mailing-list` → exit 2.
2. **Clamp the cap.** `max_drafts = clamp_max(value)` forces `[1, 50]` *before*
   reading rows — the first enforcement point of the 50-draft cap.
3. **Read the workbook** (`read_rows`): open `read_only/data_only`, take the
   **active** sheet, treat **row 1 as headers**. Find the recipient column by a
   **case-insensitive** match on `--email-column` (absent → `ValueError` → exit 2,
   listing the headers it did find). Walk data rows: build a `{header: value}`
   dict (cells coerced to strings; whole-number floats lose the `.0`; blanks →
   `""`), **skip rows whose recipient cell is blank**, render subject + body via
   `render_templates`, append the `(to, subject, body)` triple. Stop once
   `max_drafts` rows are collected.
4. **Short-circuit an empty merge.** If zero usable rows, print a clear stderr
   message and exit 2 — don't spin up Outlook/combridge for nothing. (This is the
   cleaner of the two options in the spec; the `.csx` still defends itself with
   `NOITEMS` if ever handed an empty array.)
5. **Render the `.csx`** (`render_csx`): bake the triples as a C# `List<string[]>`
   (each field escaped by `csharp_literal`, which handles `\r`/`\n` for
   multi-line bodies), bake the row count as a bare integer, bake the mailing-list
   path for the report.
6. **Run combridge** `outlook run-script <temp.csx> -`, parse the sentinel, own
   the exit code. The temp file is always unlinked in a `finally`.

### The `.csx`'s algorithm (`generate_drafts.csx.template`)

1. If `rowCount == 0` → `STATUS=NOITEMS`, return (defensive; shim usually catches
   first).
2. Resolve the Drafts folder for the report:
   `MAPIFolder drafts = olNs.GetDefaultFolder(OlDefaultFolders.olFolderDrafts);`
   (wrapped in try/catch — even if this fails, `Save()` still lands new unsent
   items in Drafts).
3. For each baked `(to, subj, body)` triple (already ≤ 50):
   ```csharp
   MailItem mi = (MailItem)olApp.CreateItem(OlItemType.olMailItem);
   mi.To = to; mi.Subject = subj; mi.Body = body;
   mi.Save();        // lands the UNSENT draft in Drafts. NO mi.Send().
   ```
   Each iteration is wrapped in try/catch so one bad address (rejected by Outlook
   on `Save`) increments `failed` rather than aborting the whole merge.
4. Emit `__OLDRAFT__ STATUS=OK created={N} failed={F}` then the markdown/text
   report: mailing-list path, recipients processed, drafts created (and the
   Drafts folder name), failure count, a capped per-draft table (`#`, To,
   Subject), and the explicit reminder: **"No emails were sent — review them in
   your Drafts folder before sending."**

### Outlook COM facts this relies on (see the office-com RAG)

* `olApp` (`Ol._Application`) is **never null**; `olNs` is the MAPI namespace.
  `olExplorer` may be null but is **not needed** here — `CreateItem` works with no
  window/explorer open.
* `olApp.CreateItem(OlItemType.olMailItem)` returns a brand-new, **unsent**
  `MailItem`. Cast the returned `object` to `MailItem`.
* `MailItem.Save()` writes the unsent item into the Drafts folder. `MailItem.To`
  / `.Subject` / `.Body` are the obvious string properties. We do **not** call
  `.Move()` — `Save()` on a never-sent new item already puts it in Drafts.
* `olNs.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)` returns the default
  store's Drafts folder (used only to name it in the report).

> **Flagged for live verification:** the `(MailItem)olApp.CreateItem(...)` cast
> and the `OlDefaultFolders.olFolderDrafts` enum name. Both match the documented
> Outlook object model and the project's confirmed Outlook plugin globals, but
> have not yet been traced against a live Outlook from this app.

---

## 4. Files in this app

| File | Role |
|---|---|
| `draft-generator.scriptree` | The form definition (6 params; an envelope-with-pencil "draft" icon embedded as PNG). Validate with `python -m scriptree validate <path>` from `D:\Dev\ScripTree`. |
| `draft-generator.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (hides IDE chrome, pops up on success/error). |
| `generate_drafts.py` | The Strategy-A shim: parses argv, **reads the .xlsx with openpyxl, renders `{Column}` tokens, bakes the rows**, runs combridge, parses the sentinel, owns the exit code, enforces the 50-draft cap. |
| `generate_drafts.csx.template` | The Roslyn C# script template with `__TOKEN__` placeholders. Creates the drafts (`CreateItem` + `Save()`); **contains no `.Send()`**. |
| `examples/` | A reproducible `make_example.py` + `sample_list.xlsx` + a scenario `README.md`. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
generate_drafts.py
  --mailing-list <path>            (required; .xlsx)
  --email-column <str>             (default "Email"; case-insensitive header match)
  --subject-template <str>         (may contain {Column} tokens)
  --body-template <str>            (required; multi-line; may contain {Column} tokens)
  --max-drafts <int>               (clamped to [1, 50])
  --output-format markdown|text
```

The form's `argument_template` emits exactly this. Every value uses a
`["--flag", "{id}"]` token group; all string params are `no_split: true` so a
multi-line body or a path/subject with spaces stays a single argv element. An
empty `--subject-template` group still emits (argparse default `""`).

---

## 5. Editing / maintenance notes

* **Preserve the drafts-only contract.** Never add `MailItem.Send()` (or
  `.SendUsingAccount`-driven send, or any auto-send) to the `.csx`. The only
  permitted mutation is `Save()` (which produces a draft). Never raise the
  `MAX_DRAFTS_CAP` above 50 without revisiting the safety story here and in the
  RAG.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree validate <path>`.
* **Offline render-check:** render the template against a sample workbook and
  grep the output for `__[A-Z_]+__` — the only legitimate match is the
  `__OLDRAFT__` sentinel. Any other `__TOKEN__` is an unfilled placeholder (a
  compile error waiting to happen). Also confirm a `{Name}` token resolved and
  that the baked `List<string[]>` has the expected number of entries. Grep the
  rendered `.csx` for `.Send(` — it must return **nothing**.
* **Re-embedding the icon** does a load→save round-trip that drops hand-edited
  `min`/`max`/`step` numeric bounds. `max_drafts` here relies on the **shim**
  clamp, not a form-level `max`, precisely so the icon round-trip can't silently
  weaken the cap. If you ever add a form-level `max`, embed the icon first and add
  the bound last.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path; this repo does not
  bundle combridge.
