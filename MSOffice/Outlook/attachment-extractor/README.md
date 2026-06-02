# Surgical Attachment Extractor (Outlook)

Save the file attachments from your **selected Outlook emails** (or every email
in the open folder) to a folder on disk — in one pass, no per-message "Save
Attachments" clicking. Classic use: select a batch of invoice emails, run the
tool, get every PDF in one folder.

> This README is written to the project's **documentation-first** standard: a
> competent engineer or LLM should be able to **reconstruct the entire tool from
> this document alone**. The prose is the logic; the code is just the syntax
> that enacts it. If you change behaviour, change this file in the same commit.

---

## 1. The safety contract (read this first)

The tool is **strictly read-only with respect to the mailbox.** The *only*
mutating operation anywhere in it is `Attachment.SaveAsFile`, which copies an
attachment's bytes to a **new file on disk**. It never calls `Attachment.Delete`,
never sets a property on a `MailItem`, never calls `.Save()` / `.Move()` /
`.Delete()` on a message. Your emails and their attachments are left exactly as
they were; the only thing that changes is that new files appear in your output
folder. That guarantee is what "Surgical" means here — verify it survives any
future edit.

---

## 2. What the user sees (end-user guide)

### The form fields

The form groups its fields into three collapsible sections — **Which emails**, **Filter**, and **Output** (fields render grouped by section, so the on-screen order follows the table below).

| Section | Field | Meaning |
|---|---|---|
| Which emails | **Which emails** | `Selected emails` (default, the surgical path) processes only the messages currently highlighted in Outlook. `Current folder` processes every item in the open folder (slower on large folders). |
| Filter | **Only these extensions** | Optional comma/space list, e.g. `pdf, docx, xlsx`. Case- and dot-insensitive. Blank = save every attachment. |
| Filter | **Include inline images** (default OFF) | When on, embedded body/signature images (logos, tracking pixels) are also saved. Off by default — these are usually noise, not real attachments. |
| Output | **Save attachments to** (required) | The output folder on disk. Created automatically if it doesn't exist. |
| Output | **If a file already exists** | `Rename` (default) keeps both by appending ` (1)`, ` (2)`, …; `Skip` keeps the existing file; `Overwrite` replaces it. Matters because many emails carry a file literally named `invoice.pdf`. |
| Output | **Output format** | `Markdown` or `Plain text` for the result summary. |

### Prerequisites

* Outlook is **running with a mailbox open** (an active explorer window).
* For `Selected emails` scope, you have one or more messages selected.
* ScripTree's bundled `lib/combridge/combridge.exe` is present.

---

## 3. The logic (reconstruct-the-tool spec)

The tool is a **Strategy-A shim**: a ScripTree form runs a Python shim
(`attachment_extractor.py`); the shim bakes the form values into a C# Roslyn
script rendered from `attachment_extractor.csx.template`, then runs it through
combridge, which owns the live COM connection to Outlook.

* **combridge `run-script` has no argv channel** — a `.csx` only sees the plugin
  globals (`olApp` / `olNs` / `olExplorer`) and environment variables. So form
  values are baked into the script text by replacing `__TOKEN__` placeholders.
* **combridge swallows the script's `return` value** — a clean run always exits 0
  (only compile=3 / unhandled-throw=4 / host=5 differ). So the script prints a
  machine-readable **sentinel** as its first stdout line and the shim translates
  it into the process exit code ScripTree reads.

### Sentinel + status → exit-code contract

First stdout line of the `.csx`:

```
__OLXTRACT__ STATUS=<code> [key=value ...]
```

The shim strips that line, prints the rest as the human report, and maps:

| STATUS | Exit | Meaning |
|---|---|---|
| `OK` | 0 | Extraction ran; report follows. **`saved` may be 0** — an empty match is still a success, not an error. |
| `NOEXPLORER` | 2 | No active Outlook explorer (or, in folder scope, no current folder). |
| `NOSEL` | 2 | Scope = selection but nothing is selected. |
| `NOOUTDIR` | 2 | Output folder missing and could not be created. |
| 3 / 4 / 5 / connect-fail | passthrough | combridge's own failure codes, surfaced verbatim. |

### The five steps the `.csx` performs

1. **Resolve the items.** If `olExplorer` is null → `NOEXPLORER`. In `selection`
   scope read `olExplorer.Selection`; empty → `NOSEL`. In `folder` scope cast
   `olExplorer.CurrentFolder` to `Folder` (null → `NOEXPLORER`) and collect every
   `folder.Items` member. (Folder scope must touch each item — there is no MAPI
   table that can *save* attachment bytes — so it is slower; selection is the
   fast path.)
2. **Ensure the output dir** with `Directory.CreateDirectory(outDir)`; failure →
   `NOOUTDIR`.
3. **Build the filename helpers.** `Sanitize` replaces
   `Path.GetInvalidFileNameChars()` with `_`. `ReserveName` enforces the clash
   policy against a `HashSet` seeded with the files already in `outDir`:
   `overwrite` returns the base name; `skip` returns null when taken; `rename`
   appends ` (n)` until unique. The extension allow-list is parsed once into a
   case-insensitive, dot-less `HashSet` (blank spec = no filter).
4. **Walk items, save matching attachments.** For each item (`dynamic` so it
   works for `MailItem` and other item types), read `SenderName`/`SenderEmailAddress`
   for the report, then iterate `item.Attachments`. Per attachment: read
   `FileName`; detect **inline** via the MAPI property `PR_ATTACH_CONTENT_ID`
   (`0x3712001F`) on its `PropertyAccessor` (present ⇒ inline) and skip unless
   "include inline" is on; apply the extension filter; resolve the final name via
   `ReserveName` (null ⇒ skip per policy); then `att.SaveAsFile(full)` — the only
   mutating call. Count saved / skipped-inline / skipped-filter / skipped-clash /
   save-errors and accumulate bytes; keep up to 200 rows for the report table.
5. **Emit** the sentinel `__OLXTRACT__ STATUS=OK saved={N} scanned={M}` followed
   by the markdown/plain-text report (scope, output folder, filter, inline policy,
   clash policy, saved count + size + how many emails contributed, the skip
   breakdown, a capped per-file table) ending with the reminder that the emails
   were not modified.

### Outlook COM facts this relies on (see the office-com RAG)

* `olExplorer` may be **null** (no window open) — guard it.
* `Selection` and `Folder.Items` are **1-based** collections but `foreach`
  handles that; items vary in type, so access `.Attachments` via `dynamic` and
  try/catch.
* `Attachment.SaveAsFile(path)` writes (and silently overwrites) a file; it does
  **not** alter the email. `Attachment.FileName`/`.Size` are the display name and
  byte size.
* Inline vs. real attachment: both are `olByValue`; the distinguishing signal is
  the **content-id** MAPI property `PR_ATTACH_CONTENT_ID` (`0x3712001F`), read via
  `Attachment.PropertyAccessor.GetProperty(...)` inside a try/catch (it throws
  when the property is absent).

---

## 4. Files in this app

| File | Role |
|---|---|
| `attachment-extractor.scriptree` | The form definition (6 params; `download` icon embedded as PNG). Validate with `python -m scriptree.cli.validate <path>` from `D:\Dev\ScripTree`. |
| `attachment-extractor.scriptree.configs.json` | Config sidecar incl. the `standalone` end-user config (hides IDE chrome, pops up on success/error). |
| `attachment_extractor.py` | The Strategy-A shim: parses argv, renders the `.csx`, runs combridge, parses the sentinel, owns the exit code. |
| `attachment_extractor.csx.template` | The Roslyn C# script template with `__TOKEN__` placeholders the shim fills. |
| `README.md` | This file. |

### argv contract (shim ⇆ form)

```
attachment_extractor.py
  --scope selection|folder
  --output-folder <path>        (required)
  --extensions <str>            (comma/space list; blank = all)
  [--include-inline]            (flag)
  --on-name-clash rename|skip|overwrite
  --output-format markdown|text
```

The form's `argument_template` emits exactly this. The boolean uses the
`{id?--flag}` conditional form; `--output-folder` and `--extensions` use
`["--flag","{id}"]` token groups (single token regardless of whitespace, so a
path with spaces or an extension list stays one argv element; the empty
`--extensions` group drops out entirely).

---

## 5. Editing / maintenance notes

* **Preserve the read-only-mailbox contract.** Adding any call that deletes,
  moves, or saves an item would break the headline safety guarantee. The only
  permitted mutation is writing files to the output folder.
* **Validate after every edit:** from `D:\Dev\ScripTree`,
  `python -m scriptree.cli.validate <path>`.
* **Offline render-check:** render the template with sample values (a Windows
  path with a quote/backslash, an extension list with mixed case/tabs) and grep
  the output for `__[A-Z_]+__` — the only legitimate matches are the `__OLXTRACT__`
  sentinel and the `__PLACEHOLDER__` mention inside a template comment; any other
  `__TOKEN__` is an unfilled placeholder (a compile error waiting to happen).
* **Re-embedding the icon** does a load→save round-trip that drops hand-edited
  `min`/`max`/`step` numeric bounds — this app has no integer params so it's moot
  here, but keep the rule in mind if you add one.
* **combridge is located at run time** by walking up from the shim to
  `lib/combridge/combridge.exe` — never bake an absolute path; this repo does not
  bundle combridge.
