# Heading-Based Document Splitter (Word)

Cut a long Word document into one file per section, splitting at every
paragraph set to a chosen heading level (Heading 1, 2, or 3). Each
section is written to its own `.docx` and/or `.pdf` in a target folder,
named after the heading text with a zero-padded sequence prefix
(`01_Introduction.docx`, `02_Methods.docx`, …). **The document you have
open is never modified.**

This document is written to the project's reconstruct-from-docs bar: a
competent engineer or LLM should be able to rebuild the tool from this
README alone. It explains the *logic*, the *contracts*, the *safety
model*, and the *edge cases* — not just the surface behaviour.

---

## 1. What it does, and when to use it

You have a long `.docx` open in Word — a book manuscript, a compiled
report, a policy binder, a thesis. You want each top-level section as a
standalone file: one chapter per `.docx`, or a folder of per-section
PDFs to hand out. You run this tool from ScripTree, pick the heading
level to cut on, and it emits one file per section.

* **Split at Heading 1** *(default)* — one file per chapter / top-level
  section. The most common case.
* **Split at Heading 2 or 3** — finer granularity (one file per
  sub-section). Useful when a document's real "unit of distribution" is
  a sub-heading.

Each output file is named `<NN>_<Heading text>.<ext>`, where `NN` is a
zero-padded sequence number (`01`, `02`, … — width grows to `001` if a
document has 100+ sections) and the heading text is sanitised into a
legal Windows filename (see §6.4).

**Front matter** (everything before the *first* heading — a title page,
a table of contents, an abstract) is, by default, emitted as
`00_Front-Matter.<ext>` so nothing is lost. Turn this off if you only
want the headed sections.

### Requirements

* The target document is **already open** in a running Word instance.
* If you don't pick an output folder, the document must have been
  **saved at least once** (the default output folder is a
  `<name>_Chapters` folder next to the original, which needs a path on
  disk). If you *do* pick an output folder, an unsaved document is fine.
* `combridge` is available — bundled with ScripTree at
  `lib/combridge/combridge.exe`. The shim finds it by walking up the
  directory tree (see §7); no absolute path is baked in.

---

## 2. The output: a folder of section files, original untouched

The tool writes into a target folder (see §3 for how it's resolved):

* `00_Front-Matter.docx` *(only if front matter exists and the option is
  on)*
* `01_<First heading>.docx`, `02_<Second heading>.docx`, … one per
  section at the chosen level.
* The same set as `.pdf` if you picked **PDF** or **both**.

**Your open document is provably untouched** — the tool only *reads*
ranges from it and *creates* new documents; it never calls `Save`/
`SaveAs2` on your document (see §5). Review the folder; if you're happy,
keep it; if not, delete it — your original is exactly as it was.

The output pane shows a summary (Markdown or plain text, your choice):
the source name, the level split on, the output folder, the file format,
how many chapters were written, how many files were created, and any
per-section failures.

---

## 3. The form (5 parameters)

The form groups its fields into two collapsible sections — **Split** and **Output**
(fields render grouped by section, so the on-screen order follows the table below).

| Section | id | label | type / widget | default | meaning |
|---|---|---|---|---|---|
| Split | `heading_level` | Split at heading level | `enum` / `radio` | `1` | `1`, `2`, or `3`. Paragraphs at this **outline level** become section boundaries (see §6.1). |
| Split | `include_front_matter` | Save content before first heading | `boolean` / `checkbox` | `true` | Emit pre-first-heading content as `00_Front-Matter`. |
| Output | `output_folder` | Output folder | `path` / `folder` | `""` | Where to write the section files. **Leave blank** to use a `<name>_Chapters` folder next to the original. |
| Output | `file_format` | Output file format | `enum` / `radio` | `docx` | `docx` (editable), `pdf` (fixed), or `both`. |
| Output | `output_format` | Output format | `enum` / `dropdown` | `markdown` | `markdown` or `text` rendering of the summary. |

### Output-folder resolution (3 cases)

1. **You pick a folder** → that folder is used verbatim (created if it
   doesn't exist).
2. **Blank + document has been saved** → a sibling
   `<DocumentName>_Chapters` folder next to the original.
3. **Blank + document never saved** → refused with `UNSAVED` (there is
   nowhere to put a default folder). Either save once (Ctrl+S) or pick a
   folder.

### argv contract (argument_template → shim)

```
split_by_heading.py
["--heading-level", "{heading_level}"]
["--output-folder", "{output_folder}"]      # token group DROPS when blank
["--file-format",   "{file_format}"]
"{include_front_matter?--include-front-matter}"   # emits the flag only when checked
["--output-format", "{output_format}"]
```

The `["--output-folder", "{output_folder}"]` token **group** disappears
entirely when `output_folder` is empty (ScripTree's drop-on-empty rule),
so the shim's `argparse` falls back to its default `""` — which the
`.csx` reads as "use the default `<name>_Chapters` folder." The
`{include_front_matter?--include-front-matter}` conditional emits the
flag only when the box is checked.

---

## 4. Exit-code contract (sentinel → exit code)

combridge's `run-script` **ignores the C# `return` value** (it exits 0 on
any clean run) and offers **no argv channel**. So the generated `.csx`
writes a **first-line sentinel** that the Python shim parses, and the
**shim** owns the exit code ScripTree finally sees:

| Sentinel first line | Exit | Meaning |
|---|---|---|
| `__WORDSPLIT__ STATUS=NODOC` | **2** | No document is open in Word. |
| `__WORDSPLIT__ STATUS=NOHEADINGS` | **2** | No paragraph at the chosen heading level → nothing to split on. |
| `__WORDSPLIT__ STATUS=UNSAVED` | **2** | No output folder given **and** the document was never saved → nowhere to create the default folder. |
| `__WORDSPLIT__ STATUS=BADFOLDER` | **2** | The output folder could not be created. |
| `__WORDSPLIT__ STATUS=OK chapters=… docx=… pdf=… failed=… front=…` | **0** | Success. |
| *(combridge's own codes)* | **3 / 4 / 5** | Compile error / script threw / host (no Word) — passed through verbatim. |

A run with **one** chapter is still exit 0 (a single Heading-1 document
legitimately splits into one file). A run where **some** sections failed
to write but others succeeded is **still exit 0** — the partial result is
real output; the failures are listed in the report (`failed=N`). The four
guarded refusal statuses are the only `STATUS=` values that map to exit 2.

---

## 5. Safety model — why the open document is never touched

This tool is **non-destructive by construction**, and the reason is
worth stating precisely because it differs from the project's *mutator*
apps (e.g. Batch Find & Replace), which need an explicit "work on a copy"
guard.

* The tool **only reads** from the open document: it enumerates
  `wdDoc.Paragraphs` to find heading boundaries and reads `wdDoc.Range(start, end)`
  spans.
* The tool **only creates new documents** for output:
  `wdApp.Documents.Add(Visible: false)`, fills them, saves *them*, closes
  *them*.
* The tool **never** calls `Save`, `SaveAs2`, or `ExportAsFixedFormat`
  on `wdDoc`.

So there is *no code path* that could write back to the user's document —
no copy guard is needed because nothing can mutate the original in the
first place. Even if the script throws midway through emitting chapters,
the worst case is a partially-populated output folder; the open document
is untouched.

`wdApp.DisplayAlerts = WdAlertLevel.wdAlertsNone` is set around the
emit loop and restored in a `finally`. A COM-attached Word **hangs
invisibly** on any modal prompt (overwrite confirmation, compatibility
notice, PDF-export dialog) — the same hazard documented for headless
Excel/PowerPoint — so alerts must be suppressed while saving/exporting.

Guards before any work:

* `wdDoc is null` → `NODOC`.
* blank output folder **and** `wdDoc.Path == ""` → `UNSAVED`.
* `Directory.CreateDirectory(targetDir)` throws → `BADFOLDER`.
* zero heading paragraphs at the chosen level → `NOHEADINGS`.

---

## 6. How the `.csx` splits the document (the core logic)

### 6.1 Split on OUTLINE LEVEL, not style name

A paragraph is a section boundary when its **outline level** equals the
chosen level — **not** when its style happens to be named "Heading 1".
Two reasons:

1. **Style names are localised.** The built-in heading style is "Heading
   1" in English, "Überschrift 1" in German, "Titre 1" in French. Matching
   on the name breaks on non-English documents.
2. **Outline level is the intent-accurate signal.** Word lets a user
   promote a body paragraph to an outline level *without* applying the
   heading style. Outline level captures "this is a level-1 heading"
   regardless of how it was achieved.

Word's enum maps cleanly: `wdOutlineLevel1..9` = `1..9`,
`wdOutlineLevelBodyText` = `10`. So a chosen level of `1`/`2`/`3`
compares directly against `(int)paragraph.OutlineLevel`. Each paragraph's
`OutlineLevel` read is wrapped in try/catch — a paragraph that throws
(rare, e.g. inside a malformed table) is skipped, never fatal.

### 6.2 Building the boundary list

```
heads = []                          # (Start char position, cleaned Title)
for each Paragraph p in wdDoc.Paragraphs:
    if (int)p.OutlineLevel == headingLevel:
        heads.add( (p.Range.Start, CleanTitle(p.Range.Text)) )

if heads is empty:  -> NOHEADINGS
```

### 6.3 Building the chapter ranges

```
docStart = wdDoc.Content.Start ;  docEnd = wdDoc.Content.End
chapters = []

# Front matter: everything before the first heading, if non-empty and enabled.
if includeFront and heads[0].Start > docStart:
    pre = wdDoc.Range(docStart, heads[0].Start)
    if pre.Text is not whitespace:
        chapters.add( (docStart, heads[0].Start, "Front-Matter") )

hasFront = (chapters.Count > 0)

# Each heading spans from its own Start to the NEXT heading's Start
# (or the end of the document for the last heading).
for i in 0 .. heads.Count-1:
    start = heads[i].Start
    end   = heads[i+1].Start  if i+1 < heads.Count  else docEnd
    chapters.add( (start, end, heads[i].Title) )
```

The sequence number is `00` for front matter (when present) and `01..N`
for the headed sections: `seq = hasFront ? i : i+1`. The zero-pad width
is `max(2, len(str(chapters.Count)))` so it's always ≥2 digits and grows
for very large documents.

### 6.4 Emitting each chapter (clipboard-free copy)

For each chapter we copy the source range into a fresh invisible document
using **`FormattedText`**, *not* the clipboard:

```csharp
Range src = wdDoc.Range(chapter.Start, chapter.End);
Document nd = wdApp.Documents.Add(Visible: false);
nd.Content.FormattedText = src.FormattedText;   // preserves char + para formatting
if (wantDocx) nd.SaveAs2(path + ".docx", WdSaveFormat.wdFormatXMLDocument);
if (wantPdf)  nd.ExportAsFixedFormat(path + ".pdf", WdExportFormat.wdExportFormatPDF);
nd.Close(WdSaveOptions.wdDoNotSaveChanges);      // in a finally
```

* **Why `FormattedText` and not Copy/Paste?** Clipboard Copy/Paste is
  racy and unreliable in a headless COM Word (the clipboard is a shared,
  asynchronous OS resource). `dest.FormattedText = src.FormattedText`
  transfers character + paragraph formatting directly, with no clipboard
  involvement.
* **What is intentionally *not* carried:** section-level constructs —
  running headers/footers, page-number fields tied to the original
  section. A chapter split wants the *body* content of each section, not
  the original document's running headers. (If you need those, that's a
  different tool.)
* Each chapter is wrapped in try/catch: a section that fails to write is
  recorded in the `failures` list and counted (`failed++`), but the loop
  continues so the other chapters still come out. The new document is
  always closed in a `finally`, even on failure, so no orphan invisible
  documents leak.

### 6.5 Filename construction (`MakeFileName`)

```
title = title or "Untitled"
strip Windows-illegal chars  \ / : * ? " < > |   -> space
collapse whitespace runs -> single space ; trim
cap title to 60 chars
stem = "<seq>_<title>"
de-dupe: if stem already used, append " (2)", " (3)", … (case-insensitive set)
```

`CleanTitle` (applied when collecting headings) first turns a raw heading
paragraph into a human title: every control character (including the
trailing paragraph mark `\r` and cell mark `\a`) becomes a space, runs of
whitespace collapse to one, and the result is trimmed. So a heading like
`"  Chapter\tOne\r"` becomes `Chapter One`, and the file stem becomes
`01_Chapter One.docx`.

---

## 7. Files in this app

| File | Role |
|---|---|
| `split_by_heading.py` | Strategy-A shim. Parses argv, escapes values, renders the `.csx` from the template, runs combridge, parses the sentinel, owns the exit code. |
| `split_by_heading.csx.template` | Roslyn script template. Token placeholders (`__HEADING_LEVEL__`, `__OUTPUT_FOLDER__`, `__FILE_FORMAT__`, `__INCLUDE_FRONT_MATTER__`, `__OUTPUT_FORMAT__`) are filled by the shim. Contains all the COM logic above. |
| `heading-splitter.scriptree` | The ScripTree form (5 params). Carries the embedded `scissors` icon (cut/split convention). |
| `heading-splitter.scriptree.configs.json` | Sidecar with the end-user `standalone` config (IDE chrome hidden, error/success popups on). |
| `README.md` | This document. |

### Token-substitution + escaping contract

The shim's `render_csx` replaces the five `__TOKEN__`s. String values
(`heading_level`, `output_folder`, `file_format`, `output_format`) go
through `csharp_literal` — backslash escaped **first**, then `"`, `\r`,
`\n`, `\t` — so an output-folder path (full of backslashes, possibly with
a quote) can't break out of the C# string literal. This escaping is the
critical one for this tool: a Windows path like
`C:\Users\Ken\My "Book"\Chapters` becomes the safe literal
`"C:\\Users\\Ken\\My \"Book\"\\Chapters"`. The boolean
(`include_front_matter`) is rendered as the bare C# literal `true`/`false`
via `cs_bool`. After substitution the only `__…__` token remaining in the
generated `.csx` is the `__WORDSPLIT__` sentinel (verified by an offline
render-check).

---

## 8. Maintenance notes

* **Supporting heading levels 4–9:** the split logic already works for any
  level (it compares `(int)OutlineLevel == headingLevel`). Only the form's
  `heading_level` choices (`1`/`2`/`3`) and the shim's `argparse choices`
  cap it at 3 — widen both to expose deeper levels.
* **Carrying headers/footers:** intentionally omitted (§6.4). Adding them
  means copying `Section.Headers`/`Footers` from the source range's
  section into the new document — substantially more complex and usually
  unwanted for a chapter split.
* **Changing the default-folder suffix:** the `_Chapters` suffix is built
  in the `.csx` (`baseName + "_Chapters"`). Keep the
  read-only-on-`wdDoc` invariant intact if you touch the emit path.
* **Validate after any form edit:** from `D:\Dev\ScripTree`, run
  `python -m scriptree validate <path>` (PowerShell:
  `Set-Location D:\Dev\ScripTree; python -m scriptree validate "<path>"`).
