# MSOffice — ScripTree apps for Word, Excel, PowerPoint & Outlook

> ⚠️ **BETA — pending live verification.** These apps are built and pass
> ScripTree's schema validation and an offline render-check of their generated
> scripts, but — except the Excel **Broken External Link Auditor**, which has
> been run against a live Excel — they have **not yet been verified against a
> real, running Office instance**. Treat them as beta: try them on copies, read
> each app's README, and expect the occasional rough edge. Bug reports welcome.

Unlike the CLI-wrapper demos in this repo, these don't wrap a command-line tool.
They drive **Microsoft Office via COM automation** through
[**combridge**](https://github.com/KenM76/combridge), so they act on the
document / workbook / presentation / mailbox you already have **open** in Office.

## How they're built (the "Strategy A" pattern)

Each app is a tiny, inspectable pipeline:

```
ScripTree form (.scriptree)
   → python shim (<tool>.py)              # takes the form values as argv
       → renders a C# script from <tool>.csx.template   # token replacement
           → combridge <plugin> run-script <tmp>.csx     # runs against live Office
```

combridge's `run-script` has no argv channel, so the shim **bakes** the form
values into a generated `.csx`, runs it through combridge, parses a first-line
sentinel from the script's output, and owns the process exit code. Everything is
plain text you can read: the form, the Python shim, and the C# template.

## What you need

1. **[ScripTree](https://github.com/KenM76/scriptree)** installed, with its
   bundled **`lib/combridge/combridge.exe`** (the shims locate combridge by
   walking up from their own folder — deploy these app folders **inside** a
   ScripTree install so that lookup succeeds).
2. **Python** (the shim runs as `executable: "python"`).
3. The **target Office app running** with the relevant item open (a document /
   workbook / presentation, or a mailbox for the Outlook tools).

## Safety model

- **Mutators default to working on a copy.** The Word/PowerPoint editors save a
  sibling (`<name>_Sanitized.docx`, `<name>_Restyled.pptx`, …) and make all
  changes there, leaving your original untouched; the Excel freezer writes a
  `_Frozen.xlsx` copy. Untick "Work on a copy" to edit in place (left unsaved
  for you to review).
- **Read-only tools never write to your data** — the auditors only read.
- **The Outlook Draft Generator only saves drafts — it never sends.**
- Several tools refuse risky preconditions (e.g. the Style Sanitizer aborts on a
  document with tracked changes).

## The apps

### Word
| App | What it does | Touches your file? |
|---|---|---|
| [`Word/batch-find-replace/`](Word/batch-find-replace/) | Find & replace every occurrence in one pass, with a clean count; match-case / whole-word / wildcards. | Mutator (copy by default) |
| [`Word/style-sanitizer/`](Word/style-sanitizer/) | Force a clean corporate look: strip rogue direct formatting back to styles, tidy whitespace, normalise fonts (or copy styles from a template doc), curly quotes, plus opt-in bullet/table-shading/per-side-margin passes. | Mutator (copy by default) |
| [`Word/heading-splitter/`](Word/heading-splitter/) | Split a long doc at every Heading-N into one `.docx`/`.pdf` per section. | Creator (reads only; writes new files) |
| [`Word/revision-processor/`](Word/revision-processor/) | Accept-all or reject-all tracked changes in one pass. | Mutator (copy by default) |

### Excel
| App | What it does | Touches your file? |
|---|---|---|
| [`Excel/link-auditor/`](Excel/link-auditor/) | Find broken/external workbook links and report OK / BROKEN / OFFLINE. ✅ *live-verified* | Read-only |
| [`Excel/sheet-segregator/`](Excel/sheet-segregator/) | Split one sheet into many by a key column (new workbook by default). | Mutator (new workbook / backup guard) |
| [`Excel/csv-aggregator/`](Excel/csv-aggregator/) | Import a folder of CSV/TXT files as new sheets, auto-detecting each file's delimiter. | Mutator (backup guard) |
| [`Excel/formula-freezer/`](Excel/formula-freezer/) | Convert all formulas to static values onto a `_Frozen.xlsx` copy; aborts on protected sheets. | Mutator (copy by default) |

### PowerPoint
| App | What it does | Touches your file? |
|---|---|---|
| [`PowerPoint/typography-enforcer/`](PowerPoint/typography-enforcer/) | Standardise fonts across a deck (swap one font or all), preserving symbol/icon fonts; or pull fonts/theme from a template deck. | Mutator (copy by default) |
| [`PowerPoint/hidden-assets-purger/`](PowerPoint/hidden-assets-purger/) | Strip speaker notes, comments, and metadata; optionally delete hidden slides. | Mutator (copy by default) |
| [`PowerPoint/deck-slimmer/`](PowerPoint/deck-slimmer/) | Remove unused slide layouts/masters to shrink a deck. (Does **not** recompress images — PowerPoint exposes that only via an interactive dialog.) | Mutator (copy by default) |
| [`PowerPoint/image-deck-builder/`](PowerPoint/image-deck-builder/) | Build a new deck from a folder of images, one per slide, scaled to fit. | Creator (writes a new deck) |

### Outlook
| App | What it does | Touches your mailbox? |
|---|---|---|
| [`Outlook/mailbox-auditor/`](Outlook/mailbox-auditor/) | Per-folder item count, size, age span, and "dead weight" — across the live mailbox **and** attached PSTs. | Read-only |
| [`Outlook/attachment-extractor/`](Outlook/attachment-extractor/) | Save attachments from selected emails (or a folder) to disk in one pass; never modifies the emails. | Read-only (writes files to disk) |
| [`Outlook/draft-generator/`](Outlook/draft-generator/) | Mail-merge from an Excel list + a body template into **drafts** — `Save()` only, **never sends**, capped at 50/run. | Creates drafts only |

## Examples

Every app has an `examples/` folder with a `make_example.py` generator, a README
stating the exact form values to use and the expected result, and (for the
file-based apps) the input fixtures. The generated sample documents themselves
aren't committed — **run `python examples/make_example.py`** in an app's folder
to (re)create them locally. The two Outlook apps ship a scenario README instead
(they read a live mailbox, so there's no input file to hand you).

## License

Same as the rest of this repo — see [`../LICENSE`](../LICENSE).
