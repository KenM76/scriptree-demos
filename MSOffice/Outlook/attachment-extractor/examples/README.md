# Example — Surgical Attachment Extractor

## Why there's no sample file here

This app takes **no input file**. Its input is your **live Outlook**: either
the messages you've **selected** in the explorer, or every message in the
**current folder**. There's nothing to hand you as a `sample.xxx` — the input
is whatever you have open and highlighted. So this folder is a **scenario
walkthrough**: a described inbox, the exact form values, and the report you'd
get back.

The tool is **surgical and read-only with respect to your mailbox**: the only
thing it ever writes is *new files on disk* (via `Attachment.SaveAsFile`). It
never deletes, edits, or moves a message. The emails are left exactly as they
were; only your output folder gains files.

## The scenario

Your Inbox has a `Vendors` folder. You select **four invoice emails** before
running the tool:

| # | From | Attachments |
|---|---|---|
| 1 | Acme Corp | `invoice.pdf` (240 KB), plus the sender's signature logo `logo.png` (inline) |
| 2 | Globex | `invoice.pdf` (180 KB) — **same filename** as #1 |
| 3 | Initech | `statement.pdf` (95 KB), `terms.docx` (40 KB) |
| 4 | Umbrella | `photo.jpg` (2.1 MB) — *not* an invoice doc |

Two things make this a good test: a **filename clash** (two `invoice.pdf`s) and
a mix of a wanted type (`pdf`), an unwanted type (`jpg`), and an **inline
signature image** (`logo.png`) that should *not* be treated as a real
attachment.

## How to try it (on your own mailbox)

1. In Outlook, open the folder and **select** the messages you want (or plan to
   use `folder` scope to take the whole folder).
2. Run **Surgical Attachment Extractor** from ScripTree.

| Field | Value | Effect |
|---|---|---|
| Which emails | `selection` | only the highlighted messages (the fast, surgical path; `folder` scans every item in the open folder) |
| Save attachments to | *(browse to an output folder, e.g. `C:\Invoices`)* | created if it doesn't exist |
| Only these extensions | `pdf, docx` | allow-list; leave blank to take **every** attachment |
| Include inline images (signature logos) | *(unchecked)* | inline/embedded body images are skipped — keeps signature logos out |
| If a file already exists | `rename` | keep both clashing files (`invoice (1).pdf`); other options are `skip` and `overwrite` |
| Output format | `markdown` | rendered report (`text` for plain monospace) |

## Expected result on disk

Output folder `C:\Invoices` now contains **four** new files:

```
invoice.pdf          (from Acme,    240 KB)
invoice (1).pdf      (from Globex,  180 KB)   ← renamed, clash policy = rename
statement.pdf        (from Initech,  95 KB)
terms.docx           (from Initech,  40 KB)
```

What was **not** saved:
* `logo.png` — inline signature image, skipped (Include inline = off).
* `photo.jpg` — extension not in the `pdf, docx` allow-list.

## Expected output (markdown mode)

```
# Attachment extraction

- **Scope:** 4 selected item(s)
- **Output folder:** `C:\Invoices`
- **Filter:** docx, pdf
- **Inline images:** skipped
- **Name clashes:** rename
- **Saved:** 4 attachment(s) (555 KB) from 3 of 4 email(s)
- **Skipped:** 1 inline image(s), 1 non-matching

| Saved file | From | Size |
|---|---|--:|
| invoice.pdf | Acme Corp | 240 KB |
| invoice (1).pdf | Globex | 180 KB |
| statement.pdf | Initech | 95 KB |
| terms.docx | Initech | 40 KB |

_The emails were not modified. Only new files were written to the output folder._
```

(Sizes/senders are illustrative; your run reports your real messages. "from 3
of 4" reflects that the Umbrella email contributed nothing after the filter.)

## Variations to try

* **Take everything:** clear **Only these extensions** → `photo.jpg` is now
  saved too (5 files).
* **Keep the signature logos:** tick **Include inline images** → `logo.png` is
  saved as well.
* **Clash = `skip`:** Globex's `invoice.pdf` is *not* written (the name is
  already taken); the report counts a "name-clash skip".
* **Clash = `overwrite`:** Globex's file **replaces** Acme's `invoice.pdf` —
  you end up with one `invoice.pdf` (the last one wins). Use with care.
* **Scope = `folder`:** processes every message in the open folder, not just
  the selection. Slower on big folders, since each item must be opened to read
  its attachments.

## What this demonstrates

* Batch-saving attachments from many emails in one pass, by selection or whole
  folder.
* The extension allow-list (dot-less, case-insensitive).
* Skipping **inline** signature/logo images by default (detected via the
  attachment's content-ID), with an opt-in to include them.
* The three name-clash policies: **rename** (keep both), **skip**, **overwrite**.
* The **read-only safety contract**: the only mutating call is writing new
  files to disk — messages are never altered.

> Pending live verification against a real Outlook profile; the scenario above
> describes the exact report shape and on-disk result to expect.
