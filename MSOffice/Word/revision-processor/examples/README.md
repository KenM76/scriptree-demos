# Example — Tracked-Changes Processing Master

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the sample. |
| `sample_review.docx` | A document with **2 real tracked changes**: one tracked **insertion** (`<w:ins>`) and one tracked **deletion** (`<w:del>`), both authored by "Reviewer" dated 2026-01-01, plus a plain paragraph with no revisions. |

### How the sample was made (approach a — OOXML injection)

`python-docx` **cannot author tracked changes** — the `<w:ins>` (insertion) and
`<w:del>`/`<w:delText>` (deletion) revision elements have no high-level API. So
the generator builds the base document with python-docx, then re-opens the
`.docx` zip and **injects genuine `<w:ins>` / `<w:del>` markup** into
`word/document.xml`, then re-zips — the same OOXML zip-rewrite technique the
Link Auditor uses to inject a real external link. The generator then re-opens
the file and asserts both `<w:ins>` and `<w:del>` are present (and that
python-docx can still parse it). This makes `wdDoc.Revisions.Count` return `2`
in a live Word, so the tool has real revisions to process.

The seeded content:

| Paragraph | Revision |
|---|---|
| "The team will deliver the prototype…" | a tracked **insertion**: "This sentence was added by the reviewer." |
| "Budget remains within the approved envelope…" | a tracked **deletion**: "This obsolete clause should be removed." |
| "Please direct questions to the project office." | none (plain text — proves mixed content is handled). |

## How to try it

1. Open `sample_review.docx` in Word. You should see the two tracked changes
   (an inserted sentence underlined, a deleted sentence struck through). **Save
   it once** (it's freshly generated; the copy-mode default needs a saved file
   to put the copy beside).
2. Run **Tracked-Changes Processing Master** from ScripTree.

### Accept run (the defaults)

| Field | Value |
|---|---|
| Action | **Accept all changes** |
| Work on a copy | *(checked)* |
| Output format | `markdown` |

**Expected result** — a sibling `sample_review_Revisions_Processed.docx` in
which:

* the inserted sentence is now **permanent body text** (the insertion was
  accepted);
* the deleted sentence is **gone** (the deletion was accepted);
* **0 tracked changes remain** (`wdDoc.Revisions.Count == 0` on the copy);
* Track Changes is **off** on the copy;
* your original `sample_review.docx` is **untouched**.

The report shows `revisions=2 action=accept copy=1` and lists the document name,
the count processed, and the copy path.

### Reject variation

Run again with **Action = Reject all changes** (copy on). Expected sibling
`sample_review_Revisions_Processed.docx` in which:

* the inserted sentence is **removed** (the insertion was rejected);
* the deleted sentence is **restored** (the deletion was rejected);
* again **0 tracked changes remain** and tracking is off;
* original untouched.

The report shows `revisions=2 action=reject copy=1`.

### The "nothing to do" case

Run the tool on the *processed* copy (which now has 0 revisions), or on any
document without tracked changes. The tool reports **success** with
`revisions=0` and a "no tracked changes — nothing to do" message, makes **no
copy**, and does not modify the document. This is exit 0, not an error — a
document already free of revisions is the desired end state.

### In-place variation

Untick **Work on a copy**. The accept/reject is applied to the open document
**in memory and left unsaved** for you to review (Ctrl+Z still undoes it) before
saving over the original.

## What this demonstrates

* One-pass force-accept / force-reject of **all** tracked changes
  (`AcceptAllRevisions()` / `RejectAllRevisions()`).
* Track Changes turned **off** afterward so the processed document isn't left
  armed.
* The work-on-a-copy default (SaveAs2-repoint) so the original is never touched.
* The `revisions=0` success-no-op (exit 0, not a failure).
* The **inverse** relationship to the Style Sanitizer, which refuses
  revision-bearing documents.

> Like every app in this catalog, the tool is pending live verification against
> a real Word. The scenarios above describe the exact behaviour to expect; the
> sample carries genuine OOXML revision markup so `Revisions.Count` is real.
