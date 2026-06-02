# Example — Heading-Based Document Splitter

## Files

| File | What it is |
|---|---|
| `make_example.py` | Generator. `python make_example.py` (re)creates the sample. Read it to see exactly what's in the document. |
| `sample_manuscript.docx` | The sample input: front matter + four Heading-1 sections, with two Heading-2 sub-sections inside "Methods". |

## How to try it

1. Open `sample_manuscript.docx` in Word (and **save it once** if you'll
   use the default output folder — a never-saved doc has no folder for the
   default `<name>_Chapters` to live in; opening this `.docx` from disk
   already satisfies that).
2. Run the **Heading-Based Document Splitter** from ScripTree.

### Run A — split at Heading 1 (the default)

Form values:

| Field | Value |
|---|---|
| Split at heading level | `1` |
| Output folder | *(leave blank)* |
| Output file format | `docx` |
| Save content before first heading | *(checked)* |
| Output format | `markdown` |

**Expected result** — a `sample_manuscript_Chapters` folder next to the
sample, containing **5** files:

```
00_Front-Matter.docx
01_Introduction.docx
02_Methods.docx          ← includes its Data Collection + Analysis sub-sections
03_Results.docx
04_Conclusion.docx
```

The report shows `chapters=5` (incl. front matter as 00), `docx=5`,
`pdf=0`, `failed=0`.

### Run B — split at Heading 2

Change **Split at heading level** to `2`. Now the boundaries are the two
Heading-2 paragraphs inside Methods, so you get:

```
00_Front-Matter.docx     ← everything before the first Heading 2
                            (title, abstract, Introduction, "Methods" intro)
01_Data Collection.docx
02_Analysis.docx
```

This is the finer-grained split: only Heading-2 paragraphs are cut points,
so the Heading-1 material that precedes the first Heading 2 all lands in
front matter.

### Run C — PDFs as well

Set **Output file format** to `both` and you get each section as both
`.docx` and `.pdf` (report shows `docx=5 pdf=5`).

## What this demonstrates

* Front matter capture (the pre-heading title/abstract → `00`).
* One-file-per-section at the chosen outline level.
* The same document splitting differently at level 1 vs level 2.
* The original `sample_manuscript.docx` is **never modified** — only read.
