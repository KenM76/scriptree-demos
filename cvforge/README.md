# cvforge — ScripTree demo

A ScripTree GUI for **[cvforge](https://github.com/SoAp9035/cvforge)** by [Ahmet Burhan Kayalı](https://github.com/SoAp9035).

## What cvforge does

cvforge turns a YAML CV definition into a clean, ATS-friendly PDF using [Typst](https://typst.app/). One command renders; another scaffolds a starter YAML; another lists the fonts you can pick from inside the YAML; another inspects an existing PDF for ATS-parseability.

## What's in this demo

Four scriptrees grouped by workflow under [`cvforge.scriptreetree`](cvforge.scriptreetree):

### Render folder

| File | Purpose |
|------|---------|
| [`build.scriptree`](build.scriptree) | `cvforge build [input]` — render the YAML CV to PDF. Optional file picker; defaults to `cv.yaml` in the current directory if blank. |
| [`init.scriptree`](init.scriptree) | `cvforge init [directory]` — create a starter `cv.yaml` template populated with every available field. Optional folder picker; defaults to the current directory. |

### Inspect folder

| File | Purpose |
|------|---------|
| [`fonts.scriptree`](fonts.scriptree) | `cvforge fonts` — list every font name CVForge can use. No arguments. |
| [`ats-check.scriptree`](ats-check.scriptree) | `cvforge ats-check <pdf>` — verify a PDF parses correctly through typical Applicant Tracking Systems. Required file picker. |

The scriptrees expect `cvforge` to be on `PATH` (i.e. installed via `pip install cvforge` or `uv tool install cvforge`). If it's not, edit each `executable` field to point at the binary.

## Why a GUI for a CLI tool?

`cvforge build` and `cvforge init` are honestly faster typed than form-filled. The GUI's value is the four-button launcher — `init` to scaffold, `build` to render, `ats-check` to validate, `fonts` to browse. One click each, no remembering the subcommand names.

## Installing cvforge

See the [upstream README](https://github.com/SoAp9035/cvforge#readme). Common paths:

```bash
pip install cvforge
# or
uv tool install cvforge
```

Requires Python 3.10+.

## Upstream

- Repository: <https://github.com/SoAp9035/cvforge>
- Author: [@SoAp9035](https://github.com/SoAp9035) (Ahmet Burhan Kayalı)
- License: see upstream repo

This demo is independent of the cvforge project; it just generates a GUI front-end for the CLI. Bug reports about cvforge's behaviour belong upstream; bug reports about the form layout belong here.
