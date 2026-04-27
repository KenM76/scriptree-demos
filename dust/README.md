# dust — ScripTree demo

A ScripTree GUI for **[dust](https://github.com/bootandy/dust)** by [Andy Boot](https://github.com/bootandy).

## What dust does

dust is a more intuitive `du`. It walks a directory and prints a sorted, percentage-bar-annotated tree of the largest entries — file or directory. One-shot CLI: defaults are sensible, every flag is a presentation knob (depth, count, units, sort order, filters, colors).

## What's in this demo

| File | Purpose |
|------|---------|
| [`dust.scriptree`](dust.scriptree) | Single-form GUI. ~25 fields across Target, Display, Filtering, Counting, and Style sections. |

The `.scriptree` expects `dust.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Repeatable flags

`-X`/`--ignore-directory` is repeatable upstream. It's exposed as a single text field where you write the full repeated flag and ScripTree splits it into multiple argv tokens at run time:

```
-X node_modules -X target -X .git
```

(See `help/LLM/argument_template.md`'s string-passthrough auto-split rule. v0.1.3+.)

## Why a GUI for a CLI tool?

`dust` with no args is honestly perfect — that's the whole point of the tool. The GUI earns its keep when you start tuning: hide files under 10M (`-z 10M`), show only the top 50 (`-n 50`), max depth 3 (`-d 3`), files-only (`-F`), and force GiB output (`-o gib`). Five flags, five widgets, no man-page round-trip.

## Installing dust

See the [upstream README](https://github.com/bootandy/dust#install). Common paths:

```bash
cargo install du-dust
# or via package manager:
brew install dust
winget install bootandy.dust
```

## Upstream

- Repository: <https://github.com/bootandy/dust>
- Author: [@bootandy](https://github.com/bootandy)
- License: see upstream repo

This demo is independent of the dust project; it just generates a GUI front-end for the CLI. Bug reports about dust's behaviour belong upstream; bug reports about the form layout belong here.
