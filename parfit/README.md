# parfit — ScripTree demo

A ScripTree GUI for **[parfit](https://github.com/caldempsey/parfit)** by [Cal Dempsey](https://github.com/caldempsey).

![parfit in ScripTree](screenshot.png)

## What parfit does

parfit reflows the prose inside source-code comments to a target column width using optimal-fit (Knuth-Plass-style) line breaking, while passing machine-readable directives — rustdoc attributes, doctest fences, `//#` / `#!`-style markers, shebangs, etc. — through unchanged.

It recognises comment syntax for: rust, python, shell, elixir, go, javascript, java, scala, c, lua, sql, lisp (plus a `text` mode for plain-text reflow).

## What's in this demo

| File | Purpose |
|------|---------|
| [`parfit.scriptree`](parfit.scriptree) | Single-form GUI for parfit. 11 params across 5 sections (Target, Formatting, Language, Filtering, Config). Defaults to `--stdout` so first-time runs are non-destructive. |
| [`parfit.scriptreetree`](parfit.scriptreetree) | Tree wrapper that surfaces `parfit.scriptree` under a "Reflow comments" folder — useful when you want the standalone, single-tool launcher view. |

The `.scriptree` expects `parfit.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field to point wherever you've installed it.

## Repeatable flags

`--include`, `--exclude`, and `--skip` are exposed as single text fields where you write the full repeated flag for each entry. ScripTree's argv emitter shlex-tokenizes string-typed fields whose placeholder fills the whole template token, so:

```
--include src/**/*.rs --include build.rs
```

becomes the four argv tokens you'd expect at run time. Quote rules are honored — `--include "path with spaces"` stays as two tokens, not three.

See `help/LLM/argument_template.md` (string-passthrough auto-split) for the full rule. ScripTree v0.1.3+ ships this behavior; earlier versions emitted the multi-flag string as one giant argument.

## Screenshots

### Tabbed standalone view

Every tool on its own tab — what you get when you single-click the tree's cell:

![Tabbed view of parfit](parfit_tabs.png)

### Forest menu

Double-click the cell in the workspace forest to get the merged tree as a popup menu:

![Forest menu for parfit](parfit_menu.png)

### Workspace cell

Docked to the forest hub:

![parfit cell docked to the forest](parfit_forest.png)

### Per-tool forms

#### parfit — reflow code comments

![parfit form](parfit_form.png)

## Installing parfit

From the upstream README:

```bash
cargo install parfit
# or
brew install caldempsey/parfit/parfit
```

## Upstream

- Repository: <https://github.com/caldempsey/parfit>
- Author: [@caldempsey](https://github.com/caldempsey)
- License: see upstream repo

This demo is independent of the parfit project; it just generates a GUI front-end for the CLI. Bug reports about parfit's behaviour belong upstream; bug reports about the form layout belong here.
