# eza — ScripTree demo

A ScripTree GUI for **[eza](https://github.com/eza-community/eza)** — the community-maintained successor to `exa`.

## What eza does

eza is a modern, more-featureful `ls`. Sensible colours, file-type icons, Git integration, recursive tree view, and a rich long-format with optional columns for blocks, inodes, hard-links, security context, mounts, xattrs, and Git status. One-shot CLI — every flag is a presentation knob.

## What's in this demo

| File | Purpose |
|------|---------|
| [`eza.scriptree`](eza.scriptree) | Single-form GUI. ~55 fields across Target, Layout, Filtering & sort, Symlinks, Appearance, Long format, Time, and Git sections. Most long-format columns and time fields only render with `--long` ticked. |

The `.scriptree` expects `eza.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Why a GUI for a CLI tool?

`eza` and `eza -l` are faster typed than form-filled. The form earns its keep when you're chasing a specific view: long format, sorted by size descending, only directories, with Git status, header, binary sizes, and a relative time-style. That's a six-flag invocation that's easy to fumble at the prompt and trivial in a typed form.

## Screenshots

### Form view

![Form view of eza](eza_form.png)

### As it appears in the workspace forest

The cell on the right is this demo, docked to the workspace forest hub:

![eza cell docked to the forest](eza_forest.png)

## Installing eza

See the [upstream README](https://github.com/eza-community/eza#installation). Common paths:

```bash
cargo install eza
# or via package manager:
brew install eza
winget install eza-community.eza
```

## Upstream

- Repository: <https://github.com/eza-community/eza>
- Maintainers: [eza-community](https://github.com/eza-community)
- License: see upstream repo

This demo is independent of the eza project; it just generates a GUI front-end for the CLI. Bug reports about eza's behaviour belong upstream; bug reports about the form layout belong here.
