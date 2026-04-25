# agg — ScripTree demo

A ScripTree GUI for **[agg](https://github.com/asciinema/agg)** by [asciinema](https://github.com/asciinema).

## What agg does

agg renders an [asciinema](https://asciinema.org/) terminal-session recording (`.cast` file) into an animated GIF. One-shot CLI — point it at a cast file, give it an output path, and it writes the GIF.

## What's in this demo

| File | Purpose |
|------|---------|
| [`agg.scriptree`](agg.scriptree) | Single-form GUI for agg. Two required positional fields (input cast, output GIF) plus four sections of optional controls: Appearance, Timing, Geometry, Logging. Dropdowns for theme and renderer; numeric fields for font-size, speed, fps-cap, idle-time-limit, etc. |

The `.scriptree` expects `agg.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field to point wherever you've installed it.

## Repeatable flag caveat

`--font-dir` is repeatable upstream, but ScripTree's current schema has no first-class repeatable-flag widget — this demo exposes a single folder picker. If you need multiple font directories, edit the `.scriptree` to add additional entries or call `agg` directly.

## Installing agg

See the [upstream README](https://github.com/asciinema/agg#installation) for current install instructions. Common paths:

```bash
cargo install --git https://github.com/asciinema/agg
# or via the prebuilt Docker image:
docker run --rm -v $PWD:/data ghcr.io/asciinema/agg demo.cast demo.gif
```

## Upstream

- Repository: <https://github.com/asciinema/agg>
- Author: [asciinema project](https://github.com/asciinema)
- License: see upstream repo

This demo is independent of the agg project; it just generates a GUI front-end for the CLI. Bug reports about agg's behaviour belong upstream; bug reports about the form layout belong here.
