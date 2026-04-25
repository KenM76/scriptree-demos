# bifrost — ScripTree demo

A ScripTree GUI for **[bifrost](https://github.com/axiom0x0/bifrost)** by [axiom0x0](https://github.com/axiom0x0).

## What bifrost does

bifrost is a single Go binary that bridges files between your computer and your phone via QR code — no cloud, no cables, no apps to install. Run it; it starts a local HTTP server on your LAN, prints a QR code, and your phone scans it to download (send mode) or upload (receive mode). Optional AES-256-GCM end-to-end encryption.

Three modes, exactly one per invocation:

- **Send a single file** — `bifrost -f path/to/file`
- **Browse a directory** — `bifrost -d path/to/folder`
- **Receive uploads** — `bifrost -r -o path/to/output`

## What's in this demo

| File | Purpose |
|------|---------|
| [`bifrost.scriptree`](bifrost.scriptree) | Single-form GUI. 6 fields across Send / Receive / Server sections. The three mode-selector controls (`-f`, `-d`, `-r`) are each in their own section as a reminder that you pick one and leave the others blank/unticked. |

The `.scriptree` expects `bifrost.exe` to live in the same folder. Drop it (or a symlink) next to the file, or edit the `executable` field.

## Caveat — mode is exclusive but the form doesn't enforce it

`-f` takes a file, `-d` takes a directory, `-r` takes nothing. ScripTree's `argument_template` only supports boolean-conditional emission, so we can't drive three different argv outputs from a single radio control — the three flags are exposed as three independent fields instead. Bifrost itself errors if you set more than one. The section headers ("Send (one of three modes)" / "Receive (one of three modes)") nudge you to pick exactly one.

## Why a GUI for a CLI tool?

`bifrost -f photo.jpg` is fast at the terminal. The GUI earns its keep when you're switching between modes a lot — pick a file from a picker, switch to "browse a directory" with a folder picker, swap to receive mode and choose where uploads land. Three modes, three widgets, no remembering which short flag is which.

## Installing bifrost

See the [upstream README](https://github.com/axiom0x0/bifrost#readme). Common paths:

```bash
go install github.com/axiom0x0/bifrost@latest
# or download a prebuilt release for your platform from:
# https://github.com/axiom0x0/bifrost/releases
```

## Upstream

- Repository: <https://github.com/axiom0x0/bifrost>
- Author: [@axiom0x0](https://github.com/axiom0x0)
- License: see upstream repo

This demo is independent of the bifrost project; it just generates a GUI front-end for the CLI. Bug reports about bifrost's behaviour belong upstream; bug reports about the form layout belong here.
