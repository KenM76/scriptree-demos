# scriptree-demos

A growing catalogue of GUI wrappers for popular command-line tools, built with [**ScripTree**](https://github.com/KenM76/scriptree) — a universal GUI generator for CLI tools.

Each folder is a ready-to-use demo: drop the upstream tool's executable next to the `.scriptree` files (or edit the `executable` field), point ScripTree at the folder, and you get a labeled form with dropdowns, file pickers, and checkboxes wrapped around the CLI.

## How to use a demo

1. **Install [ScripTree](https://github.com/KenM76/scriptree)** if you haven't already.
2. **Pick a demo folder below** and clone or download it.
3. **Install the upstream tool** (each demo's README links to its upstream repo and install instructions).
4. **Drop the executable** in the demo folder (or edit the `executable` field in the `.scriptree` file to point wherever the tool lives on your machine).
5. **Open the `.scriptree` or `.scriptreetree`** in ScripTree.

That's it. No code, no flag-memorisation, no shell-quoting headaches.

## Index

| Tool | What it does | Folder | Upstream |
|------|--------------|--------|----------|
| **parfit** | Reflows prose inside code comments using optimal-fit line breaking, leaving machine-readable directives untouched. | [`parfit/`](parfit/) | [caldempsey/parfit](https://github.com/caldempsey/parfit) |
| **jit** | Jira CLI — ticket lookup, detail views, sprint lists, and create/edit from the terminal. | [`jit/`](jit/) | [cesarferreira/jit](https://github.com/cesarferreira/jit) |
| **agg** | Renders an asciinema `.cast` recording into an animated GIF. | [`agg/`](agg/) | [asciinema/agg](https://github.com/asciinema/agg) |
| **hyperfine** | Statistical benchmarking for shell commands — runs N times, controls warmup, exports tables. | [`hyperfine/`](hyperfine/) | [sharkdp/hyperfine](https://github.com/sharkdp/hyperfine) |
| **fd** | User-friendly `find` replacement — regex/glob patterns, smart filters, parallel walk. | [`fd/`](fd/) | [sharkdp/fd](https://github.com/sharkdp/fd) |
| **dust** | More intuitive `du` — sorted tree of largest entries with percent bars. | [`dust/`](dust/) | [bootandy/dust](https://github.com/bootandy/dust) |
| **bat** | `cat` clone with syntax highlighting, line numbers, Git integration, and an automatic pager. | [`bat/`](bat/) | [sharkdp/bat](https://github.com/sharkdp/bat) |
| **eza** | Modern, more-featureful `ls` with colour, icons, Git, tree view, and rich long-format columns. | [`eza/`](eza/) | [eza-community/eza](https://github.com/eza-community/eza) |
| **dog** | Friendly DNS lookup CLI — colour output, JSON, plus UDP/TCP/TLS/HTTPS transports. | [`dog/`](dog/) | [ogham/dog](https://github.com/ogham/dog) |
| **cvforge** | YAML → ATS-friendly PDF resume via Typst. Four subcommands (build / init / fonts / ats-check). | [`cvforge/`](cvforge/) | [SoAp9035/cvforge](https://github.com/SoAp9035/cvforge) |
| **bifrost** | Bridge files between computer and phone via QR code — single Go binary, LAN HTTP server, optional AES-256-GCM. | [`bifrost/`](bifrost/) | [axiom0x0/bifrost](https://github.com/axiom0x0/bifrost) |

More on the way — every useful CLI tool I run into is a candidate.

## Suggest a tool

Open an issue with a link to the tool's repo. Good candidates:

- Have a stable CLI surface (won't be renamed next week).
- Run as a one-shot command (no long-lived REPL or daemon).
- Have flags that map to typed fields — booleans, enums, paths, numbers.
- Solve a problem real people have today (Reddit threads, HN comments, "I wish there was a GUI for X").

Less-good candidates:

- Tools whose entire value is one-keystroke speed at the terminal (a form is slower than typing).
- Heavily interactive TUIs that don't take their input from flags.

## Contributing a demo

PRs welcome. Convention for new demos:

```
<tool-name>/
  README.md                   ← describes the upstream tool + this demo + links to author's repo
  <tool-name>.scriptree       ← the form
  <tool-name>.scriptreetree   ← optional, for grouping multiple subcommands
```

If the tool has multiple distinct subcommands (like jit's `lookup` / `show` / `create` / `edit`), give each its own `.scriptree` and group them under a `.scriptreetree`. See the [`jit/`](jit/) demo for the pattern.

Per-demo `README.md` should include:

- One-line description of the tool.
- Link to the upstream repo + author.
- What's in the demo (file table).
- Install instructions for the upstream tool (or link to upstream's install docs).
- Any caveats (e.g. repeatable flags exposed as raw text fields).

## A note on attribution and license

These demos generate forms around third-party CLI tools. The demo files (this repo) are MIT-licensed — the upstream tools are licensed by their respective authors and are **not** included or redistributed here. Each demo's README links back to the upstream project so you can install the actual tool from its author.

If you're the author of a tool wrapped here and have any concerns — or you'd like the demo updated to match a new release — please open an issue.

## License

MIT. See [LICENSE](LICENSE).
