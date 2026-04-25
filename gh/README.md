# gh — ScripTree demo

A ScripTree GUI for **[GitHub CLI (gh)](https://github.com/cli/cli)** by [GitHub](https://github.com/cli).

## What gh does

`gh` is GitHub's official command-line client. It wraps the GitHub REST/GraphQL APIs in a thousand cobra-style subcommands — pull requests, issues, repos, releases, workflow runs, gists, projects, secrets, environments, raw API calls — and handles auth, pagination, and `--json + --jq` filtering for you.

The full surface is enormous. This demo wraps the eleven highest-frequency commands across four workflows.

## What's in this demo

Eleven scriptrees grouped by workflow under [`gh.scriptreetree`](gh.scriptreetree):

### Pull requests folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`pr-list.scriptree`](pr-list.scriptree) | `gh pr list` | List PRs with state / author / assignee / branch / label / search filters. |
| [`pr-view.scriptree`](pr-view.scriptree) | `gh pr view` | Show one PR (with optional comment thread). Defaults to the PR for the current branch. |
| [`pr-create.scriptree`](pr-create.scriptree) | `gh pr create` | Open a PR. Title/body/labels/reviewers + `--fill`, `--draft`, `--dry-run`. |
| [`pr-checkout.scriptree`](pr-checkout.scriptree) | `gh pr checkout` | Check out a PR locally. `--detach`, `--force`, `--recurse-submodules`. |
| [`pr-merge.scriptree`](pr-merge.scriptree) | `gh pr merge` | Merge a PR. Pick merge / squash / rebase, queue auto-merge, delete the branch. |

### Issues folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`issue-list.scriptree`](issue-list.scriptree) | `gh issue list` | List issues with state / author / assignee / mention / label / milestone / search filters. |
| [`issue-view.scriptree`](issue-view.scriptree) | `gh issue view` | Show one issue (with optional comments). |
| [`issue-create.scriptree`](issue-create.scriptree) | `gh issue create` | File a new issue. Body inline, from file, or via `$EDITOR`. |

### Repository folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`repo-view.scriptree`](repo-view.scriptree) | `gh repo view` | Show a repo's README and metadata. Defaults to the repo of the current working directory. |
| [`repo-clone.scriptree`](repo-clone.scriptree) | `gh repo clone` | Clone a repo (auth handled by gh — no SSH key juggling for private repos). |

### Releases folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`release-create.scriptree`](release-create.scriptree) | `gh release create` | Cut a release tied to a Git tag. Auto-generate notes, draft, prerelease, attach to a discussion. |

The scriptrees expect `gh` to be on `PATH` (i.e. installed via Homebrew, winget, apt, dnf, the official MSI, etc.). If it isn't, edit each file's `executable` field to point at the binary.

## What's intentionally not in this demo

A few subcommands were left out because the form factor adds little or none of the value:

- **`gh auth login`** — interactive flow with device code / browser callback. The CLI experience is the right experience.
- **`gh release upload`** — takes a positional list of asset filenames; ScripTree has no first-class repeatable-file widget. Run from the terminal.
- **`gh api`** — a raw HTTP client. Useful, but the form would amount to "type your URL and method here," which doesn't beat typing the command.
- **`gh workflow run`**, **`gh run list/view/watch`** — useful but a separate demo on its own; not yet wrapped here.

## Why a GUI for a CLI tool?

Most `gh` commands have `-R, --repo` plus another 5–15 flags. Memorising which flag is short-form which (e.g. `-A` for author vs. `-a` for assignee in the same command) is the friction point. The form factor names every flag, lays out the enum choices for `--state`, and gives you a file picker for `--body-file` / `--notes-file` instead of a relative path you have to type.

The two-click launch through the tree is also nice for the "look up one PR" / "list my open issues" / "cut a release" workflows where the flag combo is small but the chance of typing the wrong subcommand is real.

## Repeatable-flag convention

`gh` accepts repeated flags for `-l/--label`, `-a/--assignee`, `-r/--reviewer`. ScripTree's argument template emits one token per placeholder, so each repeatable flag has two fields: a typed widget for the first entry and a free-text field for "extra entries — write the full flag for each." This matches the convention used by parfit, fd, hyperfine, and dog in this repo.

## Installing gh

See the [official installation instructions](https://github.com/cli/cli#installation). Common paths:

```bash
# macOS
brew install gh

# Windows
winget install --id GitHub.cli

# Debian / Ubuntu
sudo apt install gh

# Fedora / RHEL
sudo dnf install gh
```

After install, run `gh auth login` once to authenticate.

## Upstream

- Repository: <https://github.com/cli/cli>
- Maintainer: [GitHub](https://github.com/cli)
- License: see upstream repo (MIT)

This demo is independent of the gh project; it just generates a GUI front-end for the CLI. Bug reports about gh's behaviour belong upstream; bug reports about the form layout belong here.
