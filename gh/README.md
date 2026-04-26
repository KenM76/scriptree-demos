# gh — ScripTree demo

A ScripTree GUI for **[GitHub CLI (gh)](https://github.com/cli/cli)** by [GitHub](https://github.com/cli).

## What is GitHub?

GitHub is the website where most of the world's open-source software (and a lot of private software) lives. It hosts code, tracks bugs and feature requests, lets multiple people collaborate on the same code without overwriting each other, and packages finished versions for download. `gh` is the command-line tool GitHub publishes for talking to all of that from a terminal — no browser tab, no clicking, no copy-paste.

## Glossary (for the unfamiliar)

If GitHub is new to you, these are the terms you'll see in the forms:

- **Repository (or "repo")** — A single project on GitHub. Holds the code, the history of every change, the bug tracker, and everything else. Identified as `OWNER/REPO`, e.g. `cli/cli` or `microsoft/vscode`.
- **Clone** — Download a copy of a repository to your computer so you can read it, run it, or change it locally. The local copy stays linked to the GitHub original so you can pull down updates and (if you're a contributor) push your changes back.
- **Branch** — A parallel line of development inside a repository. The "main" branch is the official version. When you want to make a change without disturbing main, you create a side branch, work on it, then merge it back. The default branch is usually called `main` (older repos: `master`).
- **Commit** — A saved snapshot of changes. Has a message ("Fix login bug") and a unique ID. Branches are made of commits.
- **Pull request (PR)** — A formal proposal to merge one branch's changes into another (usually a side branch into main). Comes with a title, description, discussion thread, code review, and automated checks. The standard way contributors propose changes.
- **Merge** — Combine the commits from a side branch into a target branch. There are three styles:
  - **Merge commit**: keep all the side-branch commits and add a special "merge" commit on top of main showing they came in together.
  - **Squash**: smash all the side-branch commits into a single commit on main. Cleanest history; loses fine-grained step-by-step.
  - **Rebase**: replay each side-branch commit on top of main one by one. Looks like the work happened linearly with no branching.
- **Draft** — A pull request marked "not ready yet". You can still push commits to it, but reviewers are signalled not to start reviewing yet.
- **Issue** — A public conversation thread about a bug, a feature request, or a question. Issues have titles, descriptions, labels, and a comment thread. They're closed when resolved.
- **Label** — A coloured tag attached to issues or PRs (e.g. "bug", "good first issue", "priority:high"). Used for filtering and triage.
- **Milestone** — A grouping for issues and PRs aimed at a particular goal or version (e.g. "v2.0", "Q3 2026").
- **Reviewer** — A person whose code-review approval is requested before a PR can be merged.
- **Assignee** — The person responsible for working on an issue or PR.
- **Tag** (Git tag) — A named pointer to a specific commit. Commonly used to mark release points like `v1.2.3`.
- **Release** — A polished, downloadable version of the project, tied to a tag. Has release notes, optional binary attachments, and shows up on the repo's "Releases" page.
- **Auto-merge** — A queue: GitHub will merge the PR automatically once all the required automated checks pass and reviews are approved.
- **Fork** — Your own personal copy of someone else's repository, hosted on your GitHub account. Lets you push changes you don't have permission to push to the original. Contributors usually fork, push, then open a pull request from their fork back to the original.
- **`@me`** — A shorthand `gh` accepts in filter fields meaning "the currently logged-in user" (you).

If you've never used Git or GitHub before, the conceptual order to learn is: clone → branch → commit → pull request → merge. Everything else (issues, releases, etc.) sits alongside that core flow.

## What's in this demo

Eleven scriptrees grouped by workflow under [`gh.scriptreetree`](gh.scriptreetree):

### Pull requests folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`pr-list.scriptree`](pr-list.scriptree) | `gh pr list` | Show all PRs in a repo, with filters for state, author, branch, labels, etc. |
| [`pr-view.scriptree`](pr-view.scriptree) | `gh pr view` | Read one PR — title, status, description, optionally comments. |
| [`pr-create.scriptree`](pr-create.scriptree) | `gh pr create` | Propose your changes by opening a new pull request. |
| [`pr-checkout.scriptree`](pr-checkout.scriptree) | `gh pr checkout` | Pull someone else's PR down onto your computer so you can run / test / read it. |
| [`pr-merge.scriptree`](pr-merge.scriptree) | `gh pr merge` | Merge an approved PR into the main branch. |

### Issues folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`issue-list.scriptree`](issue-list.scriptree) | `gh issue list` | Show all issues in a repo, with filters for state, author, labels, etc. |
| [`issue-view.scriptree`](issue-view.scriptree) | `gh issue view` | Read one issue — description and optionally comments. |
| [`issue-create.scriptree`](issue-create.scriptree) | `gh issue create` | File a new bug report or feature request. |

### Repository folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`repo-view.scriptree`](repo-view.scriptree) | `gh repo view` | Show a repo's README and metadata. |
| [`repo-clone.scriptree`](repo-clone.scriptree) | `gh repo clone` | Download a copy of a repository to your computer. |

### Releases folder

| File | Subcommand | Purpose |
|------|------------|---------|
| [`release-create.scriptree`](release-create.scriptree) | `gh release create` | Publish a tagged version of the project for users to download. |

The scriptrees expect `gh` to be on `PATH` (i.e. installed via Homebrew, winget, apt, dnf, the official MSI, etc.). If it isn't, edit each file's `executable` field to point at the binary.

## What's intentionally not in this demo

A few subcommands were left out because the form factor adds little or none of the value:

- **`gh auth login`** — the one-time login flow. Has to run in a terminal because it opens a browser callback / device code prompt.
- **`gh release upload`** — attaches files to a release. Needs a list of file paths, which ScripTree doesn't have a clean widget for; run it from the terminal after `release create`.
- **`gh api`** — a raw HTTP client. Useful, but the form would amount to "type your URL and method here," which doesn't beat typing the command.
- **`gh workflow run`**, **`gh run list/view/watch`** — useful but a separate demo on its own; not yet wrapped here.

## Why a GUI for a CLI tool?

Most `gh` commands have `-R, --repo` plus another 5–15 flags. Memorising which short-form flag is which (e.g. `-A` for author vs. `-a` for assignee, in the same command) is the friction point. The form names every flag, lays out enum choices for `--state`, gives you a file picker for `--body-file` / `--notes-file`, and groups related flags into tabs so you're not scrolling through twenty fields to find one.

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
