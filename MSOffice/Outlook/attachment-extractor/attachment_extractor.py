#!/usr/bin/env python3
"""Surgical Attachment Extractor (Outlook) — ScripTree shim (Strategy A).

================================================================================
WHAT THIS PROGRAM DOES (the logic — read this to understand/rewrite the tool)
================================================================================
The user is looking at Outlook. They want to pull the file attachments OUT of
one or more emails and drop them, as ordinary files, into a folder on disk —
without hand-saving each one through Outlook's "Save Attachments" dialog. Classic
example: select a dozen "Invoice" emails, run the tool, and get every PDF saved
into ``C:\\Invoices`` in one pass.

The defining safety property — and the reason "Surgical" is in the name — is that
this tool is **strictly read-only with respect to the mailbox**. It ONLY does two
things to Outlook: it reads the chosen items, and it calls ``Attachment.SaveAsFile``
(which copies the attachment's bytes to disk). It NEVER calls ``Attachment.Delete``,
never edits or saves a ``MailItem``, never moves or marks anything. The emails and
their attachments are left exactly as they were. The only thing that changes is
that new files appear in the user's chosen output folder.

Two scopes:

  * ``selection`` (the default, the "surgical" path) — process only the items the
    user currently has selected in the active Outlook explorer. Fast and precise.
  * ``folder`` — process every item in the currently open folder. Necessarily
    iterates every item (there is no lightweight MAPI table that can *save*
    attachment bytes), so it is slower on large folders; that's inherent to the
    task, not a defect.

Filtering:

  * An optional **extension filter** (e.g. ``pdf, docx, xlsx``). Blank = every
    attachment. Matching is case-insensitive and dot-insensitive.
  * **Include inline images** (default OFF). Email signatures and HTML bodies
    carry embedded "inline" images (logos, tracking pixels) as attachments with a
    MAPI content-id. Excluding them is the single biggest noise reducer when you
    just want the "real" attachments, so the default skips them.

Name-clash policy (many emails ship a file literally named ``invoice.pdf``):

  * ``rename`` (default) — keep both: the second becomes ``invoice (1).pdf``, etc.
  * ``skip``            — leave the first, don't save the later ones.
  * ``overwrite``       — let the later file replace the earlier one on disk.

================================================================================
WHY THIS IS A SHIM (Strategy A — the project-wide integration pattern)
================================================================================
ScripTree is a GUI form that runs a command line. It launches this script with
``executable = python`` and passes the form field values as argv. The actual
Outlook work must happen inside combridge (which owns the COM connection to the
running Outlook), and combridge's ``run-script`` subcommand executes a C# Roslyn
script (``.csx``).

combridge's ``run-script`` has **no argv channel** — a ``.csx`` can only see the
plugin globals (``olApp``/``olNs``/``olExplorer``) plus environment variables. So
this shim BAKES the form values into a generated ``.csx`` (rendered from
``attachment_extractor.csx.template`` by substituting ``__PLACEHOLDER__`` tokens),
then hands that script to::

    combridge.exe outlook run-script <temp.csx> -

================================================================================
WHY THE SENTINEL + SHIM-OWNED EXIT CODE (a hard combridge constraint)
================================================================================
combridge's ScriptHost **ignores the C# script's ``return`` value** — it exits 0
on any clean run (only a compile error = 3, an unhandled throw = 4, or a host
error = 5 produce a non-zero code). ScripTree, however, decides success/failure
from the child process exit code. So we use a two-part protocol:

  1. The ``.csx`` prints a machine-readable **sentinel** as its first stdout
     line:  ``__OLXTRACT__ STATUS=<code> [key=value ...]``  followed by the
     human-readable report.
  2. THIS shim parses that sentinel, strips it from the output, prints the
     report, and **translates the status into the process exit code** that
     ScripTree sees.

Status → exit-code contract (see ``main`` and the ``.csx`` template):

    STATUS=OK          -> exit 0   (extraction ran; report follows — even if 0 saved)
    STATUS=NOEXPLORER  -> exit 2   (no active Outlook explorer to read from)
    STATUS=NOSEL       -> exit 2   (scope=selection but nothing is selected)
    STATUS=NOOUTDIR    -> exit 2   (output folder missing and could not be created)
    (combridge's own non-zero codes — 3/4/5/connect failure — pass through verbatim)

Exit 2 is used for every "precondition not met / nothing done" case so the
caller can distinguish a *guarded refusal* (2) from a *crash* (3/4/5) from
*success* (0). An empty source folder is NOT an error — that's a clean run that
simply saved zero files, so it reports STATUS=OK.

================================================================================
HOW COMBRIDGE IS LOCATED (portability rule)
================================================================================
combridge is NOT bundled in this project repo. Apps are authored here and
DEPLOYED into a ScripTree install that ships ``lib/combridge/combridge.exe``.
This shim finds it by walking UP the directory tree from its own location
looking for ``lib/combridge/combridge.exe`` — a relative discovery, so no
absolute path is ever baked in and the app works at whatever depth it lands.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

# The first-line marker the generated .csx prints. Must match the literal used
# in attachment_extractor.csx.template. Kept deliberately unlikely to collide
# with real attachment content.
SENTINEL = "__OLXTRACT__"

# Relative location of combridge inside a deployed ScripTree install.

# combridge plugin name for Outlook automation.
PLUGIN = "outlook"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe to drop inside a C# double-quoted string.

    The generated ``.csx`` embeds user-supplied text (the output folder path,
    the extension list, the name-clash policy) directly inside ``"..."``
    literals. Without escaping, a backslash (ubiquitous in Windows paths!) or a
    double-quote in the user's input would either break compilation or — worse —
    let crafted input alter the script. We escape the four characters that are
    special inside a C# regular (non-verbatim) string literal plus the whitespace
    controls, backslash first so we don't double-escape the escapes we add after.
    """
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, scope: str, output_folder: str,
               extensions: str, include_inline: bool, on_name_clash: str,
               output_format: str) -> str:
    """Fill the ``.csx`` template with the validated form values.

    String values go through :func:`csharp_literal`. The single boolean is
    emitted as the C# keyword ``true``/``false``. Every ``__TOKEN__`` in the
    template must be replaced here — a leftover token would be a compile error,
    which is why the offline render check greps the output for ``__[A-Z_]+__``.
    """
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__SCOPE__", csharp_literal(scope))
        .replace("__OUTPUT_FOLDER__", csharp_literal(output_folder))
        .replace("__EXTENSIONS__", csharp_literal(extensions))
        .replace("__INCLUDE_INLINE__", cs_bool(include_inline))
        .replace("__ON_NAME_CLASH__", csharp_literal(on_name_clash))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    """Parse argv, generate the ``.csx``, run combridge, translate the result.

    Returns the process exit code ScripTree should see (see the status→exit
    contract in the module docstring).
    """
    parser = argparse.ArgumentParser(
        description="Save attachments from the selected (or current-folder) "
                    "Outlook emails to a folder on disk. Never deletes them "
                    "from the emails.")
    parser.add_argument(
        "--scope", dest="scope", default="selection",
        choices=["selection", "folder"],
        help="'selection' = the currently selected items (default, surgical); "
             "'folder' = every item in the open folder.")
    parser.add_argument(
        "--output-folder", dest="output_folder", required=True,
        help="Directory to save the extracted attachments into. Created if it "
             "does not exist.")
    parser.add_argument(
        "--extensions", dest="extensions", default="",
        help="Comma/space separated extension allow-list (e.g. 'pdf,docx'). "
             "Blank means every attachment. Case- and dot-insensitive.")
    parser.add_argument(
        "--include-inline", dest="include_inline", action="store_true",
        help="Also save inline/embedded images (signature logos, HTML body "
             "images). Off by default to skip signature noise.")
    parser.add_argument(
        "--on-name-clash", dest="on_name_clash", default="rename",
        choices=["rename", "skip", "overwrite"],
        help="What to do when a target filename already exists: 'rename' "
             "(append ' (n)'), 'skip', or 'overwrite'.")
    parser.add_argument(
        "--output-format", dest="output_format", default="markdown",
        choices=["markdown", "text"])
    args = parser.parse_args()

    output_folder = args.output_folder.strip()
    if not output_folder:
        print("ERROR: --output-folder must not be empty.", file=sys.stderr)
        return 2

    here = Path(__file__).resolve().parent
    template_path = here / "attachment_extractor.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        scope=args.scope,
        output_folder=output_folder,
        extensions=args.extensions,
        include_inline=args.include_inline,
        on_name_clash=args.on_name_clash,
        output_format=args.output_format,
    )

    # Write the generated script to a temp .csx, run it, and always clean up.
    # delete=False + manual unlink in finally is required on Windows because the
    # child process must be able to open the file while we hold no lock on it.
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="attachment_extractor_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", PLUGIN, "run-script", tmp.name, "-"],
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

    if proc.stderr:
        sys.stderr.write(proc.stderr)

    # combridge failed before/while running the script (compile=3, throw=4,
    # host=5, or could not connect to Outlook): surface its output verbatim and
    # propagate its exit code unchanged.
    if proc.returncode != 0:
        sys.stdout.write(proc.stdout)
        return proc.returncode

    # Clean run: parse the sentinel first line, print the rest as the report.
    lines = proc.stdout.splitlines()
    status = "OK"
    body_start = 0
    if lines and lines[0].startswith(SENTINEL):
        fields = lines[0].split()
        for f in fields[1:]:
            if f.startswith("STATUS="):
                status = f[len("STATUS="):]
        body_start = 1

    body = "\n".join(lines[body_start:])
    if body:
        print(body)

    # Any guarded refusal maps to exit 2; only a real run is 0 (even 0 saved).
    return 0 if status == "OK" else 2


if __name__ == "__main__":
    raise SystemExit(main())
