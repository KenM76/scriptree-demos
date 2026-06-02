#!/usr/bin/env python3
"""Hidden Assets & Notes Purger (PowerPoint) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``ppt_purger.csx.template``, then hand that to::

    combridge.exe powerpoint run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __PPTPURGE__ STATUS=NODECK     -> exit 2 (no presentation open)
    __PPTPURGE__ STATUS=UNSAVED    -> exit 2 (deck never saved; no folder for the copy)
    __PPTPURGE__ STATUS=OK ...     -> exit 0

This tool NEVER mutates the open deck.  The ``.csx`` produces a sanitized
sibling ``<name>_Sanitized.pptx`` via ``Presentation.SaveCopyAs`` and does
all of its stripping on that file (re-opened headless), so the deck the
user is looking at is provably untouched.

combridge is located by walking up from this file looking for
``lib/combridge/combridge.exe`` — a relative discovery so the catalog
stays portable (no absolute path baked in, per the project's path rule).
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SENTINEL = "__PPTPURGE__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, strip_notes: bool, strip_comments: bool,
               strip_metadata: bool, delete_hidden_slides: bool,
               output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__STRIP_NOTES__", cs_bool(strip_notes))
        .replace("__STRIP_COMMENTS__", cs_bool(strip_comments))
        .replace("__STRIP_METADATA__", cs_bool(strip_metadata))
        .replace("__DELETE_HIDDEN_SLIDES__", cs_bool(delete_hidden_slides))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Strip notes / comments / metadata / hidden slides from a "
                    "sanitized copy of the running PowerPoint deck.")
    parser.add_argument("--strip-notes", dest="strip_notes",
                        action="store_true")
    parser.add_argument("--strip-comments", dest="strip_comments",
                        action="store_true")
    parser.add_argument("--strip-metadata", dest="strip_metadata",
                        action="store_true")
    parser.add_argument("--delete-hidden-slides", dest="delete_hidden_slides",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "ppt_purger.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        strip_notes=args.strip_notes,
        strip_comments=args.strip_comments,
        strip_metadata=args.strip_metadata,
        delete_hidden_slides=args.delete_hidden_slides,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="ppt_purger_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", "powerpoint", "run-script", tmp.name, "-"],
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

    # combridge failed before/while running the script (compile/host/connect
    # error, or no PowerPoint session): surface its output verbatim + its code.
    if proc.returncode != 0:
        sys.stdout.write(proc.stdout)
        return proc.returncode

    # Parse the sentinel first line written by the .csx.
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

    # Precondition failures (no deck open / deck never saved) -> 2.
    return 2 if status in ("NODECK", "UNSAVED") else 0


if __name__ == "__main__":
    raise SystemExit(main())
