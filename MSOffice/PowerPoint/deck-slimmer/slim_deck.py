#!/usr/bin/env python3
"""Slide Deck Slimmer (PowerPoint) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``slim_deck.csx.template``, then hand that to::

    combridge.exe powerpoint run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __PPTSLIM__ STATUS=NODECK    -> exit 2 (no presentation open)
    __PPTSLIM__ STATUS=UNSAVED   -> exit 2 (copy mode but deck never saved)
    __PPTSLIM__ STATUS=OK ...    -> exit 0

WHAT THIS TOOL DOES (and does NOT do)
-------------------------------------
It reduces a deck's saved file size by removing UNUSED custom slide
layouts (and, opt-in, designs/masters left with no layouts).  Template-
heavy decks carry many unused layouts whose background images bloat the
file.  It does this on a ``<name>_Slimmed.pptx`` COPY by default
(SaveCopyAs -> headless Open -> edit -> Save -> Close), leaving the user's
open deck untouched.

It DOES NOT recompress images or media.  PowerPoint exposes picture
recompression ("Compress Pictures") only through an interactive dialog
that cannot be driven fire-and-forget over COM, so this tool makes no
such claim.

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

SENTINEL = "__PPTSLIM__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, remove_unused_layouts: bool,
               remove_empty_masters: bool, work_on_copy: bool,
               output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__REMOVE_UNUSED_LAYOUTS__", cs_bool(remove_unused_layouts))
        .replace("__REMOVE_EMPTY_MASTERS__", cs_bool(remove_empty_masters))
        .replace("__WORK_ON_COPY__", cs_bool(work_on_copy))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Slim a PowerPoint deck by removing unused custom slide "
                    "layouts (and optionally empty designs/masters) onto a "
                    "copy. Does NOT recompress images.")
    parser.add_argument("--remove-unused-layouts", dest="remove_unused_layouts",
                        action="store_true")
    parser.add_argument("--remove-empty-masters", dest="remove_empty_masters",
                        action="store_true")
    parser.add_argument("--work-on-copy", dest="work_on_copy",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "slim_deck.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        remove_unused_layouts=args.remove_unused_layouts,
        remove_empty_masters=args.remove_empty_masters,
        work_on_copy=args.work_on_copy,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="slim_deck_",
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

    # Precondition failures (no deck open / copy mode but deck never saved) -> 2.
    return 2 if status in ("NODECK", "UNSAVED") else 0


if __name__ == "__main__":
    raise SystemExit(main())
