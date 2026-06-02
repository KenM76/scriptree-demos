#!/usr/bin/env python3
"""Heading-Based Document Splitter (Word) — ScripTree shim (Strategy A).

PURPOSE
-------
Split the open Word document into one file per section, cutting at every
paragraph set to a chosen heading level (Heading 1, 2, or 3).  Each
chapter is written to its own ``.docx`` and/or ``.pdf`` in a target
folder, named after the heading text with a zero-padded sequence prefix
(``01_Introduction.docx`` …).  Optional "front matter" (content before
the first heading) is saved as ``00_Front-Matter`` when present.

This tool is **non-destructive by construction**: it only READS ranges
from the open document and CREATES new documents — it never calls
SaveAs2/Save on the user's document, so the original is provably
untouched (no copy guard needed, unlike the find/replace tool).

HOW IT FITS THE STRATEGY-A PATTERN
----------------------------------
ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``split_by_heading.csx.template`` (token replacement + C#-literal
escaping), then hand that to::

    combridge.exe word run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __WORDSPLIT__ STATUS=NODOC       -> exit 2 (no document open)
    __WORDSPLIT__ STATUS=NOHEADINGS  -> exit 2 (no heading at the chosen level)
    __WORDSPLIT__ STATUS=UNSAVED     -> exit 2 (no output folder given AND doc never saved)
    __WORDSPLIT__ STATUS=BADFOLDER   -> exit 2 (output folder could not be created)
    __WORDSPLIT__ STATUS=OK ...      -> exit 0
    (combridge's own non-zero codes — 3 compile / 4 throw / 5 host — pass through.)

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

SENTINEL = "__WORDSPLIT__"



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string.

    Backslash MUST be escaped first or the subsequent escapes double up
    (critical here: an output-folder path is full of backslashes).
    """
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, heading_level: str, output_folder: str,
               file_format: str, include_front_matter: bool,
               output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__HEADING_LEVEL__", csharp_literal(heading_level))
        .replace("__OUTPUT_FOLDER__", csharp_literal(output_folder))
        .replace("__FILE_FORMAT__", csharp_literal(file_format))
        .replace("__INCLUDE_FRONT_MATTER__", cs_bool(include_front_matter))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Split the open Word document into one file per heading-N "
                    "section, written to a target folder. The original "
                    "document is never modified.")
    parser.add_argument("--heading-level", dest="heading_level",
                        default="1", choices=["1", "2", "3"])
    parser.add_argument("--output-folder", dest="output_folder", default="")
    parser.add_argument("--file-format", dest="file_format",
                        default="docx", choices=["docx", "pdf", "both"])
    parser.add_argument("--include-front-matter", dest="include_front_matter",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "split_by_heading.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        heading_level=args.heading_level,
        output_folder=args.output_folder,
        file_format=args.file_format,
        include_front_matter=args.include_front_matter,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="split_by_heading_",
        delete=False, encoding="utf-8")
    try:
        tmp.write(csx)
        tmp.close()
        proc = subprocess.run(
            ["combridge.exe", "word", "run-script", tmp.name, "-"],
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
    # error, or no Word session): surface its output verbatim and its code.
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

    # Precondition failures -> 2; OK (incl. a single chapter) -> 0.
    return 2 if status in ("NODOC", "NOHEADINGS", "UNSAVED", "BADFOLDER") else 0


if __name__ == "__main__":
    raise SystemExit(main())
