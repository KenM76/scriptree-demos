#!/usr/bin/env python3
"""Global Typography Enforcer (PowerPoint) — ScripTree shim (Strategy A).

PURPOSE
-------
Standardise the fonts used across a PowerPoint deck.  Two modes:

* **specific** — replace one named font (e.g. every run set in
  "Calibri") with a new font, leaving all other fonts alone.  This is
  the precision scalpel for brand-compliance work ("our deck must use
  Aptos everywhere Calibri snuck in").
* **all** — set *every* text run to one target font, sweeping a
  "frankendeck" assembled from many sources into a single typeface.

In BOTH modes a hard-coded blocklist of **symbol/icon fonts**
(Wingdings, Wingdings 2, Wingdings 3, Webdings, Symbol, Marlett) is
ALWAYS bypassed — those fonts encode glyphs by codepoint, so reflowing
them to a text font would turn ✆/➜/✓ into meaningless letters.

HOW IT FITS THE STRATEGY-A PATTERN
----------------------------------
ScripTree launches this script (executable = ``python``) with the form
values as argv.  combridge's ``run-script`` has no argv channel — a
``.csx`` only sees the plugin globals plus environment — so we BAKE the
form values into a generated ``.csx`` rendered from
``font_enforcer.csx.template`` (token replacement + C#-literal
escaping), then hand that to::

    combridge.exe powerpoint run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0
on any clean run), so the ``.csx`` emits a first-line sentinel and THIS
shim owns the process exit code that ScripTree sees:

    __PPTFONT__ STATUS=NODECK       -> exit 2 (no presentation open)
    __PPTFONT__ STATUS=UNSAVED      -> exit 2 (deck never saved; no folder for the copy)
    __PPTFONT__ STATUS=NOTARGET     -> exit 2 (manual path: no target font given)
    __PPTFONT__ STATUS=NOSOURCE     -> exit 2 (manual specific mode but no source font given)
    __PPTFONT__ STATUS=BADTEMPLATE  -> exit 2 (template path: the named template deck does not exist)
    __PPTFONT__ STATUS=NOTHEME      -> exit 2 (template-fonts path: theme major/minor fonts unreadable)
    __PPTFONT__ STATUS=OK ...       -> exit 0
    (combridge's own non-zero codes — 3 compile / 4 throw / 5 host — pass through.)

TEMPLATE MODE (added later)
---------------------------
When ``--font-template`` names a reference ``.pptx``/``.potx``, the tool reads
styling from that deck instead of using the manual source/target font fields.
``--template-mode fonts`` reads the template's theme heading (major) + body
(minor) fonts and sweeps the working copy routing title/subtitle placeholders
to the heading font and all other text to the body font.  ``--template-mode
theme`` applies the template's WHOLE theme (fonts + colours + masters) via
``Presentation.ApplyTemplate``.  In template mode the NOTARGET / NOSOURCE
guards do not apply (no manual fonts are required).

SAFETY MODEL — the open deck is NEVER mutated
---------------------------------------------
The ``.csx`` produces a sibling ``<name>_Restyled.pptx`` via
``Presentation.SaveCopyAs`` (which, unlike Word's ``SaveAs2``, does NOT
repoint the active presentation), re-opens that copy HEADLESS
(``WithWindow=msoFalse``), does all of the font rewriting on the copy,
saves and closes it.  The deck the user is looking at is provably
untouched.

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

SENTINEL = "__PPTFONT__"
# Statuses that mean "precondition not met / nothing done" -> non-zero exit.
PRECONDITION_FAILS = ("NODECK", "UNSAVED", "NOTARGET", "NOSOURCE",
                      "BADTEMPLATE", "NOTHEME")



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string.

    Backslash MUST be escaped first or the subsequent escapes double up.
    """
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, replacement_mode: str, source_font: str,
               target_font: str, include_masters: bool,
               output_format: str, font_template: str,
               template_mode: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    return (
        template.replace("__REPLACEMENT_MODE__", csharp_literal(replacement_mode))
        .replace("__SOURCE_FONT__", csharp_literal(source_font))
        .replace("__TARGET_FONT__", csharp_literal(target_font))
        .replace("__INCLUDE_MASTERS__", cs_bool(include_masters))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
        .replace("__FONT_TEMPLATE__", csharp_literal(font_template))
        .replace("__TEMPLATE_MODE__", csharp_literal(template_mode))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Enforce one font across the running PowerPoint deck "
                    "(specific-font swap or blanket all-fonts), writing the "
                    "result to a sibling copy. Symbol/icon fonts are always "
                    "preserved.")
    parser.add_argument("--replacement-mode", dest="replacement_mode",
                        default="specific", choices=["specific", "all"])
    parser.add_argument("--source-font", dest="source_font", default="")
    parser.add_argument("--target-font", dest="target_font", default="")
    parser.add_argument("--include-masters", dest="include_masters",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    parser.add_argument("--font-template", dest="font_template", default="")
    parser.add_argument("--template-mode", dest="template_mode",
                        default="fonts", choices=["fonts", "theme"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "font_enforcer.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        replacement_mode=args.replacement_mode,
        source_font=args.source_font,
        target_font=args.target_font,
        include_masters=args.include_masters,
        output_format=args.output_format,
        font_template=args.font_template,
        template_mode=args.template_mode,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="font_enforcer_",
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

    # Precondition failures -> 2; OK (incl. zero runs changed) -> 0.
    return 2 if status in PRECONDITION_FAILS else 0


if __name__ == "__main__":
    raise SystemExit(main())
