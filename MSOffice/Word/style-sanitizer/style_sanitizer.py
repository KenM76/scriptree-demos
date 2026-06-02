#!/usr/bin/env python3
"""Corporate Style Sanitizer (Word) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv. combridge's ``run-script`` has no argv channel — a ``.csx``
only sees the plugin globals plus environment — so we BAKE the form values
into a generated ``.csx`` rendered from ``style_sanitizer.csx.template`` by
plain token replacement, then hand that to::

    combridge.exe word run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0 on
any clean run), so the ``.csx`` emits a first-line sentinel and THIS shim
owns the process exit code that ScripTree sees:

    __WORDSAN__ STATUS=NODOC     -> exit 2 (no document open)
    __WORDSAN__ STATUS=UNSAVED   -> exit 2 (copy mode needs a saved doc)
    __WORDSAN__ STATUS=TRACKED   -> exit 2 (document has tracked changes)
    __WORDSAN__ STATUS=OK ...    -> exit 0

combridge is invoked via a bare ``combridge.exe`` call -- ScripTree's runner
prepends ``<install>/lib/combridge`` to ``PATH`` on every spawned tool, so the
OS resolves the bundled binary by name.  No discovery code needed; if a tool is
launched outside ScripTree (e.g. direct ``python tool.py`` for debugging), set
``SCRIPTREE_HOME`` and prepend ``%SCRIPTREE_HOME%/lib/combridge`` to PATH
yourself, or just launch through ScripTree's editor.

WHY a shim instead of a typed plugin command: see
``_meta/rags/integration/csx_has_no_argv_channel.md`` and
``shim_generates_csx.md``. The same Strategy-A pattern backs every Office app
in this catalog.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SENTINEL = "__WORDSAN__"
# Statuses that mean "precondition not met / nothing done" -> non-zero exit.
PRECONDITION_FAILS = ("NODOC", "UNSAVED", "TRACKED", "BADTEMPLATE")



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def render_csx(template: str, *, strip_formatting: bool, collapse_spaces: bool,
               remove_blank_paragraphs: bool, normalize_fonts: bool,
               body_font: str, heading_font: str, style_template: str,
               smart_quotes: bool,
               list_normalize_bullets: bool, strip_table_shading: bool,
               enforce_margins: bool, margin_top: float, margin_bottom: float,
               margin_left: float, margin_right: float,
               work_on_copy: bool, output_format: str) -> str:
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    # Bake each margin as a C# double literal (always has a decimal point).
    return (
        template.replace("__STRIP_FORMATTING__", cs_bool(strip_formatting))
        .replace("__COLLAPSE_SPACES__", cs_bool(collapse_spaces))
        .replace("__REMOVE_BLANKS__", cs_bool(remove_blank_paragraphs))
        .replace("__NORMALIZE_FONTS__", cs_bool(normalize_fonts))
        .replace("__BODY_FONT__", csharp_literal(body_font))
        .replace("__HEADING_FONT__", csharp_literal(heading_font))
        .replace("__STYLE_TEMPLATE__", csharp_literal(style_template))
        .replace("__SMART_QUOTES__", cs_bool(smart_quotes))
        .replace("__LIST_NORMALIZE__", cs_bool(list_normalize_bullets))
        .replace("__STRIP_TABLE_SHADING__", cs_bool(strip_table_shading))
        .replace("__ENFORCE_MARGINS__", cs_bool(enforce_margins))
        .replace("__MARGIN_TOP__", repr(float(margin_top)))
        .replace("__MARGIN_BOTTOM__", repr(float(margin_bottom)))
        .replace("__MARGIN_LEFT__", repr(float(margin_left)))
        .replace("__MARGIN_RIGHT__", repr(float(margin_right)))
        .replace("__WORK_ON_COPY__", cs_bool(work_on_copy))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Sanitise the running Word document into a clean style.")
    parser.add_argument("--strip-formatting", dest="strip_formatting",
                        action="store_true")
    parser.add_argument("--collapse-spaces", dest="collapse_spaces",
                        action="store_true")
    parser.add_argument("--remove-blank-paragraphs", dest="remove_blanks",
                        action="store_true")
    parser.add_argument("--normalize-fonts", dest="normalize_fonts",
                        action="store_true")
    parser.add_argument("--body-font", dest="body_font", default="Calibri")
    parser.add_argument("--heading-font", dest="heading_font",
                        default="Calibri Light")
    parser.add_argument("--style-template", dest="style_template", default="")
    parser.add_argument("--smart-quotes", dest="smart_quotes",
                        action="store_true")
    parser.add_argument("--list-normalize-bullets", dest="list_normalize_bullets",
                        action="store_true")
    parser.add_argument("--strip-table-shading", dest="strip_table_shading",
                        action="store_true")
    parser.add_argument("--enforce-margins", dest="enforce_margins",
                        action="store_true")
    parser.add_argument("--margin-top", dest="margin_top", type=float, default=1.0)
    parser.add_argument("--margin-bottom", dest="margin_bottom", type=float,
                        default=1.0)
    parser.add_argument("--margin-left", dest="margin_left", type=float, default=1.0)
    parser.add_argument("--margin-right", dest="margin_right", type=float,
                        default=1.0)
    parser.add_argument("--work-on-copy", dest="work_on_copy",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    here = Path(__file__).resolve().parent
    template_path = here / "style_sanitizer.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        strip_formatting=args.strip_formatting,
        collapse_spaces=args.collapse_spaces,
        remove_blank_paragraphs=args.remove_blanks,
        normalize_fonts=args.normalize_fonts,
        body_font=args.body_font,
        heading_font=args.heading_font,
        style_template=args.style_template,
        smart_quotes=args.smart_quotes,
        list_normalize_bullets=args.list_normalize_bullets,
        strip_table_shading=args.strip_table_shading,
        enforce_margins=args.enforce_margins,
        margin_top=args.margin_top,
        margin_bottom=args.margin_bottom,
        margin_left=args.margin_left,
        margin_right=args.margin_right,
        work_on_copy=args.work_on_copy,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="style_sanitizer_",
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
        for f in lines[0].split()[1:]:
            if f.startswith("STATUS="):
                status = f[len("STATUS="):]
        body_start = 1

    body = "\n".join(lines[body_start:])
    if body:
        print(body)

    return 2 if status in PRECONDITION_FAILS else 0


if __name__ == "__main__":
    raise SystemExit(main())
