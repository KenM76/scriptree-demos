#!/usr/bin/env python3
"""Image-to-Slide Deck Builder (PowerPoint) — ScripTree shim (Strategy A).

ScripTree launches this script (executable = ``python``) with the form
values as argv. combridge's ``run-script`` has no argv channel — a ``.csx``
only sees the plugin globals plus environment — so we BAKE the form values
into a generated ``.csx`` rendered from ``build_image_deck.csx.template``,
then hand that to::

    combridge.exe powerpoint run-script <temp.csx> -

combridge's ScriptHost ignores the script's ``return`` value (it exits 0 on
any clean run), so the ``.csx`` emits a first-line sentinel and THIS shim owns
the process exit code that ScripTree sees:

    __PPTIMG__ STATUS=BADFOLDER  -> exit 2 (the image folder does not exist)
    __PPTIMG__ STATUS=NOIMAGES   -> exit 2 (folder had no supported images)
    __PPTIMG__ STATUS=OK ...     -> exit 0 (deck built)

WHY THE SHIM ENUMERATES THE FOLDER (not the .csx)
-------------------------------------------------
The combridge PowerPoint plugin runs the ``.csx`` inside an automation host;
doing the directory scan in Python keeps the file-system logic in one tested
place, lets the shim do path hygiene (dedupe the output name, reject a missing
folder) BEFORE we ever touch COM, and makes the generated ``.csx`` a flat,
deterministic list of absolute image paths — easy to read, easy to debug, no
``Directory.GetFiles`` quirk to reason about on the COM side. So the shim:

  1. scans ``--source-folder`` NON-recursively for supported image extensions
     (.jpg .jpeg .png .gif .bmp .tif .tiff, case-insensitive),
  2. sorts the matches by filename (stable, predictable slide order),
  3. resolves the output path (blank -> ``Slideshow.pptx`` in the folder; bare
     name -> in the folder; existing -> de-duped with " (n)"),
  4. bakes the absolute image paths as a C# string array, plus the output path
     and the caption flag, into the ``.csx`` template.

This tool only CREATES a new deck — it never modifies an existing presentation
— so there is NO work-on-a-copy guard. The only collision risk is the output
file name, which step 3 dedupes.

combridge is located by walking up from this file looking for
``lib/combridge/combridge.exe`` — a relative discovery so the catalog stays
portable (no absolute path baked in, per the project's path rule).
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SENTINEL = "__PPTIMG__"

# Supported raster image extensions (lower-case, leading dot), matched
# case-insensitively against each file's suffix.
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff"}



def csharp_literal(value: str) -> str:
    """Escape *value* so it is safe inside a C# double-quoted string."""
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def enumerate_images(folder: Path) -> list[Path]:
    """Return the supported image files directly in *folder*, sorted by name.

    NON-recursive (``iterdir``, not ``rglob``). Matching is case-insensitive on
    the extension. Sort is by lower-cased filename so re-runs are deterministic
    and the slide order matches what the user sees in Explorer.
    """
    matches = [
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTS
    ]
    matches.sort(key=lambda p: p.name.lower())
    return matches


def resolve_output_path(output_file: str, folder: Path) -> Path:
    """Resolve the deck's destination path and de-dupe against existing files.

    * blank            -> ``<folder>/Slideshow.pptx``
    * a bare file name -> ``<folder>/<name>`` (forced to a .pptx suffix)
    * a full/relative path -> used as given (forced to a .pptx suffix)

    If the chosen path already exists, append " (2)", " (3)", … before the
    extension until a free name is found, so an earlier deck is never clobbered.
    """
    output_file = output_file.strip()
    if not output_file:
        target = folder / "Slideshow.pptx"
    else:
        cand = Path(output_file)
        # A bare name (no directory part) goes inside the image folder.
        if cand.parent == Path("."):
            cand = folder / cand.name
        # Force a .pptx extension so PowerPoint writes the right format.
        if cand.suffix.lower() != ".pptx":
            cand = cand.with_suffix(".pptx")
        target = cand

    if not target.exists():
        return target

    stem, suffix, parent = target.stem, target.suffix, target.parent
    n = 2
    while True:
        candidate = parent / f"{stem} ({n}){suffix}"
        if not candidate.exists():
            return candidate
        n += 1


def render_csx(template: str, *, image_paths: list[str], output_path: str,
               add_captions: bool, output_format: str) -> str:
    """Bake the form values + enumerated image list into the .csx template.

    ``__IMAGE_PATHS__`` is replaced with the body of a C# string-array
    initialiser — each absolute path csharp_literal-escaped and double-quoted,
    comma-separated. An empty list yields an empty initialiser (the .csx then
    reports NOIMAGES). Bools render as C# ``true``/``false``.
    """
    def cs_bool(b: bool) -> str:
        return "true" if b else "false"

    array_body = ",\n    ".join(
        f'"{csharp_literal(p)}"' for p in image_paths
    )

    return (
        template.replace("__IMAGE_PATHS__", array_body)
        .replace("__OUTPUT_PATH__", csharp_literal(output_path))
        .replace("__ADD_CAPTIONS__", cs_bool(add_captions))
        .replace("__OUTPUT_FORMAT__", csharp_literal(output_format))
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a new PowerPoint deck from a folder of images, one "
                    "scaled-to-fit image per slide.")
    parser.add_argument("--source-folder", dest="source_folder", required=True)
    parser.add_argument("--output-file", dest="output_file", default="")
    parser.add_argument("--add-captions", dest="add_captions",
                        action="store_true")
    parser.add_argument("--output-format", default="markdown",
                        choices=["markdown", "text"])
    args = parser.parse_args()

    if not args.source_folder.strip():
        print("ERROR: --source-folder must not be empty.", file=sys.stderr)
        return 2

    # --- folder existence (the shim's job; see module docstring) ------------
    # Done FIRST, before the template/combridge lookups, so a bad path is a
    # clean BADFOLDER/exit-2 regardless of the environment — there is no point
    # starting PowerPoint for a folder that doesn't exist. We emit the same
    # sentinel shape the .csx would use, so the output pane is consistent
    # whether the shim or the .csx caught the problem.
    folder = Path(args.source_folder.strip())
    # ``is_dir()`` normally returns False for a missing path, but on Windows an
    # unreachable network path (or one blocked by policy) raises OSError — treat
    # any such failure to confirm a directory as BADFOLDER, not a crash.
    try:
        folder_ok = folder.is_dir()
    except OSError:
        folder_ok = False
    if not folder_ok:
        print(f"{SENTINEL} STATUS=BADFOLDER")
        print(f"The image folder does not exist, is not a directory, or is not "
              f"reachable:\n    {folder}\nCheck the path and try again.")
        return 2

    here = Path(__file__).resolve().parent
    template_path = here / "build_image_deck.csx.template"
    if not template_path.is_file():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    # --- folder enumeration -------------------------------------------------
    images = enumerate_images(folder)
    image_paths = [str(p.resolve()) for p in images]
    output_path = str(resolve_output_path(args.output_file, folder).resolve())

    template = template_path.read_text(encoding="utf-8")
    csx = render_csx(
        template,
        image_paths=image_paths,
        output_path=output_path,
        add_captions=args.add_captions,
        output_format=args.output_format,
    )

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csx", prefix="build_image_deck_",
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

    # Precondition failures (bad folder / no images) -> 2.
    return 2 if status in ("NOIMAGES", "BADFOLDER") else 0


if __name__ == "__main__":
    raise SystemExit(main())
