#!/usr/bin/env python3
"""Generate a folder of sample exports for the Multi-CSV Sheet Aggregator.

Run:  python make_example.py
Produces:  monthly_exports/   (a folder next to this script) containing
           jan.csv, feb.csv, mar.csv, europe_q1.csv, notes.txt

WHAT THE SAMPLE CONTAINS (and why)
----------------------------------
The aggregator imports every .csv and .txt directly inside a chosen folder
as one new sheet per file, auto-detecting each file's delimiter. The
sample folder is built to exercise that auto-detection:

* **jan.csv / feb.csv / mar.csv** — ordinary comma-delimited monthly
  exports (same columns). The "one workbook, a sheet per month" use case.
* **europe_q1.csv** — a **semicolon**-delimited file (as European locales
  export). A naive comma split would mangle it into one column; the tool's
  per-file auto-detect picks `;` and parses it correctly. This is the file
  that proves auto-detect matters.
* **notes.txt** — a **tab**-delimited .txt, to show that .txt files are
  imported too (not just .csv) and that tab is detected.

Subfolders are NOT scanned by the tool, so everything is placed directly
in monthly_exports/.
"""
from pathlib import Path

FILES = {
    "jan.csv": "Product,Units,Revenue\nWidget,120,2400\nGadget,80,3200\nGizmo,60,1800\n",
    "feb.csv": "Product,Units,Revenue\nWidget,95,1900\nGadget,70,2800\nGizmo,25,750\n",
    "mar.csv": "Product,Units,Revenue\nWidget,110,2200\nGadget,55,2200\nGizmo,50,1500\n",
    # Semicolon-delimited (European CSV). Note the comma is a DECIMAL here.
    "europe_q1.csv": "Produkt;Einheiten;Umsatz\nWidget;325;6500,50\nGadget;205;8200,00\nGizmo;135;4050,75\n",
    # Tab-delimited .txt.
    "notes.txt": "Region\tManager\tTarget\nEast\tAlice\t10000\nWest\tBob\t9000\nNorth\tCarol\t8500\n",
}


def main() -> None:
    folder = Path(__file__).resolve().parent / "monthly_exports"
    folder.mkdir(exist_ok=True)
    for name, content in FILES.items():
        (folder / name).write_text(content, encoding="utf-8")
        print(f"wrote {folder / name}")


if __name__ == "__main__":
    main()
