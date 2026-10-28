"""Run the scenarios of lo_vba_harness.py through excel/ALMA_xlwings.py in file
mode (openpyxl) on a copy of ALMA_xlwings.xlsx and dump every sheet as CSV."""
import csv
import os
import shutil
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "excel"))
import ALMA_xlwings as AX  # noqa: E402

LISTS = os.path.join(ROOT, "lists")


def main(out):
    os.makedirs(out, exist_ok=True)
    book = os.path.join(out, "ALMA_xlwings_test.xlsx")
    shutil.copyfile(os.path.join(ROOT, "excel", "ALMA_xlwings.xlsx"), book)
    cache = os.path.join(ROOT, "results", "pubmed_cache.sqlite")
    base = [book, "--set", "offline=TRUE", "--set", "maxdate=2020/02/10", "--set", "cachefile=" + cache]
    steps = [
        ["--action", "Self-test"],
        ["--action", "Full pipeline", "--set", "high=10", "--set", "med=2"],
        ["--action", "Text generator"],
        ["--action", "Filter matrix"],
        ["--action", "Search matrix", "--set", "high=20", "--set", "med=5",
         "--list", "A=" + os.path.join(LISTS, "kinase_inhibitors_50.txt"),
         "--list", "B=" + os.path.join(LISTS, "cancers_fig2c_8.txt")],
        ["--action", "Normalize rows", "--target", "SM~2"],
        ["--action", "Cluster sort", "--target", "SM~2_norm_rows"],
        ["--action", "Sort by sum", "--target", "SM~2"],
        ["--action", "Full pipeline", "--target", "", "--list", "B=" + os.path.join(LISTS, "cancers_fig3d_7.txt"),
         "--list", "C=" + os.path.join(out, "vars_rt.txt")],
        ["--action", "Colorize", "--target", "SM~3"],
    ]
    with open(os.path.join(out, "vars_rt.txt"), "w") as fh:
        fh.write("radiotherapy\n")
    bad = 0
    for st in steps:
        rc = AX.cli(base + st)
        bad += rc != 0
    if "xlwings" in sys.modules and hasattr(sys.modules["xlwings"], "_BOOKS"):   # fake Excel: save the open book
        sys.modules["xlwings"].Book(book).save()
    wb = openpyxl.load_workbook(book)
    sheets = os.path.join(out, "sheets")
    os.makedirs(sheets, exist_ok=True)
    for ws in wb.worksheets:
        with open(os.path.join(sheets, ws.title + ".csv"), "w", newline="") as fh:
            w = csv.writer(fh)
            for row in ws.iter_rows(values_only=True):
                w.writerow(["" if v is None else (int(v) if isinstance(v, float) and v.is_integer() else v)
                            for v in row])
    print("failed steps:", bad)
    return bad


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
