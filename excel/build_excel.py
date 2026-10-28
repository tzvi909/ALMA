#!/usr/bin/env python3
"""Build the Excel versions of ALMA.

  .venv/bin/python excel/build_excel.py

  excel/ALMA.xlsm          VBA macros embedded, buttons on the ALMA sheet, PubMed
                           cache pre-filled from results/pubmed_cache.sqlite
  excel/ALMA_template.xlsx the same sheets without macros (import excel/vba/*.bas
                           by hand and save as .xlsm if Excel rejects ALMA.xlsm)
  excel/ALMA_xlwings.xlsx  for the Python script excel/ALMA_xlwings.py (xlwings
                           "Run main"); its cache excel/ALMA_cache.sqlite starts as
                           a copy of results/pubmed_cache.sqlite
"""
import glob
import io
import os
import sqlite3
import sys

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import vbaproject  # noqa: E402
from ALMA_xlwings import ACTIONS  # noqa: E402

CACHE = os.path.join(ROOT, "results", "pubmed_cache.sqlite")
BAS = [os.path.join(HERE, "vba", n) for n in ("ALMA_PubMed.bas", "ALMA_Http.bas", "ALMA_Core.bas", "ALMA_Main.bas")]
VBA_SHEETS = ["ALMA", "Lists", "Library", "Tree", "Cache"]
XLW_SHEETS = ["ALMA", "Lists", "Library", "Tree"]

COMMON_SETTINGS = [  # key, label, default, note
    ("maxdate", "PubMed date limit (YYYY/MM/DD)", "",
     "blank = today's PubMed.  2020/02/10 reproduces the paper's searches (Feb 2020)"),
    ("datetype", "Date type", "edat", "edat = date the record was added to PubMed, pdat = publication date"),
    ("high", "High threshold (score 2)", 20, "NOP >= high is strong / green (paper: 20; Fig 3c & 5c ranking used 10)"),
    ("med", "Medium threshold (score 1)", 5, "NOP >= medium scores 1 (paper: 5; Fig 3c & 5c used 2)"),
    ("novelmax", "Novel: at most N publications", 1, "filter of Fig 3c / 5c: hypothesis has <= 1 publication ..."),
    ("pairmin", "Reasonable: every pair >=", 10, "... and every pair (LR, THR, TVR) has >= 10 publications"),
    ("cluster", "Cluster threshold", 0.9, "normalized NOP that joins a cluster (paper: 0.9)"),
    ("apikey", "NCBI API key (optional)", "", "allows 10 instead of 3 requests per second"),
]
VBA_SETTINGS = COMMON_SETTINGS + [
    ("delay", "Delay between requests (ms)", 350, "keep >= 340 without an API key"),
    ("usecache", "Use cached counts", True, "re-use counts stored on the Cache sheet"),
    ("offline", "Offline (cache only)", False, "never contact PubMed; fail if a count is not cached"),
    ("quiet", "Quiet (no message boxes)", False, "for automated runs"),
    ("laststatus", "Last run", "", "result or error of the last action"),
]
XLW_SETTINGS = [
    ("action", "Action (then click Run main)", "Full pipeline", "what 'Run main' on the xlwings tab does"),
    ("target", "Sheet to work on (optional)", "", "blank = the active sheet, else the newest suitable sheet"),
] + COMMON_SETTINGS + [
    ("cachefile", "Cache file", "ALMA_cache.sqlite", "PubMed counts are stored here (next to the workbook)"),
    ("offline", "Offline (cache only)", False, "never contact PubMed; fail if a count is not cached"),
    ("quiet", "Quiet (no message box)", False, "do not show the result in a message box"),
    ("laststatus", "Last run", "", "result or error of the last action"),
]

BUTTONS = [  # macro, caption
    ("ALMA_FullPipeline", "Full pipeline: search -> comparisons -> NRF ranking"),
    ("ALMA_NewSearchMatrix", "New search matrix (List 1 x List 2)"),
    ("ALMA_AddConstantString", "Add constant string to the search"),
    ("ALMA_ComparisonMatrix", "Comparison matrix (+ one string)"),
    ("ALMA_ComparisonList", "List of comparison matrices (+ column C)"),
    ("ALMA_NRFRanking", "NRF ranking (novelty / reasonability)"),
    ("ALMA_SortBySum", "Sort by sum"),
    ("ALMA_SortByMax", "Sort by max"),
    ("ALMA_SortByColumn", "Sort rows by a column"),
    ("ALMA_NormalizeRows", "Normalize rows (horizontal)"),
    ("ALMA_NormalizeCols", "Normalize columns (vertical)"),
    ("ALMA_ClusterSort", "Cluster sort"),
    ("ALMA_Colorize", "Colorize"),
    ("ALMA_TextGenerator", "Hypothesis text generator"),
    ("ALMA_FilterMatrix", "Temporal / geographic filter matrix"),
    ("ALMA_SelfTest", "Self-test (incl. PubMed connection)"),
    ("ALMA_ShowLists", "Show Lists"),
    ("ALMA_ShowTree", "Show Tree"),
    ("ALMA_ClearCache", "Clear cache"),
]

HELP_COMMON = [
    "Query rule (same in every ALMA version): every element in parentheses joined with AND, e.g. "
    "(Annexin A1) AND (Liposome) AND (Pancreatic cancer).  A label may hold several elements separated by ' + '.",
    "Scores: reasonability 2 if NOP >= high, 1 if NOP >= medium, else 0; novelty is the inverse.  LR = row+column, "
    "THR = row+string, TVR = column+string, HR / VR = the string with the other columns / rows.  Ranking priority "
    "N, LR, TVR, THR, VR, HR.  The '_novel' sheet keeps hypotheses with <= 1 publication whose three pairs each have "
    ">= 10 publications (the paper's Fig 3c / 5c filter).",
    "PubMed allows 3 requests per second without an API key (about 10,000 counts per hour).  Counts are cached, so "
    "repeated runs are instant.",
]
HELP_VBA = [
    "How to use (macros must be enabled: File > Info > Enable Content; if the file came from the internet, "
    "right-click it > Properties > Unblock first)",
    "1. Put your categorized terms on the Lists sheet: List 1 = rows (e.g. diseases), List 2 = columns "
    "(e.g. biomaterials), column C = strings to add for comparison matrices (e.g. molecular targets).",
    "2. Click 'Full pipeline'.  ALMA counts every combination in PubMed (search matrix), adds every string of "
    "column C (comparison matrices) and ranks all 3-element hypotheses (NRF sheet).  Every other button works on "
    "the active ALMA sheet (or the newest suitable one); all sheets are listed on the Tree sheet.  Esc stops a "
    "long search.  'Self-test' checks the macros and the connection to PubMed.",
] + HELP_COMMON
HELP_XLW = [
    "How to use (Python + xlwings; no macros in this workbook)",
    "Setup once on this computer:  pip install xlwings requests pandas numpy openpyxl   then   xlwings addin install. "
    "Keep this workbook next to ALMA_xlwings.py inside the ALMA folder.",
    "1. Put your categorized terms on the Lists sheet (List 1 = rows, List 2 = columns, column C = strings for "
    "comparison matrices).  2. Choose the Action above.  3. Click 'Run main' on the xlwings ribbon tab.",
    "Without Excel:  python excel/ALMA_xlwings.py excel/ALMA_xlwings.xlsx --action \"Full pipeline\"  edits this file "
    "directly (close it in Excel first).",
] + HELP_COMMON


def L(name):
    with open(os.path.join(ROOT, "lists", name + ".txt"), encoding="utf-8") as fh:
        return [x.strip() for x in fh if x.strip()]


def cache_rows():
    if not os.path.exists(CACHE):
        return []
    with sqlite3.connect(CACHE) as db:
        return db.execute("SELECT term, datetype, maxdate, nop, retrieved FROM counts ORDER BY maxdate, term").fetchall()


class Formats:
    def __init__(self, wb):
        self.title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E79"})
        self.sub = wb.add_format({"italic": True, "font_color": "#595959"})
        self.hdr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        self.txt = wb.add_format({"num_format": "@", "border": 1, "bg_color": "#FFF2CC"})
        self.val = wb.add_format({"border": 1, "bg_color": "#FFF2CC"})
        self.note = wb.add_format({"font_color": "#595959", "text_wrap": True, "valign": "top"})
        self.wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        self.bold = wb.add_format({"bold": True})
        self.text = wb.add_format({"num_format": "@"})


def control_sheet(a, f, subtitle, settings, help_lines, buttons=None, actions=None):
    a.set_column("A:A", 12, None, {"hidden": True})
    a.set_column("B:B", 34)
    a.set_column("C:C", 26 if actions else 16)
    a.set_column("D:D", 60)
    a.set_column("E:E", 3)
    a.set_column("F:K", 13)
    a.write("B1", "ALMA - Automated Literature Meta-Analysis", f.title)
    a.write("B2", subtitle, f.sub)
    a.write("B4", "Setting", f.hdr)
    a.write("C4", "Value", f.hdr)
    a.write("D4", "Notes", f.hdr)
    for i, (key, label, val, note) in enumerate(settings):
        r = 4 + i
        a.write_string(r, 0, key)
        a.write(r, 1, label)
        if isinstance(val, bool):
            a.write_boolean(r, 2, val, f.val)
            a.data_validation(r, 2, r, 2, {"validate": "list", "source": ["TRUE", "FALSE"]})
        elif isinstance(val, (int, float)):
            a.write_number(r, 2, val, f.val)
        else:
            a.write_string(r, 2, val, f.txt)
        a.write(r, 3, note, f.note)
        if key == "datetype":
            a.data_validation(r, 2, r, 2, {"validate": "list", "source": ["edat", "pdat", "mdat"]})
        if key == "action" and actions:
            a.data_validation(r, 2, r, 2, {"validate": "list", "source": "=Library!$Z$2:$Z$%d" % (1 + len(actions))})
    if buttons:
        a.write("F4", "Actions", f.hdr)
        for i, (macro, caption) in enumerate(buttons):
            col = 5 if i < 10 else 8
            row = 4 + 2 * (i if i < 10 else i - 10)
            a.insert_button(row, col, {"macro": macro, "caption": caption, "width": 245, "height": 30})
    r0 = 4 + len(settings) + 2
    a.write(r0, 1, help_lines[0], f.bold)
    for k, line in enumerate(help_lines[1:], 1):
        a.merge_range(r0 + k, 1, r0 + k, 3, line, f.wrap)
        a.set_row(r0 + k, 48)
    a.activate()


def lists_sheet(li, f):
    heads = ["List 1 (rows)", "List 2 (columns)", "Comparison strings (var)", "Constant strings (all queries)",
             "Text generator: 1st range", "Text generator: 2nd range", "Filter: hypotheses (a + b)",
             "Filter: values (years / countries)"]
    data = [L("diseases_fig5"), L("biomaterials_fig5"), L("targets_fig5"), [],
            L("textgen_cancers_fig1b"), L("textgen_kinases_fig1b"), L("temporal_hypotheses_fig6"),
            [str(y) for y in range(2013, 2020)]]
    for j, (h, col) in enumerate(zip(heads, data)):
        li.write(0, j, h, f.hdr)
        for i, v in enumerate(col):
            li.write_string(1 + i, j, v, f.text)
        li.set_column(j, j, 30, f.text)
    params = [("Matrix name", "SM"), ("Row title", "disease"), ("Column title", "biomaterial"),
              ("Prefix", "treating"), ("Buffer", "with"), ("Suffix", "inhibitor combined with radiotherapy"),
              ("Mode (all/single)", "all"), ("Filter tag", "[dp]"), ("String to add", ""), ("Sort column", "")]
    li.write(0, 9, "Parameter", f.hdr)
    li.write(0, 10, "Value", f.hdr)
    for i, (k, v) in enumerate(params):
        li.write_string(1 + i, 9, k)
        li.write_string(1 + i, 10, v, f.txt)
    li.data_validation(7, 10, 7, 10, {"validate": "list", "source": ["all", "single"]})
    li.set_column(8, 8, 3)
    li.set_column(9, 9, 20)
    li.set_column(10, 10, 36, f.text)
    li.write(13, 9, "Lists end at the first empty cell.  Example = paper Fig 5 (diseases x biomaterials x targets);",
             f.sub)
    li.write(14, 9, "more lists from the paper are on the Library sheet.", f.sub)
    li.freeze_panes(1, 0)


def library_sheet(lib, f, actions=None):
    files = sorted(glob.glob(os.path.join(ROOT, "lists", "*.txt")))
    for j, path in enumerate(files):
        name = os.path.splitext(os.path.basename(path))[0]
        lib.write(0, j, name, f.hdr)
        for i, v in enumerate(L(name)):
            lib.write_string(1 + i, j, v, f.text)
        lib.set_column(j, j, 26)
    if actions:
        lib.write(0, 25, "actions", f.hdr)
        for i, v in enumerate(actions):
            lib.write_string(1 + i, 25, v)
    lib.freeze_panes(1, 0)


def tree_sheet(t, f):
    t.write("A1", "ALMA tree manager", f.title)
    t.write("A2", "Every matrix / table created by ALMA is listed here (newest last).", f.sub)
    for j, h in enumerate(["Sheet", "Type", "Title", "Parent", "Added string / note", "Size", "Updated"]):
        t.write(4, j, h, f.hdr)
    for j, w in enumerate([32, 12, 36, 30, 34, 12, 17]):
        t.set_column(j, j, w)


def cache_sheet(c, f):
    for j, h in enumerate(["Query", "DateType", "MaxDate", "NOP", "Retrieved"]):
        c.write(0, j, h, f.hdr)
    rows = cache_rows()
    for i, (term, dtype, maxd, nop, ret) in enumerate(rows, 1):
        c.write_string(i, 0, term, f.text)
        c.write_string(i, 1, dtype, f.text)
        c.write_string(i, 2, maxd, f.text)
        c.write_number(i, 3, nop)
        c.write_string(i, 4, ret)
    c.set_column(0, 0, 90, f.text)
    c.set_column(1, 2, 11, f.text)
    c.set_column(3, 3, 9)
    c.set_column(4, 4, 20)
    c.freeze_panes(1, 0)
    return len(rows)


def build_vba_workbook(path, vba_bin=None):
    wb = xlsxwriter.Workbook(path)
    if vba_bin is not None:
        wb.add_vba_project(io.BytesIO(vba_bin), is_stream=True)
        wb.set_vba_name("ThisWorkbook")
    f = Formats(wb)
    ws = {name: wb.add_worksheet(name) for name in VBA_SHEETS}
    if vba_bin is not None:
        for i, name in enumerate(VBA_SHEETS, 1):
            ws[name].set_vba_name("Sheet%d" % i)
    control_sheet(ws["ALMA"], f, "Excel/VBA re-implementation of Harif et al. (Technion): generate combinatorial "
                  "hypotheses, count them in PubMed and rank the unpublished ones by novelty and reasonability.",
                  VBA_SETTINGS, HELP_VBA, buttons=BUTTONS)
    lists_sheet(ws["Lists"], f)
    library_sheet(ws["Library"], f)
    tree_sheet(ws["Tree"], f)
    n = cache_sheet(ws["Cache"], f)
    wb.close()
    return n


def build_xlwings_workbook(path):
    wb = xlsxwriter.Workbook(path)
    f = Formats(wb)
    ws = {name: wb.add_worksheet(name) for name in XLW_SHEETS}
    control_sheet(ws["ALMA"], f, "Python version for Excel (xlwings): choose an Action and click 'Run main' on the "
                  "xlwings ribbon tab.  Runs excel/ALMA_xlwings.py with the alma package.", XLW_SETTINGS, HELP_XLW,
                  actions=ACTIONS)
    lists_sheet(ws["Lists"], f)
    library_sheet(ws["Library"], f, actions=ACTIONS)
    tree_sheet(ws["Tree"], f)
    wb.close()


def main():
    codenames = ["Sheet%d" % i for i in range(1, len(VBA_SHEETS) + 1)]
    vba_bin, modules = vbaproject.build(BAS, codenames)
    with open(os.path.join(HERE, "vbaProject.bin"), "wb") as fh:
        fh.write(vba_bin)
    n = build_vba_workbook(os.path.join(HERE, "ALMA.xlsm"), vba_bin)
    build_vba_workbook(os.path.join(HERE, "ALMA_template.xlsx"))
    build_xlwings_workbook(os.path.join(HERE, "ALMA_xlwings.xlsx"))
    print("vbaProject.bin: %d bytes, modules: %s" % (len(vba_bin), ", ".join(m["name"] for m in modules)))
    print("wrote excel/ALMA.xlsm, excel/ALMA_template.xlsx, excel/ALMA_xlwings.xlsx (%d cached PubMed counts)" % n)


if __name__ == "__main__":
    main()
