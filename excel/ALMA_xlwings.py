"""ALMA in Excel through xlwings: a Python script you run from Excel.

One-time setup on the computer with Excel (Windows or macOS):

    pip install xlwings requests pandas numpy openpyxl
    xlwings addin install

Keep this file next to ALMA_xlwings.xlsx inside the ALMA folder (it imports the
``alma`` package from the folder above).  Open ALMA_xlwings.xlsx, pick an
Action on the ALMA sheet and click "Run main" on the xlwings ribbon tab.

The script also runs without Excel, editing the workbook file directly:

    python excel/ALMA_xlwings.py excel/ALMA_xlwings.xlsx --action "Full pipeline"
    python excel/ALMA_xlwings.py book.xlsx --action "Search matrix" --set maxdate=2020/02/10 \
        --list A=lists/kinase_inhibitors_50.txt --list B=lists/cancers_fig2c_8.txt

Sheet layout, list columns and results are the same as in the VBA version
(excel/ALMA.xlsm) and the Python command line (python -m alma).
"""
import argparse
import datetime
import os
import shutil
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from alma import core  # noqa: E402
from alma.pubmed import PubMed  # noqa: E402

HDR_ROW = 9
TREE_FIRST = 6
SEP = core.SEP
ACTIONS = ["Full pipeline", "Search matrix", "Comparison matrices + NRF", "NRF ranking", "Add constant string",
           "Sort by sum", "Sort by max", "Sort by column", "Normalize rows", "Normalize columns", "Cluster sort",
           "Colorize", "Text generator", "Filter matrix", "Self-test", "Clear cache"]
LIST_COLS = {"rows": 1, "cols": 2, "vars": 3, "context": 4, "first": 5, "second": 6, "hyps": 7, "values": 8}
RED, YELLOW, GREEN = (248, 105, 107), (255, 235, 132), (99, 190, 123)


# ============================================================ workbook access
class XlwingsIO:
    """The open Excel workbook (through xlwings)."""

    def __init__(self, book):
        self.book = book
        try:
            self.folder = os.path.dirname(book.fullname) if os.path.isfile(book.fullname) else HERE
        except Exception:  # noqa: BLE001
            self.folder = HERE
        self._app_state(False)

    def _app_state(self, on):
        try:
            self.book.app.screen_updating = on
        except Exception:  # noqa: BLE001
            pass

    def names(self):
        return [s.name for s in self.book.sheets]

    def active(self):
        return self.book.sheets.active.name

    def values(self, sheet):
        rng = self.book.sheets[sheet].used_range
        vals = rng.options(ndim=2).value
        r0, c0 = rng.row, rng.column
        width = c0 - 1 + max(len(r) for r in vals)
        return [[None] * width for _ in range(r0 - 1)] + [[None] * (c0 - 1) + list(r) for r in vals]

    def add(self, name):
        self.book.sheets.add(name=name, after=self.book.sheets[-1])

    def put(self, sheet, row, col, rows):
        self.book.sheets[sheet].range((row, col)).value = rows

    def fmt(self, sheet, r1, c1, r2, c2, number_format):
        self.book.sheets[sheet].range((r1, c1), (r2, c2)).number_format = number_format

    def fill(self, sheet, row, c1, c2, rgb):
        self.book.sheets[sheet].range((row, c1), (row, c2)).color = rgb

    def bold(self, sheet, r1, c1, r2, c2):
        self.book.sheets[sheet].range((r1, c1), (r2, c2)).font.bold = True

    def width(self, sheet, col, w):
        self.book.sheets[sheet].range((1, col)).column_width = w

    def activate(self, sheet):
        self.book.sheets[sheet].activate()

    def status(self, text):
        try:
            self.book.app.status_bar = text
        except Exception:  # noqa: BLE001
            pass

    def finish(self):
        self._app_state(True)
        self.status(False)


class OpenpyxlIO:
    """A workbook file (Excel closed): same operations through openpyxl."""

    def __init__(self, path):
        import openpyxl
        self.path = path
        self.folder = os.path.dirname(os.path.abspath(path))
        self.wb = openpyxl.load_workbook(path)

    def names(self):
        return list(self.wb.sheetnames)

    def active(self):
        return self.wb.active.title

    def values(self, sheet):
        ws = self.wb[sheet]
        return [list(r) for r in ws.iter_rows(min_row=1, max_row=max(1, ws.max_row), max_col=max(1, ws.max_column),
                                              values_only=True)]

    def add(self, name):
        self.wb.create_sheet(name)

    def put(self, sheet, row, col, rows):
        ws = self.wb[sheet]
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                ws.cell(row=row + i, column=col + j).value = v   # value=None in cell() would not clear

    def fmt(self, sheet, r1, c1, r2, c2, number_format):
        ws = self.wb[sheet]
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).number_format = number_format

    def fill(self, sheet, row, c1, c2, rgb):
        from openpyxl.styles import PatternFill
        f = PatternFill("solid", fgColor="%02X%02X%02X" % rgb)
        ws = self.wb[sheet]
        for c in range(c1, c2 + 1):
            ws.cell(row=row, column=c).fill = f

    def bold(self, sheet, r1, c1, r2, c2):
        from openpyxl.styles import Font
        ws = self.wb[sheet]
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).font = Font(bold=True)

    def width(self, sheet, col, w):
        from openpyxl.utils import get_column_letter
        self.wb[sheet].column_dimensions[get_column_letter(col)].width = w

    def activate(self, sheet):
        self.wb.active = self.wb.sheetnames.index(sheet)

    def status(self, text):
        if text:
            print("  " + str(text), flush=True)

    def finish(self):
        self.wb.save(self.path)


# ================================================================== helpers
def label(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y/%m/%d")
    return str(v).strip()


def number(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def cell(vals, r, c):
    """1-based access to a values grid, None outside."""
    if r - 1 < len(vals) and c - 1 < len(vals[r - 1]):
        return vals[r - 1][c - 1]
    return None


def truthy(v):
    if isinstance(v, bool):
        return v
    return label(v).lower() in ("true", "yes", "y", "1", "wahr", "vrai")


def color(v, hi):
    """Red (0) -> yellow (hi/2) -> green (>= hi), in 20 steps so equal colours form runs."""
    t = min(1.0, max(0.0, v / (hi or 1)))
    t = round(t * 20) / 20
    a, b, f = (RED, YELLOW, t / 0.5) if t < 0.5 else (YELLOW, GREEN, (t - 0.5) / 0.5)
    return tuple(int(round(a[k] + (b[k] - a[k]) * f)) for k in range(3))


class Alma:
    def __init__(self, io):
        self.io = io
        self.s = self.read_settings()
        self.lists, self.params = self.read_lists()
        self.high = self.num("high", 20)
        self.med = self.num("med", 5)

    # --------------------------------------------------------- settings
    def read_settings(self):
        vals = self.io.values("ALMA")
        return {label(r[0]).lower(): (r[2] if len(r) > 2 else None) for r in vals if r and label(r[0])}

    def set_setting(self, key, value):
        vals = self.io.values("ALMA")
        for i, r in enumerate(vals, 1):
            if r and label(r[0]).lower() == key:
                self.io.put("ALMA", i, 3, [[value]])
                self.s[key] = value
                return
        raise KeyError(key)

    def num(self, key, default):
        v = self.s.get(key)
        return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else \
            (float(label(v)) if label(v).replace(".", "", 1).isdigit() else default)

    def read_lists(self):
        vals = self.io.values("Lists")
        lists = {}
        for key, col in LIST_COLS.items():
            items = []
            for r in range(2, len(vals) + 1):
                v = label(cell(vals, r, col))
                if not v:
                    break
                items.append(v)
            lists[key] = items
        params = {}
        for r in range(1, len(vals) + 1):
            k = label(cell(vals, r, 10)).lower()
            if k:
                params[k] = label(cell(vals, r, 11))
        return lists, params

    def param(self, key, default=""):
        return self.params.get(key.lower()) or default

    def pubmed(self):
        maxd = label(self.s.get("maxdate")).replace("-", "/").replace(".", "/")
        dtype = label(self.s.get("datetype")) or "edat"
        cache = label(self.s.get("cachefile")) or "ALMA_cache.sqlite"
        if not os.path.isabs(cache):
            cache = os.path.join(self.io.folder, cache)
        seed = os.path.join(os.path.dirname(HERE), "results", "pubmed_cache.sqlite")
        if not os.path.exists(cache) and os.path.exists(seed):
            shutil.copyfile(seed, cache)          # start from the counts of the reproduction runs
        pm = PubMed(cache, max_date=maxd or None, date_type=dtype, api_key=label(self.s.get("apikey")),
                    offline=truthy(self.s.get("offline")), verbose=False)
        return pm

    def note(self):
        maxd = label(self.s.get("maxdate"))
        return ("maxdate %s (%s)" % (maxd, label(self.s.get("datetype")) or "edat")) if maxd else \
            "no date limit (searched %s)" % datetime.date.today().strftime("%Y/%m/%d")

    # ----------------------------------------------------------- sheets
    def new_sheet(self, title):
        base = "".join("_" if ch in "[]:*?/\\'" else ch for ch in title).strip()[:31] or "Sheet"
        used = {n.lower() for n in self.io.names()}
        name, i = base, 1
        while name.lower() in used:
            i += 1
            suffix = "~%d" % i
            name = base[:31 - len(suffix)] + suffix
        self.io.add(name)
        return name

    def header(self, sheet, kind, title, context="", tag="", parent="", note=""):
        rows = [["ALMA_MATRIX", kind], ["Name", title], ["Context", context], ["Column tag", tag],
                ["Parent", parent], ["Var / Note", note], ["PubMed", self.note()]]
        self.io.fmt(sheet, 1, 2, 7, 2, "@")
        self.io.put(sheet, 1, 1, rows)
        self.io.bold(sheet, 2, 2, 2, 2)

    def colour(self, sheet, r0, c0, grid, hi):
        for i, row in enumerate(grid):
            run_start, run_col = None, None
            for j in range(len(row) + 1):
                v = row[j] if j < len(row) else None
                col = color(v, hi) if isinstance(v, (int, float)) and not isinstance(v, bool) else None
                if col != run_col:
                    if run_col is not None:
                        self.io.fill(sheet, r0 + i, c0 + run_start, c0 + j - 1, run_col)
                    run_start, run_col = j, col
        return

    def tree(self, sheet, kind, title, parent, added, size):
        vals = self.io.values("Tree")
        r = TREE_FIRST
        while label(cell(vals, r, 1)):
            r += 1
        self.io.fmt("Tree", r, 1, r, 6, "@")
        self.io.put("Tree", r, 1, [[sheet, kind, title, parent, added, size,
                                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M")]])

    def write_matrix(self, m, title):
        sh = self.new_sheet(title)
        self.header(sh, m.kind, title, SEP.join(m.context), m.col_tag, m.parent, m.note)
        nr, nc = m.nop.shape
        self.io.fmt(sh, HDR_ROW, 1, HDR_ROW + nr, 1, "@")
        self.io.fmt(sh, HDR_ROW, 2, HDR_ROW, nc + 1, "@")
        grid = [[float(x) for x in row] for row in m.nop.values]
        rows = [["%s \\ %s" % (m.row_title, m.col_title)] + list(m.cols)]
        rows += [[r] + g for r, g in zip(m.rows, grid)]
        self.io.put(sh, HDR_ROW, 1, rows)
        if m.kind == "normalized":
            self.io.fmt(sh, HDR_ROW + 1, 2, HDR_ROW + nr, nc + 1, "0.00")
        self.colour(sh, HDR_ROW + 1, 2, grid, 1.0 if m.kind == "normalized" else self.high)
        self.style(sh, nc + 1)
        self.tree(sh, m.kind, title, m.parent, m.note, "%d x %d" % (nr, nc))
        return sh

    def style(self, sh, last_col):
        self.io.bold(sh, HDR_ROW, 1, HDR_ROW, last_col)
        self.io.width(sh, 1, 30)
        for c in range(2, last_col + 1):
            self.io.width(sh, c, 11)

    def write_comparison(self, c, title):
        m = c.base
        sh = self.new_sheet(title)
        self.header(sh, "comparison", title, SEP.join(m.context), m.col_tag, m.name, c.var)
        nr, nc = m.nop.shape
        top = ["var=" + c.var, c.var]
        second = [c.var, float(c.var_total)]
        for col in m.cols:
            top += [col, "& Var"]
            second += [float(c.col_var[col]), None]
        rows = [top, second]
        grid = [second[1:]]
        for r in m.rows:
            line = [float(c.row_var[r])]
            for col in m.cols:
                line += [float(m.nop.at[r, col]), float(c.triple.at[r, col])]
            rows.append([r] + line)
            grid.append(line)
        self.io.fmt(sh, HDR_ROW, 1, HDR_ROW + nr + 1, 1, "@")
        self.io.fmt(sh, HDR_ROW, 1, HDR_ROW, 2 + 2 * nc, "@")
        self.io.put(sh, HDR_ROW, 1, rows)
        self.colour(sh, HDR_ROW + 1, 2, grid, self.high)
        self.style(sh, 2 + 2 * nc)
        self.tree(sh, "comparison", title, m.name, c.var, "%d x %d" % (nr, nc))
        return sh

    def write_table(self, df, title, kind, parent="", note=""):
        sh = self.new_sheet(title)
        self.header(sh, kind, title, "", "", parent, note)
        cols = list(df.columns)
        for j, col in enumerate(cols, 1):
            if col in ("hypothesis", "row", "col", "var", "first", "second"):
                self.io.fmt(sh, HDR_ROW, j, HDR_ROW + len(df), j, "@")
        rows = [cols] + [[(v.item() if hasattr(v, "item") else v) for v in r] for r in df.itertuples(index=False)]
        self.io.put(sh, HDR_ROW, 1, rows)
        for j, col in enumerate(cols, 1):
            hi = 2 if col in ("N", "LR", "TVR", "THR", "VR", "HR") else \
                (self.high if col in ("NOP_LR", "NOP_THR", "NOP_TVR") or (col == "NOP" and kind == "text") else None)
            if hi and len(df):
                self.colour(sh, HDR_ROW + 1, j, [[float(v)] for v in df[col]], hi)
        self.io.bold(sh, HDR_ROW, 1, HDR_ROW, len(cols))
        for j, col in enumerate(cols, 1):
            self.io.width(sh, j, 36 if col in ("hypothesis", "row", "col", "var", "first", "second") else
                          (8 if col == "rank" else 10))
        self.tree(sh, kind, title, parent, note, "%d rows" % len(df))
        return sh

    # ------------------------------------------------------------ reading
    def kind(self, sheet):
        vals = self.io.values(sheet)
        return label(cell(vals, 1, 2)) if label(cell(vals, 1, 1)) == "ALMA_MATRIX" else ""

    def read_matrix(self, sheet):
        vals = self.io.values(sheet)
        if label(cell(vals, 1, 1)) != "ALMA_MATRIX":
            raise ValueError("sheet %r is not an ALMA matrix" % sheet)
        cols, c = [], 2
        while label(cell(vals, HDR_ROW, c)):
            cols.append(label(cell(vals, HDR_ROW, c)))
            c += 1
        rows, r = [], HDR_ROW + 1
        while label(cell(vals, r, 1)):
            rows.append(label(cell(vals, r, 1)))
            r += 1
        nop = pd.DataFrame([[number(cell(vals, HDR_ROW + 1 + i, 2 + j)) for j in range(len(cols))]
                            for i in range(len(rows))], index=rows, columns=cols)
        corner = label(cell(vals, HDR_ROW, 1))
        rt, ct = (corner.split(" \\ ", 1) + ["columns"])[:2] if " \\ " in corner else ("rows", "columns")
        ctx = label(cell(vals, 3, 2))
        return core.Matrix(sheet, nop, kind=label(cell(vals, 1, 2)), context=[ctx] if ctx else [],
                           col_tag=label(cell(vals, 4, 2)), row_title=rt, col_title=ct,
                           parent=label(cell(vals, 5, 2)), note=label(cell(vals, 6, 2)))

    def read_comparison(self, sheet):
        vals = self.io.values(sheet)
        cols, c = [], 3
        while label(cell(vals, HDR_ROW, c)):
            cols.append(label(cell(vals, HDR_ROW, c)))
            c += 2
        rows, r = [], HDR_ROW + 2
        while label(cell(vals, r, 1)):
            rows.append(label(cell(vals, r, 1)))
            r += 1
        g = lambda i, j: number(cell(vals, i, j))  # noqa: E731
        base = pd.DataFrame([[g(HDR_ROW + 2 + i, 3 + 2 * j) for j in range(len(cols))] for i in range(len(rows))],
                            index=rows, columns=cols)
        tri = pd.DataFrame([[g(HDR_ROW + 2 + i, 4 + 2 * j) for j in range(len(cols))] for i in range(len(rows))],
                           index=rows, columns=cols)
        ctx = label(cell(vals, 3, 2))
        m = core.Matrix(label(cell(vals, 5, 2)), base, context=[ctx] if ctx else [], col_tag=label(cell(vals, 4, 2)))
        return core.Comparison(sheet, m, label(cell(vals, 6, 2)), g(HDR_ROW + 1, 2),
                               pd.Series({rw: g(HDR_ROW + 2 + i, 2) for i, rw in enumerate(rows)}),
                               pd.Series({cl: g(HDR_ROW + 1, 3 + 2 * j) for j, cl in enumerate(cols)}), tri)

    def target(self, want):
        ok = {"matrix": ("search", "normalized"), "search": ("search",), "comparison": ("comparison",),
              "nrfsource": ("search", "comparison"), "any": None}[want]

        def good(name):
            k = self.kind(name)
            return bool(k) and (ok is None or k in ok)

        t = label(self.s.get("target"))
        if t:
            if t not in self.io.names() or not good(t):
                raise ValueError("target sheet %r is not a suitable ALMA sheet (%s)" % (t, want))
            return t
        a = self.io.active()
        if good(a):
            return a
        vals = self.io.values("Tree")
        names = [label(cell(vals, r, 1)) for r in range(TREE_FIRST, len(vals) + 1)]
        for n in reversed([n for n in names if n]):
            if n in self.io.names() and good(n):
                return n
        raise ValueError("no suitable ALMA sheet found (%s) - create or select one first" % want)

    # ------------------------------------------------------------ actions
    def search(self):
        L = self.lists
        if not L["rows"] or not L["cols"]:
            raise ValueError("fill List 1 (column A) and List 2 (column B) on the Lists sheet")
        m = core.search_matrix(self.pubmed(), L["rows"], L["cols"], self.param("Matrix name", "SM"),
                               context=[SEP.join(L["context"])] if L["context"] else [],
                               row_title=self.param("Row title", "rows"), col_title=self.param("Column title", "columns"))
        sh = self.write_matrix(m, self.param("Matrix name", "SM"))
        self.io.activate(sh)
        return sh

    def comparisons(self, sheet, variables):
        if not variables:
            raise ValueError("fill the comparison strings (column C) on the Lists sheet")
        m = self.read_matrix(sheet)
        if m.kind != "search":
            raise ValueError("comparison matrices need a search matrix (sheet %r is %s)" % (sheet, m.kind))
        pm = self.pubmed()
        comps = []
        for var in variables:
            c = core.comparison(pm, m, var)
            self.write_comparison(c, "%s_cmp_%s" % (sheet, var))
            comps.append(c)
        return self.nrf_sheets(sheet, comps)

    def nrf_sheets(self, sheet, comps):
        df = core.rank(pd.concat([core.nrf(c, self.high, self.med) for c in comps], ignore_index=True))
        nmax, pmin = self.num("novelmax", 1), self.num("pairmin", 10)
        df["novel_reasonable"] = df.index.isin(core.novel_and_reasonable(df, nmax, pmin).index)
        sh = self.write_table(df, sheet + "_NRF", "nrf", sheet, "high>=%g med>=%g" % (self.high, self.med))
        nov = df[df.novel_reasonable]
        self.write_table(nov, sheet + "_novel", "nrf", sh, "N<=%g publications and every pair >=%g" % (nmax, pmin))
        self.io.activate(sh)
        return "%d comparison matrices, NRF ranking %r, %d novel & reasonable hypotheses" % (len(comps), sh, len(nov))

    def act_pipeline(self):
        sh = self.search()
        self.write_matrix(core.sort_matrix(self.read_matrix(sh), "sum"), sh + "_sorted")
        return "search matrix %r; %s" % (sh, self.comparisons(sh, self.lists["vars"]))

    def act_search(self):
        return "search matrix %r created" % self.search()

    def act_comparisons(self):
        return self.comparisons(self.target("search"), self.lists["vars"])

    def act_nrf(self):
        t = self.target("nrfsource")
        if self.kind(t) == "comparison":
            comps = [self.read_comparison(t)]
        else:
            name2 = label(cell(self.io.values(t), 2, 2))
            comps = [self.read_comparison(n) for n in self.io.names() if self.kind(n) == "comparison"
                     and label(cell(self.io.values(n), 5, 2)) in (t, name2)]
            if not comps:
                raise ValueError("no comparison matrices of %r found - create them first" % t)
        return self.nrf_sheets(t, comps)

    def act_addstring(self):
        t = self.target("search")
        m = self.read_matrix(t)
        s = self.param("String to add")
        if not s:
            raise ValueError("put the string in 'String to add' on the Lists sheet")
        r = core.search_matrix(self.pubmed(), m.rows, m.cols, t + "+" + s, context=m.context + [s],
                               col_tag=m.col_tag, row_title=m.row_title, col_title=m.col_title)
        r.parent, r.note = t, "added string: " + s
        sh = self.write_matrix(r, t + "+" + s)
        self.io.activate(sh)
        return "search matrix %r (+ %s)" % (sh, s)

    def _derived(self, m, title):
        sh = self.write_matrix(m, title)
        self.io.activate(sh)
        return sh

    def act_sort(self, by):
        t = self.target("matrix")
        r = core.sort_matrix(self.read_matrix(t), by)
        return "sorted matrix %r (by %s)" % (self._derived(r, t + "_sorted" + ("" if by == "sum" else "_" + by)), by)

    def act_sortcol(self):
        t = self.target("matrix")
        m = self.read_matrix(t)
        col = self.param("Sort column") or m.cols[-1]
        match = [c for c in m.cols if c.lower() == col.lower()]
        if not match:
            raise ValueError("column %r not found" % col)
        return "matrix %r sorted by %s" % (self._derived(core.sort_by_column(m, match[0]), t + "_by_" + match[0]),
                                           match[0])

    def act_normalize(self, axis):
        t = self.target("matrix")
        return "normalized matrix %r" % self._derived(core.normalize(self.read_matrix(t), axis),
                                                      "%s_norm_%s" % (t, axis))

    def act_cluster(self):
        t = self.target("matrix")
        m = self.read_matrix(t)
        axis = "rows"
        if m.kind != "normalized":
            t = self.write_matrix(core.normalize(m, "rows"), t + "_norm_rows")
            m = self.read_matrix(t)
        elif "col" in m.note:
            axis = "cols"
        thr = self.num("cluster", 0.9)
        cm, _ = core.cluster_sort(m, thr, axis)
        return "cluster-sorted matrix %r (threshold %g)" % (self._derived(cm, t + "_clusters"), thr)

    def act_colorize(self):
        t = self.target("any")
        k = self.kind(t)
        if k in ("search", "normalized"):
            m = self.read_matrix(t)
            self.colour(t, HDR_ROW + 1, 2, m.nop.values.tolist(), 1.0 if k == "normalized" else self.high)
        elif k == "comparison":
            c = self.read_comparison(t)
            vals = self.io.values(t)
            grid = [[number(cell(vals, HDR_ROW + 1 + i, 2 + j)) if cell(vals, HDR_ROW + 1 + i, 2 + j) is not None
                     else None for j in range(1 + 2 * len(c.base.cols))] for i in range(1 + len(c.base.rows))]
            self.colour(t, HDR_ROW + 1, 2, grid, self.high)
        return "colours updated on %r (high threshold %g)" % (t, self.high)

    def act_textgen(self):
        L = self.lists
        if not L["first"] or not L["second"]:
            raise ValueError("fill the text generator ranges (columns E and F) on the Lists sheet")
        df = core.text_generator(self.pubmed(), self.param("Prefix"), L["first"], self.param("Buffer"),
                                 L["second"], self.param("Suffix"), (self.param("Mode (all/single)", "all")).lower())
        sh = self.write_table(df, "TextGen", "text", "", core.text_query(
            self.param("Prefix"), "<first>", self.param("Buffer"), "<second>", self.param("Suffix")))
        self.io.activate(sh)
        return "%d generated hypotheses in %r" % (len(df), sh)

    def act_filter(self):
        L = self.lists
        if not L["hyps"] or not L["values"]:
            raise ValueError("fill the filter hypotheses (column G) and values (column H) on the Lists sheet")
        tag = self.param("Filter tag", "[dp]")
        m = core.filter_matrix(self.pubmed(), L["hyps"], L["values"], tag, "FM")
        sh = self.write_matrix(m, "FM")
        sh2 = self._derived(core.normalize(self.read_matrix(sh), "rows"), sh + "_norm_rows")
        return "filter matrix %r (%d x %d, tag %s) and %r" % (sh, len(L["hyps"]), len(L["values"]), tag, sh2)

    def act_selftest(self):
        tests = [("and_query", core.and_query("palbociclib", "Head and neck cancer + radiotherapy"),
                  "(palbociclib) AND (Head and neck cancer) AND (radiotherapy)"),
                 ("text_query", core.text_query("treating", "Lung cancer", "with", "EGFR",
                                                "inhibitor combined with radiotherapy"),
                  "treating Lung cancer with EGFR inhibitor combined with radiotherapy"),
                 ("scores", "%d,%d,%d,%d" % (core.reason_score(16, 20, 5), core.reason_score(24, 20, 5),
                                             core.novelty_score(0, 20, 5), core.novelty_score(92, 20, 5)), "1,2,2,0")]
        toy = core.Matrix("toy", pd.DataFrame([[1, .1, 0], [.92, 1, 0], [0, .3, 1]], index=list("abc"),
                                              columns=list("XYZ")), kind="normalized")
        s = core.sort_matrix(toy, "sum")
        tests.append(("sort by sum", ",".join(s.rows) + "|" + ",".join(s.cols), "b,c,a|X,Y,Z"))
        s, _ = core.cluster_sort(toy, 0.9)
        tests.append(("cluster sort", ",".join(s.rows) + "|" + ",".join(s.cols), "a,b,c|X,Z,Y"))
        if truthy(self.s.get("offline")):
            tests.append(("PubMed online test", "skipped (offline mode)", "skipped (offline mode)"))
        else:
            pm = PubMed(os.path.join(self.io.folder, "ALMA_selftest.sqlite"), max_date="2020/02/10", verbose=False)
            pm.db.execute("DELETE FROM counts")
            n = pm.count("(gefitinib) AND (head and neck cancer)")
            tests.append(("PubMed online: gefitinib + HNC to 2020/02/10 (paper 420)",
                          "in range" if 380 <= n <= 480 else str(n), "in range"))
        df = pd.DataFrame(tests, columns=["test", "result", "expected"])
        df["status"] = np.where(df.result == df.expected, "PASS", "FAIL")
        sh = self.write_table(df, "SelfTest", "selftest")
        self.io.activate(sh)
        return "%d passed, %d failed (sheet %r)" % ((df.status == "PASS").sum(), (df.status == "FAIL").sum(), sh)

    def act_clearcache(self):
        pm = self.pubmed()
        n = pm.db.execute("SELECT COUNT(*) FROM counts").fetchone()[0]
        pm.db.execute("DELETE FROM counts")
        pm.db.commit()
        return "%d cached counts deleted" % n

    def run(self, action):
        table = {"full pipeline": self.act_pipeline, "search matrix": self.act_search,
                 "comparison matrices + nrf": self.act_comparisons, "nrf ranking": self.act_nrf,
                 "add constant string": self.act_addstring, "sort by sum": lambda: self.act_sort("sum"),
                 "sort by max": lambda: self.act_sort("max"), "sort by column": self.act_sortcol,
                 "normalize rows": lambda: self.act_normalize("rows"),
                 "normalize columns": lambda: self.act_normalize("cols"), "cluster sort": self.act_cluster,
                 "colorize": self.act_colorize, "text generator": self.act_textgen, "filter matrix": self.act_filter,
                 "self-test": self.act_selftest, "clear cache": self.act_clearcache}
        key = action.strip().lower()
        if key not in table:
            raise ValueError("unknown action %r; choose one of: %s" % (action, ", ".join(ACTIONS)))
        self.io.status("ALMA: %s ..." % action)
        return table[key]()


def run(io, action=None):
    """Run one action on a workbook; returns the status message (also written to 'Last run')."""
    alma = Alma(io)
    action = action or label(alma.s.get("action")) or "Full pipeline"
    try:
        msg = "OK %s: %s" % (action, alma.run(action))
    except Exception as e:  # noqa: BLE001
        msg = "ERROR in %s: %s" % (action, e)
        alma.set_setting("laststatus", msg)
        io.finish()
        raise
    alma.set_setting("laststatus", msg)
    io.finish()
    return msg


def main():
    """Entry point for xlwings 'Run main' (the workbook is the caller)."""
    import xlwings as xw
    book = xw.Book.caller()
    io = XlwingsIO(book)
    msg = run(io)
    quiet = False
    try:
        quiet = truthy(Alma(io).s.get("quiet"))
        if not quiet:
            book.app.alert(msg, title="ALMA")
    except Exception:  # noqa: BLE001
        pass
    return msg


def cli(argv=None):
    ap = argparse.ArgumentParser(description="Run an ALMA action on ALMA_xlwings.xlsx (file mode, Excel closed)")
    ap.add_argument("workbook", nargs="?", default=os.path.join(HERE, "ALMA_xlwings.xlsx"))
    ap.add_argument("--action", default=None, help="one of: " + ", ".join(ACTIONS))
    ap.add_argument("--target", default=None, help="sheet to work on (default: active / newest suitable)")
    ap.add_argument("--set", action="append", default=[], metavar="KEY=VALUE", help="change a setting first")
    ap.add_argument("--list", action="append", default=[], metavar="COL=FILE",
                    help="replace a list column (A-H) of the Lists sheet with the lines of FILE")
    ap.add_argument("--param", action="append", default=[], metavar="NAME=VALUE", help="set a Lists parameter")
    ap.add_argument("--excel", action="store_true", help="drive the workbook open in Excel through xlwings")
    a = ap.parse_args(argv)
    if a.excel:
        import xlwings as xw
        io = XlwingsIO(xw.Book(a.workbook))
    else:
        io = OpenpyxlIO(a.workbook)
    alma = Alma(io)
    for kv in a.set + (["target=" + a.target] if a.target is not None else []):
        k, v = kv.split("=", 1)
        v = True if v.upper() == "TRUE" else False if v.upper() == "FALSE" else (float(v) if _isnum(v) else v)
        alma.set_setting(k.lower(), v)
    for kv in a.list:
        col, path = kv.split("=", 1)
        c = ord(col.strip().upper()) - 64
        with open(path, encoding="utf-8") as fh:
            items = [ln.strip() for ln in fh if ln.strip()]
        vals = io.values("Lists")
        old = sum(1 for r in range(2, len(vals) + 1) if label(cell(vals, r, c)))
        io.fmt("Lists", 2, c, 1 + max(old, len(items)), c, "@")
        io.put("Lists", 2, c, [[v] for v in items] + [[None]] * max(0, old - len(items)))
    for kv in a.param:
        k, v = kv.split("=", 1)
        vals = io.values("Lists")
        for r in range(1, len(vals) + 1):
            if label(cell(vals, r, 10)).lower() == k.lower():
                io.put("Lists", r, 11, [[v]])
                break
        else:
            raise KeyError(k)
    try:
        print(run(io, a.action))
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        return 1
    return 0


def _isnum(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


if __name__ == "__main__":
    sys.exit(cli())
