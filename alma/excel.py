"""Write ALMA results as Excel workbooks.

The sheet layout is the one the Excel/VBA version of ALMA reads and writes
(excel/vba/ALMA_Core.bas), so a workbook produced here can be opened in
Excel and processed further with the macros (sort, normalize, compare, NRF ...).

  row 1   ALMA_MATRIX | <kind: search, normalized, comparison, nrf, text>
  row 2   Name        | sheet title
  row 3   Context     | constant strings added to every query (" + " separated)
  row 4   Column tag  | PubMed field tag appended to column labels ("[dp]", "[ad]")
  row 5   Parent      | sheet it was derived from
  row 6   Var / Note  | added string of a comparison matrix, or a note
  row 7   PubMed      | date limit used for the counts
  row 9   header row of the matrix / table, data from row 10
"""
import re

import pandas as pd
import xlsxwriter

HDR_ROW = 8          # 0-based row of the matrix header (row 9 in Excel)
RED, YELLOW, GREEN = "#F8696B", "#FFEB84", "#63BE7B"


def sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\']", "_", name)[:31] or "Sheet"
    s, i = base, 2
    while s.lower() in used:
        suffix = "~%d" % i
        s, i = base[:31 - len(suffix)] + suffix, i + 1
    used.add(s.lower())
    return s


class AlmaWorkbook:
    def __init__(self, path, high=20, med=5, pubmed_note="", vba_project=None):
        self.wb = xlsxwriter.Workbook(path)
        if vba_project:
            self.wb.add_vba_project(vba_project)
        self.high, self.med = high, med
        self.pubmed_note = pubmed_note
        self.used = set()
        self.f_bold = self.wb.add_format({"bold": True})
        self.f_title = self.wb.add_format({"bold": True, "font_size": 14})
        self.f_hdr = self.wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "text_wrap": True,
                                         "valign": "top"})
        self.f_lbl = self.wb.add_format({"bold": True, "border": 1})
        self.f_num = self.wb.add_format({"border": 1})
        self.f_dec = self.wb.add_format({"border": 1, "num_format": "0.00"})
        self.f_grey = self.wb.add_format({"font_color": "#808080"})
        self.tree_ws = self.wb.add_worksheet(sheet_name("Tree", self.used))
        self.tree = []

    # ------------------------------------------------------------ helpers
    def _header(self, ws, kind, name, context="", tag="", parent="", var="", ):
        rows = [("ALMA_MATRIX", kind), ("Name", name), ("Context", context), ("Column tag", tag),
                ("Parent", parent), ("Var / Note", var), ("PubMed", self.pubmed_note)]
        for i, (k, v) in enumerate(rows):
            ws.write(i, 0, k, self.f_grey)
            ws.write(i, 1, v, self.f_bold if i == 1 else None)

    def _scale(self, ws, r0, c0, r1, c1, lo=0, hi=None):
        hi = self.high if hi is None else hi
        if r1 < r0 or c1 < c0:
            return
        ws.conditional_format(r0, c0, r1, c1, {
            "type": "3_color_scale", "min_type": "num", "min_value": lo, "min_color": RED,
            "mid_type": "num", "mid_value": lo + (hi - lo) / 2.0, "mid_color": YELLOW,
            "max_type": "num", "max_value": hi, "max_color": GREEN})

    def _log(self, sheet, kind, title, parent, added, size):
        self.tree.append((sheet, kind, title, parent, added, size))

    # ------------------------------------------------------------- writers
    def add_matrix(self, m, title=None):
        sn = sheet_name(title or m.name, self.used)
        ws = self.wb.add_worksheet(sn)
        self._header(ws, m.kind, m.name, " + ".join(m.context), m.col_tag, m.parent, m.note)
        ws.write(HDR_ROW, 0, "%s \\ %s" % (m.row_title, m.col_title), self.f_hdr)
        fmt = self.f_dec if m.kind == "normalized" else self.f_num
        for j, c in enumerate(m.cols):
            ws.write(HDR_ROW, j + 1, c, self.f_hdr)
        for i, r in enumerate(m.rows):
            ws.write(HDR_ROW + 1 + i, 0, r, self.f_lbl)
            for j, c in enumerate(m.cols):
                ws.write_number(HDR_ROW + 1 + i, j + 1, float(m.nop.iat[i, j]), fmt)
        hi = 1.0 if m.kind == "normalized" else None
        self._scale(ws, HDR_ROW + 1, 1, HDR_ROW + len(m.rows), len(m.cols), hi=hi)
        ws.set_column(0, 0, max(14, min(45, max([len(str(r)) for r in m.rows] + [10]))))
        ws.set_column(1, len(m.cols), 11)
        ws.set_row(HDR_ROW, 45)
        ws.freeze_panes(HDR_ROW + 1, 1)
        self._log(sn, m.kind, m.name, m.parent, m.note, "%d x %d" % m.nop.shape)
        return sn

    def add_comparison(self, c, title=None):
        """Comparison matrix in the paper's layout (Fig 3b/3d, 4b, 5b)."""
        m = c.base
        sn = sheet_name(title or c.name, self.used)
        ws = self.wb.add_worksheet(sn)
        self._header(ws, "comparison", c.name, " + ".join(m.context), m.col_tag, m.name, c.var)
        r0 = HDR_ROW
        ws.write(r0, 0, "var=" + c.var, self.f_hdr)
        ws.write(r0, 1, c.var, self.f_hdr)
        for j, col in enumerate(m.cols):
            ws.write(r0, 2 + 2 * j, col, self.f_hdr)
            ws.write(r0, 3 + 2 * j, "& Var", self.f_hdr)
        ws.write(r0 + 1, 0, c.var, self.f_lbl)
        ws.write_number(r0 + 1, 1, c.var_total, self.f_num)
        for j, col in enumerate(m.cols):
            ws.write_number(r0 + 1, 2 + 2 * j, int(c.col_var[col]), self.f_num)
            ws.write_blank(r0 + 1, 3 + 2 * j, None, self.f_num)
        for i, r in enumerate(m.rows):
            ws.write(r0 + 2 + i, 0, r, self.f_lbl)
            ws.write_number(r0 + 2 + i, 1, int(c.row_var[r]), self.f_num)
            for j, col in enumerate(m.cols):
                ws.write_number(r0 + 2 + i, 2 + 2 * j, float(m.nop.at[r, col]), self.f_num)
                ws.write_number(r0 + 2 + i, 3 + 2 * j, int(c.triple.at[r, col]), self.f_num)
        self._scale(ws, r0 + 1, 1, r0 + 1 + len(m.rows), 1 + 2 * len(m.cols))
        ws.set_column(0, 0, max(14, min(45, max([len(str(r)) for r in m.rows] + [10]))))
        ws.set_column(1, 1 + 2 * len(m.cols), 10)
        ws.set_row(r0, 45)
        ws.freeze_panes(r0 + 2, 2)
        self._log(sn, "comparison", c.name, m.name, c.var, "%d x %d" % c.triple.shape)
        return sn

    def add_table(self, df, title, kind="table", parent="", note="", score_cols=(), nop_cols=()):
        sn = sheet_name(title, self.used)
        ws = self.wb.add_worksheet(sn)
        self._header(ws, kind, title, "", "", parent, note)
        for j, col in enumerate(df.columns):
            ws.write(HDR_ROW, j, col, self.f_hdr)
        for i, row in enumerate(df.itertuples(index=False)):
            for j, v in enumerate(row):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    ws.write_blank(HDR_ROW + 1 + i, j, None, self.f_num)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(HDR_ROW + 1 + i, j, float(v), self.f_num)
                else:
                    ws.write(HDR_ROW + 1 + i, j, str(v), self.f_num)
        n = len(df)
        cols = list(df.columns)
        for col in score_cols:
            j = cols.index(col)
            self._scale(ws, HDR_ROW + 1, j, HDR_ROW + n, j, 0, 2)
        for col in nop_cols:
            j = cols.index(col)
            self._scale(ws, HDR_ROW + 1, j, HDR_ROW + n, j)
        for j, col in enumerate(cols):
            width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
            ws.set_column(j, j, max(6, min(60, width + 2)))
        ws.freeze_panes(HDR_ROW + 1, 0)
        if n:
            ws.autofilter(HDR_ROW, 0, HDR_ROW + n, len(cols) - 1)
        self._log(sn, kind, title, parent, note, "%d rows" % n)
        return sn

    def add_nrf(self, df, title, parent="", note=""):
        return self.add_table(df, title, kind="nrf", parent=parent, note=note,
                              score_cols=[c for c in ("N", "LR", "TVR", "THR", "VR", "HR") if c in df.columns],
                              nop_cols=[c for c in ("NOP_LR", "NOP_THR", "NOP_TVR") if c in df.columns])

    def close(self):
        ws = self.tree_ws
        ws.write(0, 0, "ALMA tree manager", self.f_title)
        ws.write(1, 0, "PubMed: " + (self.pubmed_note or "current"), self.f_grey)
        ws.write(2, 0, "Thresholds: high >= %s, medium >= %s publications" % (self.high, self.med), self.f_grey)
        hdr = ["Sheet", "Type", "Title", "Parent", "Added string / note", "Size"]
        for j, h in enumerate(hdr):
            ws.write(4, j, h, self.f_hdr)
        for i, rec in enumerate(self.tree):
            for j, v in enumerate(rec):
                if j == 0:
                    ws.write_url(5 + i, 0, "internal:'%s'!A1" % v, string=v)
                else:
                    ws.write(5 + i, j, v)
        ws.set_column(0, 0, 32)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 40)
        ws.set_column(3, 3, 30)
        ws.set_column(4, 4, 30)
        self.wb.close()
