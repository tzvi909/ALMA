"""A test double for the part of the xlwings API used by excel/ALMA_xlwings.py
(XlwingsIO), backed by openpyxl and following xlwings' documented semantics:
Sheets.add(name, before, after), Sheet.range((r, c)[, (r2, c2)]), Sheet.used_range,
Range.value / options(ndim=2) / color / number_format / font.bold / column_width,
App.status_bar / screen_updating / alert, Book(fullname), Book.caller()."""
import os

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

_BOOKS = {}


class _Font:
    def __init__(self, rng):
        self._rng = rng

    @property
    def bold(self):
        return bool(self._rng._cells()[0][0].font.b)

    @bold.setter
    def bold(self, v):
        for row in self._rng._cells():
            for c in row:
                c.font = Font(bold=bool(v))


class Range:
    def __init__(self, sheet, r1, c1, r2=None, c2=None, ndim=None):
        self.sheet, self.r1, self.c1 = sheet, r1, c1
        self.r2, self.c2 = (r2 or r1), (c2 or c1)
        self._ndim = ndim

    def _cells(self):
        ws = self.sheet._ws
        return [[ws.cell(row=r, column=c) for c in range(self.c1, self.c2 + 1)] for r in range(self.r1, self.r2 + 1)]

    @property
    def row(self):
        return self.r1

    @property
    def column(self):
        return self.c1

    def options(self, convert=None, **options):
        return Range(self.sheet, self.r1, self.c1, self.r2, self.c2, ndim=options.get("ndim"))

    @property
    def value(self):
        vals = [[(float(c.value) if isinstance(c.value, int) and not isinstance(c.value, bool) else c.value)
                 for c in row] for row in self._cells()]
        if self._ndim == 2:
            return vals
        if len(vals) == 1 and len(vals[0]) == 1:
            return vals[0][0]
        if len(vals) == 1:
            return vals[0]
        if all(len(r) == 1 for r in vals):
            return [r[0] for r in vals]
        return vals

    @value.setter
    def value(self, data):
        ws = self.sheet._ws
        if not isinstance(data, (list, tuple)):
            data = [[data]]
        elif data and not isinstance(data[0], (list, tuple)):
            data = [list(data)]
        for i, row in enumerate(data):
            for j, v in enumerate(row):
                ws.cell(row=self.r1 + i, column=self.c1 + j).value = v

    @property
    def color(self):
        return None

    @color.setter
    def color(self, rgb):
        if isinstance(rgb, str):
            hexc = rgb.lstrip("#").upper()
        else:
            hexc = "%02X%02X%02X" % tuple(rgb)
        for row in self._cells():
            for c in row:
                c.fill = PatternFill("solid", fgColor=hexc)

    @property
    def number_format(self):
        return self._cells()[0][0].number_format

    @number_format.setter
    def number_format(self, fmt):
        for row in self._cells():
            for c in row:
                c.number_format = fmt

    @property
    def font(self):
        return _Font(self)

    @property
    def column_width(self):
        return self.sheet._ws.column_dimensions[get_column_letter(self.c1)].width

    @column_width.setter
    def column_width(self, w):
        for c in range(self.c1, self.c2 + 1):
            self.sheet._ws.column_dimensions[get_column_letter(c)].width = w


class Sheet:
    def __init__(self, book, ws):
        self.book, self._ws = book, ws

    @property
    def name(self):
        return self._ws.title

    @name.setter
    def name(self, v):
        self._ws.title = v

    def range(self, cell1, cell2=None):
        r1, c1 = cell1
        if cell2 is None:
            return Range(self, r1, c1)
        r2, c2 = cell2
        return Range(self, r1, c1, r2, c2)

    @property
    def used_range(self):
        ws = self._ws
        cells = [(c.row, c.column) for row in ws.iter_rows() for c in row if c.value is not None]
        if not cells:
            return Range(self, 1, 1)
        return Range(self, min(r for r, _ in cells), min(c for _, c in cells),
                     max(r for r, _ in cells), max(c for _, c in cells))

    def activate(self):
        self.book._wb.active = self.book._wb.sheetnames.index(self.name)


class Sheets:
    def __init__(self, book):
        self.book = book

    def _all(self):
        return [Sheet(self.book, ws) for ws in self.book._wb.worksheets]

    def __iter__(self):
        return iter(self._all())

    def __len__(self):
        return len(self.book._wb.worksheets)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._all()[key]
        return Sheet(self.book, self.book._wb[key])

    @property
    def active(self):
        return Sheet(self.book, self.book._wb.active)

    def add(self, name=None, before=None, after=None):
        wb = self.book._wb
        idx = len(wb.worksheets)
        if after is not None:
            idx = wb.sheetnames.index(after.name) + 1
        elif before is not None:
            idx = wb.sheetnames.index(before.name)
        ws = wb.create_sheet(title=name, index=idx)
        wb.active = idx
        return Sheet(self.book, ws)


class App:
    def __init__(self):
        self.status_bar = None
        self.screen_updating = True
        self.alerts = []

    def alert(self, prompt, title=None, buttons="ok", mode=None, callback=None):
        self.alerts.append((title, prompt))


class Book:
    def __init__(self, fullname=None, **kw):
        path = os.path.abspath(fullname)
        if path not in _BOOKS:
            _BOOKS[path] = (openpyxl.load_workbook(path), App())
        self._wb, self.app = _BOOKS[path]
        self.fullname = path
        self.sheets = Sheets(self)

    _caller = None

    @classmethod
    def caller(cls):
        return cls._caller

    def set_mock_caller(self):
        Book._caller = self

    def save(self, path=None):
        self._wb.save(path or self.fullname)
