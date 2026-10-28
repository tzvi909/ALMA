"""Command line interface.

  python -m alma matrix --rows lists/kinase_inhibitors_50.txt --cols "Head and neck cancer" \
         --var radiotherapy --var nanoparticle --out results/kis_hnc.xlsx
  python -m alma text --prefix treating --first lists/textgen_cancers_fig1b.txt --buffer with \
         --second lists/textgen_kinases_fig1b.txt --suffix "inhibitor combined with radiotherapy" --out tg.xlsx
  python -m alma filter --hypotheses lists/temporal_hypotheses_fig6.txt --values 2013-2019 --tag "[dp]" --out t.xlsx
  python -m alma cache --out cache.csv

A LIST argument is a .txt file (one item per line), a .csv file (first column),
an Excel file "book.xlsx[:Sheet[:Column]]" (items from row 2 down), or an
inline comma separated list ("Lung cancer,Breast cancer").
"""
import argparse
import os
import sys

import pandas as pd

from . import core
from .excel import AlmaWorkbook
from .pubmed import PubMed

DEFAULT_CACHE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                             "results", "pubmed_cache.sqlite")


def read_list(spec):
    path, sheet, col = spec, None, None
    if ".xlsx" in spec.lower() or ".xlsm" in spec.lower():
        parts = spec.split(":")
        path = parts[0]
        sheet = parts[1] if len(parts) > 1 and parts[1] else 0
        col = parts[2] if len(parts) > 2 else None
        df = pd.read_excel(path, sheet_name=sheet, header=0, dtype=str)
        s = df[col] if col in df.columns else df.iloc[:, (ord(col.upper()) - 65) if col else 0]
        return [x.strip() for x in s.dropna() if str(x).strip()]
    if os.path.isfile(path):
        if path.lower().endswith(".csv"):
            return [x.strip() for x in pd.read_csv(path, dtype=str).iloc[:, 0].dropna() if x.strip()]
        with open(path, encoding="utf-8") as fh:
            return [ln.strip() for ln in fh if ln.strip() and not ln.startswith("#")]
    return [x.strip() for x in spec.split(",") if x.strip()]


def values_list(spec):
    if "-" in spec and all(p.isdigit() for p in spec.split("-")):
        a, b = spec.split("-")
        return [str(y) for y in range(int(a), int(b) + 1)]
    return read_list(spec)


def make_pubmed(a):
    return PubMed(a.cache, max_date=a.max_date, date_type=a.date_type, offline=a.offline)


def pubmed_note(a):
    return "maxdate %s (%s)" % (a.max_date, a.date_type) if a.max_date else "no date limit"


def cmd_matrix(a):
    pm = make_pubmed(a)
    rows, cols = read_list(a.rows), read_list(a.cols)
    m = core.search_matrix(pm, rows, cols, a.name, context=a.context, row_title=a.row_title, col_title=a.col_title)
    wb = AlmaWorkbook(a.out, a.high, a.med, pubmed_note(a))
    wb.add_matrix(m)
    sm = core.sort_matrix(m, "sum")
    wb.add_matrix(sm)
    if a.normalize:
        nm = core.normalize(m, a.normalize)
        wb.add_matrix(nm)
        cm, clusters = core.cluster_sort(nm, a.cluster, a.normalize)
        wb.add_matrix(cm)
        wb.add_table(clusters.rename_axis("item").reset_index(), m.name + "_clusters", parent=cm.name)
    tables = []
    for var in a.var:
        c = core.comparison(pm, m, var)
        wb.add_comparison(c)
        tables.append(core.nrf(c, a.high, a.med))
    if a.csv:
        os.makedirs(a.csv, exist_ok=True)
        m.nop.to_csv(os.path.join(a.csv, m.name + "_matrix.csv"))
    if tables:
        ranked = core.rank(pd.concat(tables, ignore_index=True))
        ranked["novel_reasonable"] = ranked.index.isin(core.novel_and_reasonable(ranked, a.novel_max, a.pair_min).index)
        novel = ranked[ranked.novel_reasonable]
        wb.add_nrf(ranked, m.name + "_NRF", parent=m.name, note="high>=%s med>=%s" % (a.high, a.med))
        wb.add_nrf(novel, m.name + "_novel", parent=m.name + "_NRF",
                   note="<=%g publications and every pair >=%g" % (a.novel_max, a.pair_min))
        if a.csv:
            ranked.to_csv(os.path.join(a.csv, m.name + "_NRF.csv"), index=False)
            novel.to_csv(os.path.join(a.csv, m.name + "_novel.csv"), index=False)
        cols = ["rank", "row", "col", "var", "NOP", "NOP_LR", "NOP_THR", "NOP_TVR", "N", "LR", "TVR", "THR", "VR", "HR"]
        print("%d hypotheses ranked; %d novel & reasonable (<=%g publications, every pair >=%g). Top:"
              % (len(ranked), len(novel), a.novel_max, a.pair_min))
        print(novel[cols].head(a.top).to_string(index=False))
    else:
        print(sm.nop.to_string())
    wb.close()
    print("wrote", a.out)


def cmd_text(a):
    pm = make_pubmed(a)
    df = core.text_generator(pm, a.prefix, read_list(a.first), a.buffer, read_list(a.second), a.suffix, a.mode)
    wb = AlmaWorkbook(a.out, a.high, a.med, pubmed_note(a))
    wb.add_table(df, "text_generator", kind="text", nop_cols=["NOP"])
    wb.close()
    print(df.head(a.top).to_string(index=False))
    print("wrote", a.out)


def cmd_filter(a):
    pm = make_pubmed(a)
    m = core.filter_matrix(pm, read_list(a.hypotheses), values_list(a.values), a.tag, a.name)
    nm = core.normalize(m, "rows")
    wb = AlmaWorkbook(a.out, a.high, a.med, pubmed_note(a))
    wb.add_matrix(m)
    wb.add_matrix(nm)
    wb.add_matrix(core.sort_by_column(nm, nm.cols[-1]))
    wb.close()
    print(m.nop.to_string())
    print("wrote", a.out)


def cmd_cache(a):
    pm = make_pubmed(a)
    df = pd.DataFrame(pm.cache_rows(), columns=["query", "datetype", "maxdate", "NOP", "retrieved"])
    df.to_csv(a.out, index=False)
    print("wrote %d cached counts to %s" % (len(df), a.out))


def main(argv=None):
    p = argparse.ArgumentParser(prog="alma", description="Automated Literature Meta-Analysis (PubMed)")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
        sp.add_argument("--max-date", default=None, help="only records up to this date, YYYY/MM/DD")
        sp.add_argument("--date-type", default="edat", choices=["edat", "pdat", "mdat"],
                        help="edat = date added to PubMed (default), pdat = publication date")
        sp.add_argument("--cache", default=DEFAULT_CACHE)
        sp.add_argument("--offline", action="store_true", help="use cached counts only")
        sp.add_argument("--high", type=float, default=20, help="high NOP threshold (score 2)")
        sp.add_argument("--med", type=float, default=5, help="medium NOP threshold (score 1)")
        sp.add_argument("--top", type=int, default=20)
        sp.add_argument("--out", required=True, help="output .xlsx")

    sp = sub.add_parser("matrix", help="search matrix, optional comparison matrices + NRF ranking")
    sp.add_argument("--rows", required=True)
    sp.add_argument("--cols", required=True)
    sp.add_argument("--context", action="append", default=[], help="constant string added to every query")
    sp.add_argument("--var", action="append", default=[], help="new string for a comparison matrix (repeatable)")
    sp.add_argument("--name", default="SM")
    sp.add_argument("--row-title", default="rows")
    sp.add_argument("--col-title", default="columns")
    sp.add_argument("--novel-max", type=float, default=1, help="novel: at most this many publications (paper: 1)")
    sp.add_argument("--pair-min", type=float, default=10, help="reasonable: every pair at least this many (paper: 10)")
    sp.add_argument("--csv", default=None, help="also write the matrix, NRF ranking and novel list as CSV here")
    sp.add_argument("--normalize", choices=["rows", "cols"])
    sp.add_argument("--cluster", type=float, default=0.9)
    common(sp)
    sp.set_defaults(func=cmd_matrix)

    sp = sub.add_parser("text", help="hypothesis text generator")
    for k in ("prefix", "buffer", "suffix"):
        sp.add_argument("--" + k, default="")
    sp.add_argument("--first", required=True)
    sp.add_argument("--second", required=True)
    sp.add_argument("--mode", choices=["all", "single"], default="all")
    common(sp)
    sp.set_defaults(func=cmd_text)

    sp = sub.add_parser("filter", help="hypotheses x years/countries (temporal / geographic)")
    sp.add_argument("--hypotheses", required=True, help="list; elements of one hypothesis joined by ' + '")
    sp.add_argument("--values", required=True, help="e.g. 2013-2019 or a list of countries")
    sp.add_argument("--tag", default="[dp]", help="[dp] publication year, [ad] affiliation")
    sp.add_argument("--name", default="FM")
    common(sp)
    sp.set_defaults(func=cmd_filter)

    sp = sub.add_parser("cache", help="export the PubMed count cache as CSV")
    sp.add_argument("--out", required=True)
    sp.add_argument("--cache", default=DEFAULT_CACHE)
    sp.set_defaults(func=cmd_cache, max_date=None, date_type="edat", offline=True)

    a = p.parse_args(argv)
    a.func(a)


if __name__ == "__main__":
    sys.exit(main())
