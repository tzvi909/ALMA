#!/usr/bin/env python3
"""Re-run the ALMA analyses of Harif et al. on PubMed.

  .venv/bin/python experiments/reproduce_paper.py --max-date 2020/02/10   # PubMed as when the paper's searches ran
  .venv/bin/python experiments/reproduce_paper.py                         # PubMed today -> new novel hypotheses

Outputs (results/<label>/): ALMA_results.xlsx (all matrices, comparison
matrices and NRF rankings in ALMA's sheet layout), csv/, figs/, summary.md and
validation_vs_paper.csv (2020 snapshot only).
"""
import argparse
import datetime
import os
import sys

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

from alma import core  # noqa: E402
from alma.excel import AlmaWorkbook  # noqa: E402
from alma.plots import heatmap, lines  # noqa: E402
from alma.pubmed import PubMed  # noqa: E402
import paper_values as PV  # noqa: E402

THRESHOLDS = {"paper_fig": (10, 2), "methods_default": (20, 5)}  # (high, med)


def L(name):
    with open(os.path.join(ROOT, "lists", name + ".txt"), encoding="utf-8") as fh:
        return [x.strip() for x in fh if x.strip()]


def md_table(df, cols=None, n=None):
    d = df[cols] if cols else df
    if n:
        d = d.head(n)
    out = ["| " + " | ".join(map(str, d.columns)) + " |", "|" + "---|" * len(d.columns)]
    for row in d.itertuples(index=False):
        out.append("| " + " | ".join(str(v) for v in row) + " |")
    return "\n".join(out)


class Run:
    def __init__(self, out, pm, label, snapshot_year):
        self.out, self.pm, self.label, self.year = out, pm, label, snapshot_year
        for d in ("csv", "figs"):
            os.makedirs(os.path.join(out, d), exist_ok=True)
        note = "maxdate %s (%s)" % (pm.max_date, pm.date_type) if pm.max_date else \
            "no date limit (searched %s)" % datetime.date.today().isoformat()
        self.wb = AlmaWorkbook(os.path.join(out, "ALMA_results.xlsx"), 20, 5, note)
        self.md = ["# ALMA reproduction – %s" % label, "", "PubMed: " + note, ""]
        self.val = []

    def csv(self, df, name, index=False):
        df.to_csv(os.path.join(self.out, "csv", name + ".csv"), index=index)

    def png(self, df, name, title, vmax=20, **kw):
        heatmap(df, os.path.join(self.out, "figs", name + ".png"), title, vmax=vmax, **kw)

    def check(self, fig, item, paper, ours):
        self.val.append({"figure": fig, "item": item, "paper": paper, "reproduction": int(ours)})

    def matrix(self, m, png_title=None, vmax=20):
        self.wb.add_matrix(m)
        self.csv(m.nop, m.name, index=True)
        if png_title:
            self.png(m.nop, m.name, png_title, vmax=vmax)


def ranked_nrf(comps, high, med):
    return core.rank(pd.concat([core.nrf(c, high, med) for c in comps], ignore_index=True))


# ---------------------------------------------------------------------- figures
def fig1b(R):
    df = core.text_generator(R.pm, "treating", L("textgen_cancers_fig1b"), "with", L("textgen_kinases_fig1b"),
                             "inhibitor combined with radiotherapy")
    R.wb.add_table(df, "Fig1b_text_generator", kind="text", nop_cols=["NOP"])
    R.csv(df, "fig1b_text_generator")
    for (c, k), v in PV.FIG1B.items():
        R.check("Fig1b", "treating %s with %s inhibitor combined with radiotherapy" % (c, k), v,
                df[(df["first"] == c) & (df["second"] == k)].NOP.iloc[0])
    R.md += ["## Fig 1b – hypothesis text generator",
             "`treating <cancer> with <kinase> inhibitor combined with radiotherapy` (6 x 6 = 36 hypotheses)", "",
             md_table(df, ["hypothesis", "NOP"], 8), ""]


def fig2b(R):
    m = core.search_matrix(R.pm, L("drugs_fig2b_17"), L("cancers_fig2b_6"), "Fig2b_drugs_x_cancers",
                           row_title="drug", col_title="cancer")
    R.matrix(m, "Fig 2b drugs x cancers (NOP)")
    for d, vals in PV.FIG2B.items():
        for c, v in zip(PV.FIG2B_COLS, vals):
            R.check("Fig2b", "%s + %s" % (d, c), v, m.nop.at[d, c])
    R.md += ["## Fig 2b – drugs x cancers", md_table(m.nop.reset_index().rename(columns={"index": "drug"}), n=6), ""]


def fig2cd(R):
    m = core.search_matrix(R.pm, L("kinase_inhibitors_50"), L("cancers_fig2c_8"), "Fig2c_KIs_x_8cancers",
                           row_title="kinase inhibitor", col_title="cancer")
    R.matrix(m)
    sm = core.sort_matrix(m, "sum")
    R.matrix(sm, "Fig 2c 50 KIs x 8 cancers, sorted by sum")
    nm = core.normalize(m, "rows")
    R.matrix(nm)
    cm, clusters = core.cluster_sort(nm, 0.9)
    R.matrix(cm, "Fig 2c normalized (per KI) and cluster sorted, threshold 0.9", vmax=1)
    top = m.nop.idxmax(axis=1)
    agree = [(k, v, top[k]) for k, v in PV.FIG2D_TOP.items()]
    hits = sum(a == b for _, a, b in agree)
    t = pd.DataFrame(agree, columns=["kinase inhibitor", "paper (Fig 2d)", "reproduction"])
    t["match"] = t["paper (Fig 2d)"] == t["reproduction"]
    R.wb.add_table(t, "Fig2d_top_indication_check", parent=nm.name)
    R.csv(t, "fig2d_top_indication_check")
    R.csv(pd.DataFrame({"cluster (top cancer)": clusters}), "fig2c_clusters", index=True)
    R.md += ["## Fig 2c-d – 50 KIs x 8 cancers, normalized + cluster sorted",
             "Most studied cancer per KI agrees with the paper's Fig 2d for **%d/%d** KIs." % (hits, len(agree)), "",
             md_table(t), ""]
    return m


def fig3(R):
    kis = L("kinase_inhibitors_50")
    m7 = core.search_matrix(R.pm, kis, L("cancers_fig3d_7"), "Fig3d_KIs_x_7cancers",
                            row_title="kinase inhibitor", col_title="cancer")
    R.matrix(m7)
    hnc = m7.derive("Fig3a_KIs_x_HNC", m7.nop[["Head and neck cancer"]].sort_values(
        "Head and neck cancer", ascending=False, kind="mergesort"), parent="")
    R.matrix(hnc)
    for k, v in PV.FIG3A.items():
        R.check("Fig3a", "%s + Head and neck cancer" % k, v, hnc.nop.at[k, "Head and neck cancer"])
    comps = {}
    for var in ("radiotherapy", "nanoparticle"):
        c = core.comparison(R.pm, hnc, var, name="Fig3b_HNC_cmp_" + var)
        comps[var] = c
        R.wb.add_comparison(c)
        exp = PV.FIG3B[var]
        R.check("Fig3b", var, exp["_var"], c.var_total)
        R.check("Fig3b", "Head and neck cancer + " + var, exp["_col_var"], c.col_var["Head and neck cancer"])
        for k, (kv, kh, tri) in ((k, v) for k, v in exp.items() if not k.startswith("_")):
            R.check("Fig3b", "%s + %s" % (k, var), kv, c.row_var[k])
            R.check("Fig3b", "%s + Head and neck cancer + %s" % (k, var), tri, c.triple.at[k, "Head and neck cancer"])
    R.md += ["## Fig 3a-c – kinase inhibitors for head and neck cancer + radiotherapy / nanoparticle", ""]
    for tname, (high, med) in THRESHOLDS.items():
        for var, c in comps.items():
            r = ranked_nrf([c], high, med)
            R.wb.add_nrf(r, "Fig3c_NRF_%s_%s" % (var[:5], tname), parent=c.name, note="high>=%s med>=%s" % (high, med))
            R.csv(r, "fig3c_nrf_%s_%s" % (var, tname))
            top5 = list(r.row.head(5))
            ov = len(set(top5) & set(PV.FIG3C[var]))
            R.md += ["**%s, thresholds high>=%s / med>=%s** – top 5: %s (paper: %s; overlap %d/5)" % (
                var, high, med, ", ".join(top5), ", ".join(PV.FIG3C[var]), ov), ""]
            if tname == "paper_fig":
                R.md += [md_table(r, ["rank", "row", "NOP", "NOP_LR", "NOP_THR", "NOP_TVR", "N", "LR", "TVR", "THR",
                                      "VR", "HR"], 8), ""]
    # Fig 3d: 50 KIs x 7 cancers x radiotherapy
    c7 = core.comparison(R.pm, m7, "radiotherapy", name="Fig3d_cmp_radiotherapy")
    R.wb.add_comparison(c7)
    R.png(c7.triple, "fig3d_triples", "Fig 3d KI + cancer + radiotherapy (NOP)")
    R.md += ["## Fig 3d – 50 KIs x 7 cancers + radiotherapy", ""]
    paper = set(PV.FIG3D)
    for tname, (high, med) in THRESHOLDS.items():
        r = ranked_nrf([c7], high, med)
        R.wb.add_nrf(r, "Fig3d_NRF_" + tname, parent=c7.name, note="high>=%s med>=%s" % (high, med))
        R.csv(r, "fig3d_nrf_" + tname)
        top10 = list(zip(r.row.head(10), r.col.head(10)))
        R.md += ["**thresholds high>=%s / med>=%s** – overlap with the paper's top 10: %d/10" % (
            high, med, len(paper & set(top10))), "",
            md_table(r, ["rank", "row", "col", "NOP", "NOP_LR", "NOP_THR", "NOP_TVR", "NOP_VR", "NOP_HR",
                         "N", "LR", "TVR", "THR", "VR", "HR"], 10), ""]
    return comps, c7


def fig4(R):
    drugs = L("lab_drugs_fig4")
    m = core.search_matrix(R.pm, drugs, L("lab_cells_fig4"), "Fig4b_drugs_x_cells", row_title="drug",
                           col_title="cell line")
    R.matrix(m)
    c = core.comparison(R.pm, m, "nanoparticle", name="Fig4b_cmp_nanoparticle")
    R.wb.add_comparison(c)
    ca = core.search_matrix(R.pm, drugs, ["Osteosarcoma", "HNSCC"], "Fig4c_drugs_x_cancers", row_title="drug",
                            col_title="cancer")
    R.matrix(ca)
    t = pd.DataFrame({"drug": drugs,
                      "Nanoparticle": [c.row_var[d] for d in drugs],
                      "Osteosarcoma": [ca.nop.at[d, "Osteosarcoma"] for d in drugs],
                      "MG63": [m.nop.at[d, "MG63"] for d in drugs],
                      "MG63 + NP": [c.triple.at[d, "MG63"] for d in drugs],
                      "HNSCC": [ca.nop.at[d, "HNSCC"] for d in drugs],
                      "Fadu": [m.nop.at[d, "Fadu"] for d in drugs],
                      "Fadu + NP": [c.triple.at[d, "Fadu"] for d in drugs]})
    t["unexplored (<=5 with MG63 and Fadu)"] = (t.MG63 <= 5) & (t.Fadu <= 5)
    t["screened in paper"] = t.drug.isin(PV.FIG4_SCREENED)
    R.wb.add_table(t, "Fig4c_inventory_screen", parent=m.name, nop_cols=["Nanoparticle", "Osteosarcoma", "HNSCC"])
    R.csv(t, "fig4c_inventory_screen")
    for d, vals in PV.FIG4B.items():
        np_, mg, mgv, u2, u2v, fa, fav = vals
        R.check("Fig4b", d + " + nanoparticle", np_, c.row_var[d])
        R.check("Fig4b", d + " + MG63", mg, m.nop.at[d, "MG63"])
        R.check("Fig4b", d + " + MG63 + nanoparticle", mgv, c.triple.at[d, "MG63"])
        R.check("Fig4b", d + " + U2OS", u2, m.nop.at[d, "U2OS"])
        R.check("Fig4b", d + " + Fadu", fa, m.nop.at[d, "Fadu"])
    for cell, v in PV.FIG4B_CELL_NP.items():
        R.check("Fig4b", cell + " + nanoparticle", v, c.col_var[cell])
    for d, v in PV.FIG4C_OS.items():
        R.check("Fig4c", d + " + Osteosarcoma", v, ca.nop.at[d, "Osteosarcoma"])
    for d, v in PV.FIG4C_HNSCC.items():
        R.check("Fig4c", d + " + HNSCC", v, ca.nop.at[d, "HNSCC"])
    sel = t[t["unexplored (<=5 with MG63 and Fadu)"]].drug.tolist()
    R.md += ["## Fig 4 – lab inventory (drugs x cell lines, + osteosarcoma / HNSCC / nanoparticle)",
             "Drugs with <=5 publications with both MG63 and Fadu (the paper's screening criterion): " +
             ", ".join(sel), "", md_table(t.drop(columns=["screened in paper"])), "",
             "_The wet-lab follow-up (MTT screen, carfilzomib nanoparticles) cannot be reproduced in silico._", ""]


def fig5(R):
    m = core.search_matrix(R.pm, L("diseases_fig5"), L("biomaterials_fig5"), "Fig5b_diseases_x_biomaterials",
                           row_title="disease", col_title="biomaterial")
    R.matrix(m, "Fig 5b diseases x biomaterials (NOP)")
    for d, vals in PV.FIG5B.items():
        for b, v in zip(PV.FIG5B_COLS, vals):
            R.check("Fig5b", "%s + %s" % (d, b), v, m.nop.at[d, b])
    comps = []
    for tgt in L("targets_fig5"):
        c = core.comparison(R.pm, m, tgt, name="Fig5b_cmp_" + tgt)
        comps.append(c)
        R.wb.add_comparison(c)
        if tgt in PV.FIG5B_TARGETS:
            tot, by_b, by_d = PV.FIG5B_TARGETS[tgt]
            R.check("Fig5b", tgt, tot, c.var_total)
            for b, v in zip(PV.FIG5B_COLS, by_b):
                R.check("Fig5b", "%s + %s" % (b, tgt), v, c.col_var[b])
            for d, v in zip(m.rows, by_d):
                R.check("Fig5b", "%s + %s" % (d, tgt), v, c.row_var[d])
    high, med = THRESHOLDS["paper_fig"]
    r = ranked_nrf(comps, high, med)
    R.wb.add_nrf(r, "Fig5c_NRF_all_targets", parent=m.name, note="high>=%s med>=%s" % (high, med))
    R.csv(r, "fig5c_nrf_all_targets")
    nr = core.novel_and_reasonable(r, max_nop=1, min_pair=10).copy()
    years, cut = [], R.year - 5
    for row in nr.itertuples():
        if row.NOP == 1:
            ys = R.pm.pub_years(core.and_query(row.row, row.col, row.var))
            years.append(ys[0][0] if ys else None)
        else:
            years.append(None)
    nr["pub_year"] = years
    nr["class"] = ["novel - unpublished" if n == 0 else
                   ("novel - 1 publication >= %d (timely)" % cut if (y or 0) >= cut else
                    "novel - 1 publication < %d (untimely)" % cut) for n, y in zip(nr.NOP, nr.pub_year)]
    nr["hypothesis_text"] = ["%s targeted %s for %s" % (v, c, d) for v, c, d in zip(nr["var"], nr.col, nr.row)]
    R.wb.add_table(nr, "Fig5c_novel_reasonable", parent="Fig5c_NRF_all_targets",
                   note="N<=1 publication and every pair >=10")
    R.csv(nr, "fig5c_novel_reasonable")
    trip = set(zip(nr["var"], nr.col, nr.row))
    papers = {"unpublished": PV.FIG5C_UNPUBLISHED, "timely": PV.FIG5C_TIMELY, "untimely": PV.FIG5C_UNTIMELY}
    lines_ = ["## Fig 5 – molecularly targeted biomaterials (9 diseases x 4 biomaterials x 8 targets)",
              "%d hypotheses pass the paper's filter (<=1 publication, every pair >=10): %d unpublished, %d with one "
              "publication." % (len(nr), (nr.NOP == 0).sum(), (nr.NOP == 1).sum()), ""]
    for k, lst in papers.items():
        lines_.append("- paper's '%s' hypotheses also found by the reproduction: %d/%d" % (
            k, len(trip & set(lst)), len(lst)))
    R.md += lines_ + ["", md_table(nr, ["rank", "hypothesis_text", "NOP", "pub_year", "NOP_LR", "NOP_THR", "NOP_TVR",
                                        "class"], 25), ""]
    return comps


def fig6(R):
    last = min(R.year - 1, 2025)
    years = list(range(2013, last + 1))
    hyps = L("temporal_hypotheses_fig6")
    tm = core.filter_matrix(R.pm, hyps, years, "[dp]", "Fig6a_hypotheses_x_years")
    R.matrix(tm)
    tn = core.normalize(tm, "rows")
    ts = core.sort_by_column(tn, str(last), name="Fig6a_norm_sorted_by_%d" % last)
    R.matrix(ts, "Fig 6a hypotheses x year, normalized per hypothesis", vmax=1)
    trend = {}
    for h in hyps:
        s = tm.nop.loc[h].astype(float)
        peak = int(s.idxmax())
        if peak >= last - 1:
            trend[h] = "trending"
        elif s.iloc[-1] < 0.6 * s.max():
            trend[h] = "declining"
        else:
            trend[h] = "stable"
    lines({h: tm.nop.loc[h] for h in hyps}, os.path.join(R.out, "figs", "fig6b_temporal.png"),
          "NOP per year (%d-%d)" % (years[0], last))
    tt = pd.DataFrame({"hypothesis": hyps, "reproduction trend": [trend[h] for h in hyps],
                       "paper (to 2019)": [PV.FIG6_TRENDS.get(h, "") for h in hyps]})
    R.csv(tt, "fig6b_trends")
    gm = core.filter_matrix(R.pm, L("geo_hypotheses_fig6"), L("countries_20"), "[ad]", "Fig6d_hypotheses_x_countries")
    R.matrix(gm)
    gn = core.normalize(gm, "rows")
    R.matrix(gn, "Fig 6d hypotheses x country (normalized per hypothesis)", vmax=1)
    share = gm.nop.div(gm.nop.sum(axis=0).replace(0, np.nan), axis=1)       # hypothesis share within a country
    enrich = share.div(share.mean(axis=1).replace(0, np.nan), axis=0)       # relative to the average country
    enrich = enrich.where(gm.nop >= 5)
    gt = pd.DataFrame({"hypothesis": gm.rows, "top country": gm.nop.idxmax(axis=1).values,
                       "most over-represented country (>=5 papers)": enrich.idxmax(axis=1).values,
                       "paper": [PV.FIG6_GEO.get(h, "") for h in gm.rows]})
    R.csv(gt, "fig6d_geography")
    R.wb.add_table(tt, "Fig6b_trends", parent=tm.name)
    R.wb.add_table(gt, "Fig6d_geography", parent=gm.name)
    R.md += ["## Fig 6 – temporal and geographic analysis (demo with hypotheses named in the paper)", "",
             md_table(tt), "", md_table(gt), ""]


def validation(R):
    v = pd.DataFrame(R.val)
    v["ratio"] = (v.reproduction + 1) / (v.paper + 1)
    R.csv(v, "validation_vs_paper")
    R.wb.add_table(v, "Validation_vs_paper", note="paper numbers transcribed from the figures")
    s = []
    for fig, g in v.groupby("figure"):
        r = np.corrcoef(np.log10(g.paper + 1), np.log10(g.reproduction + 1))[0, 1] if len(g) > 2 else np.nan
        s.append({"figure": fig, "n": len(g), "pearson r (log10 NOP+1)": round(r, 3),
                  "median ratio": round(g.ratio.median(), 3),
                  "within 15%": "%d%%" % round(100 * (abs(g.ratio - 1) <= 0.15).mean()),
                  "exact zero/non-zero agreement": "%d%%" % round(100 * ((g.paper == 0) == (g.reproduction == 0)).mean())})
    s = pd.DataFrame(s)
    R.csv(s, "validation_summary")
    R.md[4:4] = ["## Agreement with the numbers printed in the paper", md_table(s), ""]
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--max-date", default=None)
    ap.add_argument("--date-type", default="edat")
    ap.add_argument("--label", default=None)
    ap.add_argument("--cache", default=os.path.join(ROOT, "results", "pubmed_cache.sqlite"))
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--only", default="", help="comma list of 1b,2b,2c,3,4,5,6")
    a = ap.parse_args()
    today = datetime.date.today()
    label = a.label or ("snapshot_%s" % a.max_date.replace("/", "-") if a.max_date else "current_%s" % today)
    year = int(a.max_date[:4]) if a.max_date else today.year
    pm = PubMed(a.cache, max_date=a.max_date, date_type=a.date_type, offline=a.offline)
    R = Run(os.path.join(ROOT, "results", label), pm, label, year)
    only = set(a.only.split(",")) if a.only else {"1b", "2b", "2c", "3", "4", "5", "6"}
    steps = [("1b", fig1b), ("2b", fig2b), ("2c", fig2cd), ("3", fig3), ("4", fig4), ("5", fig5), ("6", fig6)]
    for key, fn in steps:
        if key in only:
            print("== Fig %s" % key, flush=True)
            fn(R)
    if a.max_date:
        print(validation(R).to_string(index=False))
    R.wb.close()
    with open(os.path.join(R.out, "summary.md"), "w") as fh:
        fh.write("\n".join(R.md) + "\n")
    print("PubMed queries fetched this run: %d" % pm.n_fetched)
    print("wrote", R.out)


if __name__ == "__main__":
    main()
