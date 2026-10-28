#!/usr/bin/env python3
"""Did ALMA's 2020 'novel and reasonable' hypotheses get published afterwards?

Takes every 3-element hypothesis that had no publication in the Feb-2020
snapshot (Fig 3 HNC x radiotherapy / nanoparticle, Fig 3d 50 KIs x 7 cancers x
radiotherapy, Fig 5 biomaterials x diseases x targets), looks up its count in
today's PubMed (same query) and relates publication since 2020 to the
reasonability ALMA assigned in 2020.

  .venv/bin/python experiments/retrospective.py
"""
import glob
import os
import sys

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
from alma import core  # noqa: E402

SNAP = os.path.join(ROOT, "results", "snapshot_2020-02-10", "csv")
NOW = sorted(glob.glob(os.path.join(ROOT, "results", "current_*", "csv")))[-1]
TABLES = {"Fig 3 HNC + radiotherapy": "fig3c_nrf_radiotherapy_methods_default",
          "Fig 3 HNC + nanoparticle": "fig3c_nrf_nanoparticle_methods_default",
          "Fig 3d KIs x 7 cancers + radiotherapy": "fig3d_nrf_methods_default",
          "Fig 5 targeted biomaterials": "fig5c_nrf_all_targets"}
HIGH, MED = 20, 5


def main():
    frames = []
    for exp, name in TABLES.items():
        a = pd.read_csv(os.path.join(SNAP, name + ".csv"))
        b = pd.read_csv(os.path.join(NOW, name + ".csv"))[["hypothesis", "NOP"]].rename(columns={"NOP": "NOP_now"})
        d = a.merge(b, on="hypothesis", how="inner")
        d["experiment"] = exp
        frames.append(d)
    d = pd.concat(frames, ignore_index=True).drop_duplicates("hypothesis")
    for k in ("LR", "THR", "TVR"):
        d["s" + k] = [core.reason_score(x, HIGH, MED) for x in d["NOP_" + k]]
    d["reasonability"] = d.sLR + d.sTHR + d.sTVR
    d["novel_reasonable_2020"] = (d.NOP <= 1) & (d[["NOP_LR", "NOP_THR", "NOP_TVR"]].min(axis=1) >= 10)
    z = d[d.NOP == 0].copy()
    z["published_since"] = z.NOP_now >= 1
    by_r = z.groupby("reasonability").agg(n=("hypothesis", "size"), published=("published_since", "sum"),
                                          rate=("published_since", "mean"), median_now=("NOP_now", "median"))
    by_r["rate"] = (100 * by_r.rate).round(1)
    nr = z.groupby("novel_reasonable_2020").agg(n=("hypothesis", "size"), published=("published_since", "sum"),
                                                rate=("published_since", "mean"))
    nr["rate"] = (100 * nr.rate).round(1)
    # rank-based: within each experiment, ALMA's top 10 unpublished vs the other unpublished hypotheses
    z["rank_in_exp"] = z.groupby("experiment")["rank"].rank(method="first")
    z["top10"] = z.rank_in_exp <= 10
    top = z.groupby(["experiment", "top10"]).agg(n=("hypothesis", "size"), published=("published_since", "sum"),
                                                 rate=("published_since", "mean")).reset_index()
    top["rate"] = (100 * top.rate).round(1)
    # Spearman-free trend: share published vs reasonability (logistic-free summary)
    corr = np.corrcoef(z.reasonability, z.published_since.astype(float))[0, 1]
    out = os.path.join(ROOT, "results", "retrospective_2020_to_now")
    os.makedirs(out, exist_ok=True)
    z.sort_values(["experiment", "rank"]).to_csv(os.path.join(out, "unpublished_2020_hypotheses.csv"), index=False)
    by_r.to_csv(os.path.join(out, "by_reasonability.csv"))
    top.to_csv(os.path.join(out, "top10_vs_rest.csv"), index=False)
    lines = ["# Retrospective check: were ALMA's 2020 novel hypotheses published later?", "",
             "Hypotheses with **no publication in PubMed on 2020-02-10** (Entrez date), from the paper's Fig 3 "
             "and Fig 5 designs, re-counted in today's PubMed with the same query.  Reasonability = LR + THR + TVR "
             "scores in 2020 (each 0-2, thresholds 20 / 5).", "",
             "## Published since 2020, by 2020 reasonability", "",
             by_r.reset_index().to_markdown(index=False), "",
             "Correlation (reasonability vs published since): r = %.2f, n = %d" % (corr, len(z)), "",
             "## The paper's filter (<= 1 publication, every pair >= 10) applied in 2020", "",
             nr.reset_index().to_markdown(index=False), "",
             "## ALMA's top-10 ranked unpublished hypotheses vs the rest, per experiment", "",
             top.to_markdown(index=False), ""]
    with open(os.path.join(out, "README.md"), "w") as fh:
        fh.write("\n".join(lines))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
