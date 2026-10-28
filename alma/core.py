"""ALMA core: hypothesis generation, search matrices, sorting, normalisation,
cluster sorting, comparison matrices and NRF (novelty / reasonability /
feasibility) triangulation scoring.

Everything here mirrors the Methods of Harif et al. ("Generating and Ranking
Unpublished Biomedical Hypotheses with Automated Literature Meta-Analysis").
The Excel/VBA implementation (excel/vba/*.bas) follows the same rules so the
two give identical numbers; keep them in sync.

Query convention (shared with VBA): each element is wrapped in parentheses and
joined with " AND ", e.g. "(palbociclib) AND (head and neck cancer) AND
(radiotherapy)".  A label may hold several elements separated by " + ".
"""
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

SEP = " + "
PRIORITY = ("N", "LR", "TVR", "THR", "VR", "HR")  # ranking order used throughout the paper


# --------------------------------------------------------------------- queries
def terms_of(label):
    return [t.strip() for t in str(label).split(SEP) if t.strip()]


def and_query(*parts):
    """Build the PubMed query for a combination of elements (None/'' ignored)."""
    terms = []
    for p in parts:
        if p is None:
            continue
        if isinstance(p, (list, tuple)):
            for q in p:
                terms.extend(terms_of(q))
        else:
            terms.extend(terms_of(p))
    return " AND ".join("(%s)" % t for t in terms)


def text_query(prefix, first, buffer, second, suffix):
    """Hypothesis text generator: prefix + first + buffer + second + suffix."""
    return " ".join(s.strip() for s in (prefix, first, buffer, second, suffix) if s and s.strip())


def tagged(label, tag):
    return label + tag if tag else label


# ---------------------------------------------------------------- data classes
@dataclass
class Matrix:
    """A search matrix (NOP) or a matrix derived from one (normalized/sorted)."""
    name: str
    nop: pd.DataFrame                    # index = row labels, columns = column labels
    kind: str = "search"                 # search | normalized | text
    context: list = field(default_factory=list)  # constant strings added to every query
    col_tag: str = ""                    # PubMed field tag appended to column labels, e.g. "[dp]"
    row_title: str = "rows"
    col_title: str = "columns"
    parent: str = ""
    note: str = ""

    @property
    def rows(self):
        return list(self.nop.index)

    @property
    def cols(self):
        return list(self.nop.columns)

    def derive(self, name, nop, **kw):
        d = dict(kind=self.kind, context=list(self.context), col_tag=self.col_tag,
                 row_title=self.row_title, col_title=self.col_title, parent=self.name, note="")
        d.update(kw)
        return Matrix(name, nop, **d)


@dataclass
class Comparison:
    """Comparison matrix: a search matrix layered with the same search + one new string (var)."""
    name: str
    base: Matrix               # LR: row x col
    var: str
    var_total: int             # var (+ context) alone
    row_var: pd.Series         # THR: row x var
    col_var: pd.Series         # TVR: col x var
    triple: pd.DataFrame       # N candidate: row x col x var


# ------------------------------------------------------------ generating/searching
def search_matrix(pm, rows, cols, name, context=(), col_tag="", row_title="rows", col_title="columns"):
    """Many-vs-many (or one-vs-many if a single column) search matrix of NOPs."""
    context = list(context)
    q = {(r, c): and_query(r, tagged(c, col_tag), context) for r in rows for c in cols}
    res = pm.counts(q.values())
    nop = pd.DataFrame([[res[q[r, c]] for c in cols] for r in rows], index=rows, columns=cols)
    return Matrix(name, nop, context=context, col_tag=col_tag, row_title=row_title, col_title=col_title)


def text_generator(pm, prefix, first, buffer, second, suffix, mode="all"):
    """ALMA text generator.  mode='all' = all combinations, 'single' = i-th with i-th.

    Returns a DataFrame (hypothesis, first, second, NOP) sorted by NOP."""
    if mode == "all":
        pairs = [(a, b) for a in first for b in second]
    else:
        pairs = list(zip(first, second))
    hyps = [text_query(prefix, a, buffer, b, suffix) for a, b in pairs]
    res = pm.counts(hyps)
    df = pd.DataFrame({"hypothesis": hyps, "first": [a for a, _ in pairs],
                       "second": [b for _, b in pairs], "NOP": [res[h] for h in hyps]})
    return df.sort_values("NOP", ascending=False, kind="mergesort").reset_index(drop=True)


def comparison(pm, m, var, name=None):
    """Add a new string (var) to every hypothesis of a search matrix."""
    ctx, tag = m.context, m.col_tag
    q_var = and_query(var, ctx)
    q_row = {r: and_query(r, var, ctx) for r in m.rows}
    q_col = {c: and_query(tagged(c, tag), var, ctx) for c in m.cols}
    q_tri = {(r, c): and_query(r, tagged(c, tag), var, ctx) for r in m.rows for c in m.cols}
    res = pm.counts([q_var, *q_row.values(), *q_col.values(), *q_tri.values()])
    tri = pd.DataFrame([[res[q_tri[r, c]] for c in m.cols] for r in m.rows], index=m.rows, columns=m.cols)
    return Comparison(name or "%s_cmp_%s" % (m.name, var), m, var, res[q_var],
                      pd.Series({r: res[q_row[r]] for r in m.rows}),
                      pd.Series({c: res[q_col[c]] for c in m.cols}), tri)


def filter_matrix(pm, hypotheses, values, tag, name, context=()):
    """Temporal / geographic analysis: hypotheses x (years | countries ...).

    e.g. tag='[dp]' with values 2013..2019, or tag='[ad]' with country names."""
    return search_matrix(pm, hypotheses, [str(v) for v in values], name, context=context,
                         col_tag=tag, row_title="hypothesis", col_title=tag)


# --------------------------------------------------------------- matrix operations
def sort_matrix(m, by="sum", axis="both", name=None):
    """Sort rows and/or columns by their sum (or max / mean), highest first.

    Sorting both axes by sum puts the 'center of knowledge' at the top-left."""
    agg = {"sum": np.sum, "max": np.max, "mean": np.mean}[by]
    v = m.nop
    rows, cols = list(v.index), list(v.columns)
    if axis in ("both", "rows"):
        s = v.apply(agg, axis=1).values
        rows = [rows[i] for i in np.argsort(-s, kind="mergesort")]
    if axis in ("both", "cols"):
        s = v.apply(agg, axis=0).values
        cols = [cols[i] for i in np.argsort(-s, kind="mergesort")]
    return m.derive(name or m.name + "_sorted", v.loc[rows, cols])


def sort_by_column(m, col, name=None, ascending=False):
    nop = m.nop.sort_values(col, ascending=ascending, kind="mergesort")
    return m.derive(name or "%s_by_%s" % (m.name, col), nop)


def normalize(m, axis="rows", name=None):
    """Divide every row (horizontal) or column (vertical) by its maximum (nNOP)."""
    v = m.nop.astype(float)
    if axis == "rows":
        mx = v.max(axis=1).replace(0, np.nan)
        out = v.div(mx, axis=0).fillna(0.0)
    else:
        mx = v.max(axis=0).replace(0, np.nan)
        out = v.div(mx, axis=1).fillna(0.0)
    return m.derive(name or "%s_norm_%s" % (m.name, axis), out, kind="normalized",
                    note="normalized by %s maximum" % axis[:-1])


def cluster_sort(m, threshold=0.9, axis="rows", name=None):
    """Cluster-sort a normalized matrix.

    Every row is put in the cluster of a column where nNOP >= threshold.  Columns
    are taken greedily: the column that captures the most still-unassigned rows
    forms the next cluster.  Gives the block-diagonal picture of Fig 2c-f.
    Returns (sorted Matrix, Series row -> cluster column)."""
    v = m.nop if axis == "rows" else m.nop.T
    rows, cols = list(v.index), list(v.columns)
    a = v.values.astype(float)
    unassigned = list(range(len(rows)))
    remaining = list(range(len(cols)))
    order_r, order_c, cluster = [], [], {}
    while unassigned and remaining:
        best, best_members = None, []
        for j in remaining:
            members = [i for i in unassigned if a[i, j] >= threshold]
            if len(members) > len(best_members):
                best, best_members = j, members
        if best is None:
            break
        best_members.sort(key=lambda i: (-a[i, best], i))
        order_c.append(best)
        order_r.extend(best_members)
        for i in best_members:
            cluster[rows[i]] = cols[best]
        unassigned = [i for i in unassigned if i not in best_members]
        remaining.remove(best)
    order_c += remaining
    order_r += unassigned
    out = v.iloc[order_r, order_c]
    if axis != "rows":
        out = out.T
    sm = m.derive(name or m.name + "_clusters", out, note="cluster sorted, threshold %.2f" % threshold)
    return sm, pd.Series(cluster, name="cluster")


# ----------------------------------------------------------------- NRF scoring
def reason_score(n, high, med):
    return 2 if n >= high else (1 if n >= med else 0)


def novelty_score(n, high, med):
    return 0 if n >= high else (1 if n >= med else 2)


def nrf(cmp, high=20, med=5):
    """NRF matrix of a comparison matrix as a long table.

    LR  = row + col             (local reasonability; the original search matrix)
    THR = row + var             (total horizontal reasonability)
    TVR = col + var             (total vertical reasonability)
    HR  = row + var with the *other* columns of the matrix (feasibility; sum of triples)
    VR  = col + var with the *other* rows of the matrix    (feasibility; sum of triples)
    N   = row + col + var       (novelty; inverse score)
    """
    tri, base = cmp.triple, cmp.base.nop
    row_tot, col_tot = tri.sum(axis=1), tri.sum(axis=0)
    recs = []
    for r in tri.index:
        for c in tri.columns:
            t = int(tri.at[r, c])
            hr, vr = int(row_tot[r] - t), int(col_tot[c] - t)
            lr, thr, tvr = int(base.at[r, c]), int(cmp.row_var[r]), int(cmp.col_var[c])
            recs.append({
                "hypothesis": SEP.join([r, c, cmp.var]), "row": r, "col": c, "var": cmp.var,
                "NOP": t, "NOP_LR": lr, "NOP_THR": thr, "NOP_TVR": tvr, "NOP_HR": hr, "NOP_VR": vr,
                "N": novelty_score(t, high, med), "LR": reason_score(lr, high, med),
                "TVR": reason_score(tvr, high, med), "THR": reason_score(thr, high, med),
                "VR": reason_score(vr, high, med), "HR": reason_score(hr, high, med)})
    return pd.DataFrame(recs)


def rank(df, priority=PRIORITY):
    """Rank NRF rows: scores descending in priority order; ties broken by fewer
    publications of the full hypothesis, then by the weakest pair (higher first)."""
    d = df.copy()
    d["_minpair"] = d[["NOP_LR", "NOP_THR", "NOP_TVR"]].min(axis=1)
    d["_order"] = np.arange(len(d))
    keys = list(priority) + ["NOP", "_minpair", "_order"]
    asc = [False] * len(priority) + [True, False, True]
    d = d.sort_values(keys, ascending=asc, kind="mergesort").drop(columns=["_minpair", "_order"])
    d.insert(0, "rank", np.arange(1, len(d) + 1))
    return d.reset_index(drop=True)


def novel_and_reasonable(df, max_nop=1, min_pair=10):
    """Fig 3c / 5c filter: hypothesis has <= max_nop publications while every pair
    (LR, THR, TVR) has >= min_pair publications."""
    keep = (df.NOP <= max_nop) & (df.NOP_LR >= min_pair) & (df.NOP_THR >= min_pair) & (df.NOP_TVR >= min_pair)
    return df[keep]
