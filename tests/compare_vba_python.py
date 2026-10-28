"""Compare sheets produced by the VBA macros (dumped by lo_vba_harness.py) with
the Python engine's results for the same PubMed snapshot (2020/02/10)."""
import csv
import os
import sys

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from alma import core  # noqa: E402

HDR = 8  # 0-based header row of matrices / tables


def sheet(path):
    with open(path, newline="") as fh:
        return list(csv.reader(fh))


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return x


def matrix_from_sheet(rows):
    cols = [c for c in rows[HDR][1:] if c != ""]
    data = []
    for r in rows[HDR + 1:]:
        if not r or r[0] == "":
            break
        data.append([r[0]] + [float(v) for v in r[1:1 + len(cols)]])
    return pd.DataFrame([d[1:] for d in data], index=[d[0] for d in data], columns=cols)


def table_from_sheet(rows):
    cols = [c for c in rows[HDR] if c != ""]
    data = [r[:len(cols)] for r in rows[HDR + 1:] if r and r[0] != ""]
    df = pd.DataFrame(data, columns=cols)
    return df.apply(lambda s: s.map(num))


def same_matrix(a, b, tol=1e-9):
    if list(a.index) != list(b.index) or list(a.columns) != list(b.columns):
        return False, "labels/order differ"
    d = np.abs(a.values.astype(float) - b.values.astype(float)).max() if a.size else 0
    return d <= tol, "max abs diff %.3g" % d


def same_table(v, p, cols):
    if len(v) != len(p):
        return False, "rows %d vs %d" % (len(v), len(p))
    for c in cols:
        a, b = list(v[c]), list(p[c])
        for x, y in zip(a, b):
            if isinstance(x, float) or isinstance(y, float):
                if abs(float(x) - float(y)) > 1e-9:
                    return False, "column %s differs: %r vs %r" % (c, x, y)
            elif str(x) != str(y):
                return False, "column %s differs: %r vs %r" % (c, x, y)
    return True, "%d rows identical" % len(v)


def main(vba_dir, py_dir):
    V = lambda n: sheet(os.path.join(vba_dir, n + ".csv"))  # noqa: E731
    P = lambda n, **kw: pd.read_csv(os.path.join(py_dir, n + ".csv"), **kw)  # noqa: E731
    checks = []
    nrf_cols = ["rank", "hypothesis", "row", "col", "var", "NOP", "NOP_LR", "NOP_THR", "NOP_TVR", "NOP_HR", "NOP_VR",
                "N", "LR", "TVR", "THR", "VR", "HR"]
    checks.append(("Fig 5 search matrix", *same_matrix(matrix_from_sheet(V("SM")),
                                                       P("Fig5b_diseases_x_biomaterials", index_col=0))))
    checks.append(("Fig 5 NRF ranking (8 targets)", *same_table(table_from_sheet(V("SM_NRF")),
                                                                P("fig5c_nrf_all_targets"), nrf_cols)))
    vn, pn = table_from_sheet(V("SM_novel")), P("fig5c_novel_reasonable")
    checks.append(("Fig 5 novel & reasonable list", *same_table(vn, pn, ["hypothesis", "NOP", "NOP_LR", "NOP_THR",
                                                                          "NOP_TVR"])))
    tg = table_from_sheet(V("TextGen"))
    checks.append(("Fig 1b text generator", *same_table(tg, P("fig1b_text_generator"), ["hypothesis", "first",
                                                                                         "second", "NOP"])))
    checks.append(("Fig 6 filter matrix (years)", *same_matrix(matrix_from_sheet(V("FM")),
                                                               P("Fig6a_hypotheses_x_years", index_col=0))))
    p6 = P("Fig6a_hypotheses_x_years", index_col=0)
    p6.columns = [str(c) for c in p6.columns]
    checks.append(("Fig 6 normalized per hypothesis", *same_matrix(
        matrix_from_sheet(V("FM_norm_rows")), core.normalize(core.Matrix("x", p6), "rows").nop)))
    checks.append(("Fig 2c search matrix", *same_matrix(matrix_from_sheet(V("SM~2")),
                                                        P("Fig2c_KIs_x_8cancers", index_col=0))))
    checks.append(("Fig 2c sorted by sum", *same_matrix(matrix_from_sheet(V("SM~2_sorted")),
                                                        P("Fig2c_KIs_x_8cancers_sorted", index_col=0))))
    checks.append(("Fig 2c normalized", *same_matrix(matrix_from_sheet(V("SM~2_norm_rows")),
                                                     P("Fig2c_KIs_x_8cancers_norm_rows", index_col=0))))
    checks.append(("Fig 2c cluster sort", *same_matrix(matrix_from_sheet(V("SM~2_norm_rows_clusters")),
                                                       P("Fig2c_KIs_x_8cancers_norm_rows_clusters", index_col=0))))
    checks.append(("Fig 3d NRF ranking", *same_table(table_from_sheet(V("SM~3_NRF")),
                                                     P("fig3d_nrf_methods_default"), nrf_cols)))
    st = sheet(os.path.join(vba_dir, "SelfTest.csv"))
    fails = [r[0] for r in st[2:] if len(r) > 3 and r[3] == "FAIL"]
    checks.append(("self-test sheet", not fails, "failed: %s" % fails if fails else "all passed (online test skipped)"))
    ok = True
    for name, good, msg in checks:
        print("%-4s %-34s %s" % ("OK" if good else "FAIL", name, msg))
        ok &= bool(good)
    print("ALL IDENTICAL" if ok else "DIFFERENCES FOUND")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
