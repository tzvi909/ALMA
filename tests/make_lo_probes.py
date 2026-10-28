"""Generate probe workbooks to find VBA code that LibreOffice cannot compile.

For every module M and every procedure P of M, writes a workbook whose VBA
project holds the declaration headers of all modules plus M's header, P and a
Probe function.  tests/lo_probe_runner.py then calls Probe in each workbook;
a probe that does not answer "alive" pinpoints code LibreOffice rejects.
"""
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "excel"))
import vbaproject as vp  # noqa: E402
import xlsxwriter  # noqa: E402

PROC = re.compile(r"^(Public |Private )?(Sub|Function) (\w+)", re.I)
PROBE = 'Public Function ALMA_Probe() As String\r\n    ALMA_Probe = "alive"\r\nEnd Function\r\n'


def split(src):
    lines = src.split("\r\n")
    header, procs, cur = [], [], None
    for ln in lines:
        m = PROC.match(ln)
        if cur is None and not m:
            header.append(ln)
            continue
        if m and cur is None:
            cur = [m.group(3), [ln]]
            continue
        cur[1].append(ln)
        if re.match(r"^End (Sub|Function)", ln):
            procs.append((cur[0], "\r\n".join(cur[1]) + "\r\n"))
            cur = None
    return "\r\n".join(header) + "\r\n", procs


def workbook(path, modules):
    tmpdir = os.path.dirname(path)
    files = []
    for name, src in modules:
        f = os.path.join(tmpdir, "_%s.bas" % name)
        with open(f, "w", encoding="utf-8", newline="") as fh:
            fh.write(src)
        files.append(f)
    vba, _ = vp.build(files, ["Sheet1"])
    wb = xlsxwriter.Workbook(path)
    wb.add_vba_project(__import__("io").BytesIO(vba), is_stream=True)
    wb.set_vba_name("ThisWorkbook")
    ws = wb.add_worksheet("ALMA")
    ws.set_vba_name("Sheet1")
    wb.close()
    for f in files:
        os.remove(f)


def main(outdir):
    os.makedirs(outdir, exist_ok=True)
    mods = {}
    for p in sorted(glob.glob(os.path.join(os.path.dirname(HERE), "excel", "vba", "*.bas"))):
        name, src = vp.read_bas(p)
        mods[name] = split(src)
    manifest = []
    for name, (header, procs) in mods.items():
        others = [(n, h) for n, (h, _) in mods.items() if n != name]
        cases = [("HEADER", header + PROBE)] + [(pn, header + body + PROBE) for pn, body in procs]
        for pn, src in cases:
            path = os.path.join(outdir, "%s__%s.xlsm" % (name, pn))
            workbook(path, others + [(name, src)])
            manifest.append({"file": os.path.basename(path), "module": name, "proc": pn})
    with open(os.path.join(outdir, "manifest.json"), "w") as fh:
        json.dump(manifest, fh, indent=1)
    print(len(manifest), "probe workbooks in", outdir)


if __name__ == "__main__":
    main(sys.argv[1])
