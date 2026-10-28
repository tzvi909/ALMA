"""Run the VBA macros of excel/ALMA.xlsm inside headless LibreOffice (VBA
compatibility mode) and dump every sheet as CSV, for comparison with the
Python engine.  Runs inside the alma-libreoffice docker image:

  docker run --rm -v <repo>:/work -v <out>:/out alma-libreoffice \
      /usr/bin/python3 /work/tests/lo_vba_harness.py /work/excel/ALMA.xlsm /out

All PubMed counts come from the workbook's Cache sheet (offline mode).
"""
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

SNAPSHOT = "2020/02/10"
URL_CASES = [("(gefitinib) AND (Head and neck cancer)", "edat", SNAPSHOT),
             ("(Annexin A1) AND (Liposome) AND (Pancreatic cancer)", "", ""),
             ("treating Head and neck cancer with EGFR inhibitor combined with radiotherapy", "edat", SNAPSHOT),
             ("(osimertinib) AND (lung cancer) AND (2019[dp])", "edat", SNAPSHOT),
             ("(HIPEC) AND (ovarian cancer) AND (Italy[ad])", "", ""),
             ("(5-Fu) AND (Hepatitis c) AND (Sj\u00f6gren)", "", "")]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def start_office(profile_dir):
    cmd = ["soffice", "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
           "-env:UserInstallation=file://" + profile_dir,
           "--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(90):
        try:
            return proc, resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(1)
    raise RuntimeError("soffice did not start")


class Book:
    def __init__(self, ctx, path):
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self.doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0,
                                                (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        self.sp = self.doc.getScriptProvider()
        self.errors = []

    def sheet(self, name):
        return self.doc.Sheets.getByName(name)

    def set_setting(self, key, value):
        sh = self.sheet("ALMA")
        for r in range(100):
            if sh.getCellByPosition(0, r).String.strip().lower() == key:
                c = sh.getCellByPosition(2, r)
                if isinstance(value, str):
                    c.setString(value)
                else:
                    c.setValue(value)
                return
        raise KeyError(key)

    def setting(self, key):
        sh = self.sheet("ALMA")
        for r in range(100):
            if sh.getCellByPosition(0, r).String.strip().lower() == key:
                return sh.getCellByPosition(2, r).String
        raise KeyError(key)

    def set_list(self, col, items):
        sh = self.sheet("Lists")
        for r in range(1, 400):
            sh.getCellByPosition(col, r).setString("")
        for i, v in enumerate(items):
            sh.getCellByPosition(col, 1 + i).setString(v)

    def library_list(self, name):
        sh = self.sheet("Library")
        for c in range(60):
            if sh.getCellByPosition(c, 0).String == name:
                out, r = [], 1
                while sh.getCellByPosition(c, r).String:
                    out.append(sh.getCellByPosition(c, r).String)
                    r += 1
                return out
        raise KeyError(name)

    def run(self, action, target=""):
        script = self.sp.getScript("vnd.sun.star.script:VBAProject.ALMA_Main.RunAction?language=Basic&location=document")
        t0 = time.time()
        res = script.invoke((action, target), (), ())[0]
        print("  %-10s %-14s %5.1fs  %s" % (action, target, time.time() - t0, res), flush=True)
        if str(res).startswith("ERROR"):
            self.errors.append(res)
        return str(res)

    def dump(self, outdir):
        os.makedirs(outdir, exist_ok=True)
        for name in self.doc.Sheets.ElementNames:
            sh = self.sheet(name)
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            data = sh.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).DataArray
            with open(os.path.join(outdir, name + ".csv"), "w", newline="") as fh:
                csv.writer(fh).writerows(data)

    def color(self, sheet, col, row):
        return self.sheet(sheet).getCellByPosition(col, row).CellBackColor


def quoted(res):
    q = re.findall(r"'([^']+)'", res)
    return q if q else [""]


def main():
    xlsm, out = sys.argv[1], sys.argv[2]
    work = os.path.join(out, "ALMA_under_test.xlsm")   # LibreOffice keeps read-only files read-only
    shutil.copyfile(xlsm, work)
    proc, ctx = start_office(os.path.join(out, "lo_profile"))
    try:
        b = Book(ctx, work)
        libs = b.doc.BasicLibraries
        print("Basic libraries:", list(libs.ElementNames))
        libs.loadLibrary("VBAProject")
        print("VBAProject modules:", list(libs.getByName("VBAProject").ElementNames))
        for k, v in (("quiet", "TRUE"), ("offline", "TRUE"), ("usecache", "TRUE"), ("maxdate", SNAPSHOT),
                     ("datetype", "edat")):
            b.set_setting(k, v)
        print("== self-test")
        b.run("selftest")
        print("== Fig 5 pipeline (thresholds 10 / 2)")
        b.set_setting("high", 10)
        b.set_setting("med", 2)
        b.run("pipeline")
        print("== Fig 1b text generator")
        b.run("textgen")
        print("== Fig 6 filter matrix")
        b.run("filter")
        print("== Fig 2c search, normalize, cluster (thresholds 20 / 5)")
        b.set_setting("high", 20)
        b.set_setting("med", 5)
        b.set_list(0, b.library_list("kinase_inhibitors_50"))
        b.set_list(1, b.library_list("cancers_fig2c_8"))
        sm = quoted(b.run("search"))[0]
        nm = quoted(b.run("normrows", sm))[0]
        b.run("cluster", nm)
        b.run("sortsum", sm)
        print("== Fig 3d pipeline (KIs x 7 cancers + radiotherapy, thresholds 20 / 5)")
        b.set_list(1, b.library_list("cancers_fig3d_7"))
        b.set_list(2, ["radiotherapy"])
        b.run("pipeline")
        b.run("colorize", "SM~3")
        print("last status:", b.setting("laststatus"))
        print("ERRORS: %d" % len(b.errors))
        for e in b.errors:
            print("  " + e)
        print("colour of SM!B10 (should not be white):", b.color("SM", 1, 9))
        urls = []
        url_fn = b.sp.getScript("vnd.sun.star.script:VBAProject.ALMA_PubMed.EsearchUrl?language=Basic&location=document")
        for term, dtype, maxd in URL_CASES:
            urls.append({"term": term, "datetype": dtype, "maxdate": maxd,
                         "url": url_fn.invoke((term, dtype, maxd), (), ())[0]})
        with open(os.path.join(out, "vba_urls.json"), "w") as fh:
            json.dump(urls, fh, indent=1)
        b.dump(os.path.join(out, "sheets"))
        b.doc.storeToURL(uno.systemPathToFileUrl(os.path.join(out, "ALMA_after_macros.xlsx")),
                         (prop("FilterName", "Calc MS Excel 2007 XML"),))
        b.doc.close(True)
    finally:
        proc.terminate()


if __name__ == "__main__":
    main()
