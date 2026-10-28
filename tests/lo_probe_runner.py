"""Call ALMA_Probe in every probe workbook (inside the alma-libreoffice image)."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo_vba_harness import Book, start_office  # noqa: E402


def main(probe_dir, out):
    proc, ctx = start_office(os.path.join(out, "lo_profile_probe"))
    results = []
    try:
        for case in json.load(open(os.path.join(probe_dir, "manifest.json"))):
            b = Book(ctx, os.path.join(probe_dir, case["file"]))
            try:
                s = b.sp.getScript("vnd.sun.star.script:VBAProject.%s.ALMA_Probe?language=Basic&location=document"
                                   % case["module"])
                r = s.invoke((), (), ())[0]
            except Exception as e:  # noqa: BLE001
                r = "EXC " + str(e)[:200]
            b.doc.close(True)
            case["result"] = r
            results.append(case)
            if r != "alive":
                print("FAIL %-12s %-22s %r" % (case["module"], case["proc"], r), flush=True)
    finally:
        proc.terminate()
    ok = sum(c["result"] == "alive" for c in results)
    print("%d/%d probes compiled" % (ok, len(results)))
    json.dump(results, open(os.path.join(out, "probe_results.json"), "w"), indent=1)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
