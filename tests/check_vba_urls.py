"""Fetch the esearch URLs built by the VBA code (EsearchUrl, captured by
lo_vba_harness.py) and check PubMed returns the same counts as the Python client."""
import json
import os
import sys
import time

import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from alma.pubmed import PubMed  # noqa: E402


def main(path, cache):
    ok = True
    for case in json.load(open(path)):
        r = requests.get(case["url"], timeout=60)
        time.sleep(0.4)
        vba = int(r.json()["esearchresult"]["count"])
        pm = PubMed(cache, max_date=case["maxdate"] or None, date_type=case["datetype"] or "edat", verbose=False)
        py = pm.count(case["term"])
        good = vba == py
        ok &= good
        print("%-4s VBA %7d  Python %7d  %s" % ("OK" if good else "DIFF", vba, py, case["term"]))
    print("URLS OK" if ok else "URL MISMATCH")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, "results", "pubmed_cache.sqlite")))
