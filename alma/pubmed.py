"""PubMed E-utilities client: number of publications (NOP) per query.

ALMA (Harif et al.) searched PubMed's API in real time and used only the hit
count of every generated hypothesis.  Queries go through PubMed's normal search
(automatic term mapping / MeSH explosion), exactly like typing them in PubMed.

Counts are cached in SQLite keyed by (term, datetype, maxdate) so repeated runs
are free and the same cache can be shipped inside the Excel workbook.
"""
import datetime
import os
import sqlite3
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import requests

ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
ESUMMARY = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
MIN_DATE = "1000/01/01"


class RateLimiter:
    """Spaces request start times to stay under NCBI's per-second limit."""

    def __init__(self, per_second):
        self.interval = 1.0 / per_second
        self.lock = threading.Lock()
        self.next_t = 0.0

    def wait(self):
        with self.lock:
            now = time.monotonic()
            start = max(now, self.next_t)
            self.next_t = start + self.interval
        delay = start - time.monotonic()
        if delay > 0:
            time.sleep(delay)


class PubMed:
    """NOP lookups with caching.

    max_date/date_type restrict the search to records up to a date, e.g.
    max_date="2020/02/10", date_type="edat" reproduces the state of PubMed when
    the paper's searches were run (Entrez date = date the record was added).
    """

    def __init__(self, cache_path, max_date=None, date_type="edat", api_key=None,
                 email=None, tool="ALMA-py", workers=3, offline=False, verbose=True):
        self.max_date = max_date or ""
        self.date_type = date_type if self.max_date else ""
        self.api_key = api_key if api_key is not None else os.environ.get("NCBI_API_KEY", "")
        self.email = email if email is not None else os.environ.get("NCBI_EMAIL", "")
        self.tool = tool
        self.offline = offline
        self.verbose = verbose
        self.workers = workers
        self.limiter = RateLimiter(9.0 if self.api_key else 2.8)
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "ALMA-reproduction/1.0 (python-requests)"
        os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
        self.db = sqlite3.connect(cache_path)
        self.db.execute("CREATE TABLE IF NOT EXISTS counts (term TEXT, datetype TEXT, maxdate TEXT,"
                        " nop INTEGER, retrieved TEXT, PRIMARY KEY (term, datetype, maxdate))")
        self.db.commit()
        self.n_fetched = 0

    # ------------------------------------------------------------------ cache
    def _cached(self, term):
        row = self.db.execute("SELECT nop FROM counts WHERE term=? AND datetype=? AND maxdate=?",
                              (term, self.date_type, self.max_date)).fetchone()
        return None if row is None else row[0]

    def _store(self, results):
        now = datetime.datetime.now().isoformat(timespec="seconds")
        self.db.executemany("INSERT OR REPLACE INTO counts VALUES (?,?,?,?,?)",
                            [(t, self.date_type, self.max_date, n, now) for t, n in results.items()])
        self.db.commit()

    def cache_rows(self):
        """All cached rows (any date setting): (term, datetype, maxdate, nop, retrieved)."""
        return self.db.execute("SELECT term, datetype, maxdate, nop, retrieved FROM counts"
                               " ORDER BY maxdate, term").fetchall()

    # ---------------------------------------------------------------- network
    def _params(self, term, **extra):
        p = {"db": "pubmed", "term": term, "retmode": "json", "tool": self.tool}
        if self.max_date:
            p.update(datetype=self.date_type, mindate=MIN_DATE, maxdate=self.max_date)
        if self.api_key:
            p["api_key"] = self.api_key
        if self.email:
            p["email"] = self.email
        p.update(extra)
        return p

    def _get(self, url, params, tries=8):
        err = None
        for attempt in range(tries):
            self.limiter.wait()
            try:
                r = self.session.get(url, params=params, timeout=60)
                if r.status_code == 200:
                    data = r.json()
                    if "error" not in data:
                        return data
                    err = data["error"]
                else:
                    err = "HTTP %d" % r.status_code
            except (requests.RequestException, ValueError) as e:
                err = repr(e)
            time.sleep(min(30, 1.5 * 2 ** attempt))
        raise RuntimeError("PubMed request failed after %d tries: %s (%s)" % (tries, err, params.get("term")))

    def _fetch_count(self, term, tries=6):
        # NCBI sometimes answers 200 with an error payload and no count
        # (e.g. "Search Backend failed"); retry those.
        err = None
        for attempt in range(tries):
            res = self._get(ESEARCH, self._params(term, rettype="count")).get("esearchresult", {})
            if "count" in res:
                return int(res["count"])
            err = res.get("ERROR") or res
            time.sleep(min(30, 2 * 2 ** attempt))
        raise RuntimeError("PubMed returned no count for %r: %s" % (term, err))

    # -------------------------------------------------------------------- API
    def count(self, term):
        return self.counts([term])[term]

    def counts(self, terms):
        """Return {term: NOP}; only uncached terms hit the network."""
        out, todo = {}, []
        for t in dict.fromkeys(terms):
            n = self._cached(t)
            if n is None:
                todo.append(t)
            else:
                out[t] = n
        if todo and self.offline:
            raise RuntimeError("offline mode: %d queries are not cached, e.g. %r" % (len(todo), todo[0]))
        if todo:
            if self.verbose:
                print("  PubMed: %d new queries (%d cached)%s" % (
                    len(todo), len(out), " [maxdate %s %s]" % (self.date_type, self.max_date) if self.max_date else ""),
                    flush=True)
            batch = {}
            with ThreadPoolExecutor(self.workers) as pool:
                for i, (t, n) in enumerate(zip(todo, pool.map(self._fetch_count, todo)), 1):
                    batch[t] = n
                    if len(batch) >= 50:
                        self._store(batch)
                        batch = {}
                    if self.verbose and i % 200 == 0:
                        print("    %d/%d" % (i, len(todo)), flush=True)
            self._store(batch)
            self.n_fetched += len(todo)
            for t in todo:
                out[t] = self._cached(t)
        return out

    def pub_years(self, term, retmax=20):
        """Publication years of the (up to retmax) records matching term, oldest first."""
        ids = self._get(ESEARCH, self._params(term, retmax=retmax))["esearchresult"]["idlist"]
        if not ids:
            return []
        summ = self._get(ESUMMARY, {"db": "pubmed", "id": ",".join(ids), "retmode": "json",
                                    "tool": self.tool, **({"api_key": self.api_key} if self.api_key else {})})
        years = []
        for pmid in ids:
            rec = summ["result"].get(pmid, {})
            y = (rec.get("pubdate") or rec.get("epubdate") or "")[:4]
            if y.isdigit():
                years.append((int(y), pmid))
        return sorted(years)
