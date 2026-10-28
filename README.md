# ALMA – Automated Literature Meta-Analysis

Independent re-implementation of the ALMA programme of Harif, Azaguri, Amrusi, Harris, Avrashami, Meir,
Bloch, Dobreen, Sason & Shamay, *"Generating and Ranking Unpublished Biomedical Hypotheses with
Automated Literature Meta-Analysis (ALMA)"* (Faculty of Biomedical Engineering, Technion), written from
the manuscript's Methods and validated against the numbers in its figures.  ALMA builds every combination of
categorized terms, counts each combination in PubMed (number of publications, NOP), colours and
sorts the resulting matrices, and ranks unpublished 3-element hypotheses by novelty and
reasonability ("triangulation", NRF scoring).

The same rules are implemented three ways, and all three give identical results:

| version | where | how to run |
|---|---|---|
| Python engine + command line | `alma/` | `python -m alma matrix ...`, `experiments/reproduce_paper.py` |
| Excel / VBA (like the original) | `excel/ALMA.xlsm` | open in Excel, enable macros, click the buttons |
| Excel / Python (xlwings) | `excel/ALMA_xlwings.xlsx` + `excel/ALMA_xlwings.py` | choose an Action, click **Run main** on the xlwings tab |

## 1. Python

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python -m alma matrix --rows lists/kinase_inhibitors_50.txt --cols "Head and neck cancer" \
    --var radiotherapy --var nanoparticle --out my_run.xlsx              # search + comparison matrices + NRF ranking
.venv/bin/python -m alma text --prefix treating --first lists/textgen_cancers_fig1b.txt --buffer with \
    --second lists/textgen_kinases_fig1b.txt --suffix "inhibitor combined with radiotherapy" --out tg.xlsx
.venv/bin/python -m alma filter --hypotheses lists/temporal_hypotheses_fig6.txt --values 2013-2019 --tag "[dp]" --out t.xlsx
```
Lists can be `.txt`, `.csv`, `book.xlsx[:Sheet[:Column]]` or inline `"a,b,c"`.  `--max-date 2020/02/10`
limits PubMed to records added up to that date; `--high/--med` set the thresholds (default 20 / 5);
`--offline` uses cached counts only.  Counts are cached in `results/pubmed_cache.sqlite`.
Set `NCBI_API_KEY` to go from 3 to 10 requests per second.

Re-run the paper's analyses: `.venv/bin/python experiments/reproduce_paper.py --max-date 2020/02/10`
(literature as the paper saw it) or without `--max-date` (today), then `experiments/retrospective.py`.

## 2. Excel with VBA – `excel/ALMA.xlsm`

Sheets: **ALMA** (settings + buttons), **Lists** (your terms; pre-filled with the paper's Fig 5 example),
**Library** (all lists used in the paper), **Tree** (every matrix created), **Cache** (PubMed counts,
pre-filled with the ~3,850 counts of the runs below, so the examples run instantly).
Click **Full pipeline**: search matrix (List 1 × List 2) → sorted copy → one comparison matrix per string in
column C → NRF ranking (`SM_NRF`) → novel & reasonable list (`SM_novel`).  The other buttons (sort, normalize,
cluster sort, colorize, add string, text generator, temporal/geographic filter, NRF of existing comparison
sheets) work on the active ALMA sheet.  **Self-test** checks the macros and the PubMed connection.
Windows Excel uses MSXML2/WinHTTP; Mac Excel uses `curl` through MacScript.
If the file came from the internet: right-click → Properties → Unblock, then enable macros.
If Excel ever refuses the embedded macros, open `ALMA_template.xlsx`, import `excel/vba/*.bas`
(Alt+F11 → File → Import), save as `.xlsm`.  Rebuild everything with `.venv/bin/python excel/build_excel.py`.

## 3. Excel with Python (xlwings) – `excel/ALMA_xlwings.xlsx`

Once, on the computer with Excel: `pip install xlwings requests pandas numpy openpyxl` and
`xlwings addin install`.  Keep the workbook next to `ALMA_xlwings.py` inside this folder (it imports
`alma/`).  Pick the **Action** on the ALMA sheet (Full pipeline, Search matrix, Comparison matrices + NRF,
NRF ranking, Sort/Normalize/Cluster sort, Text generator, Filter matrix, Self-test, ...) and click
**Run main**.  The same Lists/Tree layout as the VBA workbook; counts are cached in `excel/ALMA_cache.sqlite`.
Without Excel the script edits the file directly:
`python excel/ALMA_xlwings.py excel/ALMA_xlwings.xlsx --action "Full pipeline" [--set maxdate=2020/02/10]`.
(Microsoft's built-in `=PY()` Python in Excel has no internet access, so it cannot query PubMed.)

## Scanning your own question – example: brain metastases

ALMA needs categorized lists and a hypothesis structure.  For brain metastases two designs work well
(results of both, run on 28 Oct 2026: `results/brain_metastases_2026-10-28/README.md`):

* **A – drug x primary tumour x brain metastases**: rows = drugs (`lists/brain_mets_drugs.txt`),
  columns = primary tumours (`lists/brain_mets_primaries.txt`), comparison string = brain metastases.
  Finds drugs studied in a tumour and in brain metastases, but never in brain metastases *of that tumour*
  (77 such gaps, e.g. nivolumab or trametinib for colorectal-cancer brain metastases, durvalumab for
  renal-cell-carcinoma brain metastases, palbociclib for lung-cancer brain metastases).
* **B – drug x brain metastases x strategy**: rows = drugs, column = brain metastases, comparison strings =
  `lists/brain_mets_strategies.txt` (radiosurgery, whole-brain RT, nanoparticle, intrathecal, focused
  ultrasound).  37 gaps, e.g. gefitinib-, sorafenib- or pembrolizumab-nanoparticles, temozolomide +
  focused ultrasound.

Write brain metastases as a precise PubMed expression (`lists/brain_mets_term.txt`):
`"brain metastases"[tiab] OR "brain metastasis"[tiab] OR Brain Neoplasms/secondary[mh]`.
The plain words are mapped by PubMed to *all* brain tumours (temozolomide + "brain metastases" = 6,837
hits, mostly glioma; with the expression: 315).  Any PubMed syntax is allowed inside one list cell.

* **Python**: the two commands at the end of `results/brain_metastases_2026-10-28/README.md`
  (`--csv` also writes the ranking and the novel list as CSV).
* **Excel (VBA or xlwings)**: on the Lists sheet paste the drugs into column A and the primary tumours into
  column B, and the brain-metastases expression into C2 (design A) – or the expression into B2 and the
  strategies into column C (design B).  All these lists are on the Library sheet.  Then **Full pipeline**
  (VBA button, or Action + Run main).  Results: `SM_NRF` (ranking) and `SM_novel` (filtered list).
  The workbooks' caches already hold these counts, so the example runs instantly.
* Thresholds (ALMA sheet / `--high --med --novel-max --pair-min`) decide what counts as "known": raise
  them for well-studied fields, lower them for niche ones.  Treat the output as literature gaps to review
  (biology, CNS penetration, the single existing paper), not as validated ideas.

## Method (as in the paper)

* Query of a combination: each element in parentheses joined with AND, sent to PubMed's normal search
  (automatic term mapping / MeSH), e.g. `(palbociclib) AND (head and neck cancer) AND (radiotherapy)`.
* Scores with thresholds high/medium: reasonability 2 if NOP ≥ high, 1 if ≥ medium, else 0; novelty inverse.
  LR = row+column, THR = row+string, TVR = column+string, HR / VR = the string with the *other* columns / rows
  of the matrix (sum of their triples; the paper's description is ambiguous here).  Ranking N, LR, TVR, THR, VR, HR
  (ties: fewer publications, then stronger weakest pair).  "Novel & reasonable" = ≤ 1 publication and every pair ≥ 10.
* Normalization divides a row/column by its maximum; cluster sort groups rows under the column where
  nNOP ≥ 0.9, columns taken greedily by cluster size.

## Results (details: `results/*/summary.md`, `ALMA_results.xlsx`, `csv/`, `figs/`)

**Reproduction with PubMed as of 10 Feb 2020** (Entrez date; the paper's UI screenshot shows that date):

| check | result |
|---|---|
| 347 NOPs printed in Figs 2b, 3a, 3b, 4b, 4c, 5b | log-scale r = 0.987–0.999, median ratio 1.00–1.05 |
| Fig 1b free-text sentences | same order (r = 0.91) but ~2.3× higher: PubMed's 2020 search engine now maps "treating" to the therapy subheading |
| Fig 2d most studied cancer per kinase inhibitor | 17 / 17 identical |
| Fig 3c top-5 KI + radiotherapy for HNC | identical, same order (Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib) with the Methods thresholds 20 / 5 |
| Fig 3c top-5 KI-loaded nanoparticles for HNC | 4 / 5 |
| Fig 3d top-10 KI + radiotherapy in 7 cancers | 5 / 10 |
| Fig 5c novel unpublished targeted biomaterials | 9 / 9 recovered (timely 4 / 5, untimely 1 / 3) |
| Fig 6 temporal trends / geography (demo) | 7 / 9 trends; HIFU-glioma → Netherlands, immunotherapy-esophageal → Japan |

**Did the 2020 predictions come true?** (`results/retrospective_2020_to_now/`) Of 424 hypotheses with no
publication in Feb 2020, ALMA's top-10 per experiment were published by Oct 2026 far more often than the rest:
Fig 3d 9/10 vs 27 %, HNC + radiotherapy 7/10 vs 22 %, Fig 5 3/10 vs 15 %; the paper's filter 33 % vs 18 %.
Palbociclib + HNC + radiotherapy went from 1 to 9 papers, FAP-targeted nanoparticles for pancreatic cancer
0 → 9, Annexin A1-targeted liposomes for pancreatic cancer 0 → 1.

**Novel & reasonable hypotheses in today's PubMed** (`results/current_2026-10-21/`), e.g.
KI + radiotherapy for HNC (top ranked, 0–2 papers): ibrutinib, nintedanib, regorafenib, abemaciclib; KI-loaded nanoparticles for HNC:
sirolimus, imatinib, everolimus, osimertinib (still 0 papers); KI + radiotherapy: lapatinib or palbociclib in
pancreatic cancer, lapatinib in HCC, abemaciclib / ribociclib in melanoma; targeted biomaterials:
VCAM-1 nanoparticles or liposomes for osteoarthritis, E-selectin nanoparticles for rheumatoid arthritis,
VCAM-1 nanoparticles for pancreatic cancer (104 unpublished + 36 single-paper hypotheses pass the filter).

## Assumptions and limits

* Term lists were transcribed from the figures; the paper's 50 FDA-approved kinase inhibitors are the 46 of
  Table 1 plus gilteritinib, erdafitinib, entrectinib and idelalisib.  Fig 2a (136 × 80), Fig 2e (400 genes),
  Table 1 (30 kinases) and the full Fig 6 lists are not given in the paper and were not re-run; Fig 6 is a demo with
  the hypotheses named in the text.  The wet-lab (Fig 4c–h) and Human Protein Atlas (Fig 5d–g) parts are out of scope.
* PubMed changed since 2020 (search engine, indexing), so counts differ by a few percent from the paper.
* The VBA was verified by running every macro in LibreOffice's VBA mode (results identical to Python:
  `tests/lo_vba_harness.py`, `tests/compare_vba_python.py`), by reading the generated `vbaProject.bin` back with oletools, by static checks, and by
  fetching the URLs it builds; it has **not been run in Microsoft Excel** here (no Excel on this server), and the
  Windows HTTP call and the Mac `curl` path are untested.  The xlwings script was checked against the real
  xlwings 0.37 API and run through a test double; it has not been run inside Excel either.

## Layout, tests and contributing

`alma/` engine · `experiments/` reproduction + retrospective · `lists/` term lists · `results/` outputs + cache ·
`excel/` workbooks, `vba/` sources, `vbaproject.py` (vbaProject.bin writer), `build_excel.py`, `ALMA_xlwings.py` ·
`tests/run_tests.sh` runs every check (LibreOffice part needs `docker build -t alma-libreoffice tests/docker`).
After editing `excel/vba/*.bas` or the lists, rebuild the workbooks with `python excel/build_excel.py`
(this writes `ALMA.xlsm`, `ALMA_template.xlsx`, `ALMA_xlwings.xlsx`) and run `tests/run_tests.sh`.
The manuscript PDF is deliberately not part of the repository (`.gitignore`).
