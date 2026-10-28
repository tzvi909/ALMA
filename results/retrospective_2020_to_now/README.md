# Retrospective check: were ALMA's 2020 novel hypotheses published later?

Hypotheses with **no publication in PubMed on 2020-02-10** (Entrez date), from the paper's Fig 3 and Fig 5 designs, re-counted in today's PubMed with the same query.  Reasonability = LR + THR + TVR scores in 2020 (each 0-2, thresholds 20 / 5).

## Published since 2020, by 2020 reasonability

|   reasonability |   n |   published |   rate |   median_now |
|----------------:|----:|------------:|-------:|-------------:|
|               1 |   5 |           0 |    0   |            0 |
|               2 |  93 |           9 |    9.7 |            0 |
|               3 |  88 |          21 |   23.9 |            0 |
|               4 | 109 |          19 |   17.4 |            0 |
|               5 |  75 |          25 |   33.3 |            0 |
|               6 |  54 |          20 |   37   |            0 |

Correlation (reasonability vs published since): r = 0.21, n = 424

## The paper's filter (<= 1 publication, every pair >= 10) applied in 2020

| novel_reasonable_2020   |   n |   published |   rate |
|:------------------------|----:|------------:|-------:|
| False                   | 311 |          57 |   18.3 |
| True                    | 113 |          37 |   32.7 |

## ALMA's top-10 ranked unpublished hypotheses vs the rest, per experiment

| experiment                            | top10   |   n |   published |   rate |
|:--------------------------------------|:--------|----:|------------:|-------:|
| Fig 3 HNC + nanoparticle              | False   |  34 |           4 |   11.8 |
| Fig 3 HNC + nanoparticle              | True    |  10 |           2 |   20   |
| Fig 3 HNC + radiotherapy              | False   |   9 |           2 |   22.2 |
| Fig 3 HNC + radiotherapy              | True    |  10 |           7 |   70   |
| Fig 3d KIs x 7 cancers + radiotherapy | False   | 132 |          36 |   27.3 |
| Fig 3d KIs x 7 cancers + radiotherapy | True    |  10 |           9 |   90   |
| Fig 5 targeted biomaterials           | False   | 209 |          31 |   14.8 |
| Fig 5 targeted biomaterials           | True    |  10 |           3 |   30   |
