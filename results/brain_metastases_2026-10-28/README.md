# ALMA scan: novel hypotheses for brain metastases (PubMed, 28 Oct 2026)

Brain metastases are queried with a precise PubMed expression, because PubMed maps the plain phrase
"brain metastases" to *all* brain tumours (glioma included; temozolomide + "brain metastases" = 6,837 hits,
with the expression below 315):

    "brain metastases"[tiab] OR "brain metastasis"[tiab] OR Brain Neoplasms/secondary[mh]

Thresholds: high 20 / medium 5 (paper's Methods); "novel & reasonable" = at most 1 publication for the
full hypothesis while every pair of its elements has at least 10 (paper's Fig 3c / 5c filter).
Drugs: the paper's 50 kinase inhibitors + 24 agents relevant to brain metastases (`lists/brain_mets_drugs.txt`).

## Design A – drug x primary tumour x brain metastases (77 novel & reasonable of 444)

Drug studied in the primary tumour and in brain metastases (of other tumours), never in brain metastases
of that tumour.  Rare-to-the-brain primaries (colorectal, renal) dominate, as expected.

| hypothesis                                                     |   papers (all 3) |   drug+tumour |   drug+BM |   tumour+BM |
|:---------------------------------------------------------------|-----------------:|--------------:|----------:|------------:|
| Nivolumab for brain metastases of colorectal cancer            |                0 |           456 |       406 |         803 |
| Trametinib for brain metastases of colorectal cancer           |                0 |           174 |        97 |         803 |
| Durvalumab for brain metastases of renal cell carcinoma        |                0 |            59 |        61 |        1066 |
| Cabozantinib for brain metastases of colorectal cancer         |                0 |            49 |        52 |         803 |
| Everolimus for brain metastases of colorectal cancer           |                0 |            98 |        43 |         803 |
| Pazopanib for brain metastases of colorectal cancer            |                0 |            36 |        49 |         803 |
| Binimetinib for brain metastases of colorectal cancer          |                0 |           107 |        32 |         803 |
| Sirolimus for brain metastases of colorectal cancer            |                0 |           225 |        29 |         803 |
| Axitinib for brain metastases of colorectal cancer             |                0 |            43 |        28 |         803 |
| Palbociclib for brain metastases of non-small cell lung cancer |                0 |            56 |        28 |        5131 |
| Palbociclib for brain metastases of small cell lung cancer     |                0 |            63 |        28 |        6401 |
| Palbociclib for brain metastases of colorectal cancer          |                0 |            66 |        28 |         803 |
| Abemaciclib for brain metastases of colorectal cancer          |                0 |            28 |        34 |         803 |
| Binimetinib for brain metastases of breast cancer              |                0 |            25 |        32 |        5569 |
| Osimertinib for brain metastases of renal cell carcinoma       |                0 |            21 |       356 |        1066 |

## Design B – drug x brain metastases x strategy (37 novel & reasonable of 370)

Drug studied in brain metastases and with the strategy, never the three together
(strategies: stereotactic radiosurgery, whole brain radiotherapy, nanoparticle, intrathecal, focused ultrasound;
the two radiotherapy strategies yield no gaps - they are already well studied).

| hypothesis                                             |   papers (all 3) |   drug+BM |   drug+strategy |   BM+strategy |
|:-------------------------------------------------------|-----------------:|----------:|----------------:|--------------:|
| Gefitinib + nanoparticle for brain metastases          |                0 |       321 |             149 |           177 |
| Sorafenib + nanoparticle for brain metastases          |                0 |        82 |             484 |           177 |
| Pembrolizumab + nanoparticle for brain metastases      |                0 |       322 |              62 |           177 |
| Temozolomide + focused ultrasound for brain metastases |                0 |       315 |              58 |           156 |
| Atezolizumab + nanoparticle for brain metastases       |                0 |       114 |              56 |           177 |
| Nivolumab + nanoparticle for brain metastases          |                0 |       406 |              32 |           177 |
| Sirolimus + nanoparticle for brain metastases          |                0 |        29 |             351 |           177 |
| Sirolimus + focused ultrasound for brain metastases    |                0 |        29 |              29 |           156 |
| Imatinib + nanoparticle for brain metastases           |                0 |        28 |             147 |           177 |
| Imatinib + intrathecal for brain metastases            |                0 |        28 |              30 |           206 |
| Sirolimus + intrathecal for brain metastases           |                0 |        29 |              28 |           206 |
| Sorafenib + focused ultrasound for brain metastases    |                0 |        82 |              26 |           156 |
| Crizotinib + nanoparticle for brain metastases         |                0 |       339 |              24 |           177 |
| Everolimus + focused ultrasound for brain metastases   |                0 |        43 |              20 |           156 |
| Bevacizumab + nanoparticle for brain metastases        |                1 |       435 |             240 |           177 |

Full rankings: `A_drug_x_primary_x_brain_mets.xlsx`, `B_drug_x_brain_mets_x_strategy.xlsx` (sheets `*_NRF`,
`*_novel`, comparison matrices) and `csv/`.  These are literature gaps, not validated ideas: check CNS
penetration, biology and the single existing papers before acting on them.

## Re-run

```bash
BM="$(cat lists/brain_mets_term.txt)"
python -m alma matrix --rows lists/brain_mets_drugs.txt --cols lists/brain_mets_primaries.txt --var "$BM" \
    --name BM_A_drug_x_primary --out A.xlsx --csv csv
python -m alma matrix --rows lists/brain_mets_drugs.txt --cols lists/brain_mets_term.txt \
    --var "stereotactic radiosurgery" --var "whole brain radiotherapy" --var nanoparticle --var intrathecal \
    --var "focused ultrasound" --name BM_B_drug_x_strategy --out B.xlsx --csv csv
```
