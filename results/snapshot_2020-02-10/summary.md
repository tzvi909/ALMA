# ALMA reproduction – snapshot_2020-02-10

PubMed: maxdate 2020/02/10 (edat)

## Agreement with the numbers printed in the paper
| figure | n | pearson r (log10 NOP+1) | median ratio | within 15% | exact zero/non-zero agreement |
|---|---|---|---|---|---|
| Fig1b | 23 | 0.911 | 2.25 | 4% | 100% |
| Fig2b | 102 | 0.996 | 1.025 | 93% | 100% |
| Fig3a | 25 | 0.996 | 1.052 | 80% | 100% |
| Fig3b | 54 | 0.999 | 1.023 | 83% | 98% |
| Fig4b | 68 | 0.999 | 1.0 | 91% | 100% |
| Fig4c | 20 | 0.994 | 1.0 | 95% | 95% |
| Fig5b | 78 | 0.987 | 1.045 | 68% | 99% |

## Fig 1b – hypothesis text generator
`treating <cancer> with <kinase> inhibitor combined with radiotherapy` (6 x 6 = 36 hypotheses)

| hypothesis | NOP |
|---|---|
| treating Lung cancer with EGFR inhibitor combined with radiotherapy | 302 |
| treating Head and neck cancer with EGFR inhibitor combined with radiotherapy | 295 |
| treating Breast cancer with HER2 inhibitor combined with radiotherapy | 119 |
| treating Breast cancer with EGFR inhibitor combined with radiotherapy | 114 |
| treating Lung cancer with PI3K inhibitor combined with radiotherapy | 32 |
| treating Head and neck cancer with HER2 inhibitor combined with radiotherapy | 29 |
| treating Lung cancer with HER2 inhibitor combined with radiotherapy | 28 |
| treating Head and neck cancer with PI3K inhibitor combined with radiotherapy | 25 |

## Fig 2b – drugs x cancers
| drug | Breast cancer | Glioma | Osteosarcoma | Prostate cancer | Cholangiocarcinoma | Liposarcoma |
|---|---|---|---|---|---|---|
| Doxorubicin | 11939 | 704 | 1966 | 1232 | 92 | 183 |
| 5-Fu | 8247 | 288 | 96 | 575 | 357 | 9 |
| Paclitaxel | 7393 | 416 | 106 | 949 | 32 | 14 |
| Cisplatin | 3800 | 1130 | 1438 | 1067 | 323 | 43 |
| Rapamycin | 1642 | 672 | 154 | 534 | 49 | 15 |
| Vincristine | 1395 | 1113 | 666 | 244 | 14 | 21 |

## Fig 2c-d – 50 KIs x 8 cancers, normalized + cluster sorted
Most studied cancer per KI agrees with the paper's Fig 2d for **17/17** KIs.

| kinase inhibitor | paper (Fig 2d) | reproduction | match |
|---|---|---|---|
| Dasatinib | Leukemia | Leukemia | True |
| Bosutinib | Leukemia | Leukemia | True |
| Imatinib | Leukemia | Leukemia | True |
| Gefitinib | Lung cancer | Lung cancer | True |
| Erlotinib | Lung cancer | Lung cancer | True |
| Crizotinib | Lung cancer | Lung cancer | True |
| Palbociclib | Breast cancer | Breast cancer | True |
| Neratinib | Breast cancer | Breast cancer | True |
| Lapatinib | Breast cancer | Breast cancer | True |
| Trametinib | Melanoma | Melanoma | True |
| Dabrafenib | Melanoma | Melanoma | True |
| Encorafenib | Melanoma | Melanoma | True |
| Cabozantinib | Renal cell carcinoma | Renal cell carcinoma | True |
| Everolimus | Renal cell carcinoma | Renal cell carcinoma | True |
| Sunitinib | Renal cell carcinoma | Renal cell carcinoma | True |
| Sorafenib | Hepatocellular carcinoma | Hepatocellular carcinoma | True |
| Regorafenib | Hepatocellular carcinoma | Hepatocellular carcinoma | True |

## Fig 3a-c – kinase inhibitors for head and neck cancer + radiotherapy / nanoparticle

**radiotherapy, thresholds high>=10 / med>=2** – top 5: Regorafenib, Palbociclib, Nintedanib, Ponatinib, Alectinib (paper: Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib; overlap 3/5)

| rank | row | NOP | NOP_LR | NOP_THR | NOP_TVR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Regorafenib | 0 | 14 | 27 | 59868 | 2 | 2 | 2 | 2 | 2 | 0 |
| 2 | Palbociclib | 1 | 24 | 25 | 59868 | 2 | 2 | 2 | 2 | 2 | 0 |
| 3 | Nintedanib | 1 | 11 | 10 | 59868 | 2 | 2 | 2 | 2 | 2 | 0 |
| 4 | Ponatinib | 0 | 15 | 5 | 59868 | 2 | 2 | 2 | 1 | 2 | 0 |
| 5 | Alectinib | 0 | 9 | 35 | 59868 | 2 | 1 | 2 | 2 | 2 | 0 |
| 6 | Ibrutinib | 0 | 6 | 22 | 59868 | 2 | 1 | 2 | 2 | 2 | 0 |
| 7 | Ceritinib | 0 | 5 | 21 | 59868 | 2 | 1 | 2 | 2 | 2 | 0 |
| 8 | Brigatinib | 0 | 3 | 10 | 59868 | 2 | 1 | 2 | 2 | 2 | 0 |

**nanoparticle, thresholds high>=10 / med>=2** – top 5: Sirolimus, Imatinib, Dasatinib, Everolimus, Temsirolimus (paper: Sirolimus, Imatinib, Everolimus, Dasatinib, Vandetanib; overlap 4/5)

| rank | row | NOP | NOP_LR | NOP_THR | NOP_TVR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Sirolimus | 0 | 284 | 168 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 2 | Imatinib | 0 | 177 | 79 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 3 | Dasatinib | 0 | 86 | 34 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 4 | Everolimus | 0 | 101 | 31 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 5 | Temsirolimus | 0 | 48 | 15 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 6 | Axitinib | 0 | 41 | 12 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 7 | Vemurafenib | 0 | 80 | 10 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |
| 8 | Nilotinib | 0 | 14 | 10 | 814 | 2 | 2 | 2 | 2 | 2 | 0 |

**radiotherapy, thresholds high>=20 / med>=5** – top 5: Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib (paper: Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib; overlap 5/5)

**nanoparticle, thresholds high>=20 / med>=5** – top 5: Sirolimus, Imatinib, Dasatinib, Everolimus, Sorafenib (paper: Sirolimus, Imatinib, Everolimus, Dasatinib, Vandetanib; overlap 4/5)

## Fig 3d – 50 KIs x 7 cancers + radiotherapy

**thresholds high>=10 / med>=2** – overlap with the paper's top 10: 2/10

| rank | row | col | NOP | NOP_LR | NOP_THR | NOP_TVR | NOP_VR | NOP_HR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Imatinib | Hepatocellular carcinoma | 0 | 53 | 312 | 3200 | 287 | 98 | 2 | 2 | 2 | 2 | 2 | 2 |
| 2 | Trametinib | Breast cancer | 0 | 47 | 37 | 32780 | 332 | 35 | 2 | 2 | 2 | 2 | 2 | 2 |
| 3 | Lapatinib | Pancreatic cancer | 0 | 35 | 158 | 5336 | 177 | 150 | 2 | 2 | 2 | 2 | 2 | 2 |
| 4 | Lenvatinib | Pancreatic cancer | 0 | 29 | 56 | 5336 | 177 | 56 | 2 | 2 | 2 | 2 | 2 | 2 |
| 5 | Palbociclib | Melanoma | 0 | 27 | 25 | 8193 | 243 | 18 | 2 | 2 | 2 | 2 | 2 | 2 |
| 6 | Lenvatinib | Melanoma | 0 | 20 | 56 | 8193 | 243 | 56 | 2 | 2 | 2 | 2 | 2 | 2 |
| 7 | Vandetanib | Hepatocellular carcinoma | 0 | 17 | 52 | 3200 | 287 | 28 | 2 | 2 | 2 | 2 | 2 | 2 |
| 8 | Lapatinib | Hepatocellular carcinoma | 0 | 15 | 158 | 3200 | 287 | 150 | 2 | 2 | 2 | 2 | 2 | 2 |
| 9 | Regorafenib | Head and neck cancer | 0 | 14 | 27 | 59868 | 567 | 18 | 2 | 2 | 2 | 2 | 2 | 2 |
| 10 | Palbociclib | Pancreatic cancer | 0 | 14 | 25 | 5336 | 177 | 18 | 2 | 2 | 2 | 2 | 2 | 2 |

**thresholds high>=20 / med>=5** – overlap with the paper's top 10: 5/10

| rank | row | col | NOP | NOP_LR | NOP_THR | NOP_TVR | NOP_VR | NOP_HR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Imatinib | Hepatocellular carcinoma | 0 | 53 | 312 | 3200 | 287 | 98 | 2 | 2 | 2 | 2 | 2 | 2 |
| 2 | Trametinib | Breast cancer | 0 | 47 | 37 | 32780 | 332 | 35 | 2 | 2 | 2 | 2 | 2 | 2 |
| 3 | Lapatinib | Pancreatic cancer | 0 | 35 | 158 | 5336 | 177 | 150 | 2 | 2 | 2 | 2 | 2 | 2 |
| 4 | Lenvatinib | Pancreatic cancer | 0 | 29 | 56 | 5336 | 177 | 56 | 2 | 2 | 2 | 2 | 2 | 2 |
| 5 | Lenvatinib | Melanoma | 0 | 20 | 56 | 8193 | 243 | 56 | 2 | 2 | 2 | 2 | 2 | 2 |
| 6 | Dasatinib | Melanoma | 1 | 56 | 58 | 8193 | 242 | 27 | 2 | 2 | 2 | 2 | 2 | 2 |
| 7 | Vandetanib | Breast cancer | 1 | 62 | 52 | 32780 | 331 | 27 | 2 | 2 | 2 | 2 | 2 | 2 |
| 8 | Dasatinib | Pancreatic cancer | 1 | 51 | 58 | 5336 | 176 | 27 | 2 | 2 | 2 | 2 | 2 | 2 |
| 9 | Cabozantinib | Lung cancer | 1 | 101 | 34 | 35162 | 1079 | 23 | 2 | 2 | 2 | 2 | 2 | 2 |
| 10 | Cabozantinib | Breast cancer | 1 | 35 | 34 | 32780 | 331 | 23 | 2 | 2 | 2 | 2 | 2 | 2 |

## Fig 4 – lab inventory (drugs x cell lines, + osteosarcoma / HNSCC / nanoparticle)
Drugs with <=5 publications with both MG63 and Fadu (the paper's screening criterion): Sorafenib, Nintedanib, Carfilzomib, Nilotinib, Cabozantinib, Trametinib, Palbociclib, Ponatinib, Duvelisib, Infigratinib

| drug | Nanoparticle | Osteosarcoma | MG63 | MG63 + NP | HNSCC | Fadu | Fadu + NP | unexplored (<=5 with MG63 and Fadu) |
|---|---|---|---|---|---|---|---|---|
| Rapamycin | 247 | 154 | 45 | 0 | 182 | 15 | 0 | False |
| Paclitaxel | 2486 | 106 | 15 | 4 | 523 | 37 | 2 | False |
| Gemcitabine | 420 | 108 | 11 | 1 | 95 | 3 | 0 | False |
| Sorafenib | 161 | 34 | 4 | 0 | 34 | 4 | 0 | True |
| Nintedanib | 6 | 2 | 0 | 0 | 2 | 2 | 0 | True |
| Carfilzomib | 10 | 1 | 0 | 0 | 3 | 0 | 0 | True |
| Nilotinib | 10 | 1 | 1 | 0 | 7 | 0 | 0 | True |
| Cabozantinib | 4 | 3 | 0 | 0 | 0 | 0 | 0 | True |
| Trametinib | 5 | 2 | 0 | 0 | 9 | 0 | 0 | True |
| Palbociclib | 2 | 6 | 0 | 0 | 13 | 0 | 0 | True |
| Ponatinib | 1 | 1 | 0 | 0 | 3 | 0 | 0 | True |
| Duvelisib | 0 | 0 | 0 | 0 | 0 | 0 | 0 | True |
| Infigratinib | 0 | 1 | 0 | 0 | 1 | 0 | 0 | True |

_The wet-lab follow-up (MTT screen, carfilzomib nanoparticles) cannot be reproduced in silico._

## Fig 5 – molecularly targeted biomaterials (9 diseases x 4 biomaterials x 8 targets)
107 hypotheses pass the paper's filter (<=1 publication, every pair >=10): 86 unpublished, 21 with one publication.

- paper's 'unpublished' hypotheses also found by the reproduction: 9/9
- paper's 'timely' hypotheses also found by the reproduction: 4/5
- paper's 'untimely' hypotheses also found by the reproduction: 1/3

| rank | hypothesis_text | NOP | pub_year | NOP_LR | NOP_THR | NOP_TVR | class |
|---|---|---|---|---|---|---|---|
| 1 | VCAM-1 targeted Nanoparticle for Pancreatic cancer | 0 | nan | 962 | 85 | 201 | novel - unpublished |
| 2 | VCAM-1 targeted Liposome for Rheumatoid arthritis | 1 | 2000.0 | 234 | 385 | 128 | novel - 1 publication < 2015 (untimely) |
| 3 | ICAM-1 targeted Nanoparticle for Osteoarthritis | 1 | 2016.0 | 210 | 119 | 185 | novel - 1 publication >= 2015 (timely) |
| 4 | ICAM-1 targeted Liposome for Osteoarthritis | 1 | 2003.0 | 82 | 119 | 157 | novel - 1 publication < 2015 (untimely) |
| 5 | VCAM-1 targeted Nanoparticle for Asthma | 0 | nan | 297 | 400 | 201 | novel - unpublished |
| 6 | ICAM-1 targeted Nanoparticle for Hepatitis c | 0 | nan | 167 | 105 | 185 | novel - unpublished |
| 7 | ICAM-1 targeted Liposome for Hepatitis c | 0 | nan | 123 | 105 | 157 | novel - unpublished |
| 8 | ICAM-1 targeted Nanoparticle for Pancreatic cancer | 0 | nan | 962 | 102 | 185 | novel - unpublished |
| 9 | VCAM-1 targeted Nanoparticle for Osteoarthritis | 0 | nan | 210 | 96 | 201 | novel - unpublished |
| 10 | VCAM-1 targeted Liposome for Osteoarthritis | 0 | nan | 82 | 96 | 128 | novel - unpublished |
| 11 | VCAM-1 targeted Nanoparticle for Hepatitis c | 0 | nan | 167 | 54 | 201 | novel - unpublished |
| 12 | VCAM-1 targeted Liposome for Hepatitis c | 0 | nan | 123 | 54 | 128 | novel - unpublished |
| 13 | CAV1 targeted Nanoparticle for Glaucoma | 0 | nan | 166 | 45 | 35 | novel - unpublished |
| 14 | CAV1 targeted Nanoparticle for Myocardial infarction | 0 | nan | 506 | 32 | 35 | novel - unpublished |
| 15 | CAV1 targeted Nanoparticle for Asthma | 0 | nan | 297 | 23 | 35 | novel - unpublished |
| 16 | ICAM-1 targeted Nanoparticle for Glaucoma | 0 | nan | 166 | 17 | 185 | novel - unpublished |
| 17 | ICAM-1 targeted Liposome for Glaucoma | 0 | nan | 78 | 17 | 157 | novel - unpublished |
| 18 | VCAM-1 targeted Nanoparticle for Glaucoma | 0 | nan | 166 | 12 | 201 | novel - unpublished |
| 19 | VCAM-1 targeted Liposome for Glaucoma | 0 | nan | 78 | 12 | 128 | novel - unpublished |
| 20 | VCAM-1 targeted Liposome for Asthma | 1 | 2016.0 | 123 | 400 | 128 | novel - 1 publication >= 2015 (timely) |
| 21 | ICAM-1 targeted Liposome for Pancreatic cancer | 1 | 2016.0 | 341 | 102 | 157 | novel - 1 publication >= 2015 (timely) |
| 22 | ICAM-1 targeted Radiolabeled Antibody for Breast cancer | 0 | nan | 505 | 297 | 83 | novel - unpublished |
| 23 | E-selectin targeted Liposome for Rheumatoid arthritis | 0 | nan | 234 | 191 | 71 | novel - unpublished |
| 24 | VCAM-1 targeted Radiolabeled Antibody for Pancreatic cancer | 0 | nan | 147 | 85 | 70 | novel - unpublished |
| 25 | E-selectin targeted Nanoparticle for Rheumatoid arthritis | 0 | nan | 328 | 191 | 61 | novel - unpublished |

## Fig 6 – temporal and geographic analysis (demo with hypotheses named in the paper)

| hypothesis | reproduction trend | paper (to 2019) |
|---|---|---|
| osimertinib + lung cancer | trending | trending |
| cabozantinib + hepatocellular carcinoma | trending | trending |
| olaparib + lung cancer | trending | trending |
| olaparib + ovarian cancer | trending | stable |
| olaparib + breast cancer | trending | stable |
| pembrolizumab + lung cancer | trending | trending |
| nivolumab + melanoma | trending | trending |
| pazopanib + hepatocellular carcinoma | declining | declining |
| everolimus + pancreatic cancer | declining | declining |
| sorafenib + hepatocellular carcinoma | trending |  |

| hypothesis | top country | most over-represented country (>=5 papers) | paper |
|---|---|---|---|
| HIPEC + ovarian cancer | USA | Spain | Italy/France |
| oncolytic virus + ovarian cancer | USA | Canada | USA |
| high intensity focused ultrasound + glioma | Netherlands | nan | Netherlands |
| immunotherapy + esophageal cancer | Japan | Japan | Japan |
| radiotherapy + gastrointestinal stromal tumor | USA | Switzerland | Germany |
| radiotherapy + breast cancer | USA | Sweden |  |
| chemotherapy + lung cancer | USA | Japan |  |

