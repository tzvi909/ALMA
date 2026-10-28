# ALMA reproduction – current_2026-10-21

PubMed: no date limit (searched 2026-10-21)

## Fig 1b – hypothesis text generator
`treating <cancer> with <kinase> inhibitor combined with radiotherapy` (6 x 6 = 36 hypotheses)

| hypothesis | NOP |
|---|---|
| treating Lung cancer with EGFR inhibitor combined with radiotherapy | 547 |
| treating Head and neck cancer with EGFR inhibitor combined with radiotherapy | 358 |
| treating Breast cancer with HER2 inhibitor combined with radiotherapy | 270 |
| treating Breast cancer with EGFR inhibitor combined with radiotherapy | 187 |
| treating Lung cancer with PI3K inhibitor combined with radiotherapy | 53 |
| treating Lung cancer with HER2 inhibitor combined with radiotherapy | 50 |
| treating Head and neck cancer with PI3K inhibitor combined with radiotherapy | 47 |
| treating Head and neck cancer with HER2 inhibitor combined with radiotherapy | 45 |

## Fig 2b – drugs x cancers
| drug | Breast cancer | Glioma | Osteosarcoma | Prostate cancer | Cholangiocarcinoma | Liposarcoma |
|---|---|---|---|---|---|---|
| Doxorubicin | 16874 | 1042 | 2637 | 1531 | 117 | 280 |
| 5-Fu | 9559 | 344 | 119 | 704 | 570 | 9 |
| Paclitaxel | 11025 | 578 | 155 | 1184 | 111 | 25 |
| Cisplatin | 5684 | 1370 | 1943 | 1499 | 917 | 47 |
| Rapamycin | 2465 | 927 | 240 | 721 | 93 | 27 |
| Vincristine | 1535 | 1257 | 747 | 273 | 20 | 22 |

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

**radiotherapy, thresholds high>=10 / med>=2** – top 5: Ibrutinib, Brigatinib, Nintedanib, Encorafenib, Ponatinib (paper: Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib; overlap 1/5)

| rank | row | NOP | NOP_LR | NOP_THR | NOP_TVR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Ibrutinib | 0 | 20 | 67 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 2 | Brigatinib | 0 | 13 | 31 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 3 | Nintedanib | 1 | 22 | 28 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 4 | Encorafenib | 1 | 14 | 15 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 5 | Ponatinib | 1 | 18 | 13 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 6 | Neratinib | 1 | 11 | 41 | 79230 | 2 | 2 | 2 | 2 | 2 | 0 |
| 7 | Lorlatinib | 0 | 9 | 41 | 79230 | 2 | 1 | 2 | 2 | 2 | 0 |
| 8 | Erdafitinib | 0 | 8 | 15 | 79230 | 2 | 1 | 2 | 2 | 2 | 0 |

**nanoparticle, thresholds high>=10 / med>=2** – top 5: Sirolimus, Imatinib, Everolimus, Osimertinib, Regorafenib (paper: Sirolimus, Imatinib, Everolimus, Dasatinib, Vandetanib; overlap 3/5)

| rank | row | NOP | NOP_LR | NOP_THR | NOP_TVR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Sirolimus | 0 | 365 | 351 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 2 | Imatinib | 0 | 231 | 147 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 3 | Everolimus | 0 | 142 | 59 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 4 | Osimertinib | 0 | 81 | 38 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 5 | Regorafenib | 0 | 35 | 42 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 6 | Axitinib | 0 | 68 | 32 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 7 | Nilotinib | 0 | 32 | 24 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |
| 8 | Afatinib | 0 | 156 | 22 | 1999 | 2 | 2 | 2 | 2 | 2 | 0 |

**radiotherapy, thresholds high>=20 / med>=5** – top 5: Ibrutinib, Nintedanib, Regorafenib, Abemaciclib, Ruxolitinib (paper: Palbociclib, Regorafenib, Alectinib, Ibrutinib, Ceritinib; overlap 2/5)

**nanoparticle, thresholds high>=20 / med>=5** – top 5: Sirolimus, Imatinib, Everolimus, Osimertinib, Regorafenib (paper: Sirolimus, Imatinib, Everolimus, Dasatinib, Vandetanib; overlap 3/5)

## Fig 3d – 50 KIs x 7 cancers + radiotherapy

**thresholds high>=10 / med>=2** – overlap with the paper's top 10: 1/10

| rank | row | col | NOP | NOP_LR | NOP_THR | NOP_TVR | NOP_VR | NOP_HR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Lapatinib | Pancreatic cancer | 0 | 62 | 219 | 8372 | 287 | 204 | 2 | 2 | 2 | 2 | 2 | 2 |
| 2 | Palbociclib | Pancreatic cancer | 0 | 57 | 125 | 8372 | 287 | 113 | 2 | 2 | 2 | 2 | 2 | 2 |
| 3 | Ruxolitinib | Breast cancer | 0 | 57 | 45 | 46750 | 683 | 10 | 2 | 2 | 2 | 2 | 2 | 2 |
| 4 | Ruxolitinib | Melanoma | 0 | 61 | 45 | 11209 | 404 | 10 | 2 | 2 | 2 | 2 | 2 | 2 |
| 5 | Lapatinib | Hepatocellular carcinoma | 0 | 39 | 219 | 5803 | 621 | 204 | 2 | 2 | 2 | 2 | 2 | 2 |
| 6 | Abemaciclib | Melanoma | 0 | 30 | 78 | 11209 | 404 | 75 | 2 | 2 | 2 | 2 | 2 | 2 |
| 7 | Palbociclib | Prostate cancer | 0 | 29 | 125 | 33495 | 99 | 113 | 2 | 2 | 2 | 2 | 2 | 2 |
| 8 | Ribociclib | Melanoma | 0 | 28 | 60 | 11209 | 404 | 62 | 2 | 2 | 2 | 2 | 2 | 2 |
| 9 | Nintedanib | Prostate cancer | 0 | 19 | 28 | 33495 | 99 | 19 | 2 | 2 | 2 | 2 | 2 | 2 |
| 10 | Nintedanib | Melanoma | 0 | 19 | 28 | 11209 | 404 | 19 | 2 | 2 | 2 | 2 | 2 | 2 |

**thresholds high>=20 / med>=5** – overlap with the paper's top 10: 3/10

| rank | row | col | NOP | NOP_LR | NOP_THR | NOP_TVR | NOP_VR | NOP_HR | N | LR | TVR | THR | VR | HR |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Lapatinib | Pancreatic cancer | 0 | 62 | 219 | 8372 | 287 | 204 | 2 | 2 | 2 | 2 | 2 | 2 |
| 2 | Palbociclib | Pancreatic cancer | 0 | 57 | 125 | 8372 | 287 | 113 | 2 | 2 | 2 | 2 | 2 | 2 |
| 3 | Lapatinib | Hepatocellular carcinoma | 0 | 39 | 219 | 5803 | 621 | 204 | 2 | 2 | 2 | 2 | 2 | 2 |
| 4 | Abemaciclib | Melanoma | 0 | 30 | 78 | 11209 | 404 | 75 | 2 | 2 | 2 | 2 | 2 | 2 |
| 5 | Palbociclib | Prostate cancer | 0 | 29 | 125 | 33495 | 99 | 113 | 2 | 2 | 2 | 2 | 2 | 2 |
| 6 | Ribociclib | Melanoma | 0 | 28 | 60 | 11209 | 404 | 62 | 2 | 2 | 2 | 2 | 2 | 2 |
| 7 | Dasatinib | Melanoma | 1 | 88 | 94 | 11209 | 403 | 41 | 2 | 2 | 2 | 2 | 2 | 2 |
| 8 | Imatinib | Hepatocellular carcinoma | 1 | 87 | 402 | 5803 | 620 | 125 | 2 | 2 | 2 | 2 | 2 | 2 |
| 9 | Palbociclib | Melanoma | 1 | 68 | 125 | 11209 | 403 | 112 | 2 | 2 | 2 | 2 | 2 | 2 |
| 10 | Lapatinib | Prostate cancer | 1 | 49 | 219 | 33495 | 98 | 203 | 2 | 2 | 2 | 2 | 2 | 2 |

## Fig 4 – lab inventory (drugs x cell lines, + osteosarcoma / HNSCC / nanoparticle)
Drugs with <=5 publications with both MG63 and Fadu (the paper's screening criterion): Nintedanib, Carfilzomib, Nilotinib, Cabozantinib, Trametinib, Palbociclib, Ponatinib, Duvelisib, Infigratinib

| drug | Nanoparticle | Osteosarcoma | MG63 | MG63 + NP | HNSCC | Fadu | Fadu + NP | unexplored (<=5 with MG63 and Fadu) |
|---|---|---|---|---|---|---|---|---|
| Rapamycin | 684 | 240 | 59 | 2 | 252 | 25 | 1 | False |
| Paclitaxel | 4256 | 155 | 25 | 7 | 808 | 46 | 4 | False |
| Gemcitabine | 885 | 156 | 15 | 2 | 126 | 4 | 1 | False |
| Sorafenib | 484 | 80 | 10 | 1 | 45 | 5 | 1 | False |
| Nintedanib | 62 | 7 | 0 | 0 | 5 | 2 | 0 | True |
| Carfilzomib | 30 | 5 | 0 | 0 | 7 | 0 | 0 | True |
| Nilotinib | 24 | 5 | 1 | 0 | 16 | 1 | 0 | True |
| Cabozantinib | 19 | 31 | 1 | 0 | 12 | 0 | 0 | True |
| Trametinib | 19 | 5 | 2 | 1 | 22 | 0 | 0 | True |
| Palbociclib | 38 | 23 | 1 | 0 | 43 | 1 | 0 | True |
| Ponatinib | 10 | 3 | 1 | 1 | 3 | 0 | 0 | True |
| Duvelisib | 3 | 0 | 0 | 0 | 1 | 0 | 0 | True |
| Infigratinib | 1 | 2 | 0 | 0 | 2 | 0 | 0 | True |

_The wet-lab follow-up (MTT screen, carfilzomib nanoparticles) cannot be reproduced in silico._

## Fig 5 – molecularly targeted biomaterials (9 diseases x 4 biomaterials x 8 targets)
140 hypotheses pass the paper's filter (<=1 publication, every pair >=10): 104 unpublished, 36 with one publication.

- paper's 'unpublished' hypotheses also found by the reproduction: 7/9
- paper's 'timely' hypotheses also found by the reproduction: 2/5
- paper's 'untimely' hypotheses also found by the reproduction: 0/3

| rank | hypothesis_text | NOP | pub_year | NOP_LR | NOP_THR | NOP_TVR | class |
|---|---|---|---|---|---|---|---|
| 1 | VCAM-1 targeted Nanoparticle for Osteoarthritis | 0 | nan | 1002 | 116 | 343 | novel - unpublished |
| 2 | VCAM-1 targeted Liposome for Osteoarthritis | 0 | nan | 269 | 116 | 162 | novel - unpublished |
| 3 | E-selectin targeted Nanoparticle for Rheumatoid arthritis | 0 | nan | 1115 | 206 | 105 | novel - unpublished |
| 4 | VCAM-1 targeted Nanoparticle for Pancreatic cancer | 0 | nan | 2555 | 101 | 343 | novel - unpublished |
| 5 | E-selectin targeted Nanoparticle for Osteoarthritis | 0 | nan | 1002 | 31 | 105 | novel - unpublished |
| 6 | E-selectin targeted Liposome for Osteoarthritis | 0 | nan | 269 | 31 | 92 | novel - unpublished |
| 7 | ICAM-1 targeted Liposome for Osteoarthritis | 1 | 2003.0 | 269 | 137 | 185 | novel - 1 publication < 2021 (untimely) |
| 8 | ICAM-1 targeted Nanoparticle for Pancreatic cancer | 1 | 2021.0 | 2555 | 136 | 328 | novel - 1 publication >= 2021 (timely) |
| 9 | ICAM-1 targeted Liposome for Pancreatic cancer | 1 | 2016.0 | 794 | 136 | 185 | novel - 1 publication < 2021 (untimely) |
| 10 | E-selectin targeted Liposome for Myocardial infarction | 1 | 2003.0 | 273 | 176 | 92 | novel - 1 publication < 2021 (untimely) |
| 11 | E-selectin targeted Liposome for Rheumatoid arthritis | 1 | 2026.0 | 473 | 206 | 92 | novel - 1 publication >= 2021 (timely) |
| 12 | FAP targeted Liposome for Myocardial infarction | 1 | 2025.0 | 273 | 61 | 47 | novel - 1 publication >= 2021 (timely) |
| 13 | ICAM-1 targeted Nanoparticle for Hepatitis c | 0 | nan | 291 | 118 | 328 | novel - unpublished |
| 14 | ICAM-1 targeted Liposome for Hepatitis c | 0 | nan | 146 | 118 | 185 | novel - unpublished |
| 15 | E-selectin targeted Nanoparticle for Asthma | 0 | nan | 649 | 139 | 105 | novel - unpublished |
| 16 | E-selectin targeted Liposome for Asthma | 0 | nan | 213 | 139 | 92 | novel - unpublished |
| 17 | CAV1 targeted Nanoparticle for Myocardial infarction | 0 | nan | 1214 | 65 | 72 | novel - unpublished |
| 18 | VCAM-1 targeted Nanoparticle for Hepatitis c | 0 | nan | 291 | 63 | 343 | novel - unpublished |
| 19 | VCAM-1 targeted Liposome for Hepatitis c | 0 | nan | 146 | 63 | 162 | novel - unpublished |
| 20 | FAP targeted Nanoparticle for Glaucoma | 0 | nan | 426 | 37 | 131 | novel - unpublished |
| 21 | FAP targeted Liposome for Glaucoma | 0 | nan | 150 | 37 | 47 | novel - unpublished |
| 22 | CAV1 targeted Nanoparticle for Asthma | 0 | nan | 649 | 32 | 72 | novel - unpublished |
| 23 | ICAM-1 targeted Nanoparticle for Glaucoma | 0 | nan | 426 | 31 | 328 | novel - unpublished |
| 24 | ICAM-1 targeted Liposome for Glaucoma | 0 | nan | 150 | 31 | 185 | novel - unpublished |
| 25 | VCAM-1 targeted Nanoparticle for Glaucoma | 0 | nan | 426 | 22 | 343 | novel - unpublished |

## Fig 6 – temporal and geographic analysis (demo with hypotheses named in the paper)

| hypothesis | reproduction trend | paper (to 2019) |
|---|---|---|
| osimertinib + lung cancer | trending | trending |
| cabozantinib + hepatocellular carcinoma | declining | trending |
| olaparib + lung cancer | trending | trending |
| olaparib + ovarian cancer | stable | stable |
| olaparib + breast cancer | trending | stable |
| pembrolizumab + lung cancer | stable | trending |
| nivolumab + melanoma | stable | trending |
| pazopanib + hepatocellular carcinoma | declining | declining |
| everolimus + pancreatic cancer | declining | declining |
| sorafenib + hepatocellular carcinoma | stable |  |

| hypothesis | top country | most over-represented country (>=5 papers) | paper |
|---|---|---|---|
| HIPEC + ovarian cancer | USA | Spain | Italy/France |
| oncolytic virus + ovarian cancer | USA | Canada | USA |
| high intensity focused ultrasound + glioma | USA | Taiwan | Netherlands |
| immunotherapy + esophageal cancer | China | China | Japan |
| radiotherapy + gastrointestinal stromal tumor | USA | Sweden | Germany |
| radiotherapy + breast cancer | USA | Sweden |  |
| chemotherapy + lung cancer | China | Japan |  |

