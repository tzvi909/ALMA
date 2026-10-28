"""Numbers and rankings printed in the figures of Harif et al. (searches run ~Feb 2020),
transcribed from the manuscript PDF for validating the reproduction."""

# Fig 1b: text generator output  "treating <cancer> with <kinase> inhibitor combined with radiotherapy"
FIG1B = {
    ("Head and neck cancer", "EGFR"): 82, ("Lung cancer", "EGFR"): 70, ("Breast cancer", "EGFR"): 22,
    ("Breast cancer", "HER2"): 16, ("Head and neck cancer", "PI3K"): 12, ("Lung cancer", "PI3K"): 8,
    ("Breast cancer", "PI3K"): 7, ("Prostate cancer", "EGFR"): 7, ("Head and neck cancer", "ERK"): 6,
    ("Prostate cancer", "PI3K"): 5, ("Lung cancer", "HER2"): 5, ("Breast cancer", "ERK"): 4,
    ("Hepatocellular carcinoma", "PI3K"): 4, ("Lung cancer", "ERK"): 4, ("Lung cancer", "PDGFR"): 4,
    ("Melanoma", "EGFR"): 4, ("Head and neck cancer", "PDGFR"): 3, ("Head and neck cancer", "HER2"): 2,
    ("Hepatocellular carcinoma", "EGFR"): 2, ("Hepatocellular carcinoma", "ERK"): 2,
    ("Prostate cancer", "PDGFR"): 2, ("Lung cancer", "FGFR"): 2, ("Melanoma", "ERK"): 2,
}

# Fig 2b: drugs x cancers (BC, Glioma, OS, PC, CC, LS)
FIG2B_COLS = ["Breast cancer", "Glioma", "Osteosarcoma", "Prostate cancer", "Cholangiocarcinoma", "Liposarcoma"]
FIG2B = {
    "Doxorubicin": [11669, 659, 1913, 1215, 91, 27], "5-Fu": [8149, 275, 96, 568, 346, 8],
    "Paclitaxel": [7271, 396, 105, 938, 30, 12], "Cisplatin": [3725, 1090, 1415, 1049, 311, 42],
    "Rapamycin": [1595, 641, 147, 526, 47, 14], "Vincristine": [1379, 1077, 646, 242, 13, 19],
    "Gemcitabine": [1251, 68, 107, 214, 512, 38], "Etoposide": [1155, 638, 451, 483, 12, 12],
    "Mitomycin": [873, 43, 44, 177, 56, 14], "Everolimus": [626, 116, 8, 64, 11, 2],
    "Irinotecan": [461, 376, 59, 87, 36, 0], "Imatinib": [322, 120, 47, 137, 10, 19],
    "Sunitinib": [300, 82, 10, 104, 8, 7], "Sorafenib": [278, 74, 33, 93, 64, 6],
    "Cytarabine": [248, 103, 35, 37, 0, 3], "Dacarbazine": [166, 3197, 69, 40, 0, 38],
    "Pemetrexed": [154, 8, 9, 16, 1, 0],
}

# Fig 2d: most studied cancer per kinase inhibitor after horizontal normalization (nNOP = 1)
FIG2D_TOP = {
    "Dasatinib": "Leukemia", "Bosutinib": "Leukemia", "Imatinib": "Leukemia",
    "Gefitinib": "Lung cancer", "Erlotinib": "Lung cancer", "Crizotinib": "Lung cancer",
    "Palbociclib": "Breast cancer", "Neratinib": "Breast cancer", "Lapatinib": "Breast cancer",
    "Trametinib": "Melanoma", "Dabrafenib": "Melanoma", "Encorafenib": "Melanoma",
    "Cabozantinib": "Renal cell carcinoma", "Everolimus": "Renal cell carcinoma", "Sunitinib": "Renal cell carcinoma",
    "Sorafenib": "Hepatocellular carcinoma", "Regorafenib": "Hepatocellular carcinoma",
}

# Fig 3a: KI + HNC
FIG3A = {"Gefitinib": 420, "Erlotinib": 393, "Sorafenib": 361, "Sirolimus": 272, "Vandetanib": 236,
         "Imatinib": 175, "Sunitinib": 174, "Lenvatinib": 164, "Lapatinib": 119, "Cabozantinib": 108,
         "Everolimus": 96, "Afatinib": 86, "Dasatinib": 84, "Vemurafenib": 68, "Pazopanib": 58,
         "Temsirolimus": 47, "Crizotinib": 47, "Axitinib": 41, "Trametinib": 38, "Dabrafenib": 29,
         "Dacomitinib": 22, "Palbociclib": 16, "Ponatinib": 14, "Regorafenib": 14, "Nilotinib": 13}

# Fig 3b comparison matrices: KI -> (KI+var, KI+HNC, KI+HNC+var); plus var alone and HNC+var
FIG3B = {
    "radiotherapy": {
        "_var": 354364, "_col_var": 58292,
        "Gefitinib": (462, 420, 92), "Erlotinib": (523, 393, 81), "Sorafenib": (437, 361, 65),
        "Sirolimus": (284, 272, 31), "Lenvatinib": (39, 164, 28), "Cabozantinib": (30, 108, 10),
        "Afatinib": (66, 86, 18), "Axitinib": (29, 41, 6), "Trametinib": (33, 38, 5),
        "Dacomitinib": (8, 22, 7), "Palbociclib": (24, 16, 0), "Nilotinib": (10, 13, 2),
        "Regorafenib": (23, 14, 0)},
    "nanoparticle": {
        "_var": 209356, "_col_var": 750,
        "Gefitinib": (58, 420, 4), "Erlotinib": (66, 393, 1), "Sorafenib": (150, 361, 1),
        "Sirolimus": (151, 272, 0), "Lenvatinib": (5, 164, 0), "Cabozantinib": (4, 108, 0),
        "Afatinib": (8, 86, 0), "Axitinib": (11, 41, 0), "Trametinib": (5, 38, 0),
        "Palbociclib": (2, 16, 0), "Nilotinib": (9, 13, 0), "Regorafenib": (4, 14, 0)},
}

# Fig 3c: top-5 novel & reasonable KIs for HNC
FIG3C = {"radiotherapy": ["Palbociclib", "Regorafenib", "Alectinib", "Ibrutinib", "Ceritinib"],
         "nanoparticle": ["Sirolimus", "Imatinib", "Everolimus", "Dasatinib", "Vandetanib"]}

# Fig 3d: top-10 KI + radiotherapy + cancer
FIG3D = [("Lapatinib", "Prostate cancer"), ("Lapatinib", "Pancreatic cancer"), ("Pazopanib", "Melanoma"),
         ("Trametinib", "Breast cancer"), ("Temsirolimus", "Hepatocellular carcinoma"), ("Dasatinib", "Melanoma"),
         ("Vandetanib", "Breast cancer"), ("Vandetanib", "Prostate cancer"), ("Lenvatinib", "Breast cancer"),
         ("Cabozantinib", "Lung cancer")]

# Fig 4b: drug -> (drug+NP, drug+MG63, &NP, drug+U2OS, &NP, drug+Fadu, &NP); cell+NP row
FIG4B_CELL_NP = {"MG63": 251, "U2OS": 22, "Fadu": 31}
FIG4B = {
    "Rapamycin": (238, 44, 0, 18, 0, 15, 0), "Paclitaxel": (2399, 15, 4, 18, 1, 36, 2),
    "Gemcitabine": (396, 12, 2, 5, 0, 3, 0), "Sorafenib": (155, 4, 0, 1, 0, 4, 0),
    "Nintedanib": (6, 0, 0, 0, 0, 2, 0), "Carfilzomib": (8, 0, 0, 1, 0, 0, 0),
    "Nilotinib": (10, 1, 0, 0, 0, 0, 0), "Cabozantinib": (4, 0, 0, 0, 0, 0, 0),
    "Trametinib": (5, 0, 0, 0, 0, 0, 0), "Palbociclib": (2, 0, 0, 0, 0, 0, 0),
    "Ponatinib": (1, 0, 0, 0, 0, 0, 0), "Duvelisib": (0, 0, 0, 0, 0, 0, 0),
    "Infigratinib": (0, 0, 0, 0, 0, 0, 0)}
# Fig 4c: drug + Osteosarcoma, drug + HNSCC
FIG4C_OS = {"Gemcitabine": 107, "Sorafenib": 33, "Nilotinib": 1, "Carfilzomib": 1, "Nintedanib": 2,
            "Trametinib": 2, "Cabozantinib": 3, "Ponatinib": 1, "Infigratinib": 1, "Duvelisib": 0}
FIG4C_HNSCC = {"Gemcitabine": 94, "Sorafenib": 34, "Nilotinib": 6, "Carfilzomib": 3, "Nintedanib": 2,
               "Trametinib": 9, "Cabozantinib": 0, "Ponatinib": 3, "Infigratinib": 0, "Duvelisib": 0}
FIG4_SCREENED = ["Gemcitabine", "Sorafenib", "Nilotinib", "Carfilzomib", "Nintedanib", "Trametinib",
                 "Cabozantinib", "Ponatinib", "Infigratinib", "Duvelisib"]

# Fig 5b: diseases x biomaterials (NP, Lip, HG, RAb)
FIG5B_COLS = ["Nanoparticle", "Liposome", "Hydrogel", "Radiolabeled Antibody"]
FIG5B = {"Breast cancer": [5385, 1397, 366, 379], "Lung cancer": [2776, 1309, 141, 325],
         "Pancreatic cancer": [895, 281, 57, 117], "Myocardial infarction": [483, 156, 378, 36],
         "Rheumatoid arthritis": [305, 212, 56, 61], "Asthma": [291, 116, 22, 18],
         "Osteoarthritis": [195, 74, 270, 8], "Hepatitis c": [159, 114, 10, 9], "Glaucoma": [156, 74, 109, 2]}
# target alone, target + biomaterial (NP, Lip, HG, RAb), target + disease (FIG5B row order)
FIG5B_TARGETS = {
    "Annexin A1": (1649, [11, 27, 1, 0], [73, 61, 25, 12, 31, 21, 8, 1, 3]),
    "ASGPR": (1782, [134, 109, 5, 12], [6, 12, 4, 2, 2, 0, 0, 32, 0]),
    "VCAM-1": (19761, [191, 120, 23, 53], [292, 334, 81, 336, 380, 398, 94, 53, 12]),
}
# Fig 5c: (target, biomaterial, disease)
FIG5C_UNPUBLISHED = [("Annexin A1", "Liposome", "Pancreatic cancer"), ("FAP", "Nanoparticle", "Pancreatic cancer"),
                     ("ASGPR", "Nanoparticle", "Hepatitis c"), ("CAV1", "Nanoparticle", "Myocardial infarction"),
                     ("CAV1", "Nanoparticle", "Glaucoma"), ("Annexin A1", "Liposome", "Rheumatoid arthritis"),
                     ("E-selectin", "Nanoparticle", "Osteoarthritis"), ("PSMA", "Radiolabeled Antibody", "Lung cancer"),
                     ("VCAM-1", "Hydrogel", "Pancreatic cancer")]
FIG5C_TIMELY = [("ICAM-1", "Nanoparticle", "Asthma"), ("ICAM-1", "Liposome", "Pancreatic cancer"),
                ("VCAM-1", "Liposome", "Asthma"), ("VCAM-1", "Hydrogel", "Myocardial infarction"),
                ("E-selectin", "Nanoparticle", "Myocardial infarction")]
FIG5C_UNTIMELY = [("E-selectin", "Nanoparticle", "Breast cancer"), ("FAP", "Nanoparticle", "Breast cancer"),
                  ("FAP", "Nanoparticle", "Lung cancer")]

# Fig 6 / text: temporal trends reported up to 2019
FIG6_TRENDS = {"osimertinib + lung cancer": "trending", "cabozantinib + hepatocellular carcinoma": "trending",
               "olaparib + lung cancer": "trending", "pembrolizumab + lung cancer": "trending",
               "nivolumab + melanoma": "trending", "olaparib + ovarian cancer": "stable",
               "olaparib + breast cancer": "stable", "pazopanib + hepatocellular carcinoma": "declining",
               "everolimus + pancreatic cancer": "declining"}
FIG6_GEO = {"HIPEC + ovarian cancer": "Italy/France", "oncolytic virus + ovarian cancer": "USA",
            "high intensity focused ultrasound + glioma": "Netherlands",
            "immunotherapy + esophageal cancer": "Japan", "radiotherapy + gastrointestinal stromal tumor": "Germany"}
