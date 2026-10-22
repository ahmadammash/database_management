import re

import pandas as pd

# 1. Load data
treatments_list = pd.read_excel(
    io="C:/Users/medaamm/Box/Data file (AA)/Miscellaneous Requests/Treatments List (Leeds).xlsx",
    sheet_name="Original List",
    header=0,
)


def clean_treatment_agent(text):
    if pd.isna(text):
        return ""

    text = str(text).strip()

    # Remove parenthetical notes e.g., (3 cycles), (planned for 5 years), (chest wall)
    text = re.sub(r"\([^)]*\)", "", text)

    # Standardize separator words and symbols into clean commas
    text = re.sub(
        r"\s*(\+|\/|&|\band\b|\balongside\b|\bthen\b|\bwith\b|\bplus\b)\s*",
        ", ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\s*;\s*", ", ", text)

    # Strip schedule & dosage patterns (e.g., "3-weekly", "6 months", "150 mg daily", "24h")
    text = re.sub(
        r"\b\d+[-\s]*(weekly|monthly|cycles?|months?|years?|days?)\b",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\b\d+\s*(mg|mcg|ug|gbq|g)\b", "", text, flags=re.IGNORECASE)
    text = re.sub(
        r"\b(daily|indefinite|maintenance|injections?|cover)\b",
        "",
        text,
        flags=re.IGNORECASE,
    )

    # Explicit drug typo mapping dictionary
    corrections = {
        r"\bCaboplatin\b": "Carboplatin",
        r"\bCarbopaltin\b": "Carboplatin",
        r"\bCapacitbine\b|\bCapacetabine\b|\bCapcitabine\b|\bCepecitabine\b|\bCapacitabine\b|\bCapectibine\b|\bCapecitabin\b|\bCapecitibine\b": "Capecitabine",
        r"\bCarbazitaxel\b": "Cabazitaxel",
        r"\bCarbozantinib\b": "Cabozantinib",
        r"\bCyclophoshamide\b": "Cyclophosphamide",
        r"\bDenusumab\b|\bDenosomab\b": "Denosumab",
        r"\bDegerelix\b": "Degarelix",
        r"\bEnzolutamide\b": "Enzalutamide",
        r"\bFluorouracill\b|\bFluoruracil\b|\bFlourouracil\b|\bFlurouracil\b|\bFlouoruoracil\b|\bFluorouraci\b": "Fluorouracil",
        r"\bGemicitabine\b|\bGemcitabinr\b": "Gemcitabine",
        r"\bIpililumab\b|\bIpilmumab\b|\bIpilumab\b": "Ipilimumab",
        r"\bIrinitecan\b|\bIrinorecan\b|\bIrinocap\b": "Irinotecan",
        r"\bLipoomal\b": "Liposomal",
        r"\bOxxailplaitin\b|\bOxaplatin\b|\bOxaliplatine\b": "Oxaliplatin",
        r"\bPacliataxel\b|\bPacitaxel\b|\bPaclixaxel\b": "Paclitaxel",
        r"\bPembrilizumab\b|\bPembrolizuab\b|\bPembrozulimab\b|\bPemrolizumab\b": "Pembrolizumab",
        r"\bPeterexed\b|\bPremetrexed\b": "Pemetrexed",
        r"\bRalcitrexed\b": "Raltitrexed",
        r"\bTamoxofen\b": "Tamoxifen",
        r"\bTrabetinib\b|\bTrematinib\b|\bTarametinib\b": "Trametinib",
        r"\bTrabectidin\b": "Trabectedin",
        r"\bVefurafenib\b": "Vemurafenib",
        r"\bVenorelbine\b|\bVionrelbine\b|\bVinerolbine\b": "Vinorelbine",
        r"\bVincrinstine\b": "Vincristine",
        r"\bRegorafinib\b": "Regorafenib",
        r"\bEncorafinib\b": "Encorafenib",
        r"\bIfosamide\b": "Ifosfamide",
        r"\bZolendronic\b|\bZolendronate\b": "Zoledronic Acid",
        r"\b5-?\s*Fluorouracil\b": "5-Fluorouracil",
        r"\bOxaliplatin Capecitabine\b": "Oxaliplatin, Capecitabine",
        r"\bBleomycin Electrochemotherapy\b": "Bleomycin- Electrochemotherapy",
        r"\bFolinate Filgrastim\b": "Folinate, Filgrastim",
        r"\bDoxorubicin Or Doxorubicin\b": "Doxorubicin",
        r"\bEtoposide.\b": "Etoposide, ",
        r"\bAdded To Current Zoladex Therapy\b": ", Zoladex",
        r"\bAdded To Zoladex\b": ", Zoladex",
        r"\bIn Addition To\b": ", ",
        r"\bAdded To\b": ", ",
        r"\bEtoposide. Carboplatin\b": "Etoposide, Carboplatin",
        r"\bOxaliplatin. Capecitabine\b": "Oxaliplatin, Capecitabine",
        r"\bPaclitaxel. Carboplatin\b": "Paclitaxel, Carboplatin",
        r"\b \(11\. Bicalutamide\b": ", Bicalutamide",
        r"\bZoladex  .  Zoladex\b": "Zoladex",
    }

    for pattern, replacement in corrections.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    # Standardize title casing per drug token while retaining acronyms (e.g. AZD6738, BEAM, DTIC)
    tokens = [t.strip() for t in text.split(",") if t.strip()]
    cleaned_tokens = []
    for token in tokens:
        # If token is an acronym/code (all caps or mixed letters+numbers like AZD6738), preserve upper case
        if re.match(r"^[A-Z0-9]+$", token):
            cleaned_tokens.append(token)
        else:
            cleaned_tokens.append(token.title())

    # Clean up excess spaces, dashes, and duplicate trailing commas
    res = ", ".join(cleaned_tokens)
    res = re.sub(r"\s+", " ", res).strip(" ,-")
    return res


# Apply cleaning logic
treatments_list["Treatment Agent"] = treatments_list["Treatment Agent"].apply(
    clean_treatment_agent
)

# define the strings to remove
strings_to_remove = [
    ", Subsequent To Liver Resection",
    ", Weekly",
    "Based",
    "M2,",
    "Dose Dense",
    "Dose Dense-",
    "X Sa",
    ". Note Gemcitabine Omitted In Cycle 1 Due To Urinary Tract Infection",
    ", Nb Only Of Capecitabine Given",
    "Alone",
    " . Note Diagnosed, Synchronous Lung Cancer, Treated For Primary Lung Cancer As Well As Gastric Cancer",
    "Not Recorded)",
    "- Foxtrot Trial",
    "- Post Liver Resection",
    " Modified De Gramont Schedule",
    "? Cycles",
    ". Note Tte Is Calculated From Date Of Diagnosis - Hence Minus Figures",
    "Pnet: 37.",
    "Tabs",
    "X 2",
    ", Day",
    ", 2.",
    "For Further Local Recurrence",
    "((",
    "(Planned For",
    "Added Month 10",
    "- For At Least",
    "- See Comments",
    "For Maximum Androgen Blockade",
    "Weekly",
    "(2 Weeks - Immediately Prior To Specimen Surgery",
    "6 Weeks",
    "1 Week",
    "12 X 4 Week Cycles",
    "( 1 Week Followed By",
    "7 X Cycles)",
    "- Approx 1.",
    " 2.",
    "10.",
    "11. Later",
    "11.",
    ", 3.",
    " ( Dose Of 3., Followed By 11.25 - Total Duration",
    "(11.)",
    "11. 1 X3-Monthly",
    "11. In Combination",
    "3-Monthly",
    "3-Monthly Treatment",
    "Circa",
    "Monthly",
    "Switched To",
    "Total Duration",
    "0.",
    "Titrated Down To 0.",
    "( Dose Of 3., Followed By 25 -",
    "3.",
]

# remove the defined strings from the 'Treatment Agent' column
for string in strings_to_remove:
    treatments_list["Treatment Agent"] = treatments_list["Treatment Agent"].str.replace(
        string, "", regex=False
    )

# delete specific records
treatments_list = treatments_list.loc[treatments_list["Row ID"] != 169]

# Remove duplicates based on Treatment Type and cleaned Treatment Agent
treatments_list = treatments_list.drop_duplicates(
    subset=["Treatment Type", "Treatment Agent"]
)

# Copy back to clipboard
treatments_list.to_clipboard(index=False)
