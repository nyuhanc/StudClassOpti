"""All Slovene user-facing text and the English<->Slovene data boundary.

The core package stays English (column names, subject names, constraint
labels) for maintainability/transparency; everything the user sees is Slovene.
This module is the single place that bridges the two:

* UI chrome strings (buttons, titles, status messages) as constants.
* Constraint label/description/parameter-label translations keyed by id.
* Column- and subject-name maps used to (a) rename a Slovene spreadsheet to
  the English names the core expects on load, and (b) render results/exports
  back in Slovene.
"""

from __future__ import annotations

import pandas as pd

# ---- column names: English (internal) -> Slovene (spreadsheet + display) ----
COLUMNS_EN_SL = {
    "Student": "Dijak",
    "Gender": "Spol",
    "Surname and name": "Priimek in ime",
    "Schoolmate": "Sošolec",
    "NationalTestScore": "Točke NPZ",
    # solver output columns:
    "Class": "Razred",
    "Language": "Jezik",
    "NatSci1": "Naravoslovje 1",
    "NatSci2": "Naravoslovje 2",
}

LANGUAGES_EN_SL = {
    "French": "Francoščina",
    "Italian": "Italijanščina",
    "German": "Nemščina",
    "Russian": "Ruščina",
    "Spanish": "Španščina",
}

SCIENCES_EN_SL = {
    "Biology": "Biologija",
    "Physics": "Fizika",
    "Chemistry": "Kemija",
}

SUBJECTS_EN_SL = {**LANGUAGES_EN_SL, **SCIENCES_EN_SL}
_ALL_EN_SL = {**COLUMNS_EN_SL, **SUBJECTS_EN_SL}
_SL_TO_EN = {sl: en for en, sl in _ALL_EN_SL.items()}


def subject_to_sl(name: str) -> str:
    """Slovene display name for a language/science (passthrough if unknown)."""
    return SUBJECTS_EN_SL.get(name, name)


def col_to_sl(name: str) -> str:
    """Slovene name for a spreadsheet column (passthrough if unknown)."""
    return _ALL_EN_SL.get(name, name)


def col_to_en(name: str) -> str:
    """Internal English column name for a Slovene one (passthrough if unknown)."""
    return _SL_TO_EN.get(name, name)


def excel_to_internal(df: pd.DataFrame) -> pd.DataFrame:
    """Rename a Slovene spreadsheet's columns to the English names core uses."""
    return df.rename(columns=_SL_TO_EN)


def localize_result_df(df: pd.DataFrame) -> pd.DataFrame:
    """A copy of a result frame with Slovene headers and Slovene subject cells."""
    df = df.copy()
    for col in ("Language", "NatSci1", "NatSci2"):
        if col in df.columns:
            df[col] = df[col].map(lambda v: subject_to_sl(v) if v is not None else v)
    return df.rename(columns=_ALL_EN_SL)


def localize_errors(text: str) -> str:
    """Replace English column/subject names in a core error string with Slovene."""
    # Longest first so multi-word names ("Surname and name") win over substrings.
    for en in sorted(_ALL_EN_SL, key=len, reverse=True):
        text = text.replace(en, _ALL_EN_SL[en])
    return text


# ---- constraint translations (keyed by Constraint.id) ----
# Each entry: label, description, and a {param_key: label} map.
CONSTRAINTS_SL: dict[str, dict] = {
    "c01": {
        "label": "Največja velikost razreda",
        "description": "Največ toliko dijakov na razred.",
        "params": {},
    },
    "c02": {
        "label": "Kapaciteta jezika",
        "description": "Posamezni jezik lahko izbere največ (množitelj × največja velikost razreda) dijakov.",
        "params": {"lang_capacity_multiplier": "Množitelj kapacitete"},
    },
    "c03": {
        "label": "Kapaciteta naravoslovja",
        "description": (
            "Posameznemu naravoslovnemu predmetu se dodeli največ (množitelj × največja "
            "velikost razreda) mest (vsak dijak zasede dve mesti)."
        ),
        "params": {"nat_sci_capacity_multiplier": "Množitelj kapacitete"},
    },
    "c04": {
        "label": "Različna naravoslovna predmeta",
        "description": "Dijakova naravoslovna predmeta morata biti različna.",
        "params": {},
    },
    "c05": {
        "label": "Ohrani sošolce skupaj",
        "description": (
            "Dijak, ki v stolpcu za sošolce navede partnerja, je razvrščen v isti razred "
            "kot ta partner. Opozorilo: ta omejitev praviloma poslabša kakovost razvrstitve, "
            "v skrajnem primeru pa lahko naredi problem nerešljiv."
        ),
        "params": {"schoolmate_column": "Ime stolpca za sošolce"},
    },
    "c06": {
        "label": "Združi najpogostejši naravoslovni par",
        "description": (
            "Dijaki z najpogostejšim parom naravoslovnih predmetov (samodejno zaznan ali "
            "nastavljen) se zberejo v en razred."
        ),
        "params": {},
    },
    "c08": {
        "label": "Porazdeli fante",
        "description": (
            "Fantje so zbrani v točno toliko razredov, kot je nastavljeno; vsak tak razred "
            "vsebuje med najmanj in največ fantov, ostali razredi nobenega."
        ),
        "params": {
            "male_gender_value": "Oznaka spola za fante",
            "male_classes": "Število razredov s fanti",
            "male_min_per_class": "Najmanj fantov na tak razred",
            "male_max_per_class": "Največ fantov na tak razred",
        },
    },
    "c09": {
        "label": "Zagotovi prvi jezik",
        "description": "Dijaki, katerih prva izbira jezika je izbrani jezik, ga zagotovo dobijo.",
        "params": {"forced_top_language": "Jezik"},
    },
    "c10": {
        "label": "Omejitev mest za predmet",
        "description": "Izbranemu predmetu se dodeli največ toliko mest, kolikor je največja velikost razreda.",
        "params": {"capacity_subject": "Predmet"},
    },
    "c11": {
        "label": "Vedno izberi predmet pri visoki prioriteti",
        "description": (
            "Dijaki, ki imajo izbrani predmet kot 1. ali 2. naravoslovno prioriteto, ga "
            "zagotovo dobijo v enem od svojih dveh mest."
        ),
        "params": {"gated_subject": "Predmet"},
    },
    "c12": {
        "label": "Vedno izključi predmet pri nizki prioriteti",
        "description": (
            "Dijaki, ki imajo izbrani predmet kot najnižjo naravoslovno prioriteto, ga ne "
            "dobijo. (Predmet je skupen s pravilom o obveznem predmetu.)"
        ),
        "params": {"gated_subject": "Predmet"},
    },
    "c13": {
        "label": "Poveži naravoslovna predmeta",
        "description": "Dijak, ki dobi predmet »iz«, mora v drugem mestu dobiti predmet »v«.",
        "params": {
            "paired_from": "Iz predmeta",
            "paired_to": "V predmet",
        },
    },
    "c14": {
        "label": "Izključitev jezika",
        "description": (
            "Dijaki, ki imajo izbrani jezik kot prvo prioriteto, ne smejo dobiti izključenega "
            "jezika."
        ),
        "params": {
            "exclusion_top_language": "Prva izbira jezika",
            "excluded_language": "Izključeni jezik",
        },
    },
}


def constraint_label(cid: str) -> str:
    return CONSTRAINTS_SL[cid]["label"]


def constraint_description(cid: str) -> str:
    return CONSTRAINTS_SL[cid]["description"]


def param_label(cid: str, key: str, fallback: str) -> str:
    return CONSTRAINTS_SL[cid]["params"].get(key, fallback)


# Constraints not shown in the GUI: always-on rules that are self-evident to a
# regular user (c04: a student's two sciences must differ). They keep their
# config default (enabled) since the panel never touches them.
HIDDEN_CONSTRAINTS = {"c04"}

# Constraints that accept only a single subject/language; marked with a red "*"
# whose meaning is explained once by SINGLE_CHOICE_FOOTNOTE.
SINGLE_CHOICE_CONSTRAINTS = {"c09", "c10", "c11", "c12", "c13", "c14"}

SINGLE_CHOICE_MARK = ' <span style="color:#c00; font-weight:bold;">*</span>'
SINGLE_CHOICE_FOOTNOTE = (
    '<span style="color:#c00; font-weight:bold;">*</span> Pri tej omejitvi je mogoče '
    "izbrati le eno možnost. Izbira več hkrati (npr. več predmetov ali jezikov) bi pogosto "
    "naredila razvrstitev nerešljivo."
)


# ---- UI chrome ----
WINDOW_TITLE = "Razvrščanje dijakov v razrede"

# General settings
GENERAL = "Splošno"
NUM_CLASSES = "Število razredov"
MAX_CLASS_SIZE = "Največja velikost razreda"
SHUFFLES = "Število poskusov (s premešanjem)"
TIME_LIMIT = "Časovna omejitev na poskus (s)"
SEARCH_WORKERS = "Število niti za iskanje"

NUM_CLASSES_DESC = "Na koliko vzporednih razredov naj se razporedijo dijaki."
MAX_CLASS_SIZE_DESC = "Največje dovoljeno število dijakov v posameznem razredu."
SHUFFLES_DESC = (
    "Kolikokrat naj program poskusi z naključno premešanim vrstnim redom dijakov; "
    "obdrži najboljši rezultat. Več poskusov običajno pomeni boljši rezultat, a daljši čas optimizacije."
)
TIME_LIMIT_DESC = (
    "Koliko sekund ima reševalnik na voljo za en poskus, preden se ustavi in vzame "
    "najboljšo do tedaj najdeno rešitev tega poskusa."
)
SEARCH_WORKERS_DESC = (
    "Koliko procesorskih niti naj reševalnik uporablja hkrati. Več niti pohitri iskanje "
    "(smiselno do števila jeder procesorja)."
)

# Objective weights
OBJECTIVE_BOX = "Uteži optimizacijske funkcije (napredno)"
LANG_IMPORTANCE = "Pomembnost jezika"
LANG_PENALTY = "Kazen za jezik"
NS1_IMPORTANCE = "Pomembnost naravoslovja 1"
NS2_IMPORTANCE = "Pomembnost naravoslovja 2"
NS_PENALTY = "Kazen za naravoslovje"
STRATIFICATION = "Stratifikacija"

LANG_IMPORTANCE_DESC = "Kako močno se nagradi, da dijak dobi želeni jezik (višja vrednost = jezik šteje bolj)."
LANG_PENALTY_DESC = "Kazen, kadar dijak ne dobi svoje prve izbire jezika."
NS1_IMPORTANCE_DESC = "Kako močno se nagradi prvi (glavni) naravoslovni predmet dijaka."
NS2_IMPORTANCE_DESC = "Kako močno se nagradi drugi naravoslovni predmet dijaka."
NS_PENALTY_DESC = "Kazen, kadar dijak v nobenem od obeh mest ne dobi svoje prve izbire naravoslovja."
STRATIFICATION_DESC = (
    "Kako močno se višje uvrščene želje cenijo bolj od nižjih. Višja vrednost še bolj "
    "poudari izpolnjevanje prvih izbir pred nižjimi."
)

# File / config rows
NO_FILE = "Datoteka ni naložena"
BROWSE = "Prebrskaj…"
SAVE_CONFIG_BTN = "Shrani nastavitve…"
LOAD_CONFIG_BTN = "Naloži nastavitve…"
SAVE_CONFIG_TITLE = "Shrani nastavitve"
LOAD_CONFIG_TITLE = "Naloži nastavitve"
OPEN_SPREADSHEET_TITLE = "Odpri datoteko"
CONSTRAINTS = "Omejitve"

# Constraint panel
SCIENCE_PAIR = "Naravoslovni par"
AUTO_PAIR = "Samodejno (najpogostejši)"

# Run / status
RUN = "Zaženi"
CANCEL = "Prekliči"
RUNNING = "Poteka izračun…"
CANCELLING = "Preklic po trenutnem poskusu…"
DATA_OK = "Podatki so v redu. Pripravljeno za zagon."


def could_not_read(exc) -> str:
    return f"Datoteke ni mogoče prebrati: {exc}"


def data_errors_header(n: int) -> str:
    return f"{n} napak v podatkih — popravite pred zagonom:"


def and_more(n: int) -> str:
    return f"\n  …in še {n}"


def progress_text(p) -> str:
    score = f"{p.score:.0f}" if p.score is not None else "ni rešitve"
    return (
        f"Poskus {p.shuffle}/{p.total_shuffles}: {p.status}, "
        f"rezultat={score}, najboljši={p.best_score:.0f}"
    )


def done_best(score: float) -> str:
    return f"Končano. Najboljši rezultat {score:.0f}."


DONE_INFEASIBLE = "Končano. Ni bilo mogoče najti izvedljive rešitve."


def solver_error(message: str) -> str:
    return f"Napaka reševalnika: {message}"


# Results view
NO_RESULTS = "Še ni rezultatov. Naložite datoteko in zaženite."
NO_FEASIBLE = "V nobenem poskusu ni bilo izvedljive rešitve."
CANCELLED_EARLY = " (predčasno prekinjeno)"
EXPORT_BTN = "Izvozi v Excel…"
EXPORT_TITLE = "Izvozi rezultate"
EXPORT_FAILED = "Izvoz ni uspel"
EXPORTED = "Izvoženo"


def summary_text(score: float, cancelled: str, pair: tuple[str, str], sizes_txt: str) -> str:
    return (
        f"Najboljši rezultat: {score:.0f}{cancelled}  |  "
        f"par: {subject_to_sl(pair[0])} + {subject_to_sl(pair[1])}  |  {sizes_txt}"
    )


def exported_text(path: str) -> str:
    return f"Shranjeno:\n{path}"
