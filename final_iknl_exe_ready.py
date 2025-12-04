import pandas as pd
from collections import defaultdict
import numpy as np
import warnings
import sys
import os

# ========= GUI OVERRIDES =========
# Input files (from GUI env vars)
_sleutel = os.environ.get("IKNL_SLEUTEL_PATH")
_ckv     = os.environ.get("IKNL_CKV_PATH")

if _sleutel:
    path_to_sleutel = _sleutel
if _ckv:
    path_to_iknl = _ckv

# Output directory (from GUI env var)
_outdir = os.environ.get("IKNL_OUTPUT_DIR")
if _outdir:
    path_for_export_files = _outdir
else:
    # fallback if no GUI is used
    path_for_export_files = os.path.join(os.getcwd(), "output")

os.makedirs(path_for_export_files, exist_ok=True)

def to_int_or_NA(v):
    """
    Maakt van v een nette Int64:
    - lege strings / NaN -> pd.NA
    - '9', 9, 9.0, '9.0' -> 9
    - anders (onzin) -> pd.NA
    """
    if pd.isna(v):
        return pd.NA
    s = str(v).strip()
    if s == "":
        return pd.NA
    try:
        # eerst proberen of het al een integer-string is
        return int(s)
    except ValueError:
        try:
            # vang dingen als '9.0' of '9,0' af
            s = s.replace(",", ".")
            return int(float(s))
        except Exception:
            return pd.NA


def resource_dir():
    # When frozen with PyInstaller, sys._MEIPASS is where resources are unpacked.
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # type: ignore[attr-defined]
    # Otherwise, use the directory of this file.
    return os.path.dirname(os.path.abspath(__file__))


def pick_file_dialog(title, filetypes):
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return path or None
    except Exception:
        return None


def find_or_pick_file(preferred_names, extensions, title):
    # 1) Look next to the exe/script for any of the preferred filenames
    base_dir = os.getcwd()
    for name in preferred_names:
        candidate = os.path.join(base_dir, name)
        if os.path.isfile(candidate):
            return candidate
    # 2) Otherwise, look for a single file with the right extension(s)
    matches = []
    for ext in extensions:
        for entry in os.listdir(base_dir):
            if entry.lower().endswith(ext.lower()) and os.path.isfile(os.path.join(base_dir, entry)):
                matches.append(os.path.join(base_dir, entry))
    if len(matches) == 1:
        return matches[0]
    # 3) Fall back to a picker so the user can "upload" a file at runtime
    filetypes = [(ext.upper(), f"*{ext}") for ext in extensions]
    picked = pick_file_dialog(title, filetypes)
    if picked and os.path.isfile(picked):
        return picked
    raise FileNotFoundError(
        f"Kon geen bestand vinden voor: {title}. Zet het bestand naast de .exe of kies het via de dialoog."
    )


warnings.simplefilter(action="ignore", category=pd.errors.PerformanceWarning)
warnings.simplefilter(action="ignore", category=FutureWarning)

# Pad naar het Excel-bestand met de sleutellijst
if not (_sleutel and os.path.isfile(_sleutel)):
    path_to_sleutel = find_or_pick_file(
        preferred_names=["sleutel_iknl.xlsx", "sleutel.xlsx"],
        extensions=[".xlsx", ".xls"],
        title="Kies de sleutellijst (Excel)",
    )
else:
    path_to_sleutel = _sleutel

# Pad naar het IKNL CSV-bestand
if not (_ckv and os.path.isfile(_ckv)):
    path_to_iknl = find_or_pick_file(
        preferred_names=["ckv.csv", "iknl.csv"],
        extensions=[".csv"],
        title="Kies de CKV/IKNL data (CSV met ; als scheidingsteken)",
    )
else:
    path_to_iknl = _ckv

# Pad naar dictionary-Excel (met alle mapping-tabbladen)
_dicts = os.environ.get("IKNL_DICTS_PATH")
if not (_dicts and os.path.isfile(_dicts)):
    path_to_dicts = find_or_pick_file(
        preferred_names=[
            "iknl_dictionaries.xlsx",
            "iknl_dictionaries_from_code.xlsx",
            "dictionaries.xlsx",
        ],
        extensions=[".xlsx", ".xls"],
        title="Kies dictionary-Excel (mapping-tabbladen)",
    )
else:
    path_to_dicts = _dicts

# ================== DICTIONARY-LOADER UIT EXCEL ==================

# Hier loggen we alle keys die niet in een dictionary staan
unknown_mappings = []


def log_unknown(sheet_name: str, mapping_name: str, key):
    """Log een onbekende waarde zodat we die later naar CSV kunnen schrijven."""
    if pd.isna(key):
        return
    k = str(key).strip()
    if not k:
        return
    unknown_mappings.append({"sheet": sheet_name, "mapping": mapping_name, "key": k})


def simple_dict_from_sheet(dict_sheets, sheet_name, key_cast=None, val_cast=None):
    """
    Laadt een simpele mapping-dict uit een sheet met kolommen 'key' en 'value'.
    key_cast / val_cast zijn optionele functies om types te casten.
    """
    if sheet_name not in dict_sheets:
        raise ValueError(f"Dictionary-Excel mist sheet '{sheet_name}'.")
    df = dict_sheets[sheet_name]
    out = {}
    if not {"key", "value"}.issubset(df.columns):
        raise ValueError(f"Sheet '{sheet_name}' mist kolommen 'key' en/of 'value'.")
    for _, row in df[["key", "value"]].dropna().iterrows():
        k = row["key"]
        v = row["value"]
        if key_cast is not None:
            try:
                k = key_cast(k)
            except Exception:
                pass
        if val_cast is not None:
            try:
                v = val_cast(v)
            except Exception:
                pass
        out[k] = v
    return out


def tnm_key_cast(x):
    """
    Zet keys uit t/n/m om: '0','1','2' -> int, 'X','IS' blijven string.
    """
    s = str(x).strip()
    try:
        return int(s)
    except ValueError:
        return s


def map_with_log(val, mapping_dict, sheet_name, mapping_name, default=pd.NA, key_transform=None):
    """
    Voert een dictionary-look-up uit mét logging als key niet bestaat.
    - sheet_name: welke sheet in Excel hoort hierbij
    - mapping_name: naam van de mapping (bijv. 'topo_blok_bl', 'ki67_map')
    """
    if pd.isna(val):
        return default

    key = val
    if key_transform is not None:
        try:
            key = key_transform(val)
        except Exception:
            key = val

    if key in mapping_dict:
        return mapping_dict[key]

    log_unknown(sheet_name, mapping_name, key)
    return default


# Dictionary-Excel inlezen
dict_sheets = pd.read_excel(path_to_dicts, sheet_name=None)

# 1) TNM: rijen met key '-' zijn alleen placeholders -> als missing behandelen
for sh in ("tnm_t", "tnm_n", "tnm_m"):
    if sh in dict_sheets:
        df_sh = dict_sheets[sh].copy()
        # key '-' → NaN, zodat simple_dict_from_sheet hem weggooit
        df_sh.loc[df_sh["key"].astype(str).str.strip() == "-", "key"] = np.nan
        dict_sheets[sh] = df_sh

# 2) surgery_map: tekst 'nan' in value betekent "geen mapping" -> rij droppen
if "surgery_map" in dict_sheets:
    df_sh = dict_sheets["surgery_map"].copy()
    # value 'nan' (als tekst) → NaN → valt eruit in simple_dict_from_sheet
    df_sh.loc[df_sh["value"].astype(str).str.lower() == "nan", "value"] = np.nan
    dict_sheets["surgery_map"] = df_sh

# 3) suffix_map: lege waardes *bewust* toestaan (zoals bij key 's')
if "suffix_map" in dict_sheets:
    df_sh = dict_sheets["suffix_map"].copy()
    # lege cellen in 'value' → '' zodat dropna ze NIET weghaalt
    df_sh["value"] = df_sh["value"].fillna("")
    dict_sheets["suffix_map"] = df_sh

# --- Topografie / blokken / omschrijving (sheet: topography) ---
if "topography" not in dict_sheets:
    raise ValueError("Dictionary-Excel mist sheet 'topography'.")

topo_df = dict_sheets["topography"]

topo_blok_bl = {}
unique_values_dict = {}
value_map = {}

for _, row in topo_df.iterrows():
    code = str(row.get("icd_code", "")).strip()
    if not code or code.lower() == "nan":
        continue
    blok = row.get("blok_bl", pd.NA)
    uniq = row.get("unique_val", pd.NA)
    oms  = row.get("omschrijving", pd.NA)

    if not pd.isna(blok):
        try:
            topo_blok_bl[code] = int(blok)
        except Exception:
            topo_blok_bl[code] = blok

    if not pd.isna(uniq):
        unique_values_dict[code] = str(uniq).strip()

    if not pd.isna(oms):
        value_map[code] = str(oms).strip()

# --- Overige dictionaries, 1 sheet per dict (zoals uit export-script) ---

ki67_map = simple_dict_from_sheet(
    dict_sheets,
    "ki67_map",
    key_cast=lambda x: int(float(x)),
    val_cast=lambda x: int(float(x)),
)

mitoses_map = simple_dict_from_sheet(
    dict_sheets,
    "mitoses_map",
    key_cast=lambda x: int(float(x)),
    val_cast=lambda x: int(float(x)),
)

residual_map = simple_dict_from_sheet(
    dict_sheets,
    "residual_map",
    key_cast=lambda x: str(x).strip(),
    val_cast=lambda x: int(float(x)),
)

chemo_map = simple_dict_from_sheet(
    dict_sheets,
    "chemo_map",
    key_cast=lambda x: str(x).strip().upper(),
    val_cast=lambda x: int(float(x)),
)

t = simple_dict_from_sheet(
    dict_sheets,
    "tnm_t",
    key_cast=tnm_key_cast,
    val_cast=lambda x: int(float(x)),
)
n = simple_dict_from_sheet(
    dict_sheets,
    "tnm_n",
    key_cast=tnm_key_cast,
    val_cast=lambda x: int(float(x)),
)
m = simple_dict_from_sheet(
    dict_sheets,
    "tnm_m",
    key_cast=tnm_key_cast,
    val_cast=lambda x: int(float(x)),
)

suffix_map = simple_dict_from_sheet(
    dict_sheets,
    "suffix_map",
    key_cast=lambda x: str(x).strip().lower(),
    val_cast=lambda v: ("" if pd.isna(v) else v),
)

surgery_map = simple_dict_from_sheet(
    dict_sheets,
    "surgery_map",
    key_cast=lambda x: str(x).strip(),
    val_cast=lambda x: int(float(x)),
)

morphology_map = simple_dict_from_sheet(
    dict_sheets,
    "morphology_map",
    key_cast=lambda x: int(float(x)),
    val_cast=lambda x: str(x),
)

tumor_type_map = simple_dict_from_sheet(
    dict_sheets,
    "tumor_type_map",
    key_cast=lambda x: int(float(x)),
    val_cast=lambda x: str(x),
)

behavior_map = simple_dict_from_sheet(
    dict_sheets,
    "behavior_map",
    key_cast=lambda x: int(float(x)),
    val_cast=lambda x: str(x),
)

meta_chir_code_map = simple_dict_from_sheet(
    dict_sheets,
    "meta_chir_code_map",
    key_cast=lambda x: str(x).strip(),
    val_cast=lambda x: str(x).strip(),
)

treat_type_code_map = simple_dict_from_sheet(
    dict_sheets,
    "treat_type_code_map",
    key_cast=lambda x: str(x).strip(),
    val_cast=lambda x: int(float(x)),
)

map_dict = {
    "bl_tnm_t": t,
    "bl_tnm_n": n,
    "bl_tnm_m": m,
    "p_sx_tnm_t": t,
    "p_sx_tnm_n": n,
    "p_sx_tnm_m": m,
}

"""This section maps IKNL IDs to FORCE-NEN study IDs using a lookup table."""

# Read only the two relevant columns from the Excel file
df_map = pd.read_excel(path_to_sleutel, usecols=["IKNL ID", "FORCE-NEN ID"])

# Drop rows with missing values in either of the two ID columns
df_map = df_map.dropna(subset=["IKNL ID", "FORCE-NEN ID"])

# Create a dictionary mapping: key = IKNL ID, value = FORCE-NEN ID
iknl_to_force = dict(zip(df_map["IKNL ID"], df_map["FORCE-NEN ID"]))


def label_treat_types(treat_list):
    """
    Adds a numeric suffix to each treatment type to make them unique.
    """
    counter = defaultdict(int)  # Count occurrences of each treatment code
    labeled = []

    for t_ in treat_list:
        counter[t_] += 1
        labeled.append(f"{t_}_{counter[t_]}")  # Append suffix to make label unique

    return labeled


def label_treatment_blocks(blocks):
    """
    Adds a numeric suffix to each treatment block code to make them unique.
    """
    counter = defaultdict(int)  # Count occurrences of each code
    labeled = []

    for code, start, stop in blocks:
        counter[code] += 1
        labeled.append((f"{code}_{counter[code]}", start, stop))  # Label code with suffix

    return labeled


def extract_value_and_suffix(val, value_map_local, sheet_name, mapping_name):
    # 1) echte missings én streepjes direct als leeg behandelen
    if pd.isna(val):
        return "", ""
    val_str = str(val).strip()
    if val_str in ("", "-"):
        return "", ""

    # 2) vanaf hier je bestaande logica
    # Als het een getal is:
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        mapped = value_map_local.get(int(val))
        if mapped is None:
            log_unknown(sheet_name, mapping_name, int(val))
            mapped = ""
        return mapped, ""

    val_str = val_str.upper()

    if val_str in value_map_local:
        return value_map_local[val_str], ""

    digits = "".join(c for c in val_str if c.isdigit())
    letters = "".join(c for c in val_str if c.isalpha()).lower()
    base = int(digits) if digits else val_str

    mapped = value_map_local.get(base, "")
    if mapped == "":
        log_unknown(sheet_name, mapping_name, base)

    suff = suffix_map.get(letters, "")
    if letters and letters not in suffix_map:
        log_unknown("suffix_map", "suffix_map", letters)

    return mapped, suff

def check_text_conditions(row):
    """
    Checks tumor_diagnose_txt for specific hormone-related keywords
    if either horm_prod1 or horm_prod2 equals 98, and sets corresponding flags.
    """
    # Check if hormone production code is 98
    if row["horm_prod1"] == 98 or row["horm_prod2"] == 98:
        txt = str(row["tumor_diagnose_txt"]).upper()  # Normalize text

        # Set flag if serotonin is mentioned
        if "SEROTONINE" in txt:
            row["bl_ser"] = 1

        # Set flag if any 5-HIAA related terms are found
        if any(
            keyword in txt
            for keyword in [
                "5-HIAA",
                "5-HYDROXY-INDOL-3AZIJNZUUR",
                "5-HYDROXY-INDOL-AZIJNZUUR",
                "5-HIAA1",
                "5-HYDROXYL-3-AZIJNZUUR",
            ]
        ):
            row["bl_5_hiaa_u24"] = 1

        # Set flag if gastrin is mentioned
        if "GASTRIN" in txt:
            row["bl_gastrin"] = 1

    return row


def add_to_treat_type(row, code):
    """
    Adds a treatment code to the 'treat_type' list in the row.
    If 'treat_type' is not yet a list, initializes it as a new list with the code.
    """
    if isinstance(row["treat_type"], list):
        return row["treat_type"] + [code]
    return [code]


def extract_ordered_treatment_blocks(row):
    """
    Extracts all treatment blocks with start/stop dates and orders them
    according to the order in 'treat_type'. Missing blocks are added as (code, None, None).
    """
    blocks = []

    # Step 1: Collect all available treatment start/stop dates
    raw_blocks = []

    def try_add(code, start, stop):
        if pd.notna(start) or pd.notna(stop):
            raw_blocks.append((code, start, stop))

    # Add standard treatment types if any dates are available
    try_add("chemo1", row.get("chemo_startdat1"), row.get("chemo_stopdat1"))
    try_add("chemo2", row.get("chemo_startdat2"), row.get("chemo_stopdat2"))
    try_add("horm1", row.get("horm_startdat1"), row.get("horm_stopdat1"))
    try_add("horm2", row.get("horm_startdat2"), row.get("horm_stopdat2"))
    try_add("rt", row.get("rt_startdat1"), row.get("rt_stopdat1"))
    try_add("lokchir", row.get("lok_chir_startdat1"), row.get("lok_chir_stopdat1"))
    try_add("metart1", row.get("meta_rt_startdat1"), row.get("meta_rt_stopdat1"))
    try_add("orgchir1", row.get("org_chir_startdat1"), row.get("org_chir_stopdat1"))
    try_add("orgchir2", row.get("org_chir_startdat2"), row.get("org_chir_stopdat2"))
    try_add("overigchir", row.get("overig_chir_startdat1"), row.get("overig_chir_stopdat1"))
    try_add("overigther", row.get("overig_ther_startdat1"), row.get("overig_ther_stopdat1"))

    # Handle meta-surgery codes (can occur multiple times)
    for col, (start_col, stop_col) in {
        "meta_chir_code1": ("meta_chir_startdat1", "meta_chir_stopdat1"),
        "meta_chir_code2": ("meta_chir_startdat2", "meta_chir_stopdat2"),
    }.items():
        val = row.get(col)
        if pd.notna(val):
            code = meta_chir_code_map.get(str(val).strip())
            if code is None:
                log_unknown("meta_chir_code_map", "meta_chir_code_map", str(val).strip())
            if code:
                try_add(code, row.get(start_col), row.get(stop_col))

    # Step 2: Order the blocks to match the order in 'treat_type'
    treat_list = row["treat_type"][:] if isinstance(row["treat_type"], list) else []
    temp_blocks = raw_blocks.copy()

    for code in treat_list:
        match_idx = next((i for i, (c, _, _) in enumerate(temp_blocks) if c == code), None)
        if match_idx is not None:
            blocks.append(temp_blocks.pop(match_idx))
        else:
            blocks.append((code, None, None))

    return blocks


def collect_labeled_chemo_codes(row):
    """
    Collects non-empty chemo codes and labels them as 'chemo1: <code>' or 'chemo2: <code>'.
    Returns a list of labeled codes, or pd.NA if none are present.
    """
    result = []

    if pd.notna(row["chemo_code1"]):
        result.append(f"chemo1: {str(row['chemo_code1']).strip()}")

    if pd.notna(row["chemo_code2"]):
        result.append(f"chemo2: {str(row['chemo_code2']).strip()}")

    return result if result else pd.NA


def normalize_meta_code(x):
    """
    Normalizes a meta_chir code by converting numeric values to a clean string.
    Returns pd.NA for missing values.
    """
    if pd.isna(x):
        return pd.NA
    try:
        x_float = float(x)
        if x_float.is_integer():
            return str(int(x_float))
        else:
            return str(x).strip()
    except Exception:
        return str(x).strip()


def assign_p_sx_fields(row):
    """
    Assigns values to 'p_sx_type' and 'p_sx_mets_loc' based on meta_chir codes.
    Sets 'p_sx_type' to 3 and 'p_sx_mets_loc' to 9 if either code maps to 3.
    """
    code1_raw = row.get("meta_chir_code1")
    code2_raw = row.get("meta_chir_code2")

    code1 = meta_chir_code_map.get(str(code1_raw).strip()) if pd.notna(code1_raw) else None
    code2 = meta_chir_code_map.get(str(code2_raw).strip()) if pd.notna(code2_raw) else None

    if code1 is None and pd.notna(code1_raw):
        log_unknown("meta_chir_code_map", "meta_chir_code_map", code1_raw)
    if code2 is None and pd.notna(code2_raw):
        log_unknown("meta_chir_code_map", "meta_chir_code_map", code2_raw)

    if code1 == "metachir3" or code2 == "metachir3":
        return pd.Series({"p_sx_type": 3, "p_sx_mets_loc": 9})
    else:
        return pd.Series({"p_sx_type": pd.NA, "p_sx_mets_loc": pd.NA})


def make_unique_treat_types(treat_list):
    """
    Makes treatment codes unique by appending a numeric suffix to duplicates.
    Keeps the first occurrence unchanged.
    """
    seen = {}
    result = []

    for code in treat_list:
        count = seen.get(code, 0)
        new_code = f"{code}" if count == 0 else f"{code}_{count}"
        result.append(new_code)
        seen[code] = count + 1

    return result


def create_treat_sx_radical(row):
    """
    Creates a list of labeled radical surgery flags based on surgical fields.
    Returns a list like ['residu: 1', 'lokchir: 0'], or pd.NA if no values found.
    """
    result = []

    if pd.notna(row.get("lok_chir_rad1")):
        val = 1 if float(row["lok_chir_rad1"]) > 0 else 0
        result.append(f"lokchir: {val}")

    if pd.notna(row.get("meta_chir_rad1")):
        val = 1 if float(row["meta_chir_rad1"]) > 0 else 0
        result.append(f"metachir1: {val}")

    if pd.notna(row.get("meta_chir_rad2")):
        val = 1 if float(row["meta_chir_rad2"]) > 0 else 0
        result.append(f"metachir2: {val}")

    return result if result else pd.NA


def add_staging(row, code):
    """
    Adds a staging code to the 'malign_other_staging_system' list if not already present.
    Initializes the list if it doesn't exist.
    """
    if isinstance(row["malign_other_staging_system"], list):
        if code not in row["malign_other_staging_system"]:
            return row["malign_other_staging_system"] + [code]
        return row["malign_other_staging_system"]
    return [code]


def get_matching_texts(code_string, targets):
    """
    Helper function to extract additional mapped text for specific categories (e.g., 90 or 15)
    """
    codes = [c.strip() for c in code_string.split(",") if c.strip()]
    results = []
    for code in codes:
        cat = unique_values_dict.get(code)
        if cat is None:
            log_unknown("topography", "unique_values_dict", code)
        if cat in targets:
            txt = value_map.get(code, "")
            if txt == "":
                log_unknown("topography", "value_map", code)
            results.append(txt)
    return ", ".join([r for r in results if r])


def parse_list(val):
    """
    Converts a string like '[1, 2, 3]' to a list of ints.
    Returns [] if the value is invalid or empty.
    """
    if isinstance(val, str) and val.strip().startswith("["):
        try:
            return [int(x.strip()) for x in val.strip("[]").split(",")]
        except Exception:
            return []
    elif isinstance(val, list):
        return val
    elif pd.isna(val):
        return []
    else:
        try:
            return [int(val)]
        except Exception:
            return []


def parse_treatment_blocks(val):
    """
    Converts a string like "[(3, '2021-01-01', '2021-02-01'), ...]"
    into a list of (code, start, stop) tuples.
    """
    if isinstance(val, str) and val.strip().startswith("["):
        try:
            parts = val.strip("[]").split("),")
            parsed = []
            for part in parts:
                part = part.strip().strip("()")
                if not part:
                    continue
                fields = part.split(",")
                code = int(fields[0].strip())
                start = fields[1].strip().strip("' ")
                stop = fields[2].strip().strip("' ") if len(fields) > 2 else None
                stop = None if stop in ["None", ""] else stop
                parsed.append((code, start, stop))
            return parsed
        except Exception:
            return []
    elif isinstance(val, list):
        return val
    else:
        return []


def extract_sx_code(row):
    """
    Returns 'code1' or 'code2' from treat_sx_type_prim dict
    based on treat_type (orgchir1/orgchir2).
    """
    if row["treat_type"] == "orgchir1" and isinstance(row["treat_sx_type_prim"], dict):
        return row["treat_sx_type_prim"].get("code1", pd.NA)
    elif row["treat_type"] == "orgchir2" and isinstance(row["treat_sx_type_prim"], dict):
        return row["treat_sx_type_prim"].get("code2", pd.NA)
    else:
        return pd.NA


def extract_chemo_code(row):
    """
    Searches treat_ctx_reg for an entry starting with the current treat_type,
    e.g. 'chemo1: CAPTEM', and extracts the code after ':'.
    """
    if not isinstance(row["treat_ctx_reg"], list):
        return pd.NA

    for entry in row["treat_ctx_reg"]:
        if isinstance(entry, str) and entry.startswith(f"{row['treat_type']}:"):
            return entry.split(":", 1)[1].strip()
    return pd.NA


def expand_row(row):
    """
    Expands a single patient row into multiple rows,
    one per treatment block. If some treat_type values have no
    associated block, they are added with missing dates.
    """
    results = []
    blocks = row["treatment_blocks"]
    treat_types = row["treat_type"]
    used_codes = [b[0] for b in blocks]

    for code, start, stop in blocks:
        new_row = row.copy()
        new_row["treat_type"] = code
        new_row["treat_start"] = start
        new_row["treat_stop"] = stop
        results.append(new_row)

    for code in treat_types:
        if used_codes.count(code) >= treat_types.count(code):
            continue
        used_codes.append(code)
        new_row = row.copy()
        new_row["treat_type"] = code
        new_row["treat_start"] = pd.NA
        new_row["treat_stop"] = pd.NA
        results.append(new_row)

    return results


def extract_sx_radical(row):
    """
    Extracts the radicality value (0 or 1) from the treat_sx_radical list
    based on treat_type. Handles both regular and metachir entries.
    """
    val_list = row["treat_sx_radical"]
    if not isinstance(val_list, list):
        return pd.NA

    type_ = row["treat_type"]
    if isinstance(type_, str):
        if type_.startswith("metachir"):
            match = [v for v in val_list if isinstance(v, str) and v.startswith("metachir")]
        else:
            match = [v for v in val_list if isinstance(v, str) and v.startswith(f"{type_}:")]
        if match:
            return match[0].split(":")[1].strip()
    return pd.NA


def map_treat_type_code(val):
    """
    Maps treatment code labels (as string) to een numerieke code met treat_type_code_map.
    """
    if pd.isna(val):
        return pd.NA
    key = str(val).strip()
    if key in treat_type_code_map:
        return treat_type_code_map[key]
    log_unknown("treat_type_code_map", "treat_type_code_map", key)
    return val


# --- Load & Rename ---
df = pd.read_csv(path_to_iknl, delimiter=";")
df["id"] = df["id"].map(iknl_to_force)

df.rename(
    columns={
        "vit_stat": "vs",
        "vit_stat_dat": "vs_date",
        "incdat": "bl_date",
        "diag_basis": "dx_basis",
        "topo_sublok": "pb_loc1",
        "ct": "bl_tnm_t",
        "cn": "bl_tnm_n",
        "cm": "bl_tnm_m",
        "pt": "p_sx_tnm_t",
        "pn": "p_sx_tnm_n",
        "pm": "p_sx_tnm_m",
        "beeld_diag": "bl_incident",
        "beeld_afm": "bl_primary_size",
        "cga_lab": "bl_cga_unit",
        "ond_lymf": "p_sx_prim_ln_no",
        "pos_lymf": "p_sx_prim_ln_pos",
        "meta_dia": "bl_dm",
        "ki_67": "p_sx_prim_ki67",
        "mito_aant": "p_sx_prim_mit",
        "mal_topo_sublok1": "malign_other_topography",
        "mal_morf1": "malign_other_morphology",
        "mal_tumsoort1": "malign_other_type_code",
        "mal_gedrag1": "malign_other_behavior",
        "id": "force_id",
        "residu": "p_sx_prim_r_status",
    },
    inplace=True,
)

# --- Type Conversion ---
df["dx_basis"] = df["dx_basis"].astype("Int64")
df["bl_incident"] = df["bl_incident"].astype("Int64")
df["p_sx_prim_ln_no"] = df["p_sx_prim_ln_no"].astype("Int64")
df["p_sx_prim_ln_pos"] = df["p_sx_prim_ln_pos"].astype("Int64")
df["bl_dm"] = df["bl_dm"].astype("Int64")

# --- pb_loc1 Mapping ---
df["pb_loc1_raw"] = df["pb_loc1"]
df["pb_loc1"] = df["pb_loc1_raw"].apply(
    lambda x: map_with_log(
        x,
        topo_blok_bl,
        sheet_name="topography",
        mapping_name="topo_blok_bl",
        default=pd.NA,
        key_transform=lambda v: str(v).strip(),
    )
)

df["pb_loc_add1"] = df["pb_loc1_raw"].apply(
    lambda x: value_map.get(str(x).strip(), "")
    if topo_blok_bl.get(str(x).strip(), None) == 90
    else ""
)
df["pb_loc_lnn1"] = df["pb_loc1_raw"].apply(
    lambda x: value_map.get(str(x).strip(), "")
    if topo_blok_bl.get(str(x).strip(), None) == 15
    else ""
)

# --- Insert pb_loc1 fields next to pb_loc1 ---
cols = list(df.columns)
insert_at = cols.index("pb_loc1") + 1
for col in ["pb_loc_add1", "pb_loc_lnn1"]:
    if col in cols:
        cols.remove(col)
cols[insert_at:insert_at] = ["pb_loc_add1", "pb_loc_lnn1"]
df = df[cols]

# --- Diff/Grade Handling ---
diff_index = df.columns.get_loc("diffgrad")
df.insert(diff_index, "p_sx_diff", "")
df.insert(diff_index + 1, "p_sx_res_class", "")
df.loc[df["diffgrad"] == 1, "p_sx_diff"] = 1
df.loc[df["diffgrad"] == 2, "p_sx_diff"] = 2
df.loc[df["diffgrad"] == 9, "p_sx_diff"] = 99
df.loc[df["diffgrad"] == 3, "p_sx_res_class"] = 3
df.drop(columns="diffgrad", inplace=True)

# --- TNM Mapping ---
for base in ["bl", "p_sx"]:
    for part in ["t", "n", "m"]:
        col = f"{base}_tnm_{part}"
        df[f"{col}_raw"] = df[col]

for col, mapping in map_dict.items():
    suffix_col = f"{col}suffix"
    part = col.split("_")[-1]  # t/n/m
    sheet_name = f"tnm_{part}"
    mapping_name = f"tnm_{part}"
    df[[col, suffix_col]] = df[col + "_raw"].apply(
        lambda x: pd.Series(extract_value_and_suffix(x, mapping, sheet_name, mapping_name))
    )

cols = list(df.columns)
for col in map_dict.keys():
    suffix_col = f"{col}suffix"
    if suffix_col in cols:
        cols.remove(suffix_col)
    if col in cols:
        insert_at = cols.index(col) + 1
        cols.insert(insert_at, suffix_col)
df = df[cols]

# --- pb_loc (multi) Mapping ---
df["pb_loc_raw"] = df[
    [
        "meta_topo_sublok1",
        "meta_topo_sublok2",
        "meta_topo_sublok3",
        "meta_topo_sublok4",
        "meta_topo_sublok5",
        "meta_topo_sublok6",
    ]
].apply(lambda row: ", ".join(row.dropna().astype(str)), axis=1)

def map_meta_loc_string(x):
    parts = [p.strip() for p in str(x).split(",") if p.strip()]
    mapped = []
    for code in parts:
        val = unique_values_dict.get(code)
        if val is None:
            log_unknown("topography", "unique_values_dict", code)
            mapped.append(code)
        else:
            mapped.append(val)
    return ", ".join(mapped)

df["pb_loc"] = df["pb_loc_raw"].apply(map_meta_loc_string)

df["pb_loc_add"] = df["pb_loc_raw"].apply(lambda x: get_matching_texts(x, ["90"]))
df["pb_loc_lnn"] = df["pb_loc_raw"].apply(lambda x: get_matching_texts(x, ["15"]))

cols = list(df.columns)
for col in ["pb_loc", "pb_loc_add", "pb_loc_lnn"]:
    if col in cols:
        cols.remove(col)
insert_at = cols.index("bl_dm") + 1
cols[insert_at:insert_at] = ["pb_loc", "pb_loc_add", "pb_loc_lnn"]
df = df[cols]

# --- Hormonal Profile ---
df["bl_5_hiaa_u24"] = df.apply(
    lambda x: 1 if 1 in [x["horm_prod1"], x["horm_prod2"]] else "", axis=1
)
df["bl_ser"] = df.apply(
    lambda x: 1 if 2 in [x["horm_prod1"], x["horm_prod2"]] else "", axis=1
)
df["bl_gastrin"] = df.apply(
    lambda x: 1 if 3 in [x["horm_prod1"], x["horm_prod2"]] else "", axis=1
)

df = df.apply(check_text_conditions, axis=1)

cols = list(df.columns)
for col in ["bl_5_hiaa_u24", "bl_ser", "bl_gastrin"]:
    if col in cols:
        cols.remove(col)
insert_at = cols.index("horm_prod1") + 1
cols[insert_at:insert_at] = ["bl_5_hiaa_u24", "bl_ser", "bl_gastrin"]

cols.remove("horm_prod1")
cols.remove("horm_prod2")
cols.remove("tumor_diagnose_txt")
cols.remove("tumor_opmerkingen_txt")
df = df[cols]

# --- CGA IHC ---
insert_index = df.columns.get_loc("cga_pa")
df.insert(insert_index, "p_sx_ihc", pd.NA)
df.insert(insert_index + 1, "p_sx_ihc_cga", pd.NA)
df.loc[df["cga_pa"] == 1, ["p_sx_ihc", "p_sx_ihc_cga"]] = 1

# --- Date Processing ---
df["vs_date"] = pd.to_datetime(df["vs_date"], format="%d-%m-%Y", errors="coerce")
df.loc[df["vs"] == 1, "vs_d_date_month"] = df["vs_date"].dt.month.astype("Int64").astype(str)
df.loc[df["vs"] == 1, "vs_d_date_year"] = df["vs_date"].dt.year.astype("Int64").astype(str)
df["vs_d_date_month"] = df["vs_d_date_month"].fillna("")
df["vs_d_date_year"] = df["vs_d_date_year"].fillna("")

cols = list(df.columns)
for col in ["vs_d_date_year", "vs_d_date_month"]:
    if col in cols:
        cols.remove(col)
insert_at = cols.index("vs_date") + 1
cols[insert_at:insert_at] = ["vs_d_date_year", "vs_d_date_month"]
df = df[cols]

# --- Ki-67 / Mitoses Mapping & Extension ---
df["p_sx_prim_ki67"] = df["p_sx_prim_ki67"].apply(
    lambda x: map_with_log(
        x,
        ki67_map,
        sheet_name="ki67_map",
        mapping_name="ki67_map",
        default=pd.NA,
        key_transform=lambda v: int(float(v)),
    )
).astype("Int64")

df["p_sx_prim_mit"] = df["p_sx_prim_mit"].apply(
    lambda x: map_with_log(
        x,
        mitoses_map,
        sheet_name="mitoses_map",
        mapping_name="mitoses_map",
        default=pd.NA,
        key_transform=lambda v: int(float(v)),
    )
).astype("Int64")

df["p_sx_ext"] = (
    df[["p_sx_prim_ki67", "p_sx_prim_mit"]].notna().any(axis=1).astype("Int64")
)
df.loc[df[["p_sx_prim_ki67", "p_sx_prim_mit"]].isna().all(axis=1), "p_sx_ext"] = pd.NA

cols = list(df.columns)
mit_idx = cols.index("p_sx_prim_ki67")
cols.insert(mit_idx, "p_sx_ext")
cols = [col for i, col in enumerate(cols) if col != "p_sx_ext" or i == mit_idx]
df = df[cols]

# --- Initialize treatment columns if missing ---
if "treat_sx_radical" not in df.columns:
    df["treat_sx_radical"] = [[] for _ in range(len(df))]

if "treat_type" not in df.columns:
    df.insert(df.columns.get_loc("treat_sx_radical") + 1, "treat_type", [[] for _ in range(len(df))])

# --- Add treatments from AS and Geen Therapie flags ---
if "as" in df.columns:
    df["as"] = df["as"].dropna().astype(int)
    df["treat_type"] = df.apply(
        lambda row: add_to_treat_type(row, "as") if row["as"] == 1 else row["treat_type"], axis=1
    )
    df.drop(columns="as", inplace=True)

if "geen_ther" in df.columns:
    df["geen_ther"] = df["geen_ther"].dropna().astype(int)
    df["treat_type"] = df.apply(
        lambda row: add_to_treat_type(row, "geenther")
        if row["geen_ther"] == 1
        else row["treat_type"],
        axis=1,
    )
    df.drop(columns=["geen_ther", "geen_ther_reden", "chemo"], inplace=True)

# Clean chemo_code fields
for col in ["chemo_code1", "chemo_code2"]:
    df[col] = (
        df[col]
        .astype(str)
        .str.upper()
        .str.strip()
        .replace("NAN", pd.NA)
    )

df["treat_ctx_reg"] = df.apply(collect_labeled_chemo_codes, axis=1)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "chemo1") if pd.notna(row.get("chemo_code1")) else row["treat_type"],
    axis=1,
)
df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "chemo2") if pd.notna(row.get("chemo_code2")) else row["treat_type"],
    axis=1,
)

df.drop(columns=["chemo_code1", "chemo_code2"], inplace=True)

if "treat_ctx_reg" in df.columns:
    df.insert(df.columns.get_loc("treat_type") + 1, "treat_ctx_reg", df.pop("treat_ctx_reg"))

# --- Add treatments from hormonal, radiation, and surgery fields ---
df["horm"] = pd.to_numeric(df["horm"], errors="coerce").astype("Int64")
df["rt"] = pd.to_numeric(df["rt"], errors="coerce").astype("Int64")
df["lok_chir"] = pd.to_numeric(df["lok_chir"], errors="coerce").astype("Int64")

df["treat_type"] = df.apply(
    lambda row: row["treat_type"] + ["horm1"] if pd.notna(row.get("horm_startdat1")) else row["treat_type"],
    axis=1,
)
df["treat_type"] = df.apply(
    lambda row: row["treat_type"] + ["horm2"] if pd.notna(row.get("horm_startdat2")) else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "rt")
    if pd.notna(row.get("rt")) and row.get("rt") in [1, 2, 3, 4, 5]
    else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "lokchir")
    if pd.notna(row.get("lok_chir")) and row.get("lok_chir") in [1, 2, 3, 4, 5]
    else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "metart1")
    if pd.notna(row.get("meta_rt")) and row.get("meta_rt") == 1
    else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: row["treat_type"] + ["orgchir1"] if pd.notna(row.get("org_chir_code1")) else row["treat_type"],
    axis=1,
)
df["treat_type"] = df.apply(
    lambda row: row["treat_type"] + ["orgchir2"] if pd.notna(row.get("org_chir_code2")) else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "overigchir")
    if pd.notna(row.get("overig_chir")) and row.get("overig_chir") == 1
    else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: add_to_treat_type(row, "overigther")
    if pd.notna(row.get("overig_ther")) and row.get("overig_ther") == 1
    else row["treat_type"],
    axis=1,
)

if "treat_sx_prim" not in df.columns:
    df.insert(df.columns.get_loc("treat_type") + 1, "treat_sx_prim", pd.NA)

df["treat_sx_prim"] = df.apply(
    lambda row: 1
    if pd.notna(row.get("org_chir")) and str(row["org_chir"]).strip().startswith("1")
    else row["treat_sx_prim"],
    axis=1,
)

if "treat_type_add" not in df.columns:
    df.insert(df.columns.get_loc("treat_type") + 1, "treat_type_add", pd.NA)

df["treat_type_add"] = df.apply(
    lambda row: str(row["overig_chir_code1"]).strip()
    if pd.notna(row.get("overig_chir_code1"))
    and str(row["overig_chir_code1"]).strip()
    else row["treat_type_add"],
    axis=1,
)

if "treat_prrt_type" not in df.columns:
    df.insert(df.columns.get_loc("treat_type") + 1, "treat_prrt_type", pd.NA)

df["treat_prrt_type"] = df.apply(
    lambda row: 1
    if pd.notna(row.get("rt_code1"))
    and str(row["rt_code1"]).upper().startswith("V10XX")
    else row["treat_prrt_type"],
    axis=1,
)

df["treat_ldt_type"] = df["treat_type"].apply(
    lambda lst: 5 if isinstance(lst, list) and 10 in lst else pd.NA
)

# --- Meta-chirurgical treatment codes ---
df["meta_chir_code1"] = df["meta_chir_code1"].apply(normalize_meta_code)
df["meta_chir_code2"] = df["meta_chir_code2"].apply(normalize_meta_code)

df["treat_type"] = df.apply(
    lambda row: row["treat_type"]
    + [meta_chir_code_map[str(row["meta_chir_code1"]).strip()]]
    if pd.notna(row.get("meta_chir_code1"))
    and str(row["meta_chir_code1"]).strip() in meta_chir_code_map
    else row["treat_type"],
    axis=1,
)

df["treat_type"] = df.apply(
    lambda row: row["treat_type"]
    + [meta_chir_code_map[str(row["meta_chir_code2"]).strip()]]
    if pd.notna(row.get("meta_chir_code2"))
    and str(row["meta_chir_code2"]).strip() in meta_chir_code_map
    else row["treat_type"],
    axis=1,
)

df[["p_sx_type", "p_sx_mets_loc"]] = df.apply(assign_p_sx_fields, axis=1)

insert_loc = df.columns.get_loc("chemo_startdat1")
df["treatment_blocks"] = df.apply(extract_ordered_treatment_blocks, axis=1)

df["treat_type_unique"] = df["treat_type"].apply(
    lambda x: make_unique_treat_types(x) if isinstance(x, list) else []
)

df["treatment_blocks_named"] = df.apply(
    lambda row: list(zip(row["treat_type_unique"], row["treatment_blocks"]))
    if isinstance(row["treat_type_unique"], list)
    and isinstance(row["treatment_blocks"], list)
    else [],
    axis=1,
)

cols = df.columns.tolist()
for col in ["treat_type_unique", "treatment_blocks_named"]:
    if col in cols:
        cols.remove(col)
insert_at = cols.index("treatment_blocks") + 1
cols[insert_at:insert_at] = ["treat_type_unique", "treatment_blocks_named"]
df = df[cols]

cols = df.columns.tolist()
cols.remove("treatment_blocks")
cols[insert_loc:insert_loc] = ["treatment_blocks"]
df = df[cols]

df["lok_chir_rad1"] = df["lok_chir_rad1"].replace(r"^\s*$", pd.NA, regex=True)
df["meta_chir_rad1"] = df["meta_chir_rad1"].replace(r"^\s*$", pd.NA, regex=True)
df["meta_chir_rad2"] = df["meta_chir_rad2"].replace(r"^\s*$", pd.NA, regex=True)

df["treat_sx_radical"] = df.apply(create_treat_sx_radical, axis=1)

df["org_chir_code1_mapped"] = df["org_chir_code1"].apply(
    lambda x: map_with_log(
        x,
        surgery_map,
        sheet_name="surgery_map",
        mapping_name="surgery_map",
        default=pd.NA,
        key_transform=lambda v: str(v).strip() if pd.notna(v) else v,
    )
)

df["org_chir_code2_mapped"] = df["org_chir_code2"].apply(
    lambda x: map_with_log(
        x,
        surgery_map,
        sheet_name="surgery_map",
        mapping_name="surgery_map",
        default=pd.NA,
        key_transform=lambda v: str(v).strip() if pd.notna(v) else v,
    )
)


if "treat_sx_type_prim" in df.columns:
    df.drop(columns=["treat_sx_type_prim"], inplace=True)

insert_index = df.columns.get_loc("treat_sx_prim") + 1
df.insert(
    insert_index,
    "treat_sx_type_prim",
    df.apply(
        lambda row: {
            "code1": int(row["org_chir_code1_mapped"])
            if pd.notna(row["org_chir_code1_mapped"])
            else pd.NA,
            "code2": int(row["org_chir_code2_mapped"])
            if pd.notna(row["org_chir_code2_mapped"])
            else pd.NA,
        },
        axis=1,
    ),
)

df["mal_incdat1"] = pd.to_datetime(df["mal_incdat1"], format="%d-%m-%Y", errors="coerce")

df["malign_other_date_month"] = df["mal_incdat1"].dt.month.astype("Int64")
df["malign_other_date_year"] = df["mal_incdat1"].dt.year.astype("Int64")

df["malign_other_date_month"] = df["malign_other_date_month"].astype(str).replace(
    "<NA>", ""
)
df["malign_other_date_year"] = df["malign_other_date_year"].astype(str).replace(
    "<NA>", ""
)

df["malign_other_topography"] = df["malign_other_topography"].apply(
    lambda x: value_map.get(str(x).strip(), "")
)

df["malign_other_morphology"] = pd.to_numeric(
    df["malign_other_morphology"], errors="coerce"
)
df["malign_other_morphology"] = df["malign_other_morphology"].apply(
    lambda x: map_with_log(
        x,
        morphology_map,
        sheet_name="morphology_map",
        mapping_name="morphology_map",
        default="",
        key_transform=lambda v: int(float(v)),
    )
)

df["malign_other_type_code"] = pd.to_numeric(
    df["malign_other_type_code"], errors="coerce"
).astype("Int64")
df["malign_other_behavior"] = pd.to_numeric(
    df["malign_other_behavior"], errors="coerce"
).astype("Int64")
df["malign_other_behavior"] = df["malign_other_behavior"].apply(
    lambda x: map_with_log(
        x,
        behavior_map,
        sheet_name="behavior_map",
        mapping_name="behavior_map",
        default="",
        key_transform=lambda v: int(float(v)),
    )
)

df["malign_other_tnm_t"] = df["mal_pt1"].combine_first(df["mal_ct1"])
df["malign_other_tnm_n"] = df["mal_pn1"].combine_first(df["mal_cn1"])
df["malign_other_tnm_m"] = df["mal_pm1"].combine_first(df["mal_cm1"])

df[["malign_other_tnm_t", "malign_tnm_tsuffix"]] = df["malign_other_tnm_t"].apply(
    lambda x: pd.Series(
        extract_value_and_suffix(x, t, "tnm_t", "tnm_t")
    )
)
df[["malign_other_tnm_n", "malign_tnm_nsuffix"]] = df["malign_other_tnm_n"].apply(
    lambda x: pd.Series(
        extract_value_and_suffix(x, n, "tnm_n", "tnm_n")
    )
)
df[["malign_other_tnm_m", "malign_tnm_msuffix"]] = df["malign_other_tnm_m"].apply(
    lambda x: pd.Series(
        extract_value_and_suffix(x, m, "tnm_m", "tnm_m")
    )
)

df["malign_other_staging_system"] = (
    df[["malign_other_tnm_t", "malign_other_tnm_n", "malign_other_tnm_m"]]
    .notna()
    .any(axis=1)
    .astype("Int64")
)

cols = list(df.columns)
if "malign_other_staging_system" in cols and "malign_other_tnm_t" in cols:
    cols.remove("malign_other_staging_system")
    insert_at = cols.index("malign_other_tnm_t")
    cols.insert(insert_at, "malign_other_staging_system")
    df = df[cols]

df["mal_ann_arbor1,,"] = df["mal_ann_arbor1,,"].replace(",,", pd.NA)
df["mal_ann_arbor1,,"] = df["mal_ann_arbor1,,"].replace(",", pd.NA)

df.rename(
    columns={
        "mal_ceod1": "malign_other_eod",
        "mal_ann_arbor1,,": "malign_other_aa_stage",
    },
    inplace=True,
)

if "malign_other_staging_system" not in df.columns:
    df.insert(
        df.columns.get_loc("mal_t1"),
        "malign_other_staging_system",
        [[] for _ in range(len(df))],
    )

df["malign_other_staging_system"] = df.apply(
    lambda row: add_staging(row, 3)
    if pd.notna(row["malign_other_eod"])
    else row["malign_other_staging_system"],
    axis=1,
)
df["malign_other_staging_system"] = df.apply(
    lambda row: add_staging(row, 2)
    if pd.notna(row["malign_other_aa_stage"])
    else row["malign_other_staging_system"],
    axis=1,
)

df.loc[df["malign_other_staging_system"] == 0, "malign_other_staging_system"] = ""

df["p_sx_type"] = df["treat_type_unique"].apply(
    lambda types: 3
    if isinstance(types, list)
    and any("metachir3" in str(t_) for t_ in types)
    else pd.NA
)
df["p_sx_mets_loc"] = df["treat_type_unique"].apply(
    lambda types: 9
    if isinstance(types, list)
    and any("metachir3" in str(t_) for t_ in types)
    else pd.NA
)

df = df.drop(
    columns=[
        "chemo_startdat1",
        "chemo_startdat2",
        "chemo_stopdat1",
        "chemo_stopdat2",
        "horm_startdat1",
        "horm_startdat2",
        "horm_stopdat1",
        "horm_stopdat2",
        "horm_code1",
        "horm_code2",
        "horm",
        "rt_startdat1",
        "rt_stopdat1",
        "rt_code1",
        "rt",
        "lok_chir",
        "lok_chir_startdat1",
        "lok_chir_stopdat1",
        "lok_chir_code1",
        "meta_rt_startdat1",
        "meta_rt_stopdat1",
        "org_chir",
        "meta_rt",
        "meta_rt_code1",
        "org_chir_code1",
        "org_chir_code2",
        "org_chir_startdat1",
        "org_chir_startdat2",
        "org_chir_stopdat1",
        "org_chir_stopdat2",
        "org_chir_rad1",
        "org_chir_rad2",
        "overig_chir",
        "overig_chir_code1",
        "overig_chir_startdat1",
        "overig_chir_stopdat1",
        "overig_ther",
        "overig_ther_code1",
        "overig_ther_startdat1",
        "overig_ther_stopdat1",
        "mal_incdat1",
        "mal_stadium1",
        "mal_peod1",
        "mal_ct1",
        "mal_cn1",
        "mal_cm1",
        "mal_pt1",
        "mal_pn1",
        "mal_pm1",
        "nen",
        "key_nkr",
        "key_eid",
        "incjr",
        "later",
        "morf",
        "gedrag",
        "pb_loc1_raw",
        "cstadium",
        "pstadium",
        "stadium",
        "pb_loc_raw",
        "meta_topo_sublok1",
        "meta_topo_sublok2",
        "meta_topo_sublok3",
        "meta_topo_sublok4",
        "meta_topo_sublok5",
        "meta_topo_sublok6",
        "cga_pa",
        "org_chir_code1_mapped",
        "org_chir_code2_mapped",
        "lok_chir_rad1",
        "p_sx_tnm_t_raw",
        "p_sx_tnm_n_raw",
        "p_sx_tnm_m_raw",
        "bl_tnm_t_raw",
        "bl_tnm_n_raw",
        "bl_tnm_m_raw",
    ]
)

# --- FIX: TNM-velden en suffixen als integers wegschrijven ---
# --- FIX: TNM-velden en suffixen als nette integers ---
tnm_int_cols = [
    "bl_tnm_t", "bl_tnm_n", "bl_tnm_m",
    "bl_tnm_tsuffix", "bl_tnm_nsuffix", "bl_tnm_msuffix",
    "p_sx_tnm_t", "p_sx_tnm_n", "p_sx_tnm_m",
    "p_sx_tnm_tsuffix", "p_sx_tnm_nsuffix", "p_sx_tnm_msuffix",
    "malign_other_tnm_t", "malign_other_tnm_n", "malign_other_tnm_m",
    "malign_tnm_tsuffix", "malign_tnm_nsuffix", "malign_tnm_msuffix",
]

for col in tnm_int_cols:
    if col in df.columns:
        df[col] = df[col].apply(to_int_or_NA).astype("Int64")


# --- Prepare treatment file for export ---
selected_columns = [
    "force_id",
    "treat_type",
    "treat_prrt_type",
    "treat_type_add",
    "treat_sx_prim",
    "treat_sx_type_prim",
    "treat_ctx_reg",
    "treatment_blocks",
    "treat_sx_radical",
]
df1 = df[selected_columns].copy()

df1["treat_type"] = df1["treat_type"].apply(parse_list)
df1["treatment_blocks"] = df1["treatment_blocks"].apply(parse_treatment_blocks)

expanded_rows = []
for _, row in df1.iterrows():
    expanded_rows.extend(expand_row(row))
df_expanded = pd.DataFrame(expanded_rows)

df_expanded["treat_prrt_type"] = df_expanded.apply(
    lambda row: row["treat_prrt_type"] if row["treat_type"] == "rt" else pd.NA,
    axis=1,
)
df_expanded["treat_sx_type_prim"] = df_expanded.apply(extract_sx_code, axis=1)
df_expanded["treat_ctx_reg"] = df_expanded.apply(extract_chemo_code, axis=1)
df_expanded["treat_type_add"] = df_expanded.apply(
    lambda row: row["treat_type_add"] if row["treat_type"] == "overigchir" else pd.NA,
    axis=1,
)
df_expanded["treat_sx_radical"] = df_expanded.apply(extract_sx_radical, axis=1)

df_expanded["redcap_repeat_instrument"] = "treatment"
df_expanded["redcap_repeat_instance"] = df_expanded.groupby("force_id").cumcount() + 1

first_cols = ["force_id", "redcap_repeat_instrument", "redcap_repeat_instance"]
other_cols = [col for col in df_expanded.columns if col not in first_cols]
df_expanded = df_expanded[first_cols + other_cols]

df_expanded["treat_start"] = pd.to_datetime(
    df_expanded["treat_start"], dayfirst=True, errors="coerce"
).dt.strftime("%Y-%m-%d")
df_expanded["treat_stop"] = pd.to_datetime(
    df_expanded["treat_stop"], dayfirst=True, errors="coerce"
).dt.strftime("%Y-%m-%d")

df_expanded = df_expanded.drop(columns=["treatment_blocks"])
df_expanded["treat_ctx_reg"] = df_expanded["treat_ctx_reg"].apply(
    lambda x: map_with_log(
        x,
        chemo_map,
        sheet_name="chemo_map",
        mapping_name="chemo_map",
        default=pd.NA,
        key_transform=lambda v: str(v).strip().upper(),
    )
    if isinstance(x, str)
    else pd.NA
)
df_expanded["treat_type"] = df_expanded["treat_type"].apply(map_treat_type_code)

os.makedirs(path_for_export_files, exist_ok=True)

df_expanded.to_csv(
    os.path.join(path_for_export_files, "treat_iknl_import.csv"),
    sep=";",
    index=False,
    encoding="utf-8-sig",
)
print("Export to treat csv is completed")

# --- Create 'other malignancy' file for REDCap import ---
mal_filtered = df[df["mal_aantal"] > 0].copy()

mal_cols = [
    "force_id",
    "malign_other_topography",
    "malign_other_morphology",
    "malign_other_type_code",
    "malign_other_behavior",
    "malign_other_eod",
    "malign_other_aa_stage",
    "malign_other_date_month",
    "malign_other_date_year",
    "malign_other_staging_system",
    "malign_other_tnm_t",
    "malign_tnm_tsuffix",
    "malign_other_tnm_n",
    "malign_tnm_nsuffix",
    "malign_other_tnm_m",
    "malign_tnm_msuffix",
]

mal_df = mal_filtered[mal_cols].copy()
mal_df["redcap_repeat_instrument"] = "other_malignancy"
mal_df["redcap_repeat_instance"] = mal_df.groupby("force_id").cumcount() + 1

first_cols = ["force_id", "redcap_repeat_instrument", "redcap_repeat_instance"]
other_cols = [col for col in mal_df.columns if col not in first_cols]
mal_df = mal_df[first_cols + other_cols]

mal_output_path = os.path.join(path_for_export_files, "mal_iknl_import.csv")
mal_df["force_id"] = (
    mal_df["force_id"].astype(str).replace({r"\.0": ""}, regex=True).str.zfill(6)
)
mal_df.to_csv(mal_output_path, sep=";", index=False, encoding="utf-8-sig")
print("Export to malignancy csv is completed")

# --- Create baseline file for REDCap import ---
bl_columns = [col for col in df.columns if col.startswith("bl")]

df_bl = df[
    ["force_id", "vs", "vs_date", "vs_d_date_year", "vs_d_date_month", "dx_basis"]
    + bl_columns
].copy()

df_bl["bl_date"] = pd.to_datetime(
    df_bl["bl_date"], dayfirst=True, errors="coerce"
).dt.strftime("%Y-%m-%d")

df_bl["bl_cga_abs"] = df_bl["bl_cga_unit"].apply(
    lambda x: min(x, 99999) if pd.notna(x) else pd.NA
).astype("Int64")
df_bl["bl_cga_unit"] = df_bl["bl_cga_unit"].apply(
    lambda x: 1 if pd.notna(x) else pd.NA
).astype("Int64")

first_cols = ["force_id"]
other_cols = [col for col in df_bl.columns if col not in first_cols]
df_bl = df_bl[first_cols + other_cols]

bl_output_path = os.path.join(path_for_export_files, "bl_iknl_import.csv")
df_bl["force_id"] = (
    df_bl["force_id"].astype(str).replace({r"\.0": ""}, regex=True).str.zfill(6)
)
df_bl.to_csv(bl_output_path, sep=";", index=False, encoding="utf-8-sig")
print("Export to baseline csv is completed")

# --- Create pathology biopsy (pb) file for REDCap import ---
pb_columns = [col for col in df.columns if col.startswith("pb")]
df_pb = df[["force_id"] + pb_columns].copy()

df_pb["pb_loc_combined"] = df_pb[["pb_loc1", "pb_loc"]].apply(
    lambda row: ", ".join(
        sorted(
            set(
                filter(
                    None,
                    [str(x).strip() for x in row if pd.notna(x)],
                )
            )
        )
    ),
    axis=1,
)

expanded_rows = []
for _, row in df_pb.iterrows():
    pb_vals = [v.strip() for v in row["pb_loc_combined"].split(",") if v.strip()]
    for i, val in enumerate(pb_vals):
        new_row = {
            "force_id": row["force_id"],
            "redcap_repeat_instrument": "pathology_biopsy",
            "redcap_repeat_instance": i + 1,
            "pb_loc": val,
            "pb_loc_add": "",
            "pb_loc_lnn": "",
        }
        if val == "15":
            new_row["pb_loc_lnn"] = ", ".join(
                filter(
                    None,
                    [row.get("pb_loc_lnn1", ""), row.get("pb_loc_lnn", "")],
                )
            )
        if val == "90":
            new_row["pb_loc_add"] = ", ".join(
                filter(
                    None,
                    [row.get("pb_loc_add1", ""), row.get("pb_loc_add", "")],
                )
            )
        expanded_rows.append(new_row)

df_pb_final = pd.DataFrame(expanded_rows)
# --- FIX: pb_loc als integer-code (geen '9.0' maar 9) ---
if "pb_loc" in df_pb_final.columns:
    df_pb_final["pb_loc"] = df_pb_final["pb_loc"].apply(to_int_or_NA).astype("Int64")



pb_output_path = os.path.join(path_for_export_files, "pb_iknl_import.csv")
df_pb_final["force_id"] = (
    df_pb_final["force_id"].astype(str).replace({r"\.0": ""}, regex=True).str.zfill(6)
)
df_pb_final.to_csv(pb_output_path, sep=";", index=False, encoding="utf-8-sig")
print("Export to pathology biopsy csv is completed")

# --- Create surgical pathology (p_sx) file for REDCap import ---
sx_columns = [col for col in df.columns if col.startswith("p_sx")]
df_sx = df[["force_id"] + sx_columns].copy()

df_sx[sx_columns] = df_sx[sx_columns].replace(r"^\s*$", pd.NA, regex=True)
df_sx = df_sx.dropna(subset=sx_columns, how="all")

df_sx.insert(1, "redcap_repeat_instrument", "surgical_pathology")
df_sx["redcap_repeat_instance"] = df_sx.groupby("force_id").cumcount() + 1

df_sx["p_sx_prim_r_status"] = df_sx["p_sx_prim_r_status"].apply(
    lambda x: map_with_log(
        x,
        residual_map,
        sheet_name="residual_map",
        mapping_name="residual_map",
        default=pd.NA,
        key_transform=lambda v: str(v).strip(),
    )
).astype("Int64")

sx_output_path = os.path.join(path_for_export_files, "sx_iknl_import.csv")
df_sx["force_id"] = (
    df_sx["force_id"].astype(str).replace({r"\.0": ""}, regex=True).str.zfill(6)
)
df_sx.to_csv(sx_output_path, sep=";", index=False, encoding="utf-8-sig")
print("Export to surgical pathology csv is completed")

# --- Onbekende mappings exporteren ---
if unknown_mappings:
    unk_df = pd.DataFrame(unknown_mappings).drop_duplicates()
    unk_df = unk_df.sort_values(["sheet", "mapping", "key"])
    unk_path = os.path.join(path_for_export_files, "unknown_mappings.csv")
    unk_df.to_csv(unk_path, sep=";", index=False, encoding="utf-8-sig")
    print(f"Onbekende mapping-waarden gelogd in: {unk_path}")
else:
    print("Geen onbekende mapping-waarden gevonden.")
