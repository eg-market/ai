#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Standalone program to cleanse and merge Excel files by standardizing 'الفئة' using 'الفئة.txt'.

Key fixes and features:
- Loads mapping file (canonical=alias1=alias2...) and maps aliases to canonical.
- Smart matching: exact alias, substring overlap, token matches.
- Auto-detects category from product English/Arabic names before prompting.
- Prompts interactively for unknown categories, showing both product names and fuzzy suggestions.
- Caches user decisions during a run so each unknown alias is asked only once.
- Persists confirmed mappings back to الفئة.txt.
- Appends the canonical category to العبارة الترويجية as "الفئة <canonical>" when changed.
- Logs all changes to final/category_changes_log.csv; uses a fallback log file if locked.
- Handles duplicates (type 1 exact; type 2 composite key, keep newest by تاريخ تعديل ملف الفلاير).
- Writes per-input-file output (inputfilename_with_new_categories.xlsx) and a merged output.
- Skips temporary/lock files (e.g., files starting with "~$") and non-xlsx files.
- Uses openpyxl engine explicitly for reading/writing .xlsx files to avoid ambiguous-format errors.
"""

import os
import sys
import glob
import difflib
import time
import csv
from pathlib import Path
import pandas as pd

# -------------------------
# Configuration
# -------------------------
FOLDER_PATH = r"C:\Users\t\Downloads\docstrange-main\docstrange-main\carr_merge"  # adjust if needed
FINAL_FOLDER = os.path.join(FOLDER_PATH, "final")
CATEGORY_FILE = os.path.join(FOLDER_PATH, "الفئة.txt")
EXCEL_PATTERN = os.path.join(FOLDER_PATH, "*.xlsx")

MERGED_OUTPUT = os.path.join(FINAL_FOLDER, "merged_categories_output.xlsx")
LOG_FILE = os.path.join(FINAL_FOLDER, "category_changes_log.csv")

CATEGORY_COL = "الفئة"
ENGLISH_COL = "اسم المنتج (الأساسي)"
ARABIC_COL = "الاسم بالعربية"
PROMO_COL = "العبارة الترويجية"
DATE_COL = "تاريخ تعديل ملف الفلاير"

KEY_COLS = [
    "اسم السوبر ماركت",
    "الفئة",
    "اسم العلامة التجارية",
    "اسم المنتج (الأساسي)",
    "الاسم بالعربية",
    "الوحدة"
]

FUZZY_CUTOFF = 0.55
MAX_SUGGESTIONS = 7

# Ensure final folder exists
os.makedirs(FINAL_FOLDER, exist_ok=True)

# -------------------------
# Load mapping
# -------------------------
alias_to_category = {}   # maps alias.lower() -> canonical
canonical_set = set()

def load_category_mapping(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split("=") if p.strip()]
                if not parts:
                    continue
                canonical = parts[0]
                canonical_set.add(canonical)
                alias_to_category[canonical.lower()] = canonical
                for alias in parts[1:]:
                    alias_to_category[alias.lower()] = canonical
    else:
        open(path, "a", encoding="utf-8").close()

load_category_mapping(CATEGORY_FILE)
canonical_list_sorted = sorted(list(canonical_set))

# In-memory cache for new mappings confirmed during this run
session_new_mappings = {}  # alias_lower -> canonical

def persist_mapping(canonical, alias):
    """
    Persist a mapping line canonical=alias to CATEGORY_FILE and update in-memory maps.
    """
    try:
        with open(CATEGORY_FILE, "a", encoding="utf-8") as mf:
            mf.write(f"{canonical}={alias}\n")
    except Exception as e:
        print("Warning: failed to persist mapping to file:", e)
    alias_to_category[alias.lower()] = canonical
    session_new_mappings[alias.lower()] = canonical
    canonical_set.add(canonical)
    if canonical not in canonical_list_sorted:
        canonical_list_sorted.append(canonical)

# -------------------------
# Helpers
# -------------------------
def normalize_blank_to_empty(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in {"nan", "none", "null", "-"}:
        return ""
    return s

def smart_match_category(val):
    """
    Try to match val to a canonical category using:
      1) exact alias match
      2) substring overlap (val in alias or alias in val)
      3) token match (token equals alias)
    Returns canonical or None.
    """
    s = normalize_blank_to_empty(val)
    if not s:
        return None
    s_lower = s.lower()

    # 1) exact alias
    if s_lower in alias_to_category:
        return alias_to_category[s_lower]

    # 2) substring overlap (prefer longer alias matches)
    # Build candidates sorted by alias length descending to prefer more specific matches
    candidates = sorted(alias_to_category.items(), key=lambda x: -len(x[0]))
    for alias_lower, canonical in candidates:
        if alias_lower in s_lower or s_lower in alias_lower:
            return canonical

    # 3) token match
    tokens = s_lower.split()
    for t in tokens:
        if t in alias_to_category:
            return alias_to_category[t]

    return None

def auto_detect_category(eng, ar):
    """
    Tokenize English and Arabic product names and try to match tokens to aliases.
    Returns canonical or None.
    """
    tokens = []
    if eng:
        tokens += eng.lower().split()
    if ar:
        tokens += ar.lower().split()
    # prefer longer tokens first
    tokens = sorted(set(tokens), key=lambda x: -len(x))
    for t in tokens:
        if t in alias_to_category:
            return alias_to_category[t]
    # also try substring match between tokens and aliases
    for t in tokens:
        for alias_lower, canonical in alias_to_category.items():
            if t in alias_lower or alias_lower in t:
                return canonical
    return None

def fuzzy_suggestions(name, canonical_list):
    matches = difflib.get_close_matches(name, canonical_list, n=MAX_SUGGESTIONS, cutoff=FUZZY_CUTOFF)
    return [(m, round(difflib.SequenceMatcher(None, name.lower(), m.lower()).ratio(), 3)) for m in matches]

def prompt_unknown_category(name, eng, ar, canonical_list):
    """
    Interactive prompt. Returns (chosen_canonical_or_None, save_flag_bool).
    """
    print("\nUnknown category:", name)
    print("English product name:", eng)
    print("Arabic product name:", ar)
    suggestions = fuzzy_suggestions(name, canonical_list)
    if suggestions:
        print("Suggestions:")
        for i, (s, score) in enumerate(suggestions, 1):
            print(f"  {i}. {s} (score={score})")
    else:
        print("No close suggestions found.")
    while True:
        choice = input("Enter suggestion number, type canonical, 's' skip, 'q' quit: ").strip()
        if choice.lower() == "q":
            print("Quitting as requested.")
            sys.exit(0)
        if choice.lower() == "s":
            return None, False
        if suggestions and choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(suggestions):
                selected = suggestions[idx - 1][0]
                save = input(f"Save mapping '{name}' -> '{selected}'? (y/n): ").strip().lower() == "y"
                return selected, save
            else:
                print("Number out of range. Try again.")
                continue
        typed = choice.strip()
        if not typed:
            print("Empty input. Try again.")
            continue
        save = input(f"Save mapping '{name}' -> '{typed}'? (y/n): ").strip().lower() == "y"
        return typed, save

def append_category_to_promo(existing_promo, new_category):
    existing = normalize_blank_to_empty(existing_promo)
    parts = []
    if existing:
        parts.append(existing)
    parts.append("الفئة")
    parts.append(new_category)
    return " ".join(parts).strip()

def ensure_log_header(path):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        try:
            with open(path, "w", newline="", encoding="utf-8") as lf:
                writer = csv.writer(lf)
                writer.writerow([
                    "source_file",
                    "row_index",
                    "original_category",
                    "new_category",
                    "english_name",
                    "arabic_name",
                    "action"
                ])
        except Exception:
            # If cannot create main log, we'll rely on fallback when logging
            pass

def log_change(source_file, idx, orig_cat, new_cat, eng, ar, action):
    """
    Append a log row. If main log is locked, write to a timestamped fallback log.
    """
    row = [source_file, idx, orig_cat, new_cat, eng, ar, action]
    try:
        with open(LOG_FILE, "a", newline="", encoding="utf-8") as lf:
            writer = csv.writer(lf)
            writer.writerow(row)
    except PermissionError:
        ts = time.strftime("%Y%m%d_%H%M%S")
        fallback = Path(LOG_FILE).with_name(f"category_changes_log_fallback_{ts}.csv")
        try:
            with open(fallback, "a", newline="", encoding="utf-8") as lf:
                writer = csv.writer(lf)
                writer.writerow(row)
            print("Log file locked. Wrote to fallback:", str(fallback))
        except Exception as e:
            print("Failed to write fallback log:", e)
    except Exception as e:
        print("Failed to write log:", e)

def safe_to_excel(df, path):
    path = Path(path)
    try:
        df.to_excel(path, index=False, engine="openpyxl")
        print("Saved:", str(path))
    except PermissionError:
        ts = time.strftime("%Y%m%d_%H%M%S")
        fallback = path.with_name(path.stem + f"_fallback_{ts}" + path.suffix)
        df.to_excel(fallback, index=False, engine="openpyxl")
        print("Output file locked. Wrote fallback file:", str(fallback))
    except Exception as e:
        print("Failed to save Excel:", e)

# -------------------------
# Main processing
# -------------------------
excel_files = sorted(glob.glob(EXCEL_PATTERN))
# Filter out temporary/lock files (Excel creates files starting with "~$")
excel_files = [f for f in excel_files if not os.path.basename(f).startswith("~$")]

if not excel_files:
    print("No Excel files found in", FOLDER_PATH)
    sys.exit(0)

dfs = []
duplicates_type1 = []
duplicates_type2 = []

ensure_log_header(LOG_FILE)

for file in excel_files:
    try:
        # Explicitly use openpyxl engine to avoid ambiguous-format errors
        df = pd.read_excel(file, dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"Failed to read {file}: {e}")
        continue

    df.columns = [c.strip() for c in df.columns]

    # Ensure required columns exist
    for col in [CATEGORY_COL, ENGLISH_COL, ARABIC_COL, PROMO_COL]:
        if col not in df.columns:
            df[col] = ""

    # Normalize
    df[CATEGORY_COL] = df[CATEGORY_COL].apply(normalize_blank_to_empty)
    df[ENGLISH_COL] = df[ENGLISH_COL].apply(normalize_blank_to_empty)
    df[ARABIC_COL] = df[ARABIC_COL].apply(normalize_blank_to_empty)
    df[PROMO_COL] = df[PROMO_COL].apply(normalize_blank_to_empty)

    # Process rows
    for idx in df.index:
        source_file = os.path.basename(file)
        orig_cat = normalize_blank_to_empty(df.at[idx, CATEGORY_COL])
        eng = df.at[idx, ENGLISH_COL]
        ar = df.at[idx, ARABIC_COL]

        # If this alias was resolved earlier in this run, reuse it without prompting
        if orig_cat and orig_cat.lower() in session_new_mappings:
            mapped = session_new_mappings[orig_cat.lower()]
            if mapped != orig_cat:
                df.at[idx, CATEGORY_COL] = mapped
                df.at[idx, PROMO_COL] = append_category_to_promo(df.at[idx, PROMO_COL], mapped)
                log_change(source_file, idx, orig_cat, mapped, eng, ar, "session_cached")
            continue

        # 1) smart match (exact, substring, token)
        mapped = smart_match_category(orig_cat)
        if mapped:
            if mapped != orig_cat:
                df.at[idx, CATEGORY_COL] = mapped
                df.at[idx, PROMO_COL] = append_category_to_promo(df.at[idx, PROMO_COL], mapped)
                log_change(source_file, idx, orig_cat, mapped, eng, ar, "auto_alias_or_substring")
            continue

        # 2) auto-detect from product names
        auto_cat = auto_detect_category(eng, ar)
        if auto_cat:
            df.at[idx, CATEGORY_COL] = auto_cat
            df.at[idx, PROMO_COL] = append_category_to_promo(df.at[idx, PROMO_COL], auto_cat)
            log_change(source_file, idx, orig_cat, auto_cat, eng, ar, "auto_from_product")
            continue

        # 3) interactive prompt (only once per unique unknown alias)
        # If orig_cat is empty, we still prompt but show "(blank)" as context
        display_alias = orig_cat if orig_cat else "(blank)"
        # If user already answered for this alias earlier in the run, session_new_mappings covers it.
        chosen, save_flag = prompt_unknown_category(display_alias, eng, ar, canonical_list_sorted)
        if chosen is None:
            # user skipped
            df.at[idx, CATEGORY_COL] = orig_cat
            log_change(source_file, idx, orig_cat, orig_cat, eng, ar, "skipped")
        else:
            df.at[idx, CATEGORY_COL] = chosen
            df.at[idx, PROMO_COL] = append_category_to_promo(df.at[idx, PROMO_COL], chosen)
            log_change(source_file, idx, orig_cat, chosen, eng, ar, "manual")
            # cache in session so we don't ask again
            if display_alias and display_alias != "(blank)":
                session_new_mappings[display_alias.lower()] = chosen
            # persist mapping if requested
            if save_flag:
                alias_value = orig_cat if orig_cat else chosen
                persist_mapping(chosen, alias_value)

    # Capture exact duplicates (type 1)
    dup_rows = df[df.duplicated(keep=False)]
    if not dup_rows.empty:
        dup_rows = dup_rows.copy()
        dup_rows.loc[:, "source_file"] = os.path.basename(file)
        duplicates_type1.append(dup_rows)

    # Remove exact duplicates
    df = df.drop_duplicates()

    # Duplicate type 2: keep newest by DATE_COL per composite key
    if all(col in df.columns for col in KEY_COLS) and DATE_COL in df.columns:
        try:
            df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
        except Exception:
            pass
        df = df.sort_values(by=DATE_COL)
        dup2_rows = df[df.duplicated(subset=KEY_COLS, keep="last")]
        if not dup2_rows.empty:
            dup2_rows = dup2_rows.copy()
            dup2_rows.loc[:, "source_file"] = os.path.basename(file)
            duplicates_type2.append(dup2_rows)
        df = df.drop_duplicates(subset=KEY_COLS, keep="last")

    df = df.fillna("")
    df["source_file"] = os.path.basename(file)
    dfs.append(df)

    # Save per-input-file output with new categories
    try:
        input_name = Path(file).stem
        out_path = os.path.join(FINAL_FOLDER, f"{input_name}_with_new_categories.xlsx")
        df.to_excel(out_path, index=False, engine="openpyxl")
        print("Wrote updated file:", out_path)
    except Exception as e:
        print("Failed to write per-file output for", file, ":", e)

# -------------------------
# Save merged output and duplicates
# -------------------------
if dfs:
    merged_df = pd.concat(dfs, ignore_index=True)
    safe_to_excel(merged_df, MERGED_OUTPUT)
else:
    merged_df = pd.DataFrame()
    print("No data frames to merge.")

if duplicates_type1:
    try:
        dup1_df = pd.concat(duplicates_type1, ignore_index=True)
        dup1_path = os.path.join(FINAL_FOLDER, "duplicates_type1_categories.xlsx")
        dup1_df.to_excel(dup1_path, index=False, engine="openpyxl")
        print("Duplicates type 1 saved at:", dup1_path)
    except Exception as e:
        print("Failed to save duplicates type1:", e)

if duplicates_type2:
    try:
        dup2_df = pd.concat(duplicates_type2, ignore_index=True)
        dup2_path = os.path.join(FINAL_FOLDER, "duplicates_type2_categories.xlsx")
        dup2_df.to_excel(dup2_path, index=False, engine="openpyxl")
        print("Duplicates type 2 saved at:", dup2_path)
    except Exception as e:
        print("Failed to save duplicates type2:", e)

print("Log file location (primary):", LOG_FILE)
print("All done.")
