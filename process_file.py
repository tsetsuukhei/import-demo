"""
Import File Processor
=====================
Reads pre-built CSV dictionaries (from build_dict.py) and processes
a new raw import file into categorized output sheets.

Usage:
    python process_file.py                      # Process d13.xlsx → d13_cleaned.xlsx
    python process_file.py d14.xlsx             # Process d14.xlsx → d14_cleaned.xlsx
    python process_file.py d14.xlsx output.xlsx # Custom output name

Prerequisites:
    Run `python build_dict.py` first to generate the CSV dictionaries in dicts/.

Output sheets:
    - One sheet per business (PG, IPEK, OZDEN, Old Spice, ETI, Nestle, etc.)
    - 'Deleted': rows matching the deleted dictionary (HS-relevant but not categorized)
    - 'HS Deleted': rows removed by HS code filtering (irrelevant product categories)
    - 'Uncertain': rows that passed HS filter but don't match any dictionary
"""

import pandas as pd
import sys
import os
from rapidfuzz import process, fuzz

# =========================================================
# CONFIGURATION
# =========================================================

DICT_DIR = 'dicts'
DEFAULT_INPUT = 'd13.xlsx'
THRESHOLD = 92                      # Fuzzy match score threshold (business)
DEL_THRESHOLD = 95                  # Fuzzy match score threshold (deleted dict)

# Raw column → organized column name mapping
COL_MAP = {
    'илгээгч': 'Илгээгч',
    'горим': 'Горим',
    'огноо': 'Огноо',
    'илгээгчулс': 'Илгээгч улс',
    'гаралулс': 'Гарал.улс',
    'регистр': 'Регистр',
    'ААНБнэр': 'Нэр',
    'бараакод': 'Бараа код',
    'бараанэр': 'Бараа нэр',
    'марк': 'Марк',
    'зориулалт': 'Зориулалт',
    'тоо': 'Тоо',
    'нэгж': 'Нэгж',
    'үнэдолл': 'Үнэ.$',
}

MATCH_COL_MAP = {
    'category': 'Categories',
    'sub_cat': 'sub cat',
    'brand': 'Brand',
    'sub_brand': 'Sub brand',
    'company_name': 'Company name',
}

OUTPUT_COLS = [
    'Илгээгч', 'Горим', 'Огноо', 'Илгээгч улс', 'Гарал.улс',
    'Регистр', 'Нэр', 'Company name', 'Бараа код', 'Бараа нэр',
    'Categories', 'sub cat', 'Марк', 'Brand', 'Sub brand',
    'Зориулалт', 'Тоо', 'Нэгж', 'Үнэ.$',
    'match_type', 'score',
]


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def norm(val):
    """Normalize text: lowercase, strip, handle NaN."""
    s = str(val).strip().lower()
    return '' if s in ('nan', 'none', 'nat') else s


def norm_reg(val):
    """Normalize registry number: strip .0 float suffix, lowercase."""
    s = str(val).strip()
    if s.lower() in ('nan', 'none', 'nat', ''):
        return ''
    if s.count('.') == 1:
        try:
            f = float(s)
            if f == int(f):
                s = str(int(f))
        except (ValueError, OverflowError):
            pass
    return s.lower()


def norm_code(val):
    """Normalize barcode/HS code to a plain number string."""
    s = str(val).strip()
    if s.lower() in ('nan', 'none', ''):
        return ''
    if s.count('.') == 1:
        try:
            f = float(s)
            if f == int(f):
                s = str(int(f))
        except (ValueError, OverflowError):
            pass
    return s.replace('.', '').replace(' ', '').lower()


# =========================================================
# 1. LOAD DICTIONARIES FROM CSV
# =========================================================

def load_business_dict(dict_dir):
    """
    Load business dictionary CSV and reconstruct lookup structures.

    Returns:
        biz_lookup:   {(reg, code): [{full_text, category, sub_cat, brand, sub_brand, company_name, business}]}
        biz_registry: {reg: [(full_text, (reg,code), business)]}
    """
    biz_path = os.path.join(dict_dir, 'business_dict.csv')
    df = pd.read_csv(biz_path, encoding='utf-8-sig', dtype=str, keep_default_na=False)

    biz_lookup = {}
    biz_registry = {}

    for _, row in df.iterrows():
        key = (row['reg'], row['code'])
        info = {
            'full_text': row['full_text'],
            'category': row['category'],
            'sub_cat': row['sub_cat'],
            'brand': row['brand'],
            'sub_brand': row['sub_brand'],
            'company_name': row['company_name'],
            'business': row['business'],
        }
        biz_lookup.setdefault(key, []).append(info)
        biz_registry.setdefault(row['reg'], []).append(
            (row['full_text'], key, row['business'])
        )

    # Deduplicate registry
    for reg in biz_registry:
        biz_registry[reg] = list(set(biz_registry[reg]))

    print(f"  Business dict: {len(df):,} entries, {len(biz_lookup):,} keys, {len(biz_registry):,} registries")
    return biz_lookup, biz_registry


def load_deleted_dict(dict_dir):
    """
    Load deleted dictionary CSV and reconstruct lookup structures.

    Returns:
        del_lookup:   {(reg, code): [{full_text}]}
        del_registry: {reg: [(full_text, (reg,code))]}
    """
    del_path = os.path.join(dict_dir, 'deleted_dict.csv')
    df = pd.read_csv(del_path, encoding='utf-8-sig', dtype=str, keep_default_na=False)

    del_lookup = {}
    del_registry = {}

    for _, row in df.iterrows():
        key = (row['reg'], row['code'])
        del_lookup.setdefault(key, []).append({'full_text': row['full_text']})
        del_registry.setdefault(row['reg'], []).append((row['full_text'], key))

    # Deduplicate registry
    for reg in del_registry:
        seen = set()
        deduped = []
        for entry in del_registry[reg]:
            if entry[0] not in seen:
                seen.add(entry[0])
                deduped.append(entry)
        del_registry[reg] = deduped

    print(f"  Deleted dict: {len(df):,} entries, {len(del_lookup):,} keys")
    return del_lookup, del_registry


def load_hs_prefixes(dict_dir):
    """Load HS prefixes from CSV."""
    hs_path = os.path.join(dict_dir, 'hs_prefixes.csv')
    df = pd.read_csv(hs_path, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    prefixes = df['prefix'].tolist()
    print(f"  HS prefixes: {len(prefixes)}")
    return prefixes


def load_sheet_names(dict_dir):
    """Load business sheet name order from CSV."""
    path = os.path.join(dict_dir, 'sheet_names.csv')
    df = pd.read_csv(path, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    return df['sheet'].tolist()


# =========================================================
# 2. HS CODE FILTERING
# =========================================================

def filter_by_hs(df, code_col, prefixes):
    """Split DataFrame into (kept, removed) based on HS code prefixes."""
    hs_clean = df[code_col].astype(str).str.replace('.', '', regex=False).str.strip()
    mask = pd.Series(False, index=df.index)
    for prefix in prefixes:
        mask |= hs_clean.str.startswith(prefix)
    return df[mask].copy(), df[~mask].copy()


# =========================================================
# 3. LOAD RAW DATA
# =========================================================

def load_raw(input_file):
    """Load and normalize new raw import data from all sheets."""
    frames = []
    xls = pd.ExcelFile(input_file)
    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)
        df['_source'] = sheet
        frames.append(df)

    raw = pd.concat(frames, ignore_index=True)
    raw['бараакод'] = raw['бараакод'].apply(norm_code)

    print(f"  Loaded {len(raw):,} rows from {input_file} ({len(xls.sheet_names)} sheets)")
    return raw


# =========================================================
# 4. MATCHING
# =========================================================

def _empty_result(business, match_type, score=None):
    return {
        'business': business,
        'category': '', 'sub_cat': '', 'brand': '',
        'sub_brand': '', 'company_name': '',
        'match_type': match_type, 'score': score,
    }


def _prefer_brand_match(candidates):
    """
    Among candidates with the same full_text, pick the best business:
      1. Prefer the one whose business name appears in the brand (ETI → 'Eti AG')
      2. Prefer the one with a real category (not 'Not', '', '0')
      3. Fall back to first candidate
    """
    # Priority 1: brand contains business name
    for c in candidates:
        biz = c.get('business', '').lower()
        brand = c.get('brand', '').lower()
        if biz and brand and biz in brand:
            return c
    # Priority 2: real category (not 'Not', '', '0')
    for c in candidates:
        cat = c.get('category', '').strip()
        if cat and cat not in ('Not', '0'):
            return c
    return candidates[0]


def match_row(row, biz_lookup, biz_registry, del_lookup, del_registry, threshold, del_threshold):
    """
    Match a raw row against business and deleted dictionaries.

    Priority:
      1. Business: exact (reg+code) key, single → always match
      2. Business: exact (reg+code) key, multi → fuzzy text >= threshold
      3. Business: same registry → fuzzy text >= threshold
      4. Deleted: exact (reg+code) key → fuzzy text >= del_threshold
      5. No match → Uncertain
    """
    reg = norm_reg(row['регистр'])
    code = norm_code(row['бараакод'])
    key = (reg, code)

    if not reg or not code:
        return _empty_result('Uncertain', 'missing_key')

    search_text = f"{norm(row['марк'])} {norm(row['зориулалт'])}".strip()

    # --- BUSINESS: exact (reg, code) key match ---
    if key in biz_lookup:
        candidates = biz_lookup[key]

        if len(candidates) == 1:
            score = fuzz.WRatio(search_text, candidates[0]['full_text'])
            return {k: v for k, v in candidates[0].items() if k != 'full_text'} | {
                'match_type': 'business_exact', 'score': score,
            }

        texts = [c['full_text'] for c in candidates]
        result = process.extractOne(search_text, texts, scorer=fuzz.WRatio)
        if result and result[1] >= threshold:
            same_text = [c for c in candidates if c['full_text'] == result[0]]
            matched = _prefer_brand_match(same_text)
            return {k: v for k, v in matched.items() if k != 'full_text'} | {
                'match_type': 'business_code_fuzzy', 'score': result[1],
            }

    # --- BUSINESS: registry-level fuzzy ---
    if reg in biz_registry:
        choices = biz_registry[reg]
        texts = [x[0] for x in choices]
        if texts:
            result = process.extractOne(search_text, texts, scorer=fuzz.WRatio)
            if result and result[1] >= threshold:
                same_text = [(t, k, b) for (t, k, b) in choices if t == result[0]]
                # Prefer entry whose business name matches the brand
                best_entry = same_text[0]
                for entry in same_text:
                    products = biz_lookup[entry[1]]
                    for p in products:
                        if p['business'].lower() in p.get('brand', '').lower():
                            best_entry = entry
                            break
                product = next(
                    (p for p in biz_lookup[best_entry[1]]
                     if p['business'].lower() in p.get('brand', '').lower()),
                    biz_lookup[best_entry[1]][0]
                )
                return {k: v for k, v in product.items() if k != 'full_text'} | {
                    'match_type': 'business_registry_fuzzy', 'score': result[1],
                }

    # --- DELETED: exact (reg, code) key match ---
    if key in del_lookup:
        candidates = del_lookup[key]
        texts = [c['full_text'] for c in candidates]

        if len(candidates) == 1:
            score = fuzz.WRatio(search_text, texts[0])
            if score >= del_threshold:
                return _empty_result('Deleted', 'deleted_exact', score)
        elif texts:
            result = process.extractOne(search_text, texts, scorer=fuzz.WRatio)
            if result and result[1] >= del_threshold:
                return _empty_result('Deleted', 'deleted_code_fuzzy', result[1])

    # --- NO MATCH ---
    return _empty_result('Uncertain', 'no_match')


# =========================================================
# 5. FORMAT OUTPUT
# =========================================================

def format_output(df):
    """Rename raw columns to organized format and reorder."""
    result = df.rename(columns={**COL_MAP, **MATCH_COL_MAP})
    existing = [c for c in OUTPUT_COLS if c in result.columns]
    extras = [c for c in result.columns
              if c not in OUTPUT_COLS and not c.startswith('_') and c != 'business']
    return result[existing + extras]


# =========================================================
# MAIN
# =========================================================

def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.xlsx', '_cleaned.xlsx')

    # Check dict files exist
    required_dicts = ['business_dict.csv', 'deleted_dict.csv', 'hs_prefixes.csv', 'sheet_names.csv']
    for f in required_dicts:
        path = os.path.join(DICT_DIR, f)
        if not os.path.exists(path):
            print(f"ERROR: Dictionary file not found: {path}")
            print(f"Run `python build_dict.py` first to generate dictionaries.")
            sys.exit(1)

    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        sys.exit(1)

    print("=" * 60)
    print("IMPORT FILE PROCESSOR")
    print("=" * 60)
    print(f"  Dicts:  {DICT_DIR}/")
    print(f"  Input:  {input_file}")
    print(f"  Output: {output_file}")

    # Step 1: Load dictionaries
    print(f"\n[1/4] Loading dictionaries...")
    biz_lookup, biz_registry = load_business_dict(DICT_DIR)
    del_lookup, del_registry = load_deleted_dict(DICT_DIR)
    hs_prefixes = load_hs_prefixes(DICT_DIR)
    biz_sheets = load_sheet_names(DICT_DIR)
    print(f"  Sheets: {biz_sheets}")

    # Step 2: Load raw data
    print(f"\n[2/4] Loading input...")
    raw = load_raw(input_file)

    # Step 3: HS code filtering
    print(f"\n[3/4] Filtering by HS code...")
    hs_kept, hs_deleted = filter_by_hs(raw, 'бараакод', hs_prefixes)
    print(f"  HS kept: {len(hs_kept):,} rows")
    print(f"  HS deleted: {len(hs_deleted):,} rows")

    # Step 4: Match rows
    print(f"\n[4/4] Matching {len(hs_kept):,} rows...")
    total_raw = len(raw)
    match_count = len(hs_kept)
    processed = [0]

    def _match(row):
        result = match_row(row, biz_lookup, biz_registry, del_lookup, del_registry, THRESHOLD, DEL_THRESHOLD)
        processed[0] += 1
        if processed[0] % 5000 == 0:
            print(f"  {processed[0]:,}/{match_count:,} rows processed...")
        return result

    results = hs_kept.apply(_match, axis=1)
    results_df = pd.DataFrame(list(results))
    combined = pd.concat([hs_kept.reset_index(drop=True), results_df], axis=1)
    print(f"  {match_count:,}/{match_count:,} done.")

    # --- Build output sheets ---
    print(f"\nExporting to {output_file}...")
    sheet_data = {}

    for sheet in biz_sheets:
        mask = combined['business'] == sheet
        if mask.any():
            sheet_data[sheet] = format_output(combined[mask].copy())
            print(f"  {sheet}: {mask.sum():,} rows")

    del_mask = combined['business'] == 'Deleted'
    if del_mask.any():
        sheet_data['Deleted'] = format_output(combined[del_mask].copy())
        print(f"  Deleted: {del_mask.sum():,} rows")

    if len(hs_deleted) > 0:
        sheet_data['HS Deleted'] = format_output(hs_deleted.copy())
        print(f"  HS Deleted: {len(hs_deleted):,} rows")

    unc_mask = combined['business'] == 'Uncertain'
    if unc_mask.any():
        sheet_data['Uncertain'] = format_output(combined[unc_mask].copy())
        print(f"  Uncertain: {unc_mask.sum():,} rows")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for name, df in sheet_data.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)

    # Summary
    biz_total = sum((combined['business'] == s).sum() for s in biz_sheets)
    del_total = del_mask.sum()
    hs_del_total = len(hs_deleted)
    unc_total = unc_mask.sum()

    print(f"\n{'=' * 60}")
    print(f"SUMMARY")
    print(f"{'=' * 60}")
    print(f"  Total raw rows: {total_raw:,}")
    print(f"  HS Deleted:     {hs_del_total:,} ({hs_del_total/total_raw*100:.1f}%)")
    print(f"  HS Kept:        {match_count:,} ({match_count/total_raw*100:.1f}%)")
    print(f"  Business:       {biz_total:,} ({biz_total/total_raw*100:.1f}%)")
    for sheet in biz_sheets:
        cnt = (combined['business'] == sheet).sum()
        if cnt > 0:
            print(f"    {sheet:15s} {cnt:,}")
    print(f"  Deleted (dict): {del_total:,} ({del_total/total_raw*100:.1f}%)")
    print(f"  Uncertain:      {unc_total:,} ({unc_total/total_raw*100:.1f}%)")
    print(f"\nSaved to {output_file}")


if __name__ == '__main__':
    main()
