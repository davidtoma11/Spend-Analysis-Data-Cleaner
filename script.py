import pandas as pd
import os
from difflib import SequenceMatcher
import re

# config
INPUT_FILE = 'presales_data_sample(0).csv'
FILE_1 = 'filtred_data_colored(1).xlsx'  # classic view (no yellow)
FILE_2 = 'filtred_data(2).xlsx'  # strict view (yellow for close calls)
SCORE_THRESHOLD = 60
GAP_THRESHOLD = 5  # if gap between #1 and #2/3/.. is <= this, mark yellow


if not os.path.exists(INPUT_FILE):
    print(f"error: file '{INPUT_FILE}' not found.")
    exit()

print("1. Loading data...")
df = pd.read_csv(INPUT_FILE)


# STEP 1: cleaning text
def clean_text_safe(val):
    if pd.isna(val): return ""
    val = str(val).lower().strip()
    return re.sub(r'[\x00-\x1f]', '', val)


for col in df.select_dtypes(include=['object']).columns:
    df[col] = df[col].apply(clean_text_safe)

# STEP 2: calculating SMART scores (Name + Bonuses)
print("2. Calculating smart scores (Name + Location Bonus)...")


def calculate_smart_score(row):
    # 1. Base Score (Name Similarity)
    name_input = row['input_company_name']
    name_candidate = row['company_name']

    if not name_input or not name_candidate:
        base_score = 0
    else:
        base_score = SequenceMatcher(None, name_input, name_candidate).ratio() * 100

    # 2. Location Bonuses
    bonus = 0

    # Check Country (+5 pts)
    # We use 'str' to handle potential non-string types safely
    ctry_in = str(row['input_main_country_code'])
    ctry_cand = str(row['main_country_code'])
    if ctry_in and ctry_cand and ctry_in == ctry_cand and ctry_in != "":
        bonus += 5

    # Check Region (+3 pts)
    reg_in = str(row['input_main_region'])
    reg_cand = str(row['main_region'])
    if reg_in and reg_cand and reg_in == reg_cand and reg_in != "":
        bonus += 3

    # Check City (+1 pt)
    city_in = str(row['input_main_city'])
    city_cand = str(row['main_city'])
    if city_in and city_cand and city_in == city_cand and city_in != "":
        bonus += 1

    # Final calculation (Base + Bonus)
    # We cap it at 100 (optional, but looks cleaner)
    return min(base_score + bonus, 100)


df['FINAL_SCORE'] = df.apply(calculate_smart_score, axis=1)

# STEP 3: processing logic
print("3. Selecting rows for output files...")

# prepare country codes (helper columns for logic)
df['_in_country'] = df['input_main_country_code'].astype(str).str.upper()
df['_cand_country'] = df['main_country_code'].astype(str).str.upper()

# validity check
df['_is_valid'] = (
        (df['_in_country'] == df['_cand_country']) &
        (df['_in_country'] != 'NAN') &
        (df['FINAL_SCORE'] >= SCORE_THRESHOLD)
)

# sorting (best score on top)
df = df.sort_values(by=['input_row_key', 'FINAL_SCORE'], ascending=[True, False])

# collections for file 1 (classic view)
f1_indices_best = set()   # dark green
f1_indices_valid = set()  # light green

# collections for file 2 (strict view)
f2_indices_green = set()  # dark green (clear winner)
f2_indices_yellow = set() # yellow (close race)
f2_indices_red = set()    # red (forced match)

indices_for_file_2 = []  # rows to keep in the second file

# main loop
for input_id, group in df.groupby('input_row_key'):

    # find all valid rows (green)
    valid_rows = group[group['_is_valid'] == True]

    if not valid_rows.empty:
        # SCENARIO A: WE HAVE VALID MATCHES

        # the first one is the best (sorted by score)
        best_idx = valid_rows.index[0]
        f1_indices_best.add(best_idx)
        # all valid rows are light green in file 1
        f1_indices_valid.update(valid_rows.index)

        top_score = valid_rows.iloc[0]['FINAL_SCORE']

        # finding rows that are very close to the top one
        close_contenders = valid_rows[
            (top_score - valid_rows['FINAL_SCORE']) <= GAP_THRESHOLD
            ]

        if len(close_contenders) > 1:
            # gap is small -> mark them ALL as yellow in file 2
            f2_indices_yellow.update(close_contenders.index)
            # keep all close contenders
            indices_for_file_2.extend(close_contenders.index.tolist())

        else:
            # clear winner -> mark as dark green in file 2
            f2_indices_green.add(best_idx)
            # keep ONLY the winner
            indices_for_file_2.append(best_idx)

        # Note: Valid alternatives that are NOT close contenders are dropped from File 2

    else:
        # SCENARIO B: NO VALID MATCHES (ALL RED)
        # we need to pick one "forced match"

        # try finding country match
        country_matches = group[
            (group['_in_country'] == group['_cand_country']) &
            (group['_in_country'] != 'NAN')
            ]

        if not country_matches.empty:
            forced_idx = country_matches['FINAL_SCORE'].idxmax()
        else:
            forced_idx = group['FINAL_SCORE'].idxmax()

        # update file 2 lists
        f2_indices_red.add(forced_idx)
        indices_for_file_2.append(forced_idx)

# cleanup
df = df.drop(columns=['_in_country', '_cand_country', '_is_valid'])
cols = [c for c in df.columns if c != 'FINAL_SCORE']
df = df[['FINAL_SCORE'] + cols]

# STEP 4: GENERATING FILE 1 (FULL)
print("4. Generating file 1 (all rows colored)...")


def style_file_1(row):
    # classic logic: best is dark green, valid is light green, rest is red
    if row.name in f1_indices_best:
        return ['background-color: #63be7b'] * len(row)  # dark green
    elif row.name in f1_indices_valid:
        return ['background-color: #d4edda'] * len(row)  # light green
    else:
        return ['background-color: #f8d7da'] * len(row)  # red


styler1 = df.style.apply(style_file_1, axis=1)
styler1 = styler1.format({'FINAL_SCORE': "{:.1f}"})
styler1.to_excel(FILE_1, index=False, engine='openpyxl')
print(f"   -> saved: {FILE_1}")

# STEP 5: GENERATING FILE 2 (FILTERED)
print("5. Generating file 2 (strict mode: green/yellow/red)...")

# create the filtered dataframe using our new list
df_filtered = df.loc[indices_for_file_2].copy()


def style_file_2(row):
    # strict logic: green (clear), yellow (ambiguous), red (forced)

    if row.name in f2_indices_green:
        return ['background-color: #63be7b'] * len(row)  # clear winner
    elif row.name in f2_indices_yellow:
        return ['background-color: #ffeeba'] * len(row)  # yellow (close call)
    elif row.name in f2_indices_red:
        return ['background-color: #f8d7da'] * len(row)  # forced match
    else:
        return [''] * len(row)


styler2 = df_filtered.style.apply(style_file_2, axis=1)
styler2 = styler2.format({'FINAL_SCORE': "{:.1f}"})
styler2.to_excel(FILE_2, index=False, engine='openpyxl')

print(f"   -> saved: {FILE_2}")
print("   -> bonuses applied: +5 Country, +3 Region, +1 City")
print("   -> file 2 logic: yellow applied only for close gaps (<= 10)")