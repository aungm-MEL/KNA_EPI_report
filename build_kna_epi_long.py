import pandas as pd
from pathlib import Path
import os
import warnings
import gc
import re

warnings.filterwarnings('ignore')

base_dir = Path(__file__).resolve().parent
clean_input_path = base_dir / 'KNA_cleantoreport' / 'KNA_clean.xlsx'
dataset_input_path = base_dir / 'KNA_dataset.xlsx'
env_input = os.getenv('KNA_LONG_INPUT', '').strip() or os.getenv('KNA_CLEAN_INPUT', '').strip()
input_path = Path(env_input).expanduser().resolve() if env_input else (clean_input_path if clean_input_path.exists() else dataset_input_path)
env_output = os.getenv('KNA_LONG_OUTPUT', '').strip()
output_path = Path(env_output).expanduser().resolve() if env_output else (base_dir / 'KNA_EPI_long.xlsx')
# Compact output mode: skip the large long-format sheets to reduce file size.
# Set to True to include 'child' and 'Td' long sheets in the output workbook.
INCLUDE_LONG_SHEETS_IN_OUTPUT = True

dose_mappings_child = {
    'BCG': {'dose': 'BCGD', 'age': 'BCA', 'source': 'BC_C'},
    'OPV1': {'dose': 'OP1D', 'age': 'OP1A', 'source': 'OP1C'},
    'OPV2': {'dose': 'OP2D', 'age': 'OP2A', 'source': 'OP2C'},
    'OPV3': {'dose': 'OP3D', 'age': 'OP3A', 'source': 'OP3C'},
    'Penta1': {'dose': 'Pe1D', 'age': 'Pe1A', 'source': 'PE1C'},
    'Penta2': {'dose': 'Pe2D', 'age': 'Pe2A', 'source': 'PE2C'},
    'Penta3': {'dose': 'Pe3D', 'age': 'Pe3A', 'source': 'PE3C'},
    'MMR1': {'dose': 'MM1D', 'age': 'M1A', 'source': 'MM1C'},
    'MMR2': {'dose': 'MM2D', 'age': 'M2A', 'source': 'MM2C'},
    'JE1': {'dose': 'JaED', 'age': 'JeA', 'source': 'JaEC'},
    'IPV': {'dose': 'IP1D', 'age': 'IPA', 'source': 'IP1C'},
}

completion_cols_child = [
    'completion_status_2024_Q1', 'completion_status_2024_Q2', 'completion_status_2024_Q3', 'completion_status_2024_Q4',
    'Full Dose in Q1 2025', 'Full Dose in Q2 2025', 'Full Dose in Q3 2025', 'Full Dose in Q4 2025',
    'Full Dose in Q1 2026', 'Full Dose in Q2 2026', 'Full Dose in Q3 2026', 'Full Dose in Q4 2026'
]

keep_cols_child = ['Twp', 'Reg_', 'Site', 'children_code', 'Sex_', 'DOB_', 'FSD', 'FSD A', 'Resi', 'IDP_', 'completion_status_2023', 'completion_status_2024']

pattern_mapping = {
    0: 'Not received Yet',
    1: 'Other',
    2: 'Provided by KNA',
    3: 'Stock out',
    4: 'Old age for dose',
    5: 'Young age for dose'
}

dose_mappings_td = {
    'TD1': {'dose': 'TD1D', 'source': 'TD1C'},
    'TD2': {'dose': 'TD2D', 'source': 'TD2C'},
}

keep_cols_td = ['Twp', 'Regn', 'Site', 'pw_code', 'AN Name', 'Age', 'Address', 'IDP', 'Disabled', 'FSD']

def map_pattern(val):
    if pd.isna(val):
        return val
    try:
        num = int(float(val))
        return pattern_mapping.get(num, val)
    except (ValueError, TypeError):
        return val


def _resolve_column_name(df: pd.DataFrame, expected: str):
    """Return matching column name, tolerating trailing/leading spaces."""
    if expected in df.columns:
        return expected

    target = expected.strip()
    for c in df.columns:
        if isinstance(c, str) and c.strip() == target:
            return c
    return None


def _sex_bucket(val):
    """Normalize sex labels from mixed encodings/languages to Male/Female."""
    if pd.isna(val):
        return None
    raw = str(val).strip()
    low = raw.lower()

    if raw == '\u1000\u103b\u102c\u1038' or 'male' in low or low == 'm':
        return 'Male'
    if raw == '\u1019' or 'female' in low or low == 'f':
        return 'Female'
    return None


def get_completion_status(row):
    # Priority order requested: completion_status_2024, then 2025 columns, then 2026 columns.

    # For 2024, use the new quarterly columns if present
    for q in ['Q1', 'Q2', 'Q3', 'Q4']:
        col = f'completion_status_2024_{q}'
        val_q = row.get(col, None)
        if pd.notna(val_q):
            val_q_str = str(val_q).strip()
            if val_q_str not in {'', '0', '0.0', 'nan', 'None'}:
                lower_q = val_q_str.lower()
                core_q = val_q_str
                for phrase in ['completed in', 'complete in']:
                    idx = lower_q.find(phrase)
                    if idx != -1:
                        core_q = val_q_str[:idx].strip()
                        break
                return f"{core_q} completed in {q} 2024"

    val_2024 = row.get('completion_status_2024', None)
    if pd.notna(val_2024):
        val_2024_str = str(val_2024).strip()
        if val_2024_str not in {'', '0', '0.0', 'nan', 'None'}:
            lower_2024 = val_2024_str.lower()
            y_match_2024 = re.search(r'20\d{2}', val_2024_str)
            # If 2024 completion has year only (no quarter), map to Q4 for reporting.
            if y_match_2024:
                core_2024 = val_2024_str
                for phrase in ['completed in', 'complete in']:
                    idx = lower_2024.find(phrase)
                    if idx != -1:
                        core_2024 = val_2024_str[:idx].strip()
                        break
                return f"{core_2024} completed in Q4 2024"
            return val_2024_str

    for col in completion_cols_child:
        if col not in row:
            continue
        val = row[col]
        if pd.isna(val) or val == '' or val == 0:
            continue
        val_str = str(val).strip()
        lower = val_str.lower()
        core_val = val_str
        for phrase in ['completed in', 'complete in']:
            idx = lower.find(phrase)
            if idx != -1:
                core_val = val_str[:idx].strip()
                break
        # Extract quarter (Q1-Q4) and year (20xx) directly from column name
        q_match = re.search(r'(Q[1-4])', col)
        y_match = re.search(r'(20\d{2})', col)
        if q_match and y_match:
            return f"{core_val} completed in {q_match.group(1)} {y_match.group(1)}"
    return None


def build_child_long(df_child: pd.DataFrame) -> pd.DataFrame:
    if 'Children_code' in df_child.columns and 'children_code' not in df_child.columns:
        df_child = df_child.rename(columns={'Children_code': 'children_code'})

    frames = []
    for vname, cmap in dose_mappings_child.items():
        required = [cmap['dose'], cmap['age'], cmap['source']]
        missing = [c for c in required if c not in df_child.columns]
        if missing:
            print(f"Warning missing {vname}: {missing}")
            continue
        cols = [c for c in keep_cols_child + completion_cols_child + required if c in df_child.columns]
        sub = df_child[cols].copy()
        sub.rename(columns={cmap['dose']: 'period', cmap['age']: 'age_at_dose', cmap['source']: 'receiving_pattern'}, inplace=True)
        sub['vaccine_dose'] = vname
        ordered = [c for c in keep_cols_child if c in sub.columns] + ['period', 'vaccine_dose', 'age_at_dose', 'receiving_pattern']
        frames.append(sub[ordered + [c for c in sub.columns if c not in ordered]])
    if not frames:
        raise SystemExit('No child frames created')
    long_df = pd.concat(frames, ignore_index=True)

    long_df['receiving_pattern'] = long_df['receiving_pattern'].apply(map_pattern)

    # filters
    before = len(long_df)
    long_df = long_df[~((long_df['age_at_dose'] == 1111) & (long_df['receiving_pattern'] == 'Not Received Yet'))]
    long_df = long_df[long_df['age_at_dose'] != 999]
    print(f"Child: removed {before - len(long_df)} rows from filters")

    # completion_status
    long_df['completion_status'] = long_df.apply(get_completion_status, axis=1)
    drop_cols = [c for c in completion_cols_child + ['completion_status_2024'] if c in long_df.columns]
    long_df = long_df.drop(columns=drop_cols, errors='ignore')

    # reorder period after IDP_
    if 'period' in long_df.columns and 'IDP_' in long_df.columns:
        cols = list(long_df.columns)
        cols.remove('period')
        try:
            idx = cols.index('IDP_') + 1
        except ValueError:
            idx = 0
        cols = cols[:idx] + ['period'] + cols[idx:]
        long_df = long_df[cols]

    # move completion_status_2023 to end if present
    for tail in ['completion_status_2023']:
        if tail in long_df.columns:
            cols = [c for c in long_df.columns if c != tail] + [tail]
            long_df = long_df[cols]

    return long_df


def build_td_long(df_td: pd.DataFrame) -> pd.DataFrame:
    if 'PW_code' in df_td.columns and 'pw_code' not in df_td.columns:
        df_td = df_td.rename(columns={'PW_code': 'pw_code'})

    frames = []
    for vname, cmap in dose_mappings_td.items():
        required = [cmap['dose'], cmap['source']]
        missing = [c for c in required if c not in df_td.columns]
        if missing:
            print(f"Warning missing {vname}: {missing}")
            continue
        cols = [c for c in keep_cols_td + required if c in df_td.columns]
        sub = df_td[cols].copy()
        sub.rename(columns={cmap['dose']: 'period', cmap['source']: 'receiving_pattern'}, inplace=True)
        sub['vaccine_dose'] = vname
        ordered = [c for c in keep_cols_td if c in sub.columns] + ['period', 'vaccine_dose', 'receiving_pattern']
        frames.append(sub[ordered])
    if not frames:
        raise SystemExit('No Td frames created')
    long_df = pd.concat(frames, ignore_index=True)
    long_df['receiving_pattern'] = long_df['receiving_pattern'].apply(map_pattern)
    long_df = long_df[long_df['receiving_pattern'].notna() & (long_df['receiving_pattern'] != '')]
    if 'period' in long_df.columns and 'FSD' in long_df.columns:
        cols = list(long_df.columns)
        cols.remove('period')
        try:
            idx = cols.index('FSD') + 1
        except ValueError:
            idx = 0
        cols = cols[:idx] + ['period'] + cols[idx:]
        long_df = long_df[cols]
    return long_df


def get_period(date):
    if pd.isna(date):
        return None
    d = pd.to_datetime(date)
    y = d.year
    q1_start = pd.Timestamp(year=y - 1, month=12, day=21)
    q1_end = pd.Timestamp(year=y, month=3, day=20)
    if q1_start <= d <= q1_end:
        return f"Q1_{y}"
    q2_start = pd.Timestamp(year=y, month=3, day=21)
    q2_end = pd.Timestamp(year=y, month=6, day=20)
    if q2_start <= d <= q2_end:
        return f"Q2_{y}"
    q3_start = pd.Timestamp(year=y, month=6, day=21)
    q3_end = pd.Timestamp(year=y, month=9, day=20)
    if q3_start <= d <= q3_end:
        return f"Q3_{y}"
    q4_start = pd.Timestamp(year=y, month=9, day=21)
    q4_end = pd.Timestamp(year=y, month=12, day=20)
    if q4_start <= d <= q4_end:
        return f"Q4_{y}"
    return f"Q1_{y + 1}"


def age_cat_months(age_months):
    if pd.isna(age_months):
        return None
    a = float(age_months)
    if a <= 11:
        return 'U1'
    if a <= 59:
        return 'U5'
    return '>5'


def age_cat_years(age_years):
    if pd.isna(age_years):
        return None
    a = float(age_years)
    # Treat age_at_dose as months per requirement
    # U1: 0-11 months, U5: 12-59 months, >5: >= 60 months
    if a <= 11:
        return 'U1'
    if a <= 59:
        return 'U5'
    return '>5'


# Vaccine-specific age buckets (months) for summary vaccine counts
def age_cat_vaccine(age_months, vaccine):
    if pd.isna(age_months):
        return None
    a = float(age_months)

    if vaccine == 'Penta1':
        if 2 <= a <= 11:
            return 'U1'
        if 12 <= a <= 59:
            return 'U5'
        if a >= 60:
            return '>5'
        return None

    if vaccine == 'Penta3':
        if 4 <= a <= 11:
            return 'U1'
        if 12 <= a <= 59:
            return 'U5'
        if a >= 60:
            return '>5'
        return None

    if vaccine == 'MMR1':
        if 9 <= a <= 11:
            return 'U1'
        if 12 <= a <= 59:
            return 'U5'
        if a >= 60:
            return '>5'
        return None

    if vaccine == 'MMR2':
        if 10 <= a <= 11:
            return 'U1'
        if 12 <= a <= 59:
            return 'U5'
        if a >= 60:
            return '>5'
        return None

    # Default for other vaccines
    if a <= 11:
        return 'U1'
    if a <= 59:
        return 'U5'
    return '>5'


def extract_years_from_template(template_df: pd.DataFrame, fallback=None):
    """Extract sorted unique years from common period columns in a template."""
    years = []
    for period_col in ['Period', 'period', 'Period ']:
        if period_col in template_df.columns:
            vals = pd.to_numeric(template_df[period_col], errors='coerce').dropna().astype(int).astype(str)
            years.extend([v for v in vals.tolist() if re.fullmatch(r'20\d{2}', v)])

    years = sorted(set(years))
    if years:
        return years
    return fallback if fallback is not None else ['2024', '2025', '2026']


def build_summary(child_df: pd.DataFrame, td_df: pd.DataFrame) -> pd.DataFrame:
    child_df = child_df.copy()
    td_df = td_df.copy()
    child_df_all = child_df.copy()

    # Filter for "Provided by KNA" and "Young age for dose"
    child_df = child_df[child_df['receiving_pattern'].isin(['Provided by KNA', 'Young age for dose'])].copy()
    td_df = td_df[td_df['receiving_pattern'] == 'Provided by KNA'].copy()

    child_df['quarter'] = child_df['period'].apply(get_period)
    td_df['quarter'] = td_df['period'].apply(get_period)
    child_df_all['quarter'] = child_df_all['period'].apply(get_period)

    child_df['age_category'] = child_df.apply(lambda r: age_cat_vaccine(r['age_at_dose'], r['vaccine_dose']), axis=1)
    child_df['age_category_fsd'] = child_df['FSD A'].apply(age_cat_months)
    child_df_all['age_category_fsd'] = child_df_all['FSD A'].apply(age_cat_months)

    # Completion dose counts from completion_status (CD_U1/CD_U5/CD_>5)
    cd_pivot = pd.DataFrame(columns=['quarter', 'Twp', 'Site', 'CD_U1', 'CD_U5', 'CD_>5'])
    if {'completion_status', 'children_code', 'Twp', 'Site'}.issubset(child_df_all.columns):
        comp_df = child_df_all[['completion_status', 'Twp', 'Site', 'children_code']].copy()
        comp_df = comp_df[comp_df['completion_status'].notna() & (comp_df['completion_status'] != '')]
        comp_df = comp_df[comp_df['children_code'].notna()]

        def parse_completion(val):
            text = str(val)
            lower = text.lower()
            normalized = text.replace('\u2011', '-').replace('\u2012', '-').replace('\u2013', '-').replace('\u2014', '-')
            normalized_lower = normalized.lower()

            if 'u1' in lower:
                age = 'CD_U1'
            elif '1-5' in normalized_lower or '1 - 5' in normalized_lower or 'u5' in lower:
                age = 'CD_U5'
            elif '>5' in normalized:
                age = 'CD_>5'
            else:
                age = None

            if 'q1' in lower:
                quarter = 'Q1'
            elif 'q2' in lower:
                quarter = 'Q2'
            elif 'q3' in lower:
                quarter = 'Q3'
            elif 'q4' in lower:
                quarter = 'Q4'
            else:
                quarter = None

            year_match = re.search(r'20\d{2}', normalized)
            period = f"{quarter}_{year_match.group(0)}" if quarter and year_match else None

            return age, period

        parsed = comp_df['completion_status'].apply(parse_completion)
        if len(parsed) > 0:
            parsed_df = pd.DataFrame(parsed.tolist(), columns=['cd_age', 'cd_period'], index=comp_df.index)
            comp_df = pd.concat([comp_df, parsed_df], axis=1)
        else:
            comp_df['cd_age'] = None
            comp_df['cd_period'] = None
        comp_df = comp_df.dropna(subset=['cd_age', 'cd_period'])
        comp_df = comp_df.drop_duplicates(subset=['children_code', 'Twp', 'Site', 'cd_age', 'cd_period'])

        cd_agg = comp_df.groupby(['cd_period', 'Twp', 'Site', 'cd_age'])['children_code'].nunique().reset_index(name='count')
        cd_pivot = cd_agg.pivot_table(index=['cd_period', 'Twp', 'Site'], columns='cd_age', values='count', fill_value=0).reset_index()
        cd_pivot = cd_pivot.rename(columns={'cd_period': 'quarter'})

        for col in ['CD_U1', 'CD_U5', 'CD_>5']:
            if col not in cd_pivot.columns:
                cd_pivot[col] = 0

        cd_pivot = cd_pivot[['quarter', 'Twp', 'Site', 'CD_U1', 'CD_U5', 'CD_>5']]

    # ALOD unique children_code with receiving_pattern filter
    alod_filtered = child_df_all[child_df_all['receiving_pattern'].isin(['Provided by KNA', 'Young age for dose'])].copy()
    alod_agg = alod_filtered.groupby(['quarter', 'Twp', 'Site', 'age_category_fsd'])['children_code'].nunique().reset_index(name='count')
    alod_pivot = alod_agg.pivot_table(index=['quarter', 'Twp', 'Site'], columns='age_category_fsd', values='count', fill_value=0).reset_index()
    alod_pivot.columns = ['quarter', 'Twp', 'Site'] + [f'ALOD_{c}' for c in alod_pivot.columns[3:]]

    child_agg = child_df.groupby(['quarter', 'Twp', 'Site', 'vaccine_dose', 'age_category'])['children_code'].nunique().reset_index(name='count')
    child_pivot = child_agg.pivot_table(index=['quarter', 'Twp', 'Site'], columns=['vaccine_dose', 'age_category'], values='count', fill_value=0).reset_index()
    child_pivot.columns = ['_'.join(col).strip('_') if isinstance(col, tuple) else col for col in child_pivot.columns]
    child_pivot = child_pivot[[c for c in child_pivot.columns if not c.startswith('ALOD_')]]
    # Rename JE1_* columns to JE_* to match expected output
    je_rename = {c: c.replace('JE1_', 'JE_', 1) for c in child_pivot.columns if isinstance(c, str) and c.startswith('JE1_')}
    if je_rename:
        child_pivot = child_pivot.rename(columns=je_rename)

    td_agg = td_df.groupby(['quarter', 'Twp', 'Site', 'vaccine_dose'])['pw_code'].nunique().reset_index(name='count')
    td_pivot = td_agg.pivot_table(index=['quarter', 'Twp', 'Site'], columns='vaccine_dose', values='count', fill_value=0).reset_index()

    # Td At least one dose: unique count of pw_code
    # Filters: receiving_pattern='Provided by KNA', quarterly disaggregate (21-20 day boundaries)
    # Groups by: quarter, Twp, Site
    td_alod_agg = td_df[td_df['receiving_pattern'] == 'Provided by KNA'].groupby(['quarter', 'Twp', 'Site'])['pw_code'].nunique().reset_index(name='Td At least one dose')

    summary_df = alod_pivot.merge(child_pivot, on=['quarter', 'Twp', 'Site'], how='outer')
    summary_df = summary_df.merge(td_pivot, on=['quarter', 'Twp', 'Site'], how='outer', suffixes=('', '_td'))
    summary_df = summary_df.merge(td_alod_agg, on=['quarter', 'Twp', 'Site'], how='outer')
    summary_df = summary_df.merge(cd_pivot, on=['quarter', 'Twp', 'Site'], how='outer')

    summary_df = summary_df.fillna(0)

    rename_dict = {c: c.replace('TD', 'Td') for c in summary_df.columns if c.startswith('TD')}
    summary_df = summary_df.rename(columns=rename_dict)

    summary_df = summary_df.rename(columns={'quarter': 'period', 'Twp': 'Twp_MIMU', 'Site': 'Clinic Name'})

    # Add Year column derived from period (e.g., Q1_2025 -> 2025)
    summary_df['Year'] = summary_df['period'].astype(str).str.split('_').str[-1]

    summary_df['Organization'] = 'KNA'
    summary_df['Project Name'] = 'REACH-KK'
    summary_df['District (EHO)'] = ''
    summary_df['Township_EHO'] = ''

    vaccine_columns = [
        'ALOD_U1', 'ALOD_U5', 'ALOD_>5',
        'BCG_U1', 'BCG_U5', 'BCG_>5',
        'OPV1_U1', 'OPV1_U5', 'OPV1_>5',
        'OPV2_U1', 'OPV2_U5', 'OPV2_>5',
        'OPV3_U1', 'OPV3_U5', 'OPV3_>5',
        'Penta1_U1', 'Penta1_U5', 'Penta1_>5',
        'Penta2_U1', 'Penta2_U5', 'Penta2_>5',
        'Penta3_U1', 'Penta3_U5', 'Penta3_>5',
        'MMR1_U1', 'MMR1_U5', 'MMR1_>5',
        'MMR2_U1', 'MMR2_U5', 'MMR2_>5',
        'JE_U1', 'JE_U5', 'JE_>5',
        'IPV_U1', 'IPV_U5', 'IPV_>5',
        'CD_U1', 'CD_U5', 'CD_>5',
        'Td1', 'Td2', 'Td At least one dose'
    ]

    for col in vaccine_columns:
        if col not in summary_df.columns:
            summary_df[col] = 0

    final_cols = ['Year', 'period', 'Organization', 'Project Name', 'District (EHO)', 'Township_EHO', 'Twp_MIMU', 'Clinic Name'] + vaccine_columns
    summary_df = summary_df[final_cols]
    summary_df = summary_df[summary_df['period'].notna()]
    summary_df = summary_df.sort_values(['period', 'Clinic Name']).reset_index(drop=True)
    return summary_df


def build_yearly_cumulative(child_df: pd.DataFrame, td_df: pd.DataFrame) -> pd.DataFrame:
    child_df = child_df.copy()
    td_df = td_df.copy()

    # Extract year from period column
    child_df['year'] = child_df['period'].apply(lambda d: pd.to_datetime(d).year if pd.notna(d) else None)
    td_df['year'] = td_df['period'].apply(lambda d: pd.to_datetime(d).year if pd.notna(d) else None)

    child_df['age_category_fsd'] = child_df['FSD A'].apply(age_cat_months)

    # ALOD: unique children_code with receiving_pattern='Provided by KNA' or 'Young age for dose', grouped by year
    alod_filtered = child_df[child_df['receiving_pattern'].isin(['Provided by KNA', 'Young age for dose'])].copy()
    alod_agg = alod_filtered.groupby(['year', 'Twp', 'Site', 'age_category_fsd'])['children_code'].nunique().reset_index(name='count')
    alod_pivot = alod_agg.pivot_table(index=['year', 'Twp', 'Site'], columns='age_category_fsd', values='count', fill_value=0).reset_index()
    alod_pivot.columns = ['year', 'Twp', 'Site'] + [f'ALOD-{c}' for c in alod_pivot.columns[3:]]

    # Td at least one dose: count unique pw_code with receiving_pattern='Provided by KNA', grouped by year/Twp/Site
    td_filtered = td_df[td_df['receiving_pattern'] == 'Provided by KNA'].copy()
    td_agg = td_filtered.groupby(['year', 'Twp', 'Site'])['pw_code'].nunique().reset_index(name='Td At least one dose')

    # Merge ALOD with Td
    yearly_df = alod_pivot.merge(td_agg, on=['year', 'Twp', 'Site'], how='outer')
    yearly_df = yearly_df.fillna(0)

    # Add metadata columns
    yearly_df = yearly_df.rename(columns={'year': 'period', 'Twp': 'Twp_MIMU', 'Site': 'Clinic Name'})
    yearly_df['Organization'] = 'KNA'
    yearly_df['Project Name'] = 'REACH-KK'
    yearly_df['District (EHO)'] = ''
    yearly_df['Township_EHO'] = ''

    # Ensure all ALOD columns exist
    for col in ['ALOD-U1', 'ALOD-U5', 'ALOD->5']:
        if col not in yearly_df.columns:
            yearly_df[col] = 0

    # Final column order
    final_cols = ['Organization', 'period', 'Project Name', 'District (EHO)', 'Township_EHO', 'Twp_MIMU', 'Clinic Name', 'ALOD-U1', 'ALOD-U5', 'ALOD->5', 'Td At least one dose']
    yearly_df = yearly_df[final_cols]
    
    # Remove rows without period
    yearly_df = yearly_df[yearly_df['period'].notna()]
    yearly_df = yearly_df.sort_values(['period', 'Clinic Name']).reset_index(drop=True)
    
    return yearly_df


def build_cumulative(child_df: pd.DataFrame, td_df: pd.DataFrame) -> pd.DataFrame:
    child_df = child_df.copy()
    td_df = td_df.copy()

    child_df['age_category_fsd'] = child_df['FSD A'].apply(age_cat_months)

    # ALOD: unique children_code with receiving_pattern='Provided by KNA' or 'Young age for dose', overall (no time disaggregation)
    alod_filtered = child_df[child_df['receiving_pattern'].isin(['Provided by KNA', 'Young age for dose'])].copy()
    alod_agg = alod_filtered.groupby(['Twp', 'Site', 'age_category_fsd'])['children_code'].nunique().reset_index(name='count')
    alod_pivot = alod_agg.pivot_table(index=['Twp', 'Site'], columns='age_category_fsd', values='count', fill_value=0).reset_index()
    alod_pivot.columns = ['Twp', 'Site'] + [f'ALOD-{c}' for c in alod_pivot.columns[2:]]

    # Td at least one dose: count unique pw_code with receiving_pattern='Provided by KNA', overall
    td_filtered = td_df[td_df['receiving_pattern'] == 'Provided by KNA'].copy()
    td_agg = td_filtered.groupby(['Twp', 'Site'])['pw_code'].nunique().reset_index(name='Td At least one dose')

    # Merge ALOD with Td
    cumulative_df = alod_pivot.merge(td_agg, on=['Twp', 'Site'], how='outer')
    cumulative_df = cumulative_df.fillna(0)

    # Add metadata columns
    cumulative_df = cumulative_df.rename(columns={'Twp': 'Twp_MIMU', 'Site': 'Clinic Name'})
    cumulative_df['Organization'] = 'KNA'
    cumulative_df['period'] = 'Overall'
    cumulative_df['Project Name'] = 'REACH-KK'
    cumulative_df['District (EHO)'] = ''
    cumulative_df['Township_EHO'] = ''

    # Ensure all ALOD columns exist
    for col in ['ALOD-U1', 'ALOD-U5', 'ALOD->5']:
        if col not in cumulative_df.columns:
            cumulative_df[col] = 0

    # Final column order
    final_cols = ['Organization', 'period', 'Project Name', 'District (EHO)', 'Township_EHO', 'Twp_MIMU', 'Clinic Name', 'ALOD-U1', 'ALOD-U5', 'ALOD->5', 'Td At least one dose']
    cumulative_df = cumulative_df[final_cols]
    cumulative_df = cumulative_df.sort_values(['Clinic Name']).reset_index(drop=True)
    
    return cumulative_df


def calculate_indicators(child_df: pd.DataFrame, indicators_df: pd.DataFrame) -> pd.DataFrame:
    """Calculate vaccine indicators from child sheet and update indicators sheet (with separate rows per year)"""
    child_df = child_df.copy()
    indicators_df = indicators_df.copy()
    
    # Ensure Period column exists and is set correctly depending on whether rows already have years
    if 'Period' not in indicators_df.columns:
        indicators_df['Period'] = 2025

    # Normalize matching keys once to avoid whitespace/type mismatches from templates.
    indicators_df['_indicator_norm'] = indicators_df['indicator'].astype(str).str.strip()
    indicators_df['_period_norm'] = pd.to_numeric(indicators_df['Period'], errors='coerce').astype('Int64')

    # Add quarter column
    child_df['quarter'] = child_df['period'].apply(get_period)

    # Use template years to avoid creating or logging out-of-template years
    years = extract_years_from_template(indicators_df, fallback=['2024', '2025', '2026'])
    print(f"  Years found for indicators: {years}")
    
    # If template doesn't have multiple year rows, duplicate rows and set Period column
    if indicators_df['Period'].nunique() == 1:
        original = indicators_df.copy()
        # Create rows for each year
        expanded_rows = []
        for year in years:
            year_rows = original.copy()
            year_rows['Period'] = year
            expanded_rows.append(year_rows)
        indicators_df = pd.concat(expanded_rows, ignore_index=True)
        print(f"  Expanded template to {len(indicators_df)} rows ({len(years)} years x {len(original)} indicators)")
    
    # Define vaccine indicators with age ranges
    # Format: (vaccine_name, indicator_row_name, [(age_min, age_max, col_suffix), ...])
    vaccines = [
        ('Penta3', 'Penta3 under 1-yr-old', [(4, 11, 'U1')]),
        ('Penta3', 'Penta3 under 5-yr-old', [(4, 11, 'U1'), (12, 59, '1-5')]),
        ('MMR1', 'MMR1 under 1-yr-old', [(9, 11, 'U1')]),
        ('MMR1', 'MMR1 under 5-yr-old', [(9, 11, 'U1'), (12, 59, '1-5')]),
        ('MMR2', 'MMR2 under 5-yr-old', [(10, 11, 'U1'), (12, 59, '1-5')]),
        ('Penta1', 'Penta1 under 1-yr-old', [(2, 11, 'U1')]),
        ('Penta1', 'Penta1 under 5-yr-old', [(2, 11, 'U1'), (12, 59, '1-5')]),
    ]
    
    # Process each vaccine indicator
    for vaccine_name, indicator_name, age_ranges in vaccines:
        print(f"Processing indicator: {indicator_name} ({vaccine_name})")
        indicator_rows = indicators_df[indicators_df['_indicator_norm'] == indicator_name]
        if indicator_rows.empty:
            print(f"  Skipping {indicator_name}: indicator row not found in template")
            continue
        # Filter for this vaccine
        vaccine_df = child_df[
            (child_df['vaccine_dose'] == vaccine_name) &
            (child_df['receiving_pattern'].isin(['Provided by KNA', 'Young age for dose']))
        ].copy()
        print(f"  {indicator_name}: {len(vaccine_df)} records total")
        
        if len(vaccine_df) == 0:
            continue
        
        # Process each year separately
        for year in years:
            # Filter data for this year
            vaccine_df_year = vaccine_df[vaccine_df['quarter'].astype(str).str.endswith(year, na=False)].copy()
            if len(vaccine_df_year) == 0:
                print(f"    {year}: no records")
                continue
                
            # Find the row for this year + indicator combination
            ind_idx = indicators_df[
                (indicators_df['_indicator_norm'] == indicator_name)
                & (indicators_df['_period_norm'] == int(year))
            ].index
            if len(ind_idx) == 0:
                print(f"    Warning: {indicator_name} row for {year} not found")
                continue
            idx = ind_idx[0]
            
            # Process each age range for this year
            for age_min, age_max, col_suffix in age_ranges:
                # Filter by age range for this year
                age_filtered = vaccine_df_year[(vaccine_df_year['age_at_dose'] >= age_min) & 
                                                 (vaccine_df_year['age_at_dose'] <= age_max)].copy()

                if len(age_filtered) == 0:
                    print(f"    {year} {col_suffix} ({age_min}-{age_max}): no records")
                    continue

                print(f"    {year} {col_suffix} ({age_min}-{age_max}): {len(age_filtered)} records")

                # Group by quarter and normalized sex using unique children_code counts
                age_filtered['_sex_norm'] = age_filtered['Sex_'].apply(_sex_bucket)
                age_filtered = age_filtered[age_filtered['_sex_norm'].notna()]
                sex_agg = age_filtered.groupby(['quarter', '_sex_norm'])['children_code'].nunique().reset_index(name='count')
                sex_pivot = sex_agg.pivot(index='quarter', columns='_sex_norm', values='count').fillna(0).astype(int)

                # Fill columns for each quarter
                for q in ['Q1', 'Q2', 'Q3', 'Q4']:
                    # Column names for this quarter (base, no year prefix since template doesn't use year suffixes)
                    if col_suffix == '1-5':
                        male_col = f'{q} {col_suffix} Male '
                        female_col = f'{q} {col_suffix} Female'
                    else:
                        male_col = f'{q} {col_suffix} Male'
                        female_col = f'{q} {col_suffix} Female'

                    # Find matching quarters for this Q
                    matching_quarters = [qk for qk in sex_pivot.index if isinstance(qk, str) and qk.startswith(q)]
                    male_total = 0
                    female_total = 0
                    if matching_quarters:
                        for qk in matching_quarters:
                            if 'Male' in sex_pivot.columns:
                                male_total += int(sex_pivot.loc[qk, 'Male'])
                            if 'Female' in sex_pivot.columns:
                                female_total += int(sex_pivot.loc[qk, 'Female'])

                    resolved_male = _resolve_column_name(indicators_df, male_col)
                    resolved_female = _resolve_column_name(indicators_df, female_col)

                    if resolved_male:
                        indicators_df.loc[idx, resolved_male] = male_total
                    if resolved_female:
                        indicators_df.loc[idx, resolved_female] = female_total
    
    # Calculate "Full dose under 5-yr-old" indicator — per year
    print("  Processing Full dose under 5-yr-old indicator")
    if 'completion_status' in child_df.columns:
        for year in years:
            has_completion = child_df[
                child_df['completion_status'].astype(str).str.contains(year, na=False)
                & child_df['completion_status'].notna()
            ].copy()
            print(f"    {year} full-dose completion records: {len(has_completion)}")

            if len(has_completion) == 0:
                print(f"    {year} No full dose records found")
                continue

            has_completion['age_cat'] = has_completion['completion_status'].apply(
                lambda x: 'U1' if 'U1' in str(x) else ('1-5' if '1-5' in str(x) else None)
            )

            # Extract quarter from completion_status
            def extract_q(val, yr):
                s = str(val)
                for q in ['Q1', 'Q2', 'Q3', 'Q4']:
                    if f"{q} {yr}" in s:
                        return q
                return None

            has_completion['completion_quarter'] = has_completion['completion_status'].apply(lambda x: extract_q(x, year))
            valid = has_completion[(has_completion['age_cat'].notna()) & (has_completion['completion_quarter'].notna())].copy()

            if len(valid) == 0:
                continue

            # Find the row for this year
            full_dose_idx = indicators_df[
                (indicators_df['_indicator_norm'] == 'Full dose under 5-yr-old')
         
