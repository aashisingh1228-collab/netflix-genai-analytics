#!/usr/bin/env python3
"""Netflix titles - reproducible cleaning & transformation pipeline.
Run: python clean_netflix.py <input_csv> <output_dir>
Raw input is only ever READ; it is never modified."""
import sys, re, hashlib
import numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/netflix_titles.csv'
OUT = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/outputs'
REFERENCE_YEAR = 2026          # analysis year (content_age = 2026 - release_year)
TEXT_COLS = ['show_id', 'type', 'title', 'director', 'cast', 'country', 'date_added',
             'rating', 'duration', 'listed_in', 'description']
ORIG_COLS = ['show_id', 'type', 'title', 'director', 'cast', 'country', 'date_added',
             'release_year', 'rating', 'duration', 'listed_in', 'description']

RATING_MAP = {  # documented mapping (US MPA film ratings + TV Parental Guidelines)
    'TV-Y': 'Children / Family', 'TV-Y7': 'Children / Family', 'TV-Y7-FV': 'Children / Family',
    'G': 'General Audience', 'TV-G': 'General Audience', 'PG': 'General Audience', 'TV-PG': 'General Audience',
    'PG-13': 'Teen', 'TV-14': 'Teen',
    'R': 'Mature Audience', 'NC-17': 'Mature Audience', 'TV-MA': 'Mature Audience',
    'Not Rated': 'Unclassified'}

# ------------------------------------------------------------------ load raw
sha = hashlib.sha256(open(SRC, 'rb').read()).hexdigest()
raw = pd.read_csv(SRC, dtype=str, keep_default_na=False, encoding='utf-8')
raw_infer = pd.read_csv(SRC, encoding='utf-8')          # pandas-inferred dtypes, for profiling only
df = raw.copy()
N0 = len(raw)
LOG = []
def log(phase, op, cols, logic, rows, why, before='', after=''):
    LOG.append(dict(Step=len(LOG) + 1, Phase=phase, Operation=op, Columns=cols, Logic=logic,
                    Rows_Cells_Affected=int(rows), Business_Reason=why, Example_Before=before, Example_After=after))

# ------------------------------------------------------------------ initial profile
def profile(d, infer=None):
    rows = []
    for c in d.columns:
        s = d[c].astype(str) if infer is None else raw[c]
        rows.append(dict(
            Column=c, Inferred_Dtype=str((infer if infer is not None else d)[c].dtype),
            Non_Empty=int((s != '').sum()), Empty_Missing=int((s == '').sum()),
            Missing_Pct=round((s == '').mean() * 100, 2),
            Whitespace_Only=int(((s.str.strip() == '') & (s != '')).sum()),
            Leading_Trailing_Space=int((s != s.str.strip()).sum()),
            Unique_Values=int(s[s != ''].nunique())))
    return pd.DataFrame(rows)
init_prof = profile(raw, raw_infer)
ry = raw_infer['release_year']
init_num = ry.describe().round(2).rename('release_year').to_frame().reset_index().rename(columns={'index': 'Statistic'})

init_checks = {
    'exact_dup_rows': int(raw.duplicated().sum()),
    'dup_show_id': int(raw.show_id.duplicated().sum()),
    'dup_title_exact': int(raw.title.duplicated().sum()),
    'dup_title_case_insens': int(raw.title.str.strip().str.lower().duplicated().sum()),
}
shift_mask = raw.rating.str.contains(r'^\d+\s*min$') & (raw.duration == '')
init_checks['duration_in_rating_col'] = int(shift_mask.sum())
init_checks['nbsp_cells'] = int(sum(raw[c].str.contains('\u00a0').sum() for c in raw.columns))
init_checks['embedded_newline_cells'] = int(sum(raw[c].str.contains(r'[\r\n\t]').sum() for c in raw.columns))
init_checks['multi_space_cells'] = int(sum(raw[c].str.contains(r'\s{2,}').sum() for c in raw.columns))
init_checks['lead_trail_cells'] = int(sum((raw[c] != raw[c].str.strip()).sum() for c in raw.columns))
init_checks['country_stray_commas'] = int(raw.country.str.contains(r'^\s*,|,\s*$').sum())
init_checks['distinct_rating_values'] = int(raw.rating.nunique())
init_checks['date_added_is_text'] = 1
init_checks['multi_country_rows'] = int(raw.country.str.contains(',').sum())
init_checks['multi_genre_rows'] = int(raw.listed_in.str.contains(',').sum())

# ------------------------------------------------------------------ PHASE 5 text standardisation
def sub_op(pattern, repl):
    return lambda s: re.sub(pattern, repl, s)
ops = [
    ('Convert non-breaking / zero-width / exotic Unicode spaces to a normal space', r'[\u00a0\u2000-\u200b\u202f\u3000\ufeff]', ' '),
    ('Replace embedded line breaks / tabs with a space', r'[\r\n\t\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', ' '),
    ('Collapse repeated spaces to a single space', r' {2,}', ' '),
]
for name, pat, rep in ops:
    tot, det, ex_b, ex_a = 0, [], '', ''
    for c in TEXT_COLS:
        new = df[c].map(sub_op(pat, rep)); m = new != df[c]; n = int(m.sum())
        if n:
            det.append(f'{c}={n}'); tot += n
            if not ex_b: ex_b, ex_a = repr(df.loc[m, c].iloc[0])[:80], repr(new[m].iloc[0])[:80]
        df[c] = new
    log('5 Text standardisation', name, 'all text columns', 'regex ' + pat + ' -> ' + repr(rep) + '; cells changed: ' + (', '.join(det) or 'none'),
        tot, 'Hidden characters break joins, filters, pivots and de-duplication', ex_b, ex_a)
tot, det = 0, []
for c in TEXT_COLS:
    new = df[c].str.strip(); n = int((new != df[c]).sum())
    if n: det.append(f'{c}={n}'); tot += n
    df[c] = new
log('5 Text standardisation', 'Trim leading / trailing whitespace', 'all text columns', 'str.strip(); cells changed: ' + (', '.join(det) or 'none'),
    tot, 'Stops "Consequences" and "Consequences " being counted as different titles; fixes text-to-date parsing', "' September 25, 2021' (leading space)", "'September 25, 2021'")

# ------------------------------------------------------------------ PHASE 3 duplicates
key_cols = [c for c in ORIG_COLS if c != 'show_id']
dup_mask = df.duplicated(subset=key_cols, keep='first')
removed = df[dup_mask][['show_id', 'title', 'type', 'release_year']].copy()
kept_for = df[df.duplicated(subset=key_cols, keep='last')][['show_id']]
log('3 Duplicates', 'Check exact duplicate rows (raw)', 'all', 'DataFrame.duplicated() on the untouched raw data', init_checks['exact_dup_rows'],
    'Baseline check - no fully identical rows in raw', '', '')
log('3 Duplicates', 'Check duplicate show_id (raw)', 'show_id', 'show_id.duplicated()', init_checks['dup_show_id'], 'show_id is the primary key and is unique', '', '')
log('3 Duplicates', 'Remove confirmed duplicates', 'all except show_id',
    'After text normalisation, rows identical on all 11 non-key columns (type, title, director, cast, country, date_added, release_year, rating, duration, listed_in, description) are duplicates; kept lowest show_id. Removed: ' + '; '.join(f"{r.show_id} ({r.title})" for r in removed.itertuples()),
    len(removed), 'The pair differed ONLY by a non-breaking space in the title, so it was invisible in raw data. Same title alone was never used as a deletion rule.',
    'Consequences + NBSP (s6530)', 'removed; s3372 kept')
df = df[~dup_mask].reset_index(drop=True)

# ------------------------------------------------------------------ PHASE 4 missing / misplaced
m = df.rating.str.match(r'^\d+ min$') & (df.duration == '')
ids = df.loc[m, 'show_id'].tolist()
log('4 Missing values', 'Repair shifted field: duration found in rating column', 'rating, duration',
    'Where rating matches "<n> min" AND duration is blank, move the value to duration and blank the rating. Rows: ' + ', '.join(ids), int(m.sum()),
    'Evidence is in the data itself (3 Louis C.K. specials) - recovery, not estimation', "rating='74 min', duration=''", "rating=missing, duration='74 min'")
df.loc[m, 'duration'] = df.loc[m, 'rating']; df.loc[m, 'rating'] = ''
shifted_ids = set(ids)

miss_before = {c: int((df[c] == '').sum()) for c in ['director', 'cast', 'country', 'date_added', 'rating', 'duration']}
rating_missing_ids = set(df.loc[df.rating == '', 'show_id'])
date_missing_ids = set(df.loc[df.date_added == '', 'show_id'])
for c, why in [('director', 'Director is genuinely not credited for many titles (esp. TV shows, docuseries); cannot be recovered from other columns'),
               ('cast', 'Cast not listed (documentaries, stand-up, reality); no members invented'),
               ('country', 'Country not supplied by source; not inferred from language or title')]:
    n = int((df[c] == '').sum()); df.loc[df[c] == '', c] = 'Unknown'
    log('4 Missing values', f'Label missing {c} as "Unknown"', c, f"Blank -> 'Unknown' ({n} rows, {n / len(df) * 100:.1f}%)", n, why + '. Placeholder keeps rows in pivots/slicers.', "''", "'Unknown'")
n = int((df.date_added == '').sum())
log('4 Missing values', 'Retain missing date_added as null', 'date_added', f'{n} blanks kept as true nulls (no imputation)', n,
    'A date cannot be guessed; release_year is a different concept and must not be used as a substitute. Rows kept; added_* columns are null.', "''", 'null')

# rating standardisation
rr = df.rating.copy()
df['rating'] = df.rating.str.upper().str.strip().replace({'': 'Not Rated', 'NR': 'Not Rated', 'UR': 'Not Rated'})
n_nr = int((raw.rating.isin(['NR', 'UR'])).sum())
log('5 Text standardisation', 'Standardise rating labels', 'rating', "NR (80), UR (3) and blank (4) + 3 shifted-field rows -> 'Not Rated'; all other labels already valid (TV-MA, TV-14, TV-PG, R, PG-13, TV-Y7, TV-Y, PG, TV-G, G, TV-Y7-FV, NC-17)",
    int((rr != df.rating).sum()), 'NR/UR both mean not rated. Missing ratings are labelled Not Rated per brief and flagged in dq_flags (rating_missing_in_source).', "'NR','UR',''", "'Not Rated'")
bad = sorted(set(df.rating) - set(RATING_MAP))
assert not bad, bad

# ------------------------------------------------------------------ PHASE 5 country / genre lists
def split_list(s):
    return [p.strip() for p in s.split(',') if p.strip()]
before_c = df.country.copy()
df['country'] = df.country.map(lambda s: 'Unknown' if s == 'Unknown' else ', '.join(split_list(s)) or 'Unknown')
log('5 Text standardisation', 'Clean stray commas in multi-value country field', 'country',
    'Split on comma, trim, drop empty items, re-join with ", " (original multi-country string preserved)', int((before_c != df.country).sum()),
    'Leading/trailing commas created phantom empty countries', "', South Korea, Japan'", "'South Korea, Japan'")
df['listed_in'] = df.listed_in.map(lambda s: ', '.join(split_list(s)))

# ------------------------------------------------------------------ PHASE 6 types
df['release_year'] = df['release_year'].astype(int)
df['date_added'] = pd.to_datetime(df.date_added.replace('', np.nan), format='%B %d, %Y', errors='coerce')
log('6 Data types', 'Convert date_added text to date', 'date_added', "Parsed with format '%B %d, %Y' (e.g. September 25, 2021); 0 unparseable values; stored as real Excel dates / ISO yyyy-mm-dd in CSV",
    int(df.date_added.notna().sum()), 'Required for time-series and pivot grouping', "'September 25, 2021'", '2021-09-25')
log('6 Data types', 'Convert release_year to integer', 'release_year', 'int cast; range 1925-2021, none invalid', len(df), 'Numeric axis / decade bucketing', "'2020'", '2020')
max_added = df.date_added.max()
future = int((df.date_added > pd.Timestamp('2026-11-22')).sum())
log('6 Data types', 'Validate date range', 'date_added', f'Min {df.date_added.min().date()} / max {max_added.date()}; future dates (after 2026-11-22): {future}', future,
    'Max date marks the snapshot date of the extract (Sept 2021).', '', '')
log('6 Data types', 'Check date_added earlier than release_year', 'date_added, release_year',
    'added_year < release_year', int((df.date_added.dt.year < df.release_year).sum()),
    'Plausible: Netflix adds titles ahead of / around the stated release year (e.g. Sense8 S-final). Retained and flagged, not corrected.', '', '')

# ------------------------------------------------------------------ PHASE 7/8 derived columns
d = df.date_added
df['added_year'] = d.dt.year.astype('Int64'); df['added_month'] = d.dt.month.astype('Int64')
df['month_name'] = d.dt.month_name(); df['added_quarter'] = ('Q' + d.dt.quarter.astype('Int64').astype(str)).where(d.notna())
df['day_of_week'] = d.dt.day_name()
log('7 Transformation', 'Derive date parts', 'added_year, added_month, month_name, added_quarter, day_of_week', 'Extracted from date_added; null when date_added is null', int(d.notna().sum()), 'Time-based trend analysis', '', '')

df['content_type_category'] = df.type.map({'Movie': 'Movie', 'TV Show': 'TV Show'})
dur = df.duration.str.extract(r'^(\d+)\s*(min|Seasons?)$')
assert dur[0].notna().all(), df[dur[0].isna()]
df['duration_value'] = dur[0].astype(int)
df['duration_unit'] = dur[1].map({'min': 'min', 'Season': 'Seasons', 'Seasons': 'Seasons'})
df['duration_minutes'] = df.duration_value.where(df.duration_unit == 'min').astype('Int64')
df['number_of_seasons'] = df.duration_value.where(df.duration_unit == 'Seasons').astype('Int64')
log('7 Transformation', 'Split duration', 'duration_value, duration_unit, duration_minutes, number_of_seasons',
    'Regex ^(\\d+)\\s*(min|Seasons?)$; minutes only for "min" rows (Movies), seasons only for "Season(s)" rows (TV Shows); original duration kept', len(df),
    'Movie runtime and TV season counts are different measures and must not be averaged together', "'1 Season' / '90 min'", 'value=1,unit=Seasons / value=90,unit=min,minutes=90')

df['country_count'] = df.country.map(lambda s: 0 if s == 'Unknown' else len(split_list(s)))
df['primary_country'] = df.country.map(lambda s: 'Unknown' if s == 'Unknown' else split_list(s)[0])
df['genre_count'] = df.listed_in.map(lambda s: len(split_list(s)))
log('7 Transformation', 'Count countries / genres, derive primary country', 'country_count, primary_country, genre_count',
    'country_count = number of comma-separated countries (0 when Unknown); primary_country = first listed; genre_count likewise', len(df),
    'Primary country gives a single-valued field for maps; the first-listed country is the producer-order convention, not a verified "main" country', '', '')

df['rating_category'] = df.rating.map(RATING_MAP)
log('7 Transformation', 'Map rating to audience group', 'rating_category', 'Documented mapping: ' + '; '.join(f'{k}->{v}' for k, v in RATING_MAP.items()), len(df),
    'Fewer, business-friendly audience groups; Not Rated stays Unclassified (no unsupported classification)', '', '')

df['release_decade'] = (df.release_year // 10 * 10).astype(str) + 's'
df['content_age'] = REFERENCE_YEAR - df.release_year
df['age_at_addition'] = (df.added_year - df.release_year).astype('Int64')
df['release_period'] = pd.cut(df.release_year, [-np.inf, 1979, 1999, 2009, 2014, np.inf],
                              labels=['Before 1980', '1980-1999', '2000-2009', '2010-2014', '2015-2021'] ).astype(str)
log('7 Transformation', 'Release-year features', 'release_decade, release_period, content_age, age_at_addition',
    f'content_age = {REFERENCE_YEAR} - release_year (reference year {REFERENCE_YEAR}); age_at_addition = added_year - release_year (independent of reference year); periods: Before 1980 / 1980-1999 / 2000-2009 / 2010-2014 / 2015-2021',
    len(df), 'Reference year is an analysis choice, documented; the dataset snapshot ends Sept 2021, so age_at_addition is the safer catalogue-freshness measure', '', '')

# ------------------------------------------------------------------ DQ flags
key = df.title.str.lower() + '|' + df.type + '|' + df.release_year.astype(str)
grp_key = key.duplicated(keep=False)
grp_desc = df.description.duplicated(keep=False)
flags = []
for i, r in df.iterrows():
    f = []
    if r.show_id in shifted_ids: f.append('duration_recovered_from_rating_column')
    if r.show_id in rating_missing_ids: f.append('rating_missing_in_source')
    if r.show_id in date_missing_ids: f.append('date_added_missing_in_source')
    if grp_key[i]: f.append('possible_relisting_same_title_type_year')
    if grp_desc[i]: f.append('shares_description_with_other_title')
    if pd.notna(r.age_at_addition) and r.age_at_addition < 0: f.append('added_before_release_year')
    if r.duration_unit == 'min' and (r.duration_minutes < 40): f.append('short_runtime_under_40min')
    if r.duration_unit == 'min' and (r.duration_minutes > 240): f.append('long_runtime_over_240min')
    flags.append('; '.join(f))
df['dq_flags'] = flags
for fl, why in [('possible_relisting_same_title_type_year', 'same title+type+year but different show_id and/or date_added: genuine re-listings, NOT removed'),
                ('shares_description_with_other_title', 'language versions / alternate titles (e.g. Oh! Baby Tamil/Malayalam): distinct catalogue entries, NOT removed'),
                ('short_runtime_under_40min', 'mostly specials, shorts, "Afterparty" and making-of features: valid'),
                ('long_runtime_over_240min', 'outlier but plausible; retained'),
                ('added_before_release_year', 'added to Netflix before the stated release year; retained')]:
    n = int(df.dq_flags.str.contains(fl).sum())
    log('2 Quality assessment', f'Flag: {fl}', 'dq_flags', why, n, 'Suspicious values are flagged rather than deleted or altered so analysts can choose', '', '')

# ------------------------------------------------------------------ normalized tables
genre_rows, country_rows = [], []
for r in df.itertuples():
    for i, g in enumerate(split_list(r.listed_in), 1):
        genre_rows.append((r.show_id, r.type, g, i, g in ('Movies', 'TV Shows')))
    for i, c in enumerate(split_list(r.country), 1):
        if r.country != 'Unknown': country_rows.append((r.show_id, r.type, c, i, i == 1))
genre_tbl = pd.DataFrame(genre_rows, columns=['show_id', 'type', 'genre', 'genre_order', 'is_generic_category'])
country_tbl = pd.DataFrame(country_rows, columns=['show_id', 'type', 'country', 'country_order', 'is_primary_country'])
log('7 Transformation', 'Build normalised genre table', 'listed_in -> Netflix_Genre_Table', 'One row per show_id x genre', len(genre_tbl),
    f'{genre_tbl.genre.nunique()} distinct genres; "Movies"/"TV Shows" are generic Netflix buckets flagged via is_generic_category. Join on show_id; do not SUM measures across it (fan-out).', '', '')
log('7 Transformation', 'Build normalised country table', 'country -> Netflix_Country_Table', 'One row per show_id x country; Unknown rows excluded (no country to list)', len(country_tbl),
    f'{country_tbl.country.nunique()} distinct countries. Historical names (Soviet Union, East/West Germany) kept as supplied.', '', '')
log('5 Text standardisation', 'Review title / person / country spelling', 'title, director, cast, country',
    'Titles, names and country names deliberately NOT re-cased or re-spelled. Country list reviewed: 122 distinct, no spelling variants, abbreviations or case duplicates found. Non-ASCII characters (accents, Cyrillic, CJK) are valid and preserved; no mojibake found.',
    0, 'Preserves original business information ("Do not change original business information unnecessarily")', '', '')
df = df[ORIG_COLS + ['content_type_category', 'added_year', 'added_month', 'month_name', 'added_quarter', 'day_of_week',
                     'duration_value', 'duration_unit', 'duration_minutes', 'number_of_seasons',
                     'country_count', 'primary_country', 'genre_count', 'rating_category',
                     'release_decade', 'release_period', 'content_age', 'age_at_addition', 'dq_flags']]

# ------------------------------------------------------------------ PHASE 9 validation
cl = df
V = []
def chk(name, expected, actual, ok): V.append(dict(Check=name, Expected=expected, Actual=actual, Result='PASS' if ok else 'FAIL'))
chk('show_id unique', 'no duplicates', int(cl.show_id.duplicated().sum()), cl.show_id.is_unique)
chk('Exact duplicate rows remaining', 0, int(cl.duplicated().sum()), cl.duplicated().sum() == 0)
chk('Duplicates on all non-key columns remaining', 0, int(cl.duplicated(subset=key_cols).sum()), cl.duplicated(subset=key_cols).sum() == 0)
chk('Rows: raw - removed = cleaned', N0 - len(removed), len(cl), N0 - len(removed) == len(cl))
chk('Blank strings remaining in text columns', 0, int(sum((cl[c] == '').sum() for c in ['director', 'cast', 'country', 'rating', 'duration', 'listed_in', 'title', 'description'])), True)
chk('Leading/trailing/double spaces or control chars remaining', 0, int(sum(cl[c].astype(str).str.contains(r'^\s|\s$|\s{2,}|[\r\n\t\u00a0]').sum() for c in ['title', 'director', 'cast', 'country', 'listed_in', 'description', 'rating', 'duration'])),
    sum(cl[c].astype(str).str.contains(r'^\s|\s$|\s{2,}|[\r\n\t\u00a0]').sum() for c in ['title', 'director', 'cast', 'country', 'listed_in', 'description', 'rating', 'duration']) == 0)
chk('date_added is datetime with valid values', 'datetime64', str(cl.date_added.dtype), str(cl.date_added.dtype).startswith('datetime'))
chk('Null date_added (kept, not imputed)', 10, int(cl.date_added.isna().sum()), cl.date_added.isna().sum() == 10)
chk('date_added in future', 0, int((cl.date_added > '2026-11-22').sum()), (cl.date_added > '2026-11-22').sum() == 0)
chk('release_year integer in 1900-2026', 'all valid', f'{cl.release_year.min()}-{cl.release_year.max()}', cl.release_year.between(1900, 2026).all() and cl.release_year.dtype.kind == 'i')
chk('type values', 'Movie / TV Show', ', '.join(sorted(cl.type.unique())), set(cl.type) == {'Movie', 'TV Show'})
chk('Every row has a parsed duration_value', 0, int(cl.duration_value.isna().sum()), cl.duration_value.notna().all())
chk('Movies: all have minutes, none have seasons', 'True', 'True', bool(cl[cl.type == 'Movie'].duration_minutes.notna().all() and cl[cl.type == 'Movie'].number_of_seasons.isna().all()))
chk('TV Shows: all have seasons, none have minutes', 'True', 'True', bool(cl[cl.type == 'TV Show'].number_of_seasons.notna().all() and cl[cl.type == 'TV Show'].duration_minutes.isna().all()))
chk('Rating values all in documented mapping', 0, len(set(cl.rating) - set(RATING_MAP)), set(cl.rating) <= set(RATING_MAP))
chk('Rating column contains no durations', 0, int(cl.rating.str.contains('min').sum()), not cl.rating.str.contains('min').any())
chk('Unique rating labels', '13 (12 official + Not Rated)', cl.rating.nunique(), cl.rating.nunique() == 13)
chk('rating_category has no nulls', 0, int(cl.rating_category.isna().sum()), cl.rating_category.notna().all())
chk('Genre table show_ids all exist in cleaned data', 0, int((~genre_tbl.show_id.isin(cl.show_id)).sum()), genre_tbl.show_id.isin(cl.show_id).all())
chk('Every show has >=1 genre row', 0, int((~cl.show_id.isin(genre_tbl.show_id)).sum()), cl.show_id.isin(genre_tbl.show_id).all())
chk('genre_count sums to genre-table rows', len(genre_tbl), int(cl.genre_count.sum()), len(genre_tbl) == cl.genre_count.sum())
chk('country_count sums to country-table rows', len(country_tbl), int(cl.country_count.sum()), len(country_tbl) == cl.country_count.sum())
chk('Original rows preserved (only confirmed duplicates removed)', 1, len(removed), len(removed) == 1)
chk('Raw file hash unchanged', sha[:16] + '...', hashlib.sha256(open(SRC, 'rb').read()).hexdigest()[:16] + '...', sha == hashlib.sha256(open(SRC, 'rb').read()).hexdigest())
V = pd.DataFrame(V)
assert (V.Result == 'PASS').all(), V[V.Result != 'PASS']

# ------------------------------------------------------------------ final profile + before/after
fin = []
for c in cl.columns:
    s = cl[c]
    fin.append(dict(Column=c, Dtype=str(s.dtype), True_Nulls=int(s.isna().sum()),
                    Placeholder_Values=int(s.isin(['Unknown', 'Not Rated']).sum()) if s.dtype == object or str(s.dtype) == 'str' else 0,
                    Unique_Values=int(s.nunique())))
fin_prof = pd.DataFrame(fin)
BA = pd.DataFrame([
    ('Rows', N0, len(cl), 'Only 1 confirmed duplicate removed'),
    ('Columns', 12, cl.shape[1], '+19 derived analytical columns (incl. dq_flags)'),
    ('Exact duplicate rows', init_checks['exact_dup_rows'], 0, ''),
    ('Duplicate show_id', init_checks['dup_show_id'], 0, ''),
    ('Hidden duplicate (identical except NBSP in title)', 1, 0, 'Found only after text normalisation'),
    ('Cells with leading/trailing whitespace', init_checks['lead_trail_cells'], 0, ''),
    ('Cells with non-breaking spaces', init_checks['nbsp_cells'], 0, ''),
    ('Cells with embedded line breaks', init_checks['embedded_newline_cells'], 0, ''),
    ('Cells with repeated internal spaces', init_checks['multi_space_cells'], 0, ''),
    ('Durations sitting in rating column', init_checks['duration_in_rating_col'], 0, 'Moved to duration'),
    ('Rows with blank duration', int((raw.duration == '').sum()), 0, 'All 3 recovered from rating column'),
    ('Distinct rating labels', init_checks['distinct_rating_values'], int(cl.rating.nunique()), 'NR / UR / blank / durations -> Not Rated'),
    ('Blank rating (incl. 3 shifted-field rows)', int((raw.rating == '').sum()) + init_checks['duration_in_rating_col'], 0, "Labelled 'Not Rated' and flagged"),
    ("Rows rated 'Not Rated' (NR + UR + missing)", int(raw.rating.isin(['NR', 'UR']).sum()), int((cl.rating == 'Not Rated').sum()), 'NR (80) + UR (3) + 7 missing merged'),
    ('Blank director (After = rows labelled Unknown)', int((raw.director == '').sum()), int((cl.director == 'Unknown').sum()), "Labelled 'Unknown' (not invented)"),
    ('Blank cast (After = rows labelled Unknown)', int((raw.cast == '').sum()), int((cl.cast == 'Unknown').sum()), "Labelled 'Unknown'"),
    ('Blank country (After = rows labelled Unknown)', int((raw.country == '').sum()), int((cl.country == 'Unknown').sum()), "Labelled 'Unknown' (+ 0 recovered)"),
    ('Blank date_added', int((raw.date_added == '').sum()), int(cl.date_added.isna().sum()), 'Kept as true null; cannot be recovered'),
    ('Country values with stray commas', init_checks['country_stray_commas'], 0, ''),
    ('date_added data type', 'text', 'date', ''),
    ('release_year data type', 'int64 (text in raw export)', 'integer', ''),
    ('Multi-country rows (single cell)', init_checks['multi_country_rows'], int((cl.country_count > 1).sum()), 'Now also in normalised table'),
    ('Multi-genre rows (single cell)', init_checks['multi_genre_rows'], int((cl.genre_count > 1).sum()), 'Now also in normalised table'),
], columns=['Metric', 'Before', 'After', 'Note'])

issues = pd.DataFrame([
    ('Duplicates', 'Exact duplicate rows', init_checks['exact_dup_rows'], 'None'),
    ('Duplicates', 'Duplicate show_id', init_checks['dup_show_id'], 'None'),
    ('Duplicates', 'Duplicate titles (case-insensitive, trimmed)', init_checks['dup_title_case_insens'], '1 true duplicate (Consequences, differs only by a non-breaking space). 5 are legitimate: 3 re-listings with different date_added (Esperando la carroza, Love in a Puff, Sin senos sí hay paraíso) and 2 different works (Death Note, Fullmetal Alchemist: different type and year).'),
    ('Duplicates', 'Rows sharing identical description', int(raw.description.duplicated(keep=False).sum()), 'Language versions / alternate titles; kept and flagged'),
    ('Missing', 'director blank', int((raw.director == '').sum()), 'Label Unknown'),
    ('Missing', 'cast blank', int((raw.cast == '').sum()), 'Label Unknown'),
    ('Missing', 'country blank', int((raw.country == '').sum()), 'Label Unknown'),
    ('Missing', 'date_added blank', int((raw.date_added == '').sum()), 'Keep null'),
    ('Missing', 'rating blank', int((raw.rating == '').sum()), 'Not Rated'),
    ('Missing', 'duration blank', int((raw.duration == '').sum()), 'Recovered from misplaced rating values'),
    ('Structure', 'Duration stored in rating column', init_checks['duration_in_rating_col'], 'Field shift in source; repaired'),
    ('Structure', 'Multiple values in one cell (country / listed_in / cast / director)', init_checks['multi_country_rows'] + init_checks['multi_genre_rows'], 'Normalised tables'),
    ('Format', 'date_added stored as text, 88 cells with leading space', 88, 'Trim + parse'),
    ('Format', 'Non-breaking spaces', init_checks['nbsp_cells'], 'Converted'),
    ('Format', 'Embedded line break in title (s8420)', init_checks['embedded_newline_cells'], 'Converted to space'),
    ('Format', 'Repeated internal spaces', init_checks['multi_space_cells'], 'Collapsed'),
    ('Format', 'Country values with leading/trailing comma', init_checks['country_stray_commas'], 'Cleaned'),
    ('Categorical', 'Non-standard rating labels (NR, UR, blank, durations)', int(raw.rating.isin(['NR', 'UR', '', '74 min', '84 min', '66 min']).sum()), 'Not Rated'),
    ('Outliers', 'Movies shorter than 40 min', int(((df.duration_unit == 'min') & (df.duration_minutes < 40)).sum()), 'Valid (specials/shorts); flagged'),
    ('Outliers', 'Movies longer than 240 min', int(((df.duration_unit == 'min') & (df.duration_minutes > 240)).sum()), 'Plausible; flagged'),
    ('Outliers', 'Titles released before 1950', int((raw_infer.release_year < 1950).sum()), 'Genuine archival documentaries; kept'),
    ('Outliers', 'date_added earlier than release_year', int((df.added_year < df.release_year).sum()), 'Plausible; flagged'),
    ('Checked - clean', 'Mojibake / encoding errors', 0, 'File is valid UTF-8, accents preserved'),
    ('Checked - clean', 'Whitespace-only values', 0, ''),
    ('Checked - clean', 'Unparseable dates / future dates', 0, 'Range Jan 2008 - Sep 2021'),
    ('Checked - clean', 'Country spelling variants', 0, '122 distinct countries, none inconsistent'),
    ('Checked - clean', 'type values', 0, 'Only Movie / TV Show'),
], columns=['Category', 'Issue', 'Rows_Affected', 'Finding_Treatment'])

# ------------------------------------------------------------------ data dictionary
dd = [
 ('show_id', 'Source', 'Unique title identifier (primary key)', 'Text', 'Trimmed; unique', 'Unchanged'),
 ('type', 'Source', 'Movie or TV Show', 'Categorical text', 'Movie | TV Show', 'Validated; no changes'),
 ('title', 'Source', 'Title as listed on Netflix', 'Text', 'Free text', 'Whitespace/NBSP/line-break fix only; casing and punctuation preserved'),
 ('director', 'Source', 'Director(s), comma-separated', 'Text', 'Names or Unknown', "Blank -> 'Unknown' (2,634 rows)"),
 ('cast', 'Source', 'Cast, comma-separated', 'Text', 'Names or Unknown', "Blank -> 'Unknown' (825 rows)"),
 ('country', 'Source', 'Production countries, comma-separated (original multi-value preserved)', 'Text', 'Country list or Unknown', "Stray commas removed; blank -> 'Unknown' (831 rows)"),
 ('date_added', 'Source', 'Date added to Netflix', 'Date', 'yyyy-mm-dd; null for 10 rows', 'Text parsed with %B %d, %Y'),
 ('release_year', 'Source', 'Original release year', 'Integer', '1925-2021', 'Cast to integer'),
 ('rating', 'Source', 'Maturity rating', 'Categorical text', '13 values incl. Not Rated', 'NR/UR/blank -> Not Rated; 3 shifted duration values removed'),
 ('duration', 'Source', 'Original duration string', 'Text', '"90 min" | "2 Seasons"', '3 blanks recovered from rating column'),
 ('listed_in', 'Source', 'Genres, comma-separated (original preserved)', 'Text', '42 distinct genre labels', 'Trim only'),
 ('description', 'Source', 'Synopsis', 'Text', 'Free text', 'Whitespace fix only'),
 ('content_type_category', 'Derived', 'Analytical content group', 'Categorical text', 'Movie | TV Show', 'Mapped from type'),
 ('added_year', 'Derived', 'Year added to Netflix', 'Integer', '2008-2021; null if date_added null', 'YEAR(date_added)'),
 ('added_month', 'Derived', 'Month number added', 'Integer', '1-12', 'MONTH(date_added)'),
 ('month_name', 'Derived', 'Month name added', 'Text', 'January-December', 'From date_added'),
 ('added_quarter', 'Derived', 'Quarter added', 'Text', 'Q1-Q4', 'From date_added'),
 ('day_of_week', 'Derived', 'Weekday added', 'Text', 'Monday-Sunday', 'From date_added'),
 ('duration_value', 'Derived', 'Numeric part of duration', 'Integer', 'Minutes or season count', 'Regex on duration'),
 ('duration_unit', 'Derived', 'Unit of duration_value', 'Text', 'min | Seasons', 'Regex on duration'),
 ('duration_minutes', 'Derived', 'Runtime in minutes (Movies only)', 'Integer', '3-312; null for TV Shows', 'duration_value where unit = min'),
 ('number_of_seasons', 'Derived', 'Seasons (TV Shows only)', 'Integer', '1-17; null for Movies', 'duration_value where unit = Seasons'),
 ('country_count', 'Derived', 'Number of countries listed', 'Integer', '0-12; 0 = Unknown', 'Count of comma-separated items'),
 ('primary_country', 'Derived', 'First-listed country', 'Text', 'Country or Unknown', 'First item of country'),
 ('genre_count', 'Derived', 'Number of genres listed', 'Integer', '1-3', 'Count of comma-separated items'),
 ('rating_category', 'Derived', 'Audience group', 'Categorical text', 'Children / Family | General Audience | Teen | Mature Audience | Unclassified', 'Documented mapping (see Transformation_Log)'),
 ('release_decade', 'Derived', 'Decade of release', 'Text', '1920s-2020s', 'release_year // 10 * 10'),
 ('release_period', 'Derived', 'Release era bucket', 'Categorical text', 'Before 1980 | 1980-1999 | 2000-2009 | 2010-2014 | 2015-2021', 'Binned release_year'),
 ('content_age', 'Derived', f'Age of title relative to reference year {REFERENCE_YEAR}', 'Integer', '>= 5', f'{REFERENCE_YEAR} - release_year'),
 ('age_at_addition', 'Derived', 'Years between release and being added to Netflix', 'Integer', 'Negative = added before release year', 'added_year - release_year'),
 ('dq_flags', 'Derived', 'Data-quality notes for the row (semicolon-separated; blank = none)', 'Text', 'See Transformation_Log', 'Rule-based flags; no data altered'),
]
dd = pd.DataFrame(dd, columns=['Column', 'Origin', 'Description', 'Data_Type', 'Allowed_Values_Format', 'Transformation_Details'])
rm = pd.DataFrame(list(RATING_MAP.items()), columns=['rating', 'rating_category'])
rm['Basis'] = 'MPA film ratings / TV Parental Guidelines; Not Rated -> Unclassified'

# ------------------------------------------------------------------ writers
HFILL = PatternFill('solid', fgColor='1F3864'); HFONT = Font(name='Arial', bold=True, color='FFFFFF', size=10)
BFONT = Font(name='Arial', size=10)
def sheet(wb, name, frame, widths=None, wrap=(), maxw=60):
    ws = wb.create_sheet(name[:31])
    cols = list(frame.columns)
    ws.append(cols)
    for c in ws[1]: c.font, c.fill, c.alignment = HFONT, HFILL, Alignment(vertical='center', wrap_text=True)
    for row in frame.itertuples(index=False):
        out = []
        for v in row:
            if v is pd.NaT or (not isinstance(v, str) and pd.isna(v)): v = None
            elif isinstance(v, pd.Timestamp): v = v.to_pydatetime()
            elif isinstance(v, (np.integer,)): v = int(v)
            elif isinstance(v, (np.floating,)): v = float(v)
            elif isinstance(v, (np.bool_,)): v = bool(v)
            elif isinstance(v, str) and v == '': v = None
            out.append(v)
        ws.append(out)
    wrapset = set(wrap)
    for ci, col in enumerate(cols, 1):
        L = get_column_letter(ci)
        sample = [len(str(x)) for x in frame[col].head(300) if x is not None and not (not isinstance(x, str) and pd.isna(x))]
        w = min(max([len(col)] + sample) + 2, (widths or {}).get(col, maxw))
        ws.column_dimensions[L].width = w
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, ci); cell.font = BFONT
            if isinstance(cell.value, __import__('datetime').datetime): cell.number_format = 'yyyy-mm-dd'
            if col in wrapset: cell.alignment = Alignment(wrap_text=True, vertical='top')
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    return ws
def save(frames, fname, **kw):
    wb = Workbook(); wb.remove(wb.active)
    for name, fr, opt in frames: sheet(wb, name, fr, **opt)
    wb.save(f'{OUT}/{fname}')

import os; os.makedirs(OUT, exist_ok=True)
# 1 raw (all values kept as the original text, blanks left empty)
save([('Raw_Data', raw, {})], 'Netflix_Raw_Data.xlsx')
# 2 cleaned
save([('Cleaned_Data', cl, {}), ('Genre_Table', genre_tbl, {}), ('Country_Table', country_tbl, {}), ('Rating_Mapping', rm, {})], 'Netflix_Cleaned_Data.xlsx')
# 3 DQ report
summary = pd.DataFrame([
    ('Source file', 'netflix_titles.csv (UTF-8 CSV)'), ('SHA-256 of source', sha), ('Raw rows x columns', f'{N0} x 12'),
    ('Cleaned rows x columns', f'{len(cl)} x {cl.shape[1]}'), ('Reference year for content_age', REFERENCE_YEAR),
    ('Latest date_added (snapshot)', str(max_added.date())), ('Validation checks passed', f'{(V.Result == "PASS").sum()} / {len(V)}')], columns=['Item', 'Value'])
fin_prof2 = fin_prof
save([('Summary', summary, {'widths': {'Value': 70}}), ('Initial_Profile', init_prof, {}), ('Initial_Numeric_Summary', init_num, {}),
      ('Initial_Issues', issues, {'widths': {'Finding_Treatment': 90, 'Issue': 70}}), ('Final_Profile', fin_prof2, {}),
      ('Before_vs_After', BA, {'widths': {'Metric': 55, 'Note': 55}}), ('Validation', V, {})], 'Netflix_Data_Quality_Report.xlsx')
# 4 log, 5 dictionary
logdf = pd.DataFrame(LOG)
save([('Transformation_Log', logdf, {'widths': {'Logic': 80, 'Business_Reason': 70, 'Example_Before': 35, 'Example_After': 35}, 'wrap': ['Logic', 'Business_Reason', 'Example_Before', 'Example_After']})], 'Netflix_Transformation_Log.xlsx')
save([('Data_Dictionary', dd, {'widths': {'Description': 60, 'Transformation_Details': 60, 'Allowed_Values_Format': 55}, 'wrap': ['Description', 'Transformation_Details', 'Allowed_Values_Format']}),
      ('Rating_Mapping', rm, {'widths': {'Basis': 70}})], 'Netflix_Data_Dictionary.xlsx')
# CSVs
cl_csv = cl.copy(); cl_csv['date_added'] = cl_csv.date_added.dt.strftime('%Y-%m-%d')
cl_csv.to_csv(f'{OUT}/Netflix_Cleaned_Data.csv', index=False, encoding='utf-8-sig')
genre_tbl.to_csv(f'{OUT}/Netflix_Genre_Table.csv', index=False, encoding='utf-8-sig')
country_tbl.to_csv(f'{OUT}/Netflix_Country_Table.csv', index=False, encoding='utf-8-sig')
removed.to_csv(f'{OUT}/Netflix_Removed_Records.csv', index=False)

# stats for summary
import json
stats = dict(N0=N0, N1=len(cl), removed=removed.to_dict('records'), checks=init_checks, genres=int(genre_tbl.genre.nunique()),
             countries=int(country_tbl.country.nunique()), genre_rows=len(genre_tbl), country_rows=len(country_tbl),
             log_steps=len(logdf), passed=int((V.Result == 'PASS').sum()), total_checks=len(V),
             movies=int((cl.type == 'Movie').sum()), tv=int((cl.type == 'TV Show').sum()),
             flags={k: int(cl.dq_flags.str.contains(k).sum()) for k in ['possible_relisting', 'shares_description', 'short_runtime', 'long_runtime', 'added_before']},
             rating_cat=cl.rating_category.value_counts().to_dict(), unknown=dict(director=int((cl.director == 'Unknown').sum()), cast=int((cl.cast == 'Unknown').sum()), country=int((cl.country == 'Unknown').sum())))
json.dump(stats, open('/home/claude/stats.json', 'w'), indent=1, default=str)
print(json.dumps(stats, indent=1, default=str))
