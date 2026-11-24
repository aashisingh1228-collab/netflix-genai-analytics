# Netflix Movies & TV Shows – Data Cleaning & Transformation: Project Summary

## 1. Business objective
Prepare a trustworthy, analysis-ready Netflix catalogue dataset for Excel pivots/dashboards, Power BI, SQL and EDA, without deleting valid records or inventing data.

## 2. Dataset overview
- Source: `netflix_titles.csv` (UTF-8 CSV), 8,807 rows x 12 columns (SHA-256 recorded in the Quality Report; raw file never modified).
- 6,131 Movies and 2,676 TV Shows in the raw file (6,130 Movies after the duplicate was removed); release years 1925-2021; titles added Jan 2008 - Sep 25 2021 (the extract's snapshot date).

## 3. Initial data-quality issues (all verified on the data)
| Issue | Count |
|---|---|
| Exact duplicate rows / duplicate show_id | 0 / 0 |
| Hidden duplicate (identical except a non-breaking space in the title) | 1 pair |
| Blank director / cast / country / date_added | 2,634 / 825 / 831 / 10 |
| Blank rating / blank duration | 4 / 3 |
| Duration values sitting in the rating column (3 Louis C.K. specials) | 3 |
| Non-standard rating labels (NR, UR, blank, durations) | 90 |
| Cells with leading/trailing spaces / NBSP / line breaks / repeated spaces | 89 / 14 / 2 / 17 |
| Country values with stray leading/trailing commas | 7 |
| date_added stored as text | all |
| Multi-value cells (country / genre) | 1,320 / 6,787 rows |

Checked and found clean: encoding (no mojibake), whitespace-only values, unparseable or future dates, country spelling variants, `type` values.

## 4. Cleaning operations
1. Normalised hidden Unicode spaces, embedded line breaks, repeated and edge spaces.
2. Removed 1 confirmed duplicate (s6530 "Consequences"; identical on all 11 non-key columns once the NBSP was fixed). No record was removed for sharing a title alone.
3. Repaired 3 shifted fields (duration moved from rating to duration).
4. Missing values: director/cast/country -> `Unknown`; rating -> `Not Rated` (flagged); date_added kept null; nothing imputed.
5. Rating labels standardised (NR/UR/blank -> Not Rated); country comma errors fixed.
6. Titles, names and country spellings intentionally not re-cased.

## 5. Transformations
date_added -> date; release_year -> integer; added_year/month/month_name/quarter/day_of_week; duration_value/unit/minutes (Movies only)/number_of_seasons (TV only); country_count, primary_country, genre_count; rating_category (documented mapping); release_decade, release_period, content_age (reference year 2026), age_at_addition; dq_flags; normalised Genre table (19,320 rows, 42 genres) and Country table (10,011 rows, 122 countries).

## 6. Validation
24 of 24 checks pass (unique keys, no remaining duplicates, valid dates and integers, movies vs seasons separated, ratings in mapping, normalised tables reconcile with count columns, raw hash unchanged). See `Netflix_Data_Quality_Report.xlsx`.

## 7. Before vs after
Records 8,807 -> 8,806. Rating labels 18 -> 13. Whitespace / NBSP / line-break / repeated-space cells: all fixed (0 remaining). Misplaced durations 3 -> 0. Columns 12 -> 31.

## 8. Judgement calls to be aware of
- Missing rating is labelled `Not Rated`, as the brief asks, but it is not the same as an official NR rating; `dq_flags` marks the 7 rows where it was missing in the source.
- 3 re-listed titles (same title/type/year, different date_added) and 57 rows sharing a description (language versions) are flagged, not removed.
- `content_age` uses 2026 as the analysis year; the data ends in Sept 2021, so `age_at_addition` is the better freshness measure.
- `primary_country` is simply the first-listed country.
- Genre and country tables fan out rows: count titles with DISTINCT show_id, never sum measures across them.
- 170 short films/specials (<40 min) and 14 titles added before their release year are valid and flagged.

## 9. Recommended next steps
Content growth by added_year and type; country and genre mix over time; rating_category by country; movie runtime and season-count distributions; release-to-addition lag by decade; director/cast network tables (not built here: they would need the same normalisation); Power BI star schema using show_id.

## 10. Reproducibility
`python clean_netflix.py <netflix_titles.csv> <output_dir>` regenerates every file.
