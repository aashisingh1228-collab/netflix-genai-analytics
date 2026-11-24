# Netflix Movies & TV Shows: Data Cleaning, Transformation & Interactive Dashboard

An end-to-end data analyst portfolio project: a raw Netflix catalogue export is audited, cleaned, documented and turned into an interactive analytics dashboard.

## Project overview

| Stage | What was done |
|---|---|
| 1. Profiling | Row/column counts, data types, missing values, whitespace and duplicate checks on the raw file |
| 2. Cleaning | Hidden duplicate removed, shifted fields repaired, text normalised, missing values handled without inventing data |
| 3. Transformation | Date parts, duration split, rating groups, release-era columns, normalised genre and country tables |
| 4. Validation | 24 automated checks (all pass), plus a before/after quality report |
| 5. Dashboard | Single-file interactive HTML dashboard (Chart.js) built from the cleaned data |

## Key results

- **8,807 → 8,806 records.** Only one confirmed duplicate was removed: two "Consequences" rows that differed only by a non-breaking space in the title. No record was removed just for sharing a title.
- **3 shifted fields repaired.** Runtimes sitting in the `rating` column were moved to `duration`.
- **Missing values handled, not invented.** Director (2,634), cast (825) and country (831) blanks became `Unknown`. The 10 missing `date_added` values stayed null. 7 missing ratings became `Not Rated` and are flagged in `dq_flags`.
- **Final mix:** 6,130 Movies and 2,676 TV Shows.
- **Validation:** 24 of 24 checks pass, and the raw file's hash is unchanged.

## Repository contents

| File | Purpose |
|---|---|
| `netflix_titles.csv` | Original source file (input for `clean_netflix.py`) |
| `Netflix_Raw_Data.xlsx` | Original data, preserved without modification |
| `Netflix_Cleaned_Data.xlsx` / `.csv` | Cleaned and transformed dataset (31 columns) |
| `Netflix_Genre_Table.csv` | Normalised table, one row per title and genre |
| `Netflix_Country_Table.csv` | Normalised table, one row per title and country |
| `Netflix_Removed_Records.csv` | The records removed during cleaning |
| `Netflix_Data_Quality_Report.xlsx` | Initial profile, issue list, final profile, before/after and validation |
| `Netflix_Transformation_Log.xlsx` | Every cleaning and transformation step with logic, rows affected and reason |
| `Netflix_Data_Dictionary.xlsx` | Column descriptions, types and transformation details |
| `Netflix_Project_Summary.md` | Project write-up |
| `clean_netflix.py` | Reproducible cleaning pipeline |
| `Netflix_Analytics_Dashboard.html` | Interactive dashboard (self-contained) |

## Dashboard

Open `Netflix_Analytics_Dashboard.html` in any modern browser. It needs an internet connection because Chart.js loads from a CDN. All data is embedded in the file.

- **KPIs:** Total Titles, Movies, TV Shows, average movie runtime, average seasons per show
- **Charts:** Movies vs TV Shows, content added by year, release-year trend, rating distribution, top 10 genres, top 10 countries, movie duration, TV season counts
- **Filters:** content type, release-year range, year added, country, genre, rating, plus a reset button. Clicking a genre, country or rating bar also filters.
- **Insights:** generated in the page from the filtered data, with threshold-based recommendations

Notes on the numbers:
- Total Titles counts unique `show_id` values. Country and genre bars count a title once per country or genre, so they can add up to more than the total.
- The data ends on 25 Sep 2021, so 2021 is a partial year.
- Movie runtime (minutes) and TV seasons are analysed separately.

To publish it with GitHub Pages, copy the file to `index.html` and enable Pages in the repository settings.

## Reproduce the cleaning

Requires Python 3.9+ with `pandas`, `numpy` and `openpyxl`.

```bash
pip install pandas numpy openpyxl
python clean_netflix.py path/to/netflix_titles.csv output_folder
```

The script reads the raw CSV (never modifies it) and regenerates the cleaned data, normalised tables, reports, log and dictionary.

## Design decisions

- Titles, names and country spellings were not re-cased, to preserve the original business information.
- `content_age` uses 2026 as its reference year. `age_at_addition` (year added minus release year) does not depend on that choice.
- Suspicious but plausible records (short runtimes, titles added before release year, re-listed titles, shared descriptions) are flagged in `dq_flags`, not deleted.

## Tools

Python (pandas, openpyxl), Excel, HTML/CSS/JavaScript, Chart.js
