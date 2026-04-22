# Grad School Pricing Analysis

Analysis of Emory graduate program tuition, scholarship discounting, and net revenue, with peer benchmarking against Columbia, NYU, and Georgia Tech.

## Overview

The project ingests raw student-level tuition and scholarship records from institutional Excel exports, filters to a defined set of LGS (Laney Graduate School) master's programs, and produces:

1. **Student-program-year-level data** — one row per (student ID, academic plan, academic year) with billed tuition, kept scholarship, net tuition, and a per-year term count. Enables cohort-year breakdowns.
2. **Student-program-level data** — one row per (student ID, academic plan) with gross tuition, kept scholarship, net tuition, international flag, and `first_ay / last_ay / ay_count / term_count` for career span.
3. **Program-level summary** — aggregated metrics per program pooled across all years: student counts, international %, avg gross/net tuition, discount rate, totals.
4. **Peer comparison** — Emory gross tuition annualized per program using `terms_per_year` from the Information Sheet, grouped into Data/CS, Economics, and General buckets, plotted against hardcoded peer tuition.

All analysis lives in a single Jupyter notebook; outputs are written back to the project root as `.xlsx`.

## Project Structure

```
grad-school-pricing/
├── analysis.ipynb                                                    # main notebook — all cleaning, aggregation, charts
│
├── Tuition Data.xlsx                                                 # primary input; 4 sheets, 2 used by the pipeline:
│                                                                     #   - "Gross Tuition Billed to Student"   (used)
│                                                                     #   - "Tuition Scholarship given to st"   (used)
│                                                                     #   - "Cheat Sheet - Programs - Term"     (reference — program keep/ignore flags, term-code → semester map)
│                                                                     #   - "Information Sheet"                 (reference — published tuition rates, program term counts)
├── Multi-Year Tuition-Tuition scholarship by Program.xlsx            # BYTE-IDENTICAL DUPLICATE of Tuition Data.xlsx (same md5); kept for naming/history but not a separate source
├── 2026 Bioethics Masters Tuition Rates and Degree Completion Time Summary.xlsx
├── Students' IDs.xlsx                                                # student ID reference
│
├── Information - Process.pdf                                         # process/methodology notes
├── 260219 Conversation with Surabhi Vittal.pdf                       # stakeholder conversation notes
│
├── program_summary_output.xlsx                                       # GENERATED — program-level summary (pooled across all years)
├── student_program_level_output.xlsx                                 # GENERATED — per-student, per-program rows (years collapsed, with first_ay/last_ay/ay_count/term_count)
└── student_program_year_output.xlsx                                  # GENERATED — per-student, per-program, per-academic-year rows (for cohort-year analysis)
```

## Data Pipeline (analysis.ipynb)

1. **Load** — read `Tuition Data.xlsx` sheets into `tuition` and `scholarship` DataFrames; strip whitespace from columns.
2. **Decode academic year** — `term_to_academic_year()` parses Emory PeopleSoft term codes `5YYS` (YY = last 2 digits of calendar year, S = 1=Spring / 6=Summer / 9=Fall). Fall rolls into the next AY end-year so the tuition side matches the scholarship sheet's `Aid Yr` convention (e.g. Fall 2018 + Spring 2019 + Summer 2019 all → AY 2019).
3. **Clean tuition** — keep `ID`, `acad_plan`, `term_code`, `tuition`, derived `academic_year`; coerce tuition to numeric; drop rows missing plan or amount.
4. **Clean scholarships** — sum `fall_sch + spring_sch + summer_sch` into a single `scholarship` column per row; rename `Aid Yr` → `academic_year`; keep `ID`, `acad_plan`, `descr`, `scholarship`, `academic_year`.
5. **Program filter** — both frames are filtered to `program_map` (see below).
6. **Scholarship filter** — `keep_scholarship()` retains descriptions containing `"lgs"` and drops external/federal awards and generic LGS "special" awards (`nih`, `nsf`, `grfp`, `training grant`, `special-scholarship`, `yellow ribbon`, `vet`, `americorps`, `pell`, `hope`, `zell`, `woodruff scholar-grad tuition`).
7. **Aggregate year-level (`df_year`)** — group to `(ID, acad_plan, academic_year)` for both tuition (with `term_count`) and scholarship; left-merge; unmatched scholarship rows get `0`.
8. **Derive year-level** — `net_tuition = max(tuition - scholarship, 0)` (clipped at zero — a kept scholarship occasionally exceeds billed tuition for a partial-term student; we report the student as paying $0 rather than a negative amount); `program` column via `program_map`.
9. **Collapse to student-program (`df`)** — group `df_year` to `(ID, acad_plan)` summing tuition/scholarship/term_count and computing `first_ay`, `last_ay`, `ay_count`. Recompute `net_tuition` with the same clip.
10. **International flag** — re-load scholarship sheet, mark `EU CC IPEDS == "Non US Citizen"` as `intl=1`, collapse to max per (ID, plan), merge into `df`.
11. **Program summary** — group `df` by `program`, compute student counts, international counts/%, avg gross/scholarship/net, totals, and `discount_rate = total_scholarship / total_gross_tuition`.
12. **Peer benchmark** — annualize Emory gross tuition per (ID, acad_plan) using `annual = total_tuition * terms_per_year / term_count` (program-specific terms/year from the Information Sheet + empirical seasonal distribution; see `terms_per_year` dict). Group programs into Data/CS / Economics / General, concatenate with hardcoded `peer_data` (Columbia, NYU, Georgia Tech).
13. **Charts** — discount rate by program, avg net tuition by program, international % vs discount rate scatter, gross-vs-net line chart, grouped bar chart of annual gross tuition by school × program group.
14. **Save** — `program_summary_output.xlsx`, `student_program_level_output.xlsx`, `student_program_year_output.xlsx`.

## Program Map

`acad_plan` codes are mapped to human-readable program names:

| Code          | Program              |
|---------------|----------------------|
| COMPSCIMS     | Computer Science     |
| CS4P1MS       | Computer Science 4+1 |
| DATASCIMS     | Data Science         |
| QTMMS         | Data Science         |
| ECONMS        | Economics            |
| ECON4P1MS     | Economics            |
| MATHMS        | Math                 |
| BIOETHMA      | Bioethics            |
| BIOETH4P1     | Bioethics 4+1        |
| DEVPRACMDP    | MDP                  |
| HUMANRTCRT    | MDP                  |
| BMIDMS        | MBID                 |
| BBS4P1MS      | Cancer Biology 4+1   |

Rows with any other `acad_plan` are dropped. Small-n programs are merged into a parent bucket where appropriate (e.g. `ECONMS` + `ECON4P1MS` → "Economics", `DATASCIMS` + `QTMMS` → "Data Science") so that per-program averages are not dominated by a handful of students.

Programs with no kept scholarship rows (e.g. `QTMMS`, `DATASCIMS`, `BMIDMS` in the current inputs) are **retained** in the program summary with a 0% discount rate. This is intentional — surfacing these zero-discount programs is part of the pricing analysis and flags potential discrepancies vs. programs that do give LGS aid.

## Terms-Per-Year (Annualization)

The peer-comparison annualization uses program-specific term counts rather than a blanket 2-term assumption. Sourced from the Information Sheet where listed; otherwise inferred from the empirical seasonal distribution (Fall / Spring / Summer row counts per plan).

| `acad_plan` | terms/yr | source |
|---|---|---|
| `COMPSCIMS`, `CS4P1MS`, `MATHMS`, `ECON4P1MS` | 2 | Info Sheet (Fall-Spring) |
| `BIOETHMA`, `BIOETH4P1`, `BMIDMS` | 2 | empirical (Fall-Spring dominant, minimal summer) |
| `ECONMS` | 3 | Info Sheet (Summer-Fall-Spring) |
| `DATASCIMS` | 3 | Info Sheet (Fall-Spring-Summer) |
| `QTMMS`, `DEVPRACMDP`, `HUMANRTCRT`, `BBS4P1MS` | 3 | empirical (all three seasons ~equal) |

Annualization is computed at the `(ID, acad_plan)` grain (not at the merged `program` grain) so that mixed-terms buckets like "Economics" (ECONMS=3 + ECON4P1MS=2) apply the correct rate to each student before averaging.

## Peer Group (hardcoded)

Annual tuition for Columbia, NYU, and Georgia Tech across three program buckets (Data/CS, Economics, General) is hardcoded in the notebook for benchmarking — not sourced from the Excel inputs.

## Running

Open `analysis.ipynb` and run top to bottom. Dependencies: `pandas`, `matplotlib`, `openpyxl` (for `.xlsx` I/O). Working directory must contain `Tuition Data.xlsx`; outputs overwrite in place.
