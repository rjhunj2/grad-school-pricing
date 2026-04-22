# Grad School Pricing Analysis

Analysis of Emory graduate program tuition, scholarship discounting, and net revenue, with peer benchmarking against Columbia, NYU, and Georgia Tech.

## Overview

The project ingests raw student-level tuition and scholarship records from institutional Excel exports, filters to a defined set of LGS (Laney Graduate School) master's programs, and produces:

1. **Student-program-level data** — one row per (student ID, academic plan) with gross tuition, kept scholarship, net tuition, and an international flag.
2. **Program-level summary** — aggregated metrics per program: student counts, international %, avg gross/net tuition, discount rate, totals.
3. **Peer comparison** — annualized Emory gross tuition grouped into Data/CS, Economics, and General buckets, plotted against hardcoded peer tuition.

All analysis lives in a single Jupyter notebook; outputs are written back to the project root as `.xlsx`.

## Project Structure

```
grad-school-pricing/
├── analysis.ipynb                                                    # main notebook — all cleaning, aggregation, charts
│
├── Tuition Data.xlsx                                                 # primary input; two sheets:
│                                                                     #   - "Gross Tuition Billed to Student"
│                                                                     #   - "Tuition Scholarship given to st"
├── Multi-Year Tuition-Tuition scholarship by Program.xlsx            # multi-year view of tuition + scholarships
├── 2026 Bioethics Masters Tuition Rates and Degree Completion Time Summary.xlsx
├── Students' IDs.xlsx                                                # student ID reference
│
├── Information - Process.pdf                                         # process/methodology notes
├── 260219 Conversation with Surabhi Vittal.pdf                       # stakeholder conversation notes
│
├── program_summary_output.xlsx                                       # GENERATED — program-level summary
└── student_program_level_output.xlsx                                 # GENERATED — per-student, per-program rows
```

## Data Pipeline (analysis.ipynb)

1. **Load** — read `Tuition Data.xlsx` sheets into `tuition` and `scholarship` DataFrames; strip whitespace from columns.
2. **Clean tuition** — keep `ID`, `acad_plan`, `tuition`; coerce tuition to numeric; drop rows missing plan or amount.
3. **Clean scholarships** — sum `fall_sch + spring_sch + summer_sch` into a single `scholarship` column per row; keep `ID`, `acad_plan`, `descr`.
4. **Program filter** — both frames are filtered to `program_map` (see below).
5. **Scholarship filter** — `keep_scholarship()` retains descriptions containing `"lgs"` and drops external/federal awards (`nih`, `nsf`, `grfp`, `training grant`, `yellow ribbon`, `vet`, `americorps`, `pell`, `hope`, `zell`, `woodruff scholar-grad tuition`, etc.).
6. **Aggregate** — group to `(ID, acad_plan)` level and merge tuition with kept scholarships; unmatched scholarship rows get `0`.
7. **Derive** — `net_tuition = tuition - scholarship`; `program` column via `program_map`.
8. **International flag** — re-load scholarship sheet, mark `EU CC IPEDS == "Non US Citizen"` as `intl=1`, collapse to max per (ID, plan), merge into `df`.
9. **Program summary** — group by `program`, compute student counts, international counts/%, avg gross/scholarship/net, totals, and `discount_rate = total_scholarship / total_gross_tuition`.
10. **Peer benchmark** — annualize Emory gross tuition (`annual = total * 2 / term_count`), group programs into Data/CS / Economics / General, concatenate with hardcoded `peer_data` (Columbia, NYU, Georgia Tech).
11. **Charts** — discount rate by program, avg net tuition by program, international % vs discount rate scatter, gross-vs-net line chart, grouped bar chart of annual gross tuition by school × program group.
12. **Save** — `program_summary_output.xlsx` and `student_program_level_output.xlsx`.

## Program Map

`acad_plan` codes are mapped to human-readable program names:

| Code          | Program              |
|---------------|----------------------|
| COMPSCIMS     | Computer Science     |
| CS4P1MS       | Computer Science 4+1 |
| DATASCIMS     | Data Science         |
| QTMMS         | Data Science         |
| ECONMS        | Economics            |
| ECON4P1MS     | Economics 4+1        |
| MATHMS        | Math                 |
| BIOETHMA      | Bioethics            |
| BIOETH4P1     | Bioethics 4+1        |
| DEVPRACMDP    | MDP                  |
| HUMANRTCRT    | MDP                  |
| BMIDMS        | MBID                 |
| BBS4P1MS      | Cancer Biology 4+1   |

Rows with any other `acad_plan` are dropped.

## Peer Group (hardcoded)

Annual tuition for Columbia, NYU, and Georgia Tech across three program buckets (Data/CS, Economics, General) is hardcoded in the notebook for benchmarking — not sourced from the Excel inputs.

## Running

Open `analysis.ipynb` and run top to bottom. Dependencies: `pandas`, `matplotlib`, `openpyxl` (for `.xlsx` I/O). Working directory must contain `Tuition Data.xlsx`; outputs overwrite in place.
