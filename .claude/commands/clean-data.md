---
description: Run the tuition data cleaning pipeline end-to-end and regenerate the output workbooks
---

Execute the cleaning pipeline described in CLAUDE.md against `Tuition Data.xlsx` and regenerate `program_summary_output.xlsx` and `student_program_level_output.xlsx`. Use the existing logic in `analysis.ipynb` — do not invent a new pipeline.

Steps:

1. **Load.** Read both sheets from `Tuition Data.xlsx`:
   - `Gross Tuition Billed to Student` → `tuition`
   - `Tuition Scholarship given to st` → `scholarship`
   Strip whitespace from column names on both frames.

2. **Clean tuition.** Keep `ID`, `acad_plan`, `tuition`. Coerce `tuition` to numeric (`errors="coerce"`). Drop rows where `acad_plan` or `tuition` is null. Report rows dropped and reason.

3. **Clean scholarships.** Compute `scholarship = fall_sch + spring_sch + summer_sch` (treating nulls as 0). Keep `ID`, `acad_plan`, `descr`, `scholarship`.

4. **Program filter.** Filter both frames to the `program_map` keys in CLAUDE.md:
   `COMPSCIMS, CS4P1MS, DATASCIMS, QTMMS, ECONMS, ECON4P1MS, MATHMS, BIOETHMA, BIOETH4P1, DEVPRACMDP, HUMANRTCRT, BMIDMS, BBS4P1MS`.
   Report counts kept vs. dropped per frame.

5. **Scholarship keep filter.** Apply `keep_scholarship()`: retain descriptions containing `"lgs"` (case-insensitive) and drop any containing excluded tokens (`nih`, `nsf`, `grfp`, `training grant`, `yellow ribbon`, `vet`, `americorps`, `pell`, `hope`, `zell`, `woodruff scholar-grad tuition`). Report counts kept/dropped and list any ambiguous descriptions for user review.

6. **Aggregate.** Group scholarship to `(ID, acad_plan)` with `scholarship.sum()`. Left-merge onto tuition (also grouped to `(ID, acad_plan)`). Fill unmatched scholarship rows with `0`.

7. **Derive.** Add `net_tuition = tuition - scholarship` and `program` via `program_map`. Flag any negative `net_tuition` rows in the report (do not drop).

8. **International flag.** Re-load the scholarship sheet, mark `EU CC IPEDS == "Non US Citizen"` as `intl=1` (else `0`), collapse to max per `(ID, acad_plan)`, and merge into the student-program frame. Fill missing with `0`.

9. **Program summary.** Group by `program` and compute: student count, international count, international %, avg gross / scholarship / net tuition, total gross / scholarship / net, and `discount_rate = total_scholarship / total_gross_tuition`.

10. **Write outputs.** Save `student_program_level_output.xlsx` and `program_summary_output.xlsx` in the project root (overwrite in place).

11. **Report.** Print a short summary: rows in each output, programs represented, total gross/net/scholarship across all programs, overall discount rate, and any warnings surfaced in earlier steps.

Run this by executing the relevant cells in `analysis.ipynb` (or an equivalent script) — prefer the notebook so the user sees the same outputs. Do not change the program map, excluded-token list, or annualization logic without asking first.
