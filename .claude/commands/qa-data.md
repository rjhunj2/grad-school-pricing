---
description: QA the cleaned tuition outputs — sanity checks, outlier detection, reconciliation against raw inputs
---

Validate the outputs of the tuition pipeline against the raw inputs and flag anything suspicious. Read-only — do not modify data or notebook cells. Report findings in a compact punch list.

Inputs to check:
- `student_program_level_output.xlsx`
- `program_summary_output.xlsx`
- Raw: `Tuition Data.xlsx` (both sheets)

Checks:

1. **Reconciliation vs. raw.**
   - Sum of `tuition` in `student_program_level_output.xlsx` should equal the sum of gross tuition in the raw sheet after the program filter. Report any gap.
   - Sum of `scholarship` should equal the sum of kept (LGS, non-excluded) scholarships from the raw sheet. Report any gap.
   - Distinct `(ID, acad_plan)` count in the student-level output should equal distinct tuition pairs after filtering.

2. **Row-level sanity.**
   - Any `net_tuition < 0` (scholarship exceeds tuition). List offending IDs and programs.
   - Any `tuition == 0` or `tuition` below a plausibility floor (e.g. < $1,000). List them.
   - Any rows missing `program` (i.e. `acad_plan` slipped through the program filter).
   - Any duplicate `(ID, acad_plan)` rows in the student-level output.

3. **Scholarship filter audit.**
   - List scholarship descriptions present in the raw sheet but **not** represented in the kept totals (i.e. filtered out). Confirm each matches an excluded token or lacks `"lgs"`.
   - Flag any description that was kept but looks external/federal (possible false positive).

4. **International flag.**
   - Count of `intl == 1` per program; compare to the raw `EU CC IPEDS == "Non US Citizen"` count after the program filter.
   - Any `(ID, acad_plan)` with conflicting `intl` values across scholarship rows that collapsed to `max` — note how many conflicts existed.

5. **Program summary integrity.**
   - Recompute `discount_rate = total_scholarship / total_gross_tuition` per program and compare to the value in `program_summary_output.xlsx`.
   - Confirm `avg_net = avg_gross - avg_scholarship` (within rounding).
   - Flag any program with student count below a small threshold (e.g. < 5) — averages are noisy.
   - Flag any program with discount rate > 100% or < 0%.

6. **Peer benchmark cross-check.**
   - Confirm the annualized Emory gross tuition used in the peer chart (`annual = total * 2 / term_count`) matches what the notebook currently produces.
   - Flag Data/CS, Economics, or General bucket assignments that don't match the program_map in CLAUDE.md.

7. **Output file hygiene.**
   - `program_summary_output.xlsx` and `student_program_level_output.xlsx` mtimes newer than `Tuition Data.xlsx`? If not, outputs are stale.
   - Column set matches what CLAUDE.md describes for each output.

Report format — for each check, one of: `OK`, `WARN: <detail>`, or `FAIL: <detail>`. End with a one-line overall verdict and a suggested next step (re-run `/clean-data`, ask the user about a specific ambiguity, etc.).
