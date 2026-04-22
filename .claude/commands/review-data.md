---
description: Review the raw tuition inputs and the current state of the analysis pipeline
---

Review the tuition data inputs and the state of the analysis pipeline in `analysis.ipynb`. Report findings — do not modify files.

Steps:

1. **Inventory inputs.** Confirm these files exist in the project root and note their sizes / last-modified dates:
   - `Tuition Data.xlsx` (sheets: `Gross Tuition Billed to Student`, `Tuition Scholarship given to st`)
   - `Multi-Year Tuition-Tuition scholarship by Program.xlsx`
   - `2026 Bioethics Masters Tuition Rates and Degree Completion Time Summary.xlsx`
   - `Students' IDs.xlsx`
   - Generated outputs: `program_summary_output.xlsx`, `student_program_level_output.xlsx`

2. **Inspect the tuition and scholarship sheets.** For each sheet in `Tuition Data.xlsx`, report:
   - Row count and column list (flag leading/trailing whitespace in column names)
   - Distinct `acad_plan` values and how many match the `program_map` in CLAUDE.md vs. would be dropped
   - Null counts for `ID`, `acad_plan`, `tuition` (tuition sheet) and `fall_sch`, `spring_sch`, `summer_sch`, `descr`, `EU CC IPEDS` (scholarship sheet)
   - Non-numeric values in tuition/scholarship amount columns

3. **Scan scholarship descriptions.** List the distinct `descr` values and flag:
   - Which would be kept by `keep_scholarship()` (contains `"lgs"`, not excluded)
   - Which would be excluded (nih, nsf, grfp, training grant, yellow ribbon, vet, americorps, pell, hope, zell, woodruff scholar-grad tuition, etc.)
   - Any ambiguous descriptions worth asking the user about

4. **Notebook state.** Open `analysis.ipynb` and report:
   - Whether cells have stale outputs vs. current inputs (compare output file mtimes to input file mtimes)
   - Any hardcoded values worth flagging (peer tuition table, term count for annualization, program_map)
   - Cells that import or write files — confirm paths resolve

5. **Summary.** End with a short punch list:
   - Data quality issues to resolve before re-running
   - Questions for the user (ambiguous scholarships, missing plans, etc.)
   - Suggested next step (`/clean-data` to re-run the pipeline, `/qa-data` to validate outputs, or ask the user)

Keep the report tight — bullets over prose. Reference `file:line` when pointing at notebook cells or code.
