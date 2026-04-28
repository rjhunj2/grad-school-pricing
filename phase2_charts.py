import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd


NAVY = "#012169"
GOLD = "#F2A900"

PROGRAM_PALETTE = [
    "#012169", "#F2A900", "#4A6FA5", "#C69214", "#6B8BC2",
    "#E8C170", "#2C3E50", "#8B7355", "#1F4788", "#A9835A",
]

# Per CLAUDE.md terms-per-year table (Info Sheet + empirical).
TERMS_PER_YEAR = {
    "COMPSCIMS": 2, "CS4P1MS": 2, "MATHMS": 2,
    "ECON4P1MS": 2, "BIOETHMA": 2, "BIOETH4P1": 2, "BMIDMS": 2,
    "ECONMS": 3, "DATASCIMS": 3, "QTMMS": 3,
    "DEVPRACMDP": 3, "HUMANRTCRT": 3, "BBS4P1MS": 3,
}

PROGRAM_MAP = {
    "COMPSCIMS": "Computer Science",
    "CS4P1MS":   "Computer Science 4+1",
    "DATASCIMS": "Data Science",
    "QTMMS":     "Data Science",
    "ECONMS":    "Economics",
    "ECON4P1MS": "Economics",
    "MATHMS":    "Math",
    "BIOETHMA":  "Bioethics",
    "BIOETH4P1": "Bioethics 4+1",
    "DEVPRACMDP":"MDP",
    "HUMANRTCRT":"MDP",
    "BMIDMS":    "MBID",
    "BBS4P1MS":  "Cancer Biology 4+1",
}

# Mirrors keep_scholarship() in analysis.ipynb — LGS-only, drop external/federal
# and generic LGS "special" awards.
_SCH_EXCLUDE = [
    "nih", "nsf", "grfp", "training grant", "special-scholarship",
    "yellow ribbon", "vet", "americorps", "pell", "hope", "zell",
    "woodruff scholar-grad tuition",
]


def _keep_scholarship(descr) -> bool:
    if pd.isna(descr):
        return False
    d = str(descr).lower().strip()
    if any(term in d for term in _SCH_EXCLUDE):
        return False
    return "lgs" in d


PEER_DATA = pd.DataFrame([
    {"school": "Columbia",     "program_group": "Data/CS",   "tuition": 64800},
    {"school": "NYU",          "program_group": "Data/CS",   "tuition": 75750},
    {"school": "Georgia Tech", "program_group": "Data/CS",   "tuition": 39500},
    {"school": "Columbia",     "program_group": "Economics", "tuition": 90732},
    {"school": "NYU",          "program_group": "Economics", "tuition": 70000},
    {"school": "Georgia Tech", "program_group": "Economics", "tuition": 41390},
    {"school": "Columbia",     "program_group": "General",   "tuition": 73456},
    {"school": "NYU",          "program_group": "General",   "tuition": 70000},
    {"school": "Georgia Tech", "program_group": "General",   "tuition": 31210},
])


def map_group(program: str) -> str:
    p = program.lower()
    if any(x in p for x in ["computer", "data"]):
        return "Data/CS"
    if any(x in p for x in ["econ", "math"]):
        return "Economics"
    return "General"


def save(fig, path: str) -> None:
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  wrote {path}")


def chart_discount_rate(ps: pd.DataFrame, out: str) -> None:
    d = ps.sort_values("discount_rate", ascending=True)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(d["program"], d["discount_rate"], color=NAVY)
    for i, v in enumerate(d["discount_rate"]):
        ax.text(v + 0.6, i, f"{v:.1f}%", va="center", fontsize=9)
    ax.set_xlabel("Discount Rate (%)")
    ax.set_title("Discount Rate by Program", color=NAVY, fontweight="bold")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, out)


def chart_gross_vs_net(ps: pd.DataFrame, out: str) -> None:
    d = ps.sort_values("avg_tuition", ascending=False)
    x = range(len(d))
    w = 0.4
    fig, ax = plt.subplots(figsize=(11, 6))
    ax.bar([i - w / 2 for i in x], d["avg_tuition"],      width=w, label="Avg Gross Tuition", color=NAVY)
    ax.bar([i + w / 2 for i in x], d["avg_net_tuition"],  width=w, label="Avg Net Tuition",   color=GOLD)
    ax.set_xticks(list(x))
    ax.set_xticklabels(d["program"], rotation=45, ha="right")
    ax.set_ylabel("Amount ($)")
    ax.set_title("Average Gross vs Net Tuition by Program", color=NAVY, fontweight="bold")
    ax.legend()
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, out)


def chart_intl_vs_discount(ps: pd.DataFrame, out: str) -> None:
    fig, ax = plt.subplots(figsize=(10, 7))
    sizes = ps["students"] * 10
    ax.scatter(
        ps["intl_pct"], ps["discount_rate"],
        s=sizes, color=NAVY, alpha=0.6,
        edgecolors=GOLD, linewidths=2,
    )
    for _, r in ps.iterrows():
        ax.annotate(
            f"{r['program']} (n={r['students']})",
            (r["intl_pct"], r["discount_rate"]),
            xytext=(7, 7), textcoords="offset points", fontsize=9,
        )
    ax.set_xlabel("International Students (%)")
    ax.set_ylabel("Discount Rate (%)")
    ax.set_title("International % vs Discount Rate\n(bubble size = enrollment)", color=NAVY, fontweight="bold")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.3)
    save(fig, out)


def chart_enrollment_trend(spy: pd.DataFrame, out: str) -> None:
    e = spy.groupby(["program", "academic_year"])["ID"].nunique().reset_index(name="students")
    pivot = e.pivot(index="academic_year", columns="program", values="students").fillna(0)
    # Order programs by total enrollment (largest first) so the legend is readable.
    order = pivot.sum(axis=0).sort_values(ascending=False).index.tolist()
    pivot = pivot[order]

    fig, ax = plt.subplots(figsize=(12, 6))
    for i, col in enumerate(pivot.columns):
        ax.plot(
            pivot.index, pivot[col],
            marker="o", linewidth=2,
            color=PROGRAM_PALETTE[i % len(PROGRAM_PALETTE)],
            label=col,
        )
    ax.set_xlabel("Academic Year")
    ax.set_ylabel("Students Enrolled")
    ax.set_title("Enrollment by Program over Academic Year", color=NAVY, fontweight="bold")
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), fontsize=9, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.3)
    save(fig, out)


def chart_discount_trend(spy: pd.DataFrame, out: str) -> None:
    t = (
        spy.groupby(["program", "academic_year"])
        .agg(tuition=("tuition", "sum"), scholarship=("scholarship", "sum"))
        .reset_index()
    )
    t["discount_rate"] = (t["scholarship"] / t["tuition"]) * 100
    pivot = t.pivot(index="academic_year", columns="program", values="discount_rate")
    # Match the ordering used in the enrollment chart.
    enroll = spy.groupby(["program", "academic_year"])["ID"].nunique().reset_index(name="n")
    order = enroll.groupby("program")["n"].sum().sort_values(ascending=False).index.tolist()
    pivot = pivot[[c for c in order if c in pivot.columns]]

    fig, ax = plt.subplots(figsize=(12, 6))
    for i, col in enumerate(pivot.columns):
        ax.plot(
            pivot.index, pivot[col],
            marker="o", linewidth=2,
            color=PROGRAM_PALETTE[i % len(PROGRAM_PALETTE)],
            label=col,
        )
    ax.set_xlabel("Academic Year")
    ax.set_ylabel("Discount Rate (%)")
    ax.set_title("Discount Rate by Program over Academic Year", color=NAVY, fontweight="bold")
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), fontsize=9, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.3)
    save(fig, out)


def chart_peer_benchmark(spl: pd.DataFrame, out: str) -> None:
    s = spl.copy()
    s["terms_per_year"] = s["acad_plan"].map(TERMS_PER_YEAR)
    s["annual_tuition"] = s["tuition"] * s["terms_per_year"] / s["term_count"]

    emory_prog_avg = s.groupby("program", as_index=False)["annual_tuition"].mean()
    emory_prog_avg["program_group"] = emory_prog_avg["program"].apply(map_group)
    emory_group = (
        emory_prog_avg.groupby("program_group", as_index=False)["annual_tuition"].mean()
        .rename(columns={"annual_tuition": "tuition"})
    )
    emory_group["school"] = "Emory"

    combined = pd.concat(
        [PEER_DATA, emory_group[["school", "program_group", "tuition"]]],
        ignore_index=True,
    )
    schools = ["Emory", "Columbia", "NYU", "Georgia Tech"]
    groups = ["Data/CS", "Economics", "General"]
    combined["school"] = pd.Categorical(combined["school"], categories=schools, ordered=True)
    combined["program_group"] = pd.Categorical(combined["program_group"], categories=groups, ordered=True)
    pivot = combined.pivot(index="school", columns="program_group", values="tuition").loc[schools, groups]

    fig, ax = plt.subplots(figsize=(11, 6))
    pivot.plot(kind="bar", ax=ax, color=[NAVY, GOLD, "#6B8BC2"], width=0.75)
    ax.set_title(
        "Annual Gross Tuition by Program Group — Emory vs Peers\n"
        "(Emory annualized with program-specific terms-per-year)",
        color=NAVY, fontweight="bold",
    )
    ax.set_ylabel("Annual Tuition ($)")
    ax.set_xlabel("")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    plt.xticks(rotation=0)
    ax.legend(title="Program Group", frameon=False)
    for container in ax.containers:
        ax.bar_label(container, fmt="$%.0f", padding=3, fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, out)


def _build_semester_data(tuition_xlsx: str) -> pd.DataFrame:
    """Per-(student, program, academic_year, semester) gross + scholarship.

    Filtered to the df_pricing population (career-total tuition >= $20k at the
    student-program grain) so partial-term billings don't drag the averages
    down. Mirrors the cleaning rules in analysis.ipynb.
    """
    t = pd.read_excel(tuition_xlsx, sheet_name="Gross Tuition Billed to Student")
    s = pd.read_excel(tuition_xlsx, sheet_name="Tuition Scholarship given to st")
    t.columns = t.columns.str.strip()
    s.columns = s.columns.str.strip()

    t = t.rename(columns={
        "Gross Tuition": "tuition",
        "Acad Plan": "acad_plan",
        "Item Term": "term_code",
    })[["ID", "acad_plan", "term_code", "tuition"]].copy()
    t["tuition"] = pd.to_numeric(t["tuition"], errors="coerce")
    t = t.dropna(subset=["acad_plan", "tuition"])
    t = t[t["acad_plan"].isin(PROGRAM_MAP)].copy()
    t["term_code"] = t["term_code"].astype(int)
    # Term code 5YYS: S = 1=Spring, 6=Summer, 9=Fall.
    t["season_code"] = t["term_code"] % 10
    # AY end-year: Fall rolls into next year (matches the scholarship sheet's Aid Yr).
    t["academic_year"] = (2000 + (t["term_code"] // 10) % 100) + (t["season_code"] == 9).astype(int)
    t["program"] = t["acad_plan"].map(PROGRAM_MAP)

    # df_pricing population: career-total >= $20k at student-program grain.
    career = t.groupby(["ID", "acad_plan"], as_index=False)["tuition"].sum()
    keep = career[career["tuition"] >= 20000][["ID", "acad_plan"]]
    t = t.merge(keep, on=["ID", "acad_plan"], how="inner")

    # Tuition by (ID, plan, year, season). Scholarship comes from the per-season columns.
    tuition_sem = (
        t.groupby(["ID", "acad_plan", "program", "academic_year", "season_code"], as_index=False)
        ["tuition"].sum()
    )

    s = s.rename(columns={
        "Academic Plan": "acad_plan",
        "Descr": "descr",
        "Fall Tuition Scholarship": "fall_sch",
        "Spring Tuiiton Scholarship": "spring_sch",
        "Summer Tuition Scholarship": "summer_sch",
        "Aid Yr": "academic_year",
    })
    for col in ["fall_sch", "spring_sch", "summer_sch"]:
        s[col] = pd.to_numeric(s[col], errors="coerce").fillna(0)
    s["academic_year"] = pd.to_numeric(s["academic_year"], errors="coerce").astype("Int64")
    s = s[s["acad_plan"].isin(PROGRAM_MAP)].copy()
    s = s[s["descr"].apply(_keep_scholarship)].copy()
    sch_year = s.groupby(["ID", "acad_plan", "academic_year"], as_index=False)[
        ["fall_sch", "spring_sch", "summer_sch"]
    ].sum()

    season_to_col = {9: "fall_sch", 1: "spring_sch", 6: "summer_sch"}
    tuition_sem["scholarship"] = 0.0
    for code, col in season_to_col.items():
        mask = tuition_sem["season_code"] == code
        m = tuition_sem.loc[mask].merge(
            sch_year[["ID", "acad_plan", "academic_year", col]],
            on=["ID", "acad_plan", "academic_year"], how="left",
        )
        tuition_sem.loc[mask, "scholarship"] = m[col].fillna(0).to_numpy()

    tuition_sem["net_tuition"] = (tuition_sem["tuition"] - tuition_sem["scholarship"]).clip(lower=0)
    return tuition_sem


def chart_gross_vs_net_by_semester(sem_df: pd.DataFrame, season_label: str, season_code: int, out: str) -> None:
    rows = sem_df[sem_df["season_code"] == season_code]
    if rows.empty:
        print(f"  skip {season_label}: no rows")
        return
    agg = (
        rows.groupby("program", as_index=False)
        .agg(
            students=("ID", "nunique"),
            avg_tuition=("tuition", "mean"),
            avg_net_tuition=("net_tuition", "mean"),
        )
        .sort_values("avg_tuition", ascending=False)
    )

    x = range(len(agg))
    w = 0.4
    fig, ax = plt.subplots(figsize=(11, 6))
    ax.bar([i - w / 2 for i in x], agg["avg_tuition"],     width=w, label="Avg Gross Tuition", color=NAVY)
    ax.bar([i + w / 2 for i in x], agg["avg_net_tuition"], width=w, label="Avg Net Tuition",   color=GOLD)
    ax.set_xticks(list(x))
    ax.set_xticklabels(agg["program"], rotation=45, ha="right")
    ax.set_ylabel("Amount per Semester ($)")
    ax.set_title(
        f"Average Gross vs Net Tuition by Program — {season_label}\n"
        "(per-semester billing, df_pricing population)",
        color=NAVY, fontweight="bold",
    )
    ax.legend()
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    save(fig, out)


def main() -> None:
    os.makedirs("charts", exist_ok=True)

    ps      = pd.read_excel("program_summary_output.xlsx", sheet_name="program_summary")
    ps_pric = pd.read_excel("program_summary_output.xlsx", sheet_name="pricing_summary")
    spl     = pd.read_excel("student_program_level_output.xlsx")
    spy     = pd.read_excel("student_program_year_output.xlsx")

    # Mirrors df_pricing in analysis.ipynb: drop partial-term rows that
    # distort discount and tuition averages. Used only for the three
    # pricing-focused charts.
    spl_pric = spl[spl["tuition"] >= 20000].copy()

    print(f"Inputs: program_summary={len(ps)} rows, "
          f"pricing_summary={len(ps_pric)} rows, "
          f"student_program={len(spl)} rows ({len(spl_pric)} after pricing filter), "
          f"student_program_year={len(spy)} rows")

    chart_discount_rate(ps_pric,    "charts/01_discount_rate_by_program.png")
    chart_gross_vs_net(ps_pric,     "charts/02_avg_gross_vs_net_tuition.png")
    chart_intl_vs_discount(ps,      "charts/03_intl_pct_vs_discount_rate.png")
    chart_enrollment_trend(spy,     "charts/04_enrollment_trend_by_program.png")
    chart_discount_trend(spy,       "charts/05_discount_rate_trend_by_program.png")
    chart_peer_benchmark(spl_pric,  "charts/06_peer_benchmark_emory_vs_peers.png")

    sem = _build_semester_data("Tuition Data.xlsx")
    print(f"  semester rows: {len(sem)} ({sem['ID'].nunique()} students)")
    chart_gross_vs_net_by_semester(sem, "Fall",   9, "charts/gross_net_fall.png")
    chart_gross_vs_net_by_semester(sem, "Spring", 1, "charts/gross_net_spring.png")
    chart_gross_vs_net_by_semester(sem, "Summer", 6, "charts/gross_net_summer.png")


if __name__ == "__main__":
    main()
