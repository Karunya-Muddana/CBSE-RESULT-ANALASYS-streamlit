# student_analysis.py
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from io import BytesIO

sns.set(style="whitegrid")


def parse_student_data_from_lines(lines):
    """
    Parse CBSE-style two-line per student text (list of lines).
    Returns a DataFrame with columns:
      Roll Number, Name, Gender, Subject1_Marks, Subject1_Grade, ..., Subject5_Marks, Subject5_Grade
    """
    students = []
    lines = [ln.strip() for ln in lines if ln and ln.strip()]  # normalize

    for i in range(0, len(lines) - 1, 2):
        line1 = lines[i]
        line2 = lines[i + 1]

        match = re.match(r"(\d+)\s+([MF])\s+(.+?)\s+\d{3}", line1)
        if not match:
            match = re.match(r"(\d+)\s+([MF])\s+(.+)$", line1)
            if not match:
                continue

        roll, gender, name = match.groups()

        # parse marks + grades
        marks_grades = re.findall(
            r"(\d{1,3})\s+([A-D][1-2]|E1|E2|F|C2|C1|B2|B1|A2|A1)", line2
        )

        student = {"Roll Number": roll, "Name": name.strip(), "Gender": gender}

        for idx, (marks, grade) in enumerate(marks_grades, start=1):
            try:
                m_int = int(re.sub(r"\D", "", marks))
            except Exception:
                m_int = np.nan
            student[f"Subject{idx}_Marks"] = m_int
            student[f"Subject{idx}_Grade"] = grade.strip()

        students.append(student)

    df = pd.DataFrame(students)
    return df


def parse_student_data_from_file(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.read().splitlines()
    return parse_student_data_from_lines(lines)


def map_subjects_to_names(df, mapping=None):
    """
    Rename Subject1..Subject5 to friendly names.
    If mapping is None, default to ENG, LANG II, MATH, SCI, SOC.
    """
    if mapping is None:
        mapping = {
            "Subject1_Marks": "ENG",
            "Subject1_Grade": "ENG GRADE",
            "Subject2_Marks": "LANG II",
            "Subject2_Grade": "LANG II GRADE",
            "Subject3_Marks": "MATH",
            "Subject3_Grade": "MATH GRADE",
            "Subject4_Marks": "SCI",
            "Subject4_Grade": "SCI GRADE",
            "Subject5_Marks": "SOC",
            "Subject5_Grade": "SOC GRADE",
        }
    df2 = df.rename(columns=mapping)
    return df2


def get_subject_mark_cols(df):
    """Return list of mark columns in the DataFrame."""
    mark_cols = [
        c
        for c in df.columns
        if pd.api.types.is_numeric_dtype(df[c])
        and "GRADE" not in c.upper()
        and c not in ["Roll Number", "Gender"]
    ]
    return mark_cols


def add_total_marks(df):
    """Adds a Total column = sum of all subject marks."""
    mark_cols = get_subject_mark_cols(df)
    df = df.copy()
    df["Total"] = df[mark_cols].sum(axis=1, skipna=True)
    return df


def analyze_subject(df, subject_col, subject_grade_col=None, produce_figs=True):
    """
    Analyze one subject. Returns dict with top/bottom10, grade_counts, stats, and figs.
    """
    result = {}
    if subject_col not in df.columns:
        raise KeyError(f"{subject_col} not in DataFrame")

    if subject_grade_col is None:
        candidates = [
            c
            for c in df.columns
            if c.lower().startswith(subject_col.lower().split()[0])
            and "grade" in c.lower()
        ]
        subject_grade_col = candidates[0] if candidates else None

    marks_series = pd.to_numeric(df[subject_col], errors="coerce")

    subject_table = df.loc[
        :,
        ["Name", "Roll Number", subject_col]
        + ([subject_grade_col] if subject_grade_col in df.columns else []),
    ].copy()
    subject_sorted = subject_table.sort_values(by=subject_col, ascending=False)
    result["top10"] = subject_sorted.head(10)
    result["bottom10"] = subject_sorted.tail(10)

    if subject_grade_col and subject_grade_col in df.columns:
        result["grade_counts"] = df[subject_grade_col].value_counts().sort_index()
    else:
        result["grade_counts"] = None

    stats = {
        "mean": float(marks_series.mean()) if not marks_series.isna().all() else None,
        "median": float(marks_series.median())
        if not marks_series.isna().all()
        else None,
        "std": float(marks_series.std()) if not marks_series.isna().all() else None,
        "max": int(marks_series.max()) if not marks_series.isna().all() else None,
        "min": int(marks_series.min()) if not marks_series.isna().all() else None,
        "count": int(marks_series.count()),
    }
    result["stats"] = stats

    figs = {}
    if produce_figs:
        if result["grade_counts"] is not None and not result["grade_counts"].empty:
            fig1, ax1 = plt.subplots(figsize=(5, 5))
            result["grade_counts"].plot.pie(
                autopct="%1.1f%%", startangle=90, shadow=True, ax=ax1
            )
            ax1.set_ylabel("")
            ax1.set_title(f"{subject_col} Grade Distribution")
            figs["pie"] = fig1

        fig2, ax2 = plt.subplots(figsize=(7, 4))
        sns.histplot(marks_series.dropna(), bins=10, kde=True, ax=ax2)
        ax2.set_title(f"{subject_col} Marks Distribution")
        ax2.set_xlabel("Marks")
        figs["hist"] = fig2

        fig3, ax3 = plt.subplots(figsize=(4, 4))
        sns.boxplot(y=marks_series.dropna(), ax=ax3)
        ax3.set_title(f"{subject_col} Boxplot")
        figs["box"] = fig3

        comb = pd.concat([result["top10"], result["bottom10"]])
        if not comb.empty:
            fig4, ax4 = plt.subplots(figsize=(10, 4))
            sns.barplot(
                x="Name",
                y=subject_col,
                data=comb,
                hue=subject_grade_col if subject_grade_col in comb.columns else None,
                dodge=False,
                ax=ax4,
            )
            ax4.set_xticklabels(ax4.get_xticklabels(), rotation=45, ha="right")
            ax4.set_title(f"Top & Bottom 10 in {subject_col}")
            figs["top_bottom"] = fig4

    result["figs"] = figs
    return result


def analyze_total(df, produce_figs=True):
    """Analyze overall total marks for all students."""
    if "Total" not in df.columns:
        df = add_total_marks(df)

    result = {}
    totals = df["Total"]

    total_table = df.loc[:, ["Name", "Roll Number", "Total"]].copy()
    total_sorted = total_table.sort_values(by="Total", ascending=False)
    result["top10"] = total_sorted.head(10)
    result["bottom10"] = total_sorted.tail(10)

    stats = {
        "mean": float(totals.mean()),
        "median": float(totals.median()),
        "std": float(totals.std()),
        "max": int(totals.max()),
        "min": int(totals.min()),
        "count": int(totals.count()),
    }
    result["stats"] = stats

    figs = {}
    if produce_figs:
        fig1, ax1 = plt.subplots(figsize=(7, 4))
        sns.histplot(totals.dropna(), bins=15, kde=True, ax=ax1)
        ax1.set_title("Total Marks Distribution")
        ax1.set_xlabel("Total Marks")
        figs["hist"] = fig1

        fig2, ax2 = plt.subplots(figsize=(4, 4))
        sns.boxplot(y=totals.dropna(), ax=ax2)
        ax2.set_title("Total Marks Boxplot")
        figs["box"] = fig2

        comb = pd.concat([result["top10"], result["bottom10"]])
        fig3, ax3 = plt.subplots(figsize=(10, 4))
        sns.barplot(x="Name", y="Total", data=comb, dodge=False, ax=ax3)
        ax3.set_xticklabels(ax3.get_xticklabels(), rotation=45, ha="right")
        ax3.set_title("Top & Bottom 10 by Total Marks")
        figs["top_bottom"] = fig3

    result["figs"] = figs
    return result


def save_full_report_excel_bytes(df):
    """
    Create an Excel report containing subject-wise and total analysis.
    """
    df = add_total_marks(df)
    import matplotlib.pyplot as plt
    from io import BytesIO

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="All Data", index=False)
        workbook = writer.book
        charts_ws = workbook.add_worksheet("Charts")
        row = 0

        # Subject-wise
        for subject_col in get_subject_mark_cols(df):
            if subject_col == "Total":
                continue
            grade_col = (
                subject_col.replace("Marks", "Grade")
                if f"{subject_col.split('_')[0]}_Grade" in df.columns
                else None
            )
            report = analyze_subject(df, subject_col, grade_col, produce_figs=True)

            report["top10"].to_excel(
                writer, sheet_name=f"{subject_col}_Top", index=False
            )
            report["bottom10"].to_excel(
                writer, sheet_name=f"{subject_col}_Bottom", index=False
            )

            if report["grade_counts"] is not None:
                gc = report["grade_counts"].reset_index()
                gc.columns = [grade_col or "Grade", "Count"]
                gc.to_excel(
                    writer, sheet_name=f"{subject_col}_Grades", index=False
                )

            for name, fig in report["figs"].items():
                img_bytes = BytesIO()
                fig.savefig(img_bytes, format="png", bbox_inches="tight")
                plt.close(fig)
                img_bytes.seek(0)
                charts_ws.insert_image(
                    row, 0, f"{subject_col}_{name}.png", {"image_data": img_bytes}
                )
                row += 25

        # Total report
        total_report = analyze_total(df, produce_figs=True)
        total_report["top10"].to_excel(
            writer, sheet_name="Total_Top", index=False
        )
        total_report["bottom10"].to_excel(
            writer, sheet_name="Total_Bottom", index=False
        )

        for name, fig in total_report["figs"].items():
            img_bytes = BytesIO()
            fig.savefig(img_bytes, format="png", bbox_inches="tight")
            plt.close(fig)
            img_bytes.seek(0)
            charts_ws.insert_image(
                row, 0, f"Total_{name}.png", {"image_data": img_bytes}
            )
            row += 25

    output.seek(0)
    return output.getvalue()
