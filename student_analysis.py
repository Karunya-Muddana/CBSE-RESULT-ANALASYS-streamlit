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
    # normalize lines: drop empty lines
    lines = [ln.strip() for ln in lines if ln and ln.strip()]

    for i in range(0, len(lines) - 1, 2):
        line1 = lines[i]
        line2 = lines[i + 1]

        # robust regex: capture roll, gender, name (stops before first 3-digit code)
        match = re.match(r"(\d+)\s+([MF])\s+(.+?)\s+\d{3}", line1)
        if not match:
            # try slightly different fallback (if subject codes are missing)
            match = re.match(r"(\d+)\s+([MF])\s+(.+)$", line1)
            if not match:
                continue

        roll, gender, name = match.groups()

        # parse marks + grades from second line
        marks_grades = re.findall(r"(\d{2,3})\s+([A-D][1-2]|D2|D1|C2|C1|B2|B1|A2|A1)", line2)
        # if pattern returns as strings we still handle it

        student = {
            "Roll Number": roll,
            "Name": name.strip(),
            "Gender": gender
        }

        for idx, (marks, grade) in enumerate(marks_grades, start=1):
            # safe conversion
            try:
                m_int = int(re.sub(r"\D", "", marks))
            except:
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
    mapping: dict like {"Subject1_Marks": "ENG", "Subject1_Grade":"ENG GRADE", ...}
    If mapping is None, default to ENG, LANG II, MATH, SCI, SOC.
    Returns a new DataFrame (copy).
    """
    if mapping is None:
        mapping = {
            'Subject1_Marks': 'ENG',
            'Subject1_Grade': 'ENG GRADE',
            'Subject2_Marks': 'LANG II',
            'Subject2_Grade': 'LANG II GRADE',
            'Subject3_Marks': 'MATH',
            'Subject3_Grade': 'MATH GRADE',
            'Subject4_Marks': 'SCI',
            'Subject4_Grade': 'SCI GRADE',
            'Subject5_Marks': 'SOC',
            'Subject5_Grade': 'SOC GRADE'
        }
    df2 = df.rename(columns=mapping)
    return df2


def get_subject_mark_cols(df):
    """Return list of mark columns in the DataFrame (like ENG, MATH, or Subject1_Marks)."""
    mark_cols = [c for c in df.columns if c.upper().endswith("MARKS") or (isinstance(df[c].dtype, (np.number,)) and "GRADE" not in c.upper() and c not in ['Roll Number','Gender','Name'])]
    # fallback: explicitly match our mapped common names
    if not mark_cols:
        candidates = ["ENG", "LANG II", "MATH", "SCI", "SOC", "Subject1_Marks"]
        mark_cols = [c for c in df.columns if c in candidates]
    return mark_cols


def analyze_subject(df, subject_col, subject_grade_col=None, produce_figs=True):
    """
    Analyze one subject. Returns a dict:
      {
       "top10": DataFrame,
       "bottom10": DataFrame,
       "grade_counts": Series,
       "stats": dict(mean, median, std, max, min),
       "figs": {"pie": fig, "hist": fig, ...}  # matplotlib Figure objects (if produce_figs)
      }
    """
    result = {}
    # defensive checks
    if subject_col not in df.columns:
        raise KeyError(f"{subject_col} not in DataFrame columns: {df.columns.tolist()}")

    if subject_grade_col is None:
        # try to guess grade column
        candidates = [c for c in df.columns if c.lower().startswith(subject_col.lower().split()[0]) and "grade" in c.lower()]
        subject_grade_col = candidates[0] if candidates else None

    # ensure numeric marks
    marks_series = pd.to_numeric(df[subject_col], errors='coerce')

    # core tables
    subject_table = df.loc[:, ['Name', 'Roll Number', subject_col, subject_grade_col] if subject_grade_col in df.columns else ['Name','Roll Number',subject_col]].copy()
    subject_sorted = subject_table.sort_values(by=subject_col, ascending=False)
    result["top10"] = subject_sorted.head(10)
    result["bottom10"] = subject_sorted.tail(10)

    # grade counts
    if subject_grade_col and subject_grade_col in df.columns:
        result["grade_counts"] = df[subject_grade_col].value_counts().sort_index()
    else:
        result["grade_counts"] = None

    # stats
    stats = {
        "mean": float(marks_series.mean()) if not marks_series.isna().all() else None,
        "median": float(marks_series.median()) if not marks_series.isna().all() else None,
        "std": float(marks_series.std()) if not marks_series.isna().all() else None,
        "max": int(marks_series.max()) if not marks_series.isna().all() else None,
        "min": int(marks_series.min()) if not marks_series.isna().all() else None,
        "count": int(marks_series.count())
    }
    result["stats"] = stats

    figs = {}
    if produce_figs:
        # Pie (grade distribution) - only if grade_counts exists
        if result["grade_counts"] is not None and not result["grade_counts"].empty:
            fig1, ax1 = plt.subplots(figsize=(5,5))
            result["grade_counts"].plot.pie(autopct='%1.1f%%', startangle=90, shadow=True, ax=ax1)
            ax1.set_ylabel('')
            ax1.set_title(f"{subject_col} Grade Distribution")
            figs["pie"] = fig1

        # Histogram
        fig2, ax2 = plt.subplots(figsize=(7,4))
        sns.histplot(marks_series.dropna(), bins=10, kde=True, ax=ax2)
        ax2.set_title(f"{subject_col} Marks Distribution")
        ax2.set_xlabel("Marks")
        figs["hist"] = fig2

        # Boxplot
        fig3, ax3 = plt.subplots(figsize=(4,4))
        sns.boxplot(y=marks_series.dropna(), ax=ax3)
        ax3.set_title(f"{subject_col} Boxplot")
        figs["box"] = fig3

        # Top vs Bottom bar (names may be long -> rotate)
        comb = pd.concat([result["top10"], result["bottom10"]])
        if not comb.empty:
            fig4, ax4 = plt.subplots(figsize=(10,4))
            sns.barplot(x="Name", y=subject_col, data=comb, hue=subject_grade_col if subject_grade_col in comb.columns else None, dodge=False, ax=ax4)
            ax4.set_xticklabels(ax4.get_xticklabels(), rotation=45, ha="right")
            ax4.set_title(f"Top & Bottom {len(result['top10'])} in {subject_col}")
            figs["top_bottom"] = fig4

    result["figs"] = figs
    return result


def save_report_excel_bytes(df, subject_col, subject_grade_col=None):
    """
    Create an xlsx report in-memory and return bytes. Embeds generated charts as images.
    """
    import xlsxwriter
    from io import BytesIO
    import matplotlib.pyplot as plt

    # Generate report (figures included)
    report = analyze_subject(df, subject_col, subject_grade_col, produce_figs=True)

    output = BytesIO()
    # Use xlsxwriter engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write raw data and tables
        df.to_excel(writer, sheet_name='All Data', index=False)

        if "top10" in report and not report["top10"].empty:
            report["top10"].to_excel(writer, sheet_name='Top 10', index=False)
        if "bottom10" in report and not report["bottom10"].empty:
            report["bottom10"].to_excel(writer, sheet_name='Bottom 10', index=False)
        if report.get("grade_counts") is not None:
            # grade_counts is a Series — convert to DataFrame for Excel
            gc = report["grade_counts"].reset_index()
            gc.columns = [subject_grade_col if subject_grade_col else "Grade", "Count"]
            gc.to_excel(writer, sheet_name='Grade Counts', index=False)

        # Get workbook handle and create a chart sheet
        workbook = writer.book
        # create a worksheet specifically for charts
        try:
            charts_ws = workbook.add_worksheet('Charts')
        except Exception:
            # if worksheet exists for some reason, get it from writer.sheets
            charts_ws = writer.sheets.get('Charts') or workbook.add_worksheet('Charts')

        row = 0
        col = 0
        # Insert each figure as an image into the Charts sheet
        for name, fig in report.get("figs", {}).items():
            if fig is None:
                continue
            img_bytes = BytesIO()
            fig.savefig(img_bytes, format='png', bbox_inches='tight')
            plt.close(fig)
            img_bytes.seek(0)
            # insert_image(row, col, filename, options) supports 'image_data'
            charts_ws.insert_image(row, col, f"{name}.png", {'image_data': img_bytes})
            # move down sufficiently for the next image (adjust as needed)
            row += 25

        # exit 'with' block -> ExcelWriter finalizes and writes to output
    output.seek(0)
    return output.getvalue()
