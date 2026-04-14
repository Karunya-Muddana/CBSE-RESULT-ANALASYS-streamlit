import re
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
import numpy as np
from io import BytesIO

sns.set_theme(style="whitegrid", palette="muted")

# ── CBSE Subject Code → Name ─────────────────────────────────────────────────
SUBJECT_CODE_MAP = {
    "184": "English",
    "085": "Hindi",
    "018": "Telugu",
    "089": "Tamil",
    "041": "Mathematics",
    "086": "Science",
    "087": "Social Science",
    "049": "Painting",
    "165": "Hindi Elective",
    "002": "Hindi",
    "241": "Mathematics Basic",
}

LANG2_CODES = {"085", "018", "089", "165", "002"}  # 2nd language codes

GRADE_ORDER = ["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2", "E1", "E2", "F"]
GRADE_COLORS = {
    "A1": "#2ecc71", "A2": "#27ae60",
    "B1": "#3498db", "B2": "#2980b9",
    "C1": "#f39c12", "C2": "#e67e22",
    "D1": "#e74c3c", "D2": "#c0392b",
    "E1": "#95a5a6", "E2": "#7f8c8d",
    "F":  "#2c3e50",
}


# ── PARSING ───────────────────────────────────────────────────────────────────

def parse_student_data_from_lines(lines):
    """
    Parse CBSE gazette format. Handles variable 2nd languages.
    Returns DataFrame with columns:
      Roll Number, Name, Gender, Result,
      English_Marks, English_Grade,
      Lang2_Code, Lang2_Name, Lang2_Marks, Lang2_Grade,
      Mathematics_Marks, Mathematics_Grade,
      Science_Marks, Science_Grade,
      SocialScience_Marks, SocialScience_Grade,
      [optional extra subjects]
    """
    students = []
    clean = [ln.rstrip("\r") for ln in lines]

    i = 0
    while i < len(clean) - 1:
        line1 = clean[i].strip()

        # Match student header line
        m = re.match(r"(\d{7,10})\s+([MF])\s+(.+?)\s{3,}((?:\s*\d{3})+)\s+(PASS|FAIL|COMP|ESSEN|ABSE)", line1)
        if not m:
            i += 1
            continue

        roll = m.group(1).strip()
        gender = m.group(2)
        name = m.group(3).strip()
        subject_codes = re.findall(r"\d{3}", m.group(4))
        result = m.group(5)

        # Next non-empty line has marks+grades
        j = i + 1
        marks_line = ""
        while j < len(clean):
            candidate = clean[j].strip()
            if candidate and re.search(r"\d{2,3}\s+[A-F]\d", candidate):
                marks_line = candidate
                break
            j += 1

        marks_grades = re.findall(r"(\d{1,3})\s+([A-D][1-2]|E1|E2|F)", marks_line)

        student = {
            "Roll Number": roll,
            "Name": name,
            "Gender": gender,
            "Result": result,
        }

        # Map codes to marks/grades
        lang2_code = None
        lang2_name = "Lang II"
        for pos, (code, (marks_str, grade)) in enumerate(
            zip(subject_codes, marks_grades), start=1
        ):
            subj_name = SUBJECT_CODE_MAP.get(code, f"Sub{code}")
            col_key = subj_name.replace(" ", "")

            # Detect 2nd language slot
            if code in LANG2_CODES and lang2_code is None:
                lang2_code = code
                lang2_name = subj_name
                student["Lang2_Code"] = code
                student["Lang2_Name"] = lang2_name

            try:
                marks_val = int(marks_str)
            except Exception:
                marks_val = np.nan

            student[f"{col_key}_Marks"] = marks_val
            student[f"{col_key}_Grade"] = grade

        students.append(student)
        i = j + 1

    df = pd.DataFrame(students)
    return df


def parse_student_data_from_file(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.read().splitlines()
    return parse_student_data_from_lines(lines)


# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_mark_cols(df):
    return [
        c for c in df.columns
        if c.endswith("_Marks")
        and pd.api.types.is_numeric_dtype(df[c])
        and c != "Total_Marks"
    ]


def get_grade_col(df, mark_col):
    g = mark_col.replace("_Marks", "_Grade")
    return g if g in df.columns else None


def add_total_marks(df):
    df = df.copy()
    mc = get_mark_cols(df)
    df["Total"] = df[mc].sum(axis=1, skipna=True)
    return df


def get_lang2_breakdown(df):
    """Returns a dict: lang_name → sub-DataFrame"""
    if "Lang2_Name" not in df.columns:
        return {}
    groups = {}
    for lang, sub in df.groupby("Lang2_Name"):
        groups[lang] = sub.reset_index(drop=True)
    return groups


def pass_fail_counts(df):
    if "Result" in df.columns:
        return df["Result"].value_counts()
    return pd.Series(dtype=int)


def grade_summary(df):
    """Wide table: subject × grade counts"""
    rows = []
    for mc in get_mark_cols(df):
        gc = get_grade_col(df, mc)
        subj = mc.replace("_Marks", "")
        if gc and gc in df.columns:
            counts = df[gc].value_counts()
            row = {"Subject": subj}
            for g in GRADE_ORDER:
                row[g] = counts.get(g, 0)
            rows.append(row)
    return pd.DataFrame(rows).set_index("Subject") if rows else pd.DataFrame()


# ── FIGURES ───────────────────────────────────────────────────────────────────

def fig_grade_distribution(df, mark_col):
    gc = get_grade_col(df, mark_col)
    if not gc or gc not in df.columns:
        return None
    counts = df[gc].value_counts().reindex(GRADE_ORDER).dropna()
    if counts.empty:
        return None

    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.bar(
        counts.index,
        counts.values,
        color=[GRADE_COLORS.get(g, "#95a5a6") for g in counts.index],
        edgecolor="white", linewidth=0.8
    )
    for bar, val in zip(bars, counts.values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                str(int(val)), ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_title(f"Grade Distribution — {mark_col.replace('_Marks', '')}", fontsize=12, pad=10)
    ax.set_xlabel("Grade")
    ax.set_ylabel("Number of Students")
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def fig_marks_histogram(df, mark_col):
    vals = pd.to_numeric(df[mark_col], errors="coerce").dropna()
    if vals.empty:
        return None
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.hist(vals, bins=15, color="#3498db", edgecolor="white", alpha=0.85, linewidth=0.8)
    ax.axvline(vals.mean(), color="#e74c3c", lw=2, linestyle="--", label=f"Mean: {vals.mean():.1f}")
    ax.axvline(vals.median(), color="#2ecc71", lw=2, linestyle="--", label=f"Median: {vals.median():.1f}")
    ax.set_title(f"Marks Distribution — {mark_col.replace('_Marks', '')}", fontsize=12, pad=10)
    ax.set_xlabel("Marks")
    ax.set_ylabel("Students")
    ax.legend(fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def fig_subject_comparison(df):
    mc = get_mark_cols(df)
    if not mc:
        return None
    means = {c.replace("_Marks", ""): pd.to_numeric(df[c], errors="coerce").mean() for c in mc}
    labels = list(means.keys())
    vals = list(means.values())
    colors = ["#3498db", "#2ecc71", "#e74c3c", "#f39c12", "#9b59b6", "#1abc9c"]

    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.bar(labels, vals, color=colors[:len(labels)], edgecolor="white", linewidth=0.8)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                f"{v:.1f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_title("Average Marks by Subject", fontsize=12, pad=10)
    ax.set_ylabel("Average Marks")
    ax.set_ylim(0, 105)
    ax.spines[["top", "right"]].set_visible(False)
    plt.xticks(rotation=20, ha="right")
    fig.tight_layout()
    return fig


def fig_gender_comparison(df):
    mc = get_mark_cols(df)
    if "Gender" not in df.columns or not mc:
        return None
    rows = []
    for g in ["M", "F"]:
        sub = df[df["Gender"] == g]
        for c in mc:
            rows.append({
                "Gender": "Male" if g == "M" else "Female",
                "Subject": c.replace("_Marks", ""),
                "Mean": pd.to_numeric(sub[c], errors="coerce").mean()
            })
    gdf = pd.DataFrame(rows)
    if gdf.empty:
        return None

    fig, ax = plt.subplots(figsize=(8, 4))
    subjects = gdf["Subject"].unique()
    x = np.arange(len(subjects))
    w = 0.35
    males = gdf[gdf["Gender"] == "Male"].set_index("Subject").reindex(subjects)["Mean"]
    females = gdf[gdf["Gender"] == "Female"].set_index("Subject").reindex(subjects)["Mean"]

    ax.bar(x - w/2, males.values, w, label="Male", color="#3498db", alpha=0.85)
    ax.bar(x + w/2, females.values, w, label="Female", color="#e91e8c", alpha=0.85)
    ax.set_xticks(x)
    ax.set_xticklabels(subjects, rotation=20, ha="right")
    ax.set_title("Gender-wise Average Marks by Subject", fontsize=12, pad=10)
    ax.set_ylabel("Average Marks")
    ax.set_ylim(0, 105)
    ax.legend()
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def fig_lang2_comparison(df):
    """Compare performance across 2nd language groups."""
    if "Lang2_Name" not in df.columns:
        return None
    groups = df["Lang2_Name"].unique()
    if len(groups) < 2:
        return None

    # Compare by Total
    if "Total" not in df.columns:
        df = add_total_marks(df)

    fig, ax = plt.subplots(figsize=(7, 4))
    colors = ["#3498db", "#2ecc71", "#e74c3c", "#f39c12"]
    for idx, lang in enumerate(sorted(groups)):
        sub = df[df["Lang2_Name"] == lang]["Total"].dropna()
        ax.hist(sub, bins=12, alpha=0.6, label=f"{lang} (n={len(sub)})",
                color=colors[idx % len(colors)])
    ax.set_title("Total Marks Distribution by 2nd Language", fontsize=12, pad=10)
    ax.set_xlabel("Total Marks")
    ax.set_ylabel("Students")
    ax.legend()
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def fig_top_bottom(df, mark_col, n=10):
    gc = get_grade_col(df, mark_col)
    subj_label = mark_col.replace("_Marks", "")
    cols = ["Name", "Roll Number", mark_col] + ([gc] if gc and gc in df.columns else [])
    sub = df[cols].copy()
    sub[mark_col] = pd.to_numeric(sub[mark_col], errors="coerce")
    sub = sub.dropna(subset=[mark_col]).sort_values(mark_col, ascending=False)

    top = sub.head(n).copy()
    top["Group"] = "Top"
    bot = sub.tail(n).copy()
    bot["Group"] = "Bottom"
    comb = pd.concat([top, bot]).reset_index(drop=True)

    if comb.empty:
        return None
    fig, ax = plt.subplots(figsize=(11, 4))
    colors = ["#2ecc71" if g == "Top" else "#e74c3c" for g in comb["Group"]]
    bars = ax.bar(comb["Name"], comb[mark_col], color=colors, edgecolor="white", linewidth=0.6)
    for bar, row in zip(bars, comb.itertuples()):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                str(int(getattr(row, mark_col.replace(" ", "_"), bar.get_height()))),
                ha="center", va="bottom", fontsize=7.5)
    ax.set_xticklabels(comb["Name"], rotation=45, ha="right", fontsize=8)
    ax.set_title(f"Top & Bottom {n} — {subj_label}", fontsize=12, pad=10)
    ax.set_ylabel("Marks")
    top_patch = mpatches.Patch(color="#2ecc71", label="Top")
    bot_patch = mpatches.Patch(color="#e74c3c", label="Bottom")
    ax.legend(handles=[top_patch, bot_patch])
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def fig_pass_fail_pie(df):
    pf = pass_fail_counts(df)
    if pf.empty:
        return None
    colors = {"PASS": "#2ecc71", "FAIL": "#e74c3c", "COMP": "#f39c12",
              "ESSEN": "#3498db", "ABSE": "#95a5a6"}
    fig, ax = plt.subplots(figsize=(5, 5))
    wedge_colors = [colors.get(k, "#95a5a6") for k in pf.index]
    wedges, texts, autotexts = ax.pie(
        pf.values, labels=pf.index, autopct="%1.1f%%",
        colors=wedge_colors, startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 1.5}
    )
    for at in autotexts:
        at.set_fontsize(10)
    ax.set_title("Pass / Fail Breakdown", fontsize=12, pad=10)
    fig.tight_layout()
    return fig


def fig_correlation_heatmap(df):
    mc = get_mark_cols(df)
    if len(mc) < 2:
        return None
    num_df = df[mc].apply(pd.to_numeric, errors="coerce")
    corr = num_df.corr()
    labels = [c.replace("_Marks", "") for c in corr.columns]
    corr.columns = labels
    corr.index = labels

    fig, ax = plt.subplots(figsize=(6, 5))
    mask = np.triu(np.ones_like(corr, dtype=bool))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="RdYlGn",
                vmin=-1, vmax=1, ax=ax, mask=mask,
                linewidths=0.5, linecolor="white",
                annot_kws={"size": 10})
    ax.set_title("Subject Marks Correlation", fontsize=12, pad=10)
    fig.tight_layout()
    return fig


def fig_scatter_two_subjects(df, subj_x, subj_y):
    x = pd.to_numeric(df[subj_x], errors="coerce")
    y = pd.to_numeric(df[subj_y], errors="coerce")
    mask = x.notna() & y.notna()
    if mask.sum() < 2:
        return None
    fig, ax = plt.subplots(figsize=(6, 5))
    colors = ["#3498db" if g == "M" else "#e91e8c" for g in df.loc[mask, "Gender"]] \
        if "Gender" in df.columns else "#3498db"
    ax.scatter(x[mask], y[mask], c=colors, alpha=0.6, edgecolors="white", s=60)
    m, b = np.polyfit(x[mask], y[mask], 1)
    xs = np.linspace(x[mask].min(), x[mask].max(), 100)
    ax.plot(xs, m * xs + b, color="#e74c3c", lw=1.5, linestyle="--", label="Trend")
    ax.set_xlabel(subj_x.replace("_Marks", ""))
    ax.set_ylabel(subj_y.replace("_Marks", ""))
    ax.set_title(f"Correlation: {subj_x.replace('_Marks','')} vs {subj_y.replace('_Marks','')}", fontsize=11)
    if "Gender" in df.columns:
        male_patch = mpatches.Patch(color="#3498db", label="Male")
        female_patch = mpatches.Patch(color="#e91e8c", label="Female")
        ax.legend(handles=[male_patch, female_patch, plt.Line2D([0],[0], color="#e74c3c", lw=1.5, linestyle="--", label="Trend")])
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


# ── EXCEL EXPORT ──────────────────────────────────────────────────────────────

def build_excel_report(df):
    """
    Teacher-friendly Excel report with multiple formatted sheets.
    Returns bytes.
    """
    from openpyxl import Workbook
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                  numbers)
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.series import DataPoint
    import copy

    df = add_total_marks(df)
    mark_cols = get_mark_cols(df)

    wb = Workbook()

    # ── Colour palette ──
    HDR_BG  = "1F4E79"
    HDR_FG  = "FFFFFF"
    SUBHDR  = "2E75B6"
    ALT_ROW = "DEEAF1"
    PASS_BG = "E2EFDA"
    FAIL_BG = "FFE6E6"

    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_style(cell, bg=HDR_BG, fg=HDR_FG, size=10):
        cell.font = Font(bold=True, color=fg, size=size)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    def data_style(cell, bold=False, center=False, bg=None, size=10):
        cell.font = Font(bold=bold, size=size)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
        cell.border = border

    def col_width(ws, col, width):
        ws.column_dimensions[get_column_letter(col)].width = width

    # ════════════════════════════════════════════════════════════
    # SHEET 1: Complete Student Data
    # ════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "📋 All Students"
    ws1.row_dimensions[1].height = 30
    ws1.freeze_panes = "A3"

    # Build columns
    base_cols = ["Roll Number", "Name", "Gender", "Result"]
    subject_cols = []
    for mc in mark_cols:
        gc = get_grade_col(df, mc)
        subject_cols.append(mc)
        if gc and gc in df.columns:
            subject_cols.append(gc)
    extra = ["Lang2_Name"] if "Lang2_Name" in df.columns else []
    all_cols = base_cols + extra + subject_cols + ["Total"]

    # Title row
    title_cell = ws1.cell(1, 1, "CBSE Class X Results — Complete Student Data")
    title_cell.font = Font(bold=True, size=13, color=HDR_FG)
    title_cell.fill = PatternFill("solid", fgColor=HDR_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_cols))

    # Header row
    for col_idx, col in enumerate(all_cols, 1):
        cell = ws1.cell(2, col_idx, col.replace("_", " ").replace("Marks", "Mks").replace("Grade", "Gr"))
        hdr_style(cell, bg=SUBHDR)

    # Data rows
    for r_idx, (_, row) in enumerate(df[all_cols].iterrows(), 3):
        bg = PASS_BG if str(row.get("Result", "")).upper() == "PASS" else \
             (FAIL_BG if str(row.get("Result", "")).upper() == "FAIL" else None)
        is_alt = r_idx % 2 == 0
        for c_idx, col in enumerate(all_cols, 1):
            val = row.get(col, "")
            if pd.isna(val):
                val = ""
            cell = ws1.cell(r_idx, c_idx, val)
            row_bg = bg or (ALT_ROW if is_alt else None)
            data_style(cell, center=(c_idx > 2), bg=row_bg)
            if col == "Total":
                cell.font = Font(bold=True, size=10)

    # Column widths
    widths = {"Roll Number": 14, "Name": 28, "Gender": 8, "Result": 8, "Lang2_Name": 14, "Total": 9}
    for c_idx, col in enumerate(all_cols, 1):
        w = widths.get(col, 10 if "Grade" in col else 9)
        col_width(ws1, c_idx, w)

    # Auto-filter
    ws1.auto_filter.ref = f"A2:{get_column_letter(len(all_cols))}2"

    # ════════════════════════════════════════════════════════════
    # SHEET 2: Class Summary Dashboard
    # ════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("📊 Class Summary")
    ws2.row_dimensions[1].height = 28
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 16
    ws2.column_dimensions["F"].width = 16
    ws2.column_dimensions["G"].width = 16

    t = ws2.cell(1, 1, "Class Summary — Subject-wise Statistics")
    t.font = Font(bold=True, size=13, color=HDR_FG)
    t.fill = PatternFill("solid", fgColor=HDR_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.merge_cells("A1:G1")

    for c_idx, hdr in enumerate(["Subject", "Mean", "Median", "Std Dev", "Highest", "Lowest", "Students"], 1):
        cell = ws2.cell(2, c_idx, hdr)
        hdr_style(cell, bg=SUBHDR)

    for r_idx, mc in enumerate(mark_cols, 3):
        vals = pd.to_numeric(df[mc], errors="coerce").dropna()
        subj = mc.replace("_Marks", "")
        data = [subj, round(vals.mean(), 1), round(vals.median(), 1),
                round(vals.std(), 1), int(vals.max()), int(vals.min()), len(vals)]
        for c_idx, v in enumerate(data, 1):
            cell = ws2.cell(r_idx, c_idx, v)
            data_style(cell, center=(c_idx > 1), bg=ALT_ROW if r_idx % 2 == 0 else None,
                       bold=(c_idx == 1))

    # Total row
    total_vals = pd.to_numeric(df["Total"], errors="coerce").dropna()
    r_tot = 3 + len(mark_cols)
    total_data = ["TOTAL", round(total_vals.mean(), 1), round(total_vals.median(), 1),
                  round(total_vals.std(), 1), int(total_vals.max()), int(total_vals.min()), len(total_vals)]
    for c_idx, v in enumerate(total_data, 1):
        cell = ws2.cell(r_tot, c_idx, v)
        hdr_style(cell, bg="2E75B6")

    # Pass/Fail section
    ws2.cell(r_tot + 2, 1, "Pass/Fail Breakdown").font = Font(bold=True, size=11)
    pf = pass_fail_counts(df)
    for r_idx, (status, count) in enumerate(pf.items(), r_tot + 3):
        ws2.cell(r_idx, 1, status)
        ws2.cell(r_idx, 2, count)
        ws2.cell(r_idx, 3, f"=B{r_idx}/B{r_tot+2+len(pf)+1}")

    # Lang2 breakdown if exists
    if "Lang2_Name" in df.columns:
        groups = get_lang2_breakdown(df)
        r_lang = r_tot + 3 + len(pf) + 2
        ws2.cell(r_lang, 1, "2nd Language Breakdown").font = Font(bold=True, size=11)
        for c_idx, hdr in enumerate(["Language", "Students", "Avg Total"], 1):
            cell = ws2.cell(r_lang + 1, c_idx, hdr)
            hdr_style(cell, bg=SUBHDR)
        for r_idx, (lang, sub) in enumerate(groups.items(), r_lang + 2):
            ws2.cell(r_idx, 1, lang)
            ws2.cell(r_idx, 2, len(sub))
            avg = pd.to_numeric(sub.get("Total", pd.Series(dtype=float)), errors="coerce").mean()
            ws2.cell(r_idx, 3, round(avg, 1) if not pd.isna(avg) else "")

    # ════════════════════════════════════════════════════════════
    # SHEET 3: Grade Distribution
    # ════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("📈 Grade Distribution")
    gs = grade_summary(df)
    if not gs.empty:
        ws3.cell(1, 1, "Grade Distribution by Subject").font = Font(bold=True, size=13, color=HDR_FG)
        ws3.cell(1, 1).fill = PatternFill("solid", fgColor=HDR_BG)
        ws3.cell(1, 1).alignment = Alignment(horizontal="center")
        present_grades = [g for g in GRADE_ORDER if g in gs.columns]
        headers = ["Subject"] + present_grades + ["Total Students"]
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        for c_idx, h in enumerate(headers, 1):
            cell = ws3.cell(2, c_idx, h)
            hdr_style(cell, bg=SUBHDR)
        for r_idx, (subj, row) in enumerate(gs.iterrows(), 3):
            ws3.cell(r_idx, 1, subj).font = Font(bold=True)
            total_students = 0
            for c_idx, g in enumerate(present_grades, 2):
                val = int(row.get(g, 0))
                cell = ws3.cell(r_idx, c_idx, val if val > 0 else "")
                total_students += val
                bg = ALT_ROW if r_idx % 2 == 0 else None
                if g in ("A1", "A2") and val > 0:
                    bg = "E2EFDA"
                elif g in ("E1", "E2", "F") and val > 0:
                    bg = "FFE6E6"
                data_style(cell, center=True, bg=bg)
            cell_total = ws3.cell(r_idx, len(headers), total_students)
            cell_total.font = Font(bold=True)
            cell_total.alignment = Alignment(horizontal="center")
        ws3.column_dimensions["A"].width = 22
        for c in range(2, len(headers) + 1):
            ws3.column_dimensions[get_column_letter(c)].width = 9

    # ════════════════════════════════════════════════════════════
    # SHEET 4: Top Performers
    # ════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("🏆 Top Performers")
    ws4.cell(1, 1, "Top 10 Students by Total Marks").font = Font(bold=True, size=13, color=HDR_FG)
    ws4.cell(1, 1).fill = PatternFill("solid", fgColor=HDR_BG)
    ws4.cell(1, 1).alignment = Alignment(horizontal="center")

    df_sorted = df.sort_values("Total", ascending=False).reset_index(drop=True)
    top_cols = ["Name", "Roll Number", "Gender", "Result"] + \
               ([mc for mc in mark_cols]) + ["Total"]
    if "Lang2_Name" in df.columns:
        top_cols.insert(4, "Lang2_Name")

    available_top_cols = [c for c in top_cols if c in df_sorted.columns]
    ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(available_top_cols))
    for c_idx, col in enumerate(available_top_cols, 1):
        cell = ws4.cell(2, c_idx, col.replace("_Marks", "").replace("_", " "))
        hdr_style(cell, bg=SUBHDR)

    for r_idx, (_, row) in enumerate(df_sorted.head(20)[available_top_cols].iterrows(), 3):
        medal = "🥇" if r_idx == 3 else "🥈" if r_idx == 4 else "🥉" if r_idx == 5 else ""
        for c_idx, col in enumerate(available_top_cols, 1):
            val = row.get(col, "")
            if pd.isna(val):
                val = ""
            if col == "Name" and medal:
                val = f"{medal} {val}"
            cell = ws4.cell(r_idx, c_idx, val)
            data_style(cell, center=(c_idx > 2),
                       bg=PASS_BG if r_idx % 2 == 0 else None,
                       bold=(col == "Total"))

    ws4.column_dimensions["A"].width = 28
    ws4.column_dimensions["B"].width = 14
    for c in range(3, len(available_top_cols) + 1):
        ws4.column_dimensions[get_column_letter(c)].width = 10

    ws4.auto_filter.ref = f"A2:{get_column_letter(len(available_top_cols))}2"
    ws4.freeze_panes = "A3"

    # ════════════════════════════════════════════════════════════
    # SHEET 5: Needs Attention (bottom / failed)
    # ════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet("⚠️ Needs Attention")
    ws5.cell(1, 1, "Students Needing Academic Support").font = Font(bold=True, size=13, color=HDR_FG)
    ws5.cell(1, 1).fill = PatternFill("solid", fgColor="C00000")
    ws5.cell(1, 1).alignment = Alignment(horizontal="center")

    # Flag students: any fail/E grade OR total below class average
    avg_total = df["Total"].mean()
    flagged = df[
        (df.get("Result", pd.Series(["PASS"] * len(df))).isin(["FAIL", "COMP", "ESSEN"])) |
        (df["Total"] < avg_total * 0.75)
    ].sort_values("Total").reset_index(drop=True)

    flagged_cols = ["Name", "Roll Number", "Gender", "Result"] + mark_cols + ["Total"]
    flagged_cols = [c for c in flagged_cols if c in flagged.columns]
    ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(flagged_cols))
    for c_idx, col in enumerate(flagged_cols, 1):
        cell = ws5.cell(2, c_idx, col.replace("_Marks", "").replace("_", " "))
        hdr_style(cell, bg="C00000")

    for r_idx, (_, row) in enumerate(flagged[flagged_cols].iterrows(), 3):
        for c_idx, col in enumerate(flagged_cols, 1):
            val = row.get(col, "")
            if pd.isna(val):
                val = ""
            cell = ws5.cell(r_idx, c_idx, val)
            data_style(cell, center=(c_idx > 2), bg=FAIL_BG if r_idx % 2 == 0 else "FFF2CC")

    ws5.column_dimensions["A"].width = 28
    ws5.column_dimensions["B"].width = 14
    for c in range(3, len(flagged_cols) + 1):
        ws5.column_dimensions[get_column_letter(c)].width = 10
    ws5.freeze_panes = "A3"

    # ════════════════════════════════════════════════════════════
    # SHEET 6: Subject-wise Rank Lists
    # ════════════════════════════════════════════════════════════
    for mc in mark_cols:
        subj = mc.replace("_Marks", "")
        sheet_name = f"🔢 {subj[:12]}"
        wsx = wb.create_sheet(sheet_name)
        wsx.cell(1, 1, f"Rank List — {subj}").font = Font(bold=True, size=13, color=HDR_FG)
        wsx.cell(1, 1).fill = PatternFill("solid", fgColor=HDR_BG)

        gc = get_grade_col(df, mc)
        rank_cols = ["Name", "Roll Number", "Gender", mc] + ([gc] if gc and gc in df.columns else []) + ["Total"]
        rank_cols = [c for c in rank_cols if c in df.columns]
        wsx.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(rank_cols) + 1)

        headers = ["Rank"] + [c.replace("_Marks", "").replace("_Grade", " Grade").replace("_", " ") for c in rank_cols]
        for c_idx, h in enumerate(headers, 1):
            cell = wsx.cell(2, c_idx, h)
            hdr_style(cell, bg=SUBHDR)

        ranked = df[rank_cols].copy()
        ranked[mc] = pd.to_numeric(ranked[mc], errors="coerce")
        ranked = ranked.sort_values(mc, ascending=False).reset_index(drop=True)

        for r_idx, (_, row) in enumerate(ranked.iterrows(), 3):
            wsx.cell(r_idx, 1, r_idx - 2)
            for c_idx, col in enumerate(rank_cols, 2):
                val = row.get(col, "")
                if pd.isna(val):
                    val = ""
                cell = wsx.cell(r_idx, c_idx, val)
                data_style(cell, center=(c_idx > 3), bg=ALT_ROW if r_idx % 2 == 0 else None)

        wsx.column_dimensions["A"].width = 8
        wsx.column_dimensions["B"].width = 28
        wsx.column_dimensions["C"].width = 14
        for c in range(4, len(rank_cols) + 2):
            wsx.column_dimensions[get_column_letter(c)].width = 11
        wsx.freeze_panes = "A3"
        wsx.auto_filter.ref = f"A2:{get_column_letter(len(rank_cols)+1)}2"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
