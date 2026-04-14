import re
import pandas as pd
import numpy as np
from io import BytesIO

SUBJECT_CODE_MAP = {
    "184": "English", "085": "Hindi", "089": "Telugu", "018": "Telugu",
    "041": "Maths", "086": "Science", "087": "Social", "049": "Painting",
    "165": "Hindi(B)", "241": "Maths(B)", "002": "Hindi",
}

GRADE_ORDER = ["A1","A2","B1","B2","C1","C2","D1","D2","E1","E2","F"]

GRADE_COLORS = {
    "A1":"1a9850","A2":"66bd63",
    "B1":"3182bd","B2":"6baed6",
    "C1":"fdae61","C2":"f46d43",
    "D1":"d73027","D2":"a50026",
    "E1":"969696","E2":"636363","F":"252525",
}

SUBJECTS     = ["English","Lang2","Maths","Science","Social"]
SUBJ_LABELS  = {"English":"English","Lang2":"2nd Language",
                "Maths":"Mathematics","Science":"Science","Social":"Social Science"}


def parse(lines_or_path):
    if isinstance(lines_or_path, str):
        with open(lines_or_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.read().splitlines()
    else:
        lines = lines_or_path

    records = []
    i = 0
    while i < len(lines):
        line = lines[i].rstrip("\r")
        m = re.match(
            r"(\d{8})\s+([MF])\s+(.+?)\s{3,}((?:\s*\d{3})+)\s+(PASS|FAIL|COMP|ESSEN|ABSE)",
            line
        )
        if m:
            roll, gender, name = m.group(1), m.group(2), m.group(3).strip()
            codes  = re.findall(r"\d{3}", m.group(4))
            result = m.group(5)
            j = i + 1
            while j < len(lines) and not re.search(r"\d{2,3}\s+[A-F]\d?", lines[j]):
                j += 1
            mg = re.findall(r"(\d{1,3})\s+([A-F]\d?)", lines[j]) if j < len(lines) else []

            rec = {
                "Roll":      roll,
                "Name":      name,
                "Gender":    gender,
                "Result":    result,
                "Lang2_Name": SUBJECT_CODE_MAP.get(codes[1], "Lang2") if len(codes) > 1 else "Lang2",
            }
            for idx, subj in enumerate(SUBJECTS):
                rec[f"{subj}_M"] = int(mg[idx][0]) if idx < len(mg) else np.nan
                rec[f"{subj}_G"] = mg[idx][1]       if idx < len(mg) else ""
            records.append(rec)
            i = j + 1
        else:
            i += 1

    df = pd.DataFrame(records)
    mark_cols = [f"{s}_M" for s in SUBJECTS]
    df["Total"] = df[mark_cols].sum(axis=1)
    df["Rank"]  = df["Total"].rank(method="min", ascending=False).astype(int)
    return df


def build_excel(df):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule

    wb = Workbook()

    # ── palette ────────────────────────────────────────────────────────────────
    C_DARK   = "1B3A5C"
    C_MID    = "2E75B6"
    C_LIGHT  = "D6E4F0"
    C_ACCENT = "E8A838"
    C_GREEN  = "1A7A4A"
    C_GREEN2 = "D8F0E4"
    C_RED    = "C0392B"
    C_RED2   = "FAE0DD"
    C_WHITE  = "FFFFFF"
    C_GRAY   = "F4F6F9"
    C_TEXT   = "1A1A2E"

    thin = Side(style="thin",   color="CBD5E0")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    gf = {g: PatternFill("solid", fgColor=c) for g, c in GRADE_COLORS.items()}

    def hdr(cell, text, bg=C_DARK, fg=C_WHITE, size=10, bold=True):
        cell.value = text
        cell.font  = Font(name="Calibri", bold=bold, color=fg, size=size)
        cell.fill  = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd

    def dat(cell, val, bold=False, center=True, bg=None, color=C_TEXT, size=10):
        cell.value = val
        cell.font  = Font(name="Calibri", bold=bold, color=color, size=size)
        if bg: cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
        cell.border = brd

    def title_row(ws, text, ncols, row=1, bg=C_DARK):
        c = ws.cell(row, 1, text)
        c.font = Font(name="Calibri", bold=True, color=C_WHITE, size=14)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = brd
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        ws.row_dimensions[row].height = 32

    def col_w(ws, col, w):
        ws.column_dimensions[get_column_letter(col)].width = w

    def color_scale(ws, ref):
        ws.conditional_formatting.add(ref, ColorScaleRule(
            start_type="min",  start_color="FA8072",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max",    end_color="63BE7B"
        ))

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 1 — All Students
    # ══════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "All Students"
    ws1.sheet_view.showGridLines = False

    cols1   = ["Rank","Roll","Name","Gender","Lang2_Name","Result",
               "English_M","English_G","Lang2_M","Lang2_G",
               "Maths_M","Maths_G","Science_M","Science_G","Social_M","Social_G","Total"]
    hdrs1   = ["Rank","Roll No","Name","Gender","2nd Lang","Result",
               "Eng","Eng Gr","Lang2","L2 Gr",
               "Maths","Math Gr","Science","Sci Gr","Social","Soc Gr","Total/500"]
    n1      = len(cols1)

    title_row(ws1, "CBSE Class X Results 2023 — Ganges Valley School, Bachupally", n1)
    ws1.row_dimensions[2].height = 24
    for ci, lbl in enumerate(hdrs1, 1):
        hdr(ws1.cell(2, ci), lbl, bg=C_MID)

    df1 = df.sort_values("Total", ascending=False).reset_index(drop=True)
    for ri, (_, row) in enumerate(df1[cols1].iterrows(), 3):
        alt = ri % 2 == 0
        rbg = C_GRAY if alt else C_WHITE
        for ci, col in enumerate(cols1, 1):
            v = row[col]
            if pd.isna(v): v = ""
            cell = ws1.cell(ri, ci)
            if col.endswith("_M") and v != "":
                cell.value = int(v)
            elif col == "Total" and v != "":
                cell.value = int(v)
            elif col == "Rank":
                medal = {1:"🥇",2:"🥈",3:"🥉"}.get(int(v),"")
                cell.value = f"{medal} {int(v)}" if medal else int(v)
            else:
                cell.value = v
            dat(cell, cell.value, bg=rbg, center=(ci != 3))
            if col.endswith("_G") and v in gf:
                cell.fill = gf[v]
                cell.font = Font(name="Calibri", bold=True, color=C_WHITE, size=10)
            if col == "Total":
                cell.font = Font(name="Calibri", bold=True, color=C_DARK, size=11)

    widths1 = [9,13,30,7,10,7,7,8,7,8,7,8,8,8,7,8,10]
    for i, w in enumerate(widths1, 1): col_w(ws1, i, w)
    ws1.freeze_panes = "C3"
    ws1.auto_filter.ref = f"A2:{get_column_letter(n1)}2"
    color_scale(ws1, f"{get_column_letter(n1)}3:{get_column_letter(n1)}{2+len(df)}")

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 2 — Dashboard
    # ══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Dashboard")
    ws2.sheet_view.showGridLines = False
    for c in range(1, 14): col_w(ws2, c, 14)
    ws2.column_dimensions["A"].width = 24

    title_row(ws2, "Class Dashboard — Performance Summary", 12)

    # KPI row
    kpis = [
        ("Total Students", len(df),               C_MID,    1),
        ("Pass Rate",      "100%",                 C_GREEN,  3),
        ("Class Average",  f"{df['Total'].mean():.1f}/500", C_ACCENT, 5),
        ("Highest Score",  f"{int(df['Total'].max())}/500", "6B2FBE", 7),
        ("Boys / Girls",   f"{int((df.Gender=='M').sum())} / {int((df.Gender=='F').sum())}", "0D6EFD", 9),
        ("School",         "57643",                "555555", 11),
    ]
    ws2.row_dimensions[3].height = 14
    ws2.row_dimensions[4].height = 32
    for label, val, color, col in kpis:
        lc = ws2.cell(3, col, label)
        lc.font = Font(name="Calibri", size=9, color="777777")
        lc.alignment = Alignment(horizontal="center")
        lc.fill = PatternFill("solid", fgColor="F8FAFF")
        ws2.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col+1)
        vc = ws2.cell(4, col, val)
        vc.font = Font(name="Calibri", bold=True, size=16, color=color)
        vc.fill = PatternFill("solid", fgColor="F8FAFF")
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = Border(bottom=Side(style="thick", color=color))
        ws2.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)

    # Subject stats table
    ws2.row_dimensions[6].height = 22
    ws2.merge_cells("A6:H6")
    title_row(ws2, "Subject-wise Statistics", 8, row=6, bg=C_MID)
    for ci, h in enumerate(["Subject","Average","Median","Highest","Lowest","Std Dev","A1+A2 %","A1+A2 Count"], 1):
        hdr(ws2.cell(7, ci), h, bg=C_DARK)
    for ri, subj in enumerate(SUBJECTS, 8):
        s = pd.to_numeric(df[f"{subj}_M"], errors="coerce").dropna()
        a1a2 = df[f"{subj}_G"].isin(["A1","A2"]).sum()
        pct  = f"{a1a2/len(df)*100:.0f}%"
        vals = [SUBJ_LABELS[subj], round(s.mean(),1), round(s.median(),1),
                int(s.max()), int(s.min()), round(s.std(),1), pct, int(a1a2)]
        alt  = ri % 2 == 0
        for ci, v in enumerate(vals, 1):
            dat(ws2.cell(ri, ci), v, bg=C_GRAY if alt else C_WHITE,
                bold=(ci == 1), center=(ci != 1))
    # Total row
    ts = df["Total"]
    tr = ["TOTAL  (/500)", round(ts.mean(),1), round(ts.median(),1),
          int(ts.max()), int(ts.min()), round(ts.std(),1), "100%", len(df)]
    ws2.merge_cells("A13:H13") if False else None
    for ci, v in enumerate(tr, 1):
        hdr(ws2.cell(13, ci), v, bg=C_ACCENT, fg=C_DARK)

    # Gender comparison
    ws2.row_dimensions[15].height = 22
    ws2.merge_cells("A15:H15")
    title_row(ws2, "Gender Performance Comparison", 8, row=15, bg=C_MID)
    for ci, h in enumerate(["Gender","English","Lang2","Maths","Science","Social","Avg Total","Count"], 1):
        hdr(ws2.cell(16, ci), h, bg=C_DARK)
    for ri, (g, grp) in enumerate(df.groupby("Gender"), 17):
        avgs = [round(pd.to_numeric(grp[f"{s}_M"], errors="coerce").mean(), 1) for s in SUBJECTS]
        row_data = [("Female 👩" if g=="F" else "Male 👦")] + avgs + [round(grp["Total"].mean(),1), len(grp)]
        for ci, v in enumerate(row_data, 1):
            dat(ws2.cell(ri, ci), v, bg="FFF0F5" if g=="F" else "EFF6FF",
                bold=(ci==1), center=(ci!=1))

    # Lang2 breakdown
    ws2.row_dimensions[20].height = 22
    ws2.merge_cells("A20:F20")
    title_row(ws2, "2nd Language Group Analysis", 6, row=20, bg=C_MID)
    for ci, h in enumerate(["Language","Students","% of Class","Avg Total","Avg Lang2 Marks","A1+A2 %"], 1):
        hdr(ws2.cell(21, ci), h, bg=C_DARK)
    for ri, (lang, grp) in enumerate(df.groupby("Lang2_Name"), 22):
        l2_avg  = pd.to_numeric(grp["Lang2_M"], errors="coerce").mean()
        a1a2_pct= f"{grp['Lang2_G'].isin(['A1','A2']).sum()/len(grp)*100:.0f}%"
        vals = [lang, len(grp), f"{len(grp)/len(df)*100:.1f}%",
                round(grp["Total"].mean(),1), round(l2_avg,1), a1a2_pct]
        for ci, v in enumerate(vals, 1):
            dat(ws2.cell(ri, ci), v, bg=C_GRAY if ri%2==0 else C_WHITE,
                bold=(ci==1), center=(ci!=1))

    # Bar chart: subject averages
    _CR = 27  # chart data start row
    ws2.cell(_CR, 1, "Subject"); ws2.cell(_CR, 2, "Average")
    for ri, subj in enumerate(SUBJECTS, _CR+1):
        ws2.cell(ri, 1, SUBJ_LABELS[subj])
        ws2.cell(ri, 2, round(pd.to_numeric(df[f"{subj}_M"], errors="coerce").mean(), 1))

    bar = BarChart()
    bar.type  = "col"; bar.style = 10
    bar.title = "Class Average — Subject Comparison"
    bar.y_axis.title = "Average Marks"; bar.x_axis.title = "Subject"
    bar.width = 16; bar.height = 11
    bar.y_axis.scaling.min = 50; bar.y_axis.scaling.max = 100
    bar.add_data(Reference(ws2, min_col=2, min_row=_CR, max_row=_CR+5), titles_from_data=True)
    bar.set_categories(Reference(ws2, min_col=1, min_row=_CR+1, max_row=_CR+5))
    bar.series[0].graphicalProperties.solidFill = C_MID
    ws2.add_chart(bar, "J1")

    # Gender pie
    _GR = _CR + 7
    ws2.cell(_GR, 1,"Gender"); ws2.cell(_GR,2,"Count")
    for ri,(g,cnt) in enumerate(df["Gender"].value_counts().items(), _GR+1):
        ws2.cell(ri,1,"Male" if g=="M" else "Female"); ws2.cell(ri,2,cnt)
    pie = PieChart(); pie.style=10
    pie.title="Gender Split"; pie.width=10; pie.height=8
    pie.add_data(Reference(ws2,min_col=2,min_row=_GR,max_row=_GR+2),titles_from_data=True)
    pie.set_categories(Reference(ws2,min_col=1,min_row=_GR+1,max_row=_GR+2))
    ws2.add_chart(pie, "J14")

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 3 — Grade Distribution
    # ══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Grade Distribution")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 20
    for i in range(2, 16): col_w(ws3, i, 9)

    present = [g for g in GRADE_ORDER
               if any(df[f"{s}_G"].eq(g).any() for s in SUBJECTS)]
    n3      = 1 + len(present) + 2

    title_row(ws3, "Grade Distribution — Subject × Grade Matrix", n3)
    hdrs3 = ["Subject"] + present + ["Total Stdnts","A1+A2 %"]
    for ci, h in enumerate(hdrs3, 1):
        cell = ws3.cell(2, ci, h)
        bg   = GRADE_COLORS.get(h, C_MID)
        hdr(cell, h, bg=bg, fg="FFFFFF" if h in GRADE_COLORS else C_WHITE)
    ws3.row_dimensions[2].height = 22

    for ri, subj in enumerate(SUBJECTS, 3):
        ws3.cell(ri, 1, SUBJ_LABELS[subj]).font = Font(name="Calibri", bold=True, size=10, color=C_DARK)
        ws3.cell(ri, 1).alignment = Alignment(horizontal="left", vertical="center")
        ws3.cell(ri, 1).border = brd
        ws3.cell(ri, 1).fill   = PatternFill("solid", fgColor=C_GRAY if ri%2==0 else C_WHITE)
        total_n = 0; a1a2_n = 0
        for ci, g in enumerate(present, 2):
            cnt  = int(df[f"{subj}_G"].eq(g).sum())
            cell = ws3.cell(ri, ci, cnt if cnt > 0 else "")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = brd
            if cnt > 0 and g in gf:
                cell.fill = gf[g]
                cell.font = Font(name="Calibri", bold=True, color=C_WHITE, size=10)
            else:
                cell.fill = PatternFill("solid", fgColor=C_GRAY if ri%2==0 else C_WHITE)
                cell.font = Font(name="Calibri", color="BBBBBB", size=10)
            total_n += cnt
            if g in ("A1","A2"): a1a2_n += cnt
        dat(ws3.cell(ri, len(present)+2), total_n, bold=True, bg=C_LIGHT)
        pct = f"{a1a2_n/total_n*100:.0f}%" if total_n else ""
        dat(ws3.cell(ri, len(present)+3), pct, bold=True,
            bg=C_GREEN2 if a1a2_n/max(total_n,1) >= 0.5 else C_RED2)

    # Totals row
    for ci, g in enumerate(present, 2):
        cnt = sum(int(df[f"{s}_G"].eq(g).sum()) for s in SUBJECTS)
        hdr(ws3.cell(8, ci), cnt, bg=C_DARK)
    hdr(ws3.cell(8, 1), "CLASS TOTAL", bg=C_DARK)
    hdr(ws3.cell(8, len(present)+2), len(df)*5, bg=C_DARK)

    # stacked bar chart
    _GD = 10
    ws3.cell(_GD, 1, "Subject")
    use_g = present[:7]
    for ci, g in enumerate(use_g, 2): ws3.cell(_GD, ci, g)
    for ri, subj in enumerate(SUBJECTS, _GD+1):
        ws3.cell(ri, 1, SUBJ_LABELS[subj])
        for ci, g in enumerate(use_g, 2):
            ws3.cell(ri, ci, int(df[f"{subj}_G"].eq(g).sum()))

    bar2 = BarChart(); bar2.type="col"; bar2.grouping="stacked"; bar2.style=10
    bar2.title="Grade Distribution per Subject"
    bar2.y_axis.title="Number of Students"; bar2.width=18; bar2.height=12
    bar2.add_data(Reference(ws3, min_col=2, min_row=_GD, max_col=len(use_g)+1, max_row=_GD+5), titles_from_data=True)
    bar2.set_categories(Reference(ws3, min_col=1, min_row=_GD+1, max_row=_GD+5))
    grade_fills = ["1a9850","66bd63","3182bd","6baed6","fdae61","f46d43","d73027"]
    for i, s in enumerate(bar2.series):
        if i < len(grade_fills): s.graphicalProperties.solidFill = grade_fills[i]
    ws3.add_chart(bar2, "A16")

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 4 — Rank List
    # ══════════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("Rank List")
    ws4.sheet_view.showGridLines = False

    cols4 = ["Rank","Roll","Name","Gender","Lang2_Name",
             "English_M","Lang2_M","Maths_M","Science_M","Social_M","Total"]
    hdrs4 = ["Rank","Roll No","Name","Gender","2nd Lang",
             "English","Lang2","Maths","Science","Social","Total /500"]
    n4    = len(cols4)

    title_row(ws4, "Overall Rank List — CBSE Class X 2023", n4)
    ws4.row_dimensions[2].height = 22
    for ci, h in enumerate(hdrs4, 1): hdr(ws4.cell(2, ci), h, bg=C_MID)

    df4 = df.sort_values("Total", ascending=False).reset_index(drop=True)
    for ri, (_, row) in enumerate(df4[cols4].iterrows(), 3):
        rank = int(row["Rank"])
        rbg  = {"1":"FFF9E6","2":"F5F5F5","3":"FFF0E6"}.get(str(rank), C_GRAY if ri%2==0 else C_WHITE)
        for ci, col in enumerate(cols4, 1):
            v    = row[col]
            if pd.isna(v): v = ""
            cell = ws4.cell(ri, ci)
            if col == "Rank":
                medal = {1:"🥇",2:"🥈",3:"🥉"}.get(rank,"")
                cell.value = f"{medal} {rank}" if medal else rank
            elif col in ("Total",) + tuple(f"{s}_M" for s in SUBJECTS) and v != "":
                cell.value = int(v)
            else:
                cell.value = v
            dat(cell, cell.value, bg=rbg,
                bold=(col=="Total" or rank<=3), center=(ci!=3))
            if col == "Total":
                cell.font = Font(name="Calibri", bold=True, color=C_DARK, size=11)

    ws4.column_dimensions["A"].width = 9
    ws4.column_dimensions["B"].width = 13
    ws4.column_dimensions["C"].width = 30
    ws4.column_dimensions["D"].width = 8
    ws4.column_dimensions["E"].width = 10
    for i in range(6, n4+1): col_w(ws4, i, 10)
    ws4.freeze_panes = "A3"
    ws4.auto_filter.ref = f"A2:{get_column_letter(n4)}2"
    color_scale(ws4, f"{get_column_letter(n4)}3:{get_column_letter(n4)}{2+len(df)}")

    # ══════════════════════════════════════════════════════════════════════════
    # SHEETS 5–9 — Per-subject rank lists
    # ══════════════════════════════════════════════════════════════════════════
    for subj in SUBJECTS:
        label = SUBJ_LABELS[subj]
        ws    = wb.create_sheet(label[:12])
        ws.sheet_view.showGridLines = False

        scols = ["Roll","Name","Gender","Lang2_Name",f"{subj}_M",f"{subj}_G","Total"]
        shdrs = ["Roll No","Name","Gender","2nd Lang","Marks","Grade","Total"]
        ns    = len(scols) + 1

        title_row(ws, f"Subject Rank — {label}", ns)
        ws.row_dimensions[2].height = 22
        hdr(ws.cell(2,1), "Rank", bg=C_MID)
        for ci, h in enumerate(shdrs, 2): hdr(ws.cell(2, ci), h, bg=C_MID)

        dfs = df.sort_values(f"{subj}_M", ascending=False).reset_index(drop=True)
        for ri, (_, row) in enumerate(dfs.iterrows(), 3):
            sr   = ri - 2
            alt  = ri % 2 == 0
            rbg  = {"1":"FFF9E6","2":"F5F5F5","3":"FFF0E6"}.get(str(sr), C_GRAY if alt else C_WHITE)
            medal = {1:"🥇",2:"🥈",3:"🥉"}.get(sr,"")
            c = ws.cell(ri, 1, f"{medal} {sr}" if medal else sr)
            dat(c, c.value, bg=rbg, bold=(sr<=3))
            for ci, col in enumerate(scols, 2):
                v    = row[col]
                if pd.isna(v): v = ""
                cell = ws.cell(ri, ci)
                cell.value = int(v) if isinstance(v, float) and col.endswith("_M") else (int(v) if col=="Total" and v != "" else v)
                dat(cell, cell.value, bg=rbg, center=(ci!=3))
                if col == f"{subj}_G" and v in gf:
                    cell.fill = gf[v]
                    cell.font = Font(name="Calibri", bold=True, color=C_WHITE, size=10)

        ws.column_dimensions["A"].width = 9
        ws.column_dimensions["B"].width = 13
        ws.column_dimensions["C"].width = 30
        for i in range(4, ns+1): col_w(ws, i, 11)
        ws.freeze_panes = "A3"

        # Stats box
        sr  = len(df) + 4
        s   = pd.to_numeric(df[f"{subj}_M"], errors="coerce").dropna()
        a1a2= df[f"{subj}_G"].isin(["A1","A2"]).sum()
        title_row(ws, f"{label} — Summary Statistics", ns, row=sr, bg=C_MID)
        stats_data = [
            ("Average Marks",  round(s.mean(),1)),
            ("Median",         round(s.median(),1)),
            ("Highest",        int(s.max())),
            ("Lowest",         int(s.min())),
            ("Std Deviation",  round(s.std(),1)),
            ("A1 + A2 Count",  int(a1a2)),
            ("A1 + A2 %",      f"{a1a2/len(df)*100:.0f}%"),
        ]
        for i, (lbl, v) in enumerate(stats_data, sr+1):
            lc = ws.cell(i, 1, lbl)
            lc.font  = Font(name="Calibri", bold=True, size=10)
            lc.alignment = Alignment(horizontal="left", vertical="center")
            lc.fill = PatternFill("solid", fgColor=C_GRAY if i%2==0 else C_WHITE)
            lc.border= brd
            vc = ws.cell(i, 2, v)
            vc.font  = Font(name="Calibri", bold=True, size=10, color=C_MID)
            vc.alignment = Alignment(horizontal="center", vertical="center")
            vc.fill = PatternFill("solid", fgColor=C_GRAY if i%2==0 else C_WHITE)
            vc.border= brd

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 10 — Needs Attention
    # ══════════════════════════════════════════════════════════════════════════
    ws10 = wb.create_sheet("Needs Attention")
    ws10.sheet_view.showGridLines = False

    avg_t  = df["Total"].mean()
    std_t  = df["Total"].std()
    thr    = avg_t - std_t
    flag   = (
        df[["English_M","Lang2_M","Maths_M","Science_M","Social_M"]].lt(60).any(axis=1) |
        df["Total"].lt(thr)
    )
    df_na  = df[flag].sort_values("Total").reset_index(drop=True)

    na_cols = ["Rank","Roll","Name","Gender","English_M","Lang2_M",
               "Maths_M","Science_M","Social_M","Total"]
    na_hdrs = ["Rank","Roll No","Name","Gender","English","Lang2",
               "Maths","Science","Social","Total"]
    n10     = len(na_cols)

    title_row(ws10, f"⚠  Needs Attention — {len(df_na)} Students  ⚠", n10, bg=C_RED)
    sub = ws10.cell(2, 1, f"Criteria: Any subject below 60 marks  OR  Total below {thr:.0f}  (class avg {avg_t:.1f} − 1 SD {std_t:.1f})")
    sub.font = Font(name="Calibri", italic=True, size=9, color="888888")
    ws10.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n10)
    ws10.row_dimensions[3].height = 22
    for ci, h in enumerate(na_hdrs, 1): hdr(ws10.cell(3, ci), h, bg="8B0000")

    for ri, (_, row) in enumerate(df_na[na_cols].iterrows(), 4):
        for ci, col in enumerate(na_cols, 1):
            v    = row[col]
            if pd.isna(v): v = ""
            cell = ws10.cell(ri, ci)
            cell.value = int(v) if isinstance(v, float) and v != "" else v
            is_low = col.endswith("_M") and isinstance(v, (int,float)) and v < 60
            bg = "FFD5D5" if is_low else (C_GRAY if ri%2==0 else C_WHITE)
            dat(cell, cell.value, bg=bg, center=(ci!=3))
            if is_low:
                cell.font = Font(name="Calibri", bold=True, color=C_RED, size=10)

    ws10.column_dimensions["A"].width = 7
    ws10.column_dimensions["B"].width = 13
    ws10.column_dimensions["C"].width = 30
    ws10.column_dimensions["D"].width = 8
    for i in range(5, n10+1): col_w(ws10, i, 10)
    ws10.freeze_panes = "A4"
    color_scale(ws10, f"{get_column_letter(n10)}4:{get_column_letter(n10)}{3+len(df_na)}")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
