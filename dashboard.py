"""
CBSE Class X Result Analysis Dashboard
Teacher-friendly | Multi-language 2nd subject support
"""

import streamlit as st
import pandas as pd
import student_analysis as sa
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="CBSE Result Analysis",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #f0f4f8; }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 16px 20px;
        border-left: 5px solid #1F4E79;
        margin-bottom: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }
    .metric-val { font-size: 2rem; font-weight: 700; color: #1F4E79; }
    .metric-lbl { font-size: 0.85rem; color: #555; margin-top: 2px; }
    .section-header {
        background: #1F4E79;
        color: white;
        padding: 8px 16px;
        border-radius: 6px;
        font-size: 1.1rem;
        font-weight: 600;
        margin: 18px 0 10px 0;
    }
    .pass-badge { background:#E2EFDA; color:#375623; padding:2px 8px;
                  border-radius:12px; font-size:0.8rem; font-weight:600; }
    .fail-badge { background:#FFE6E6; color:#C00000; padding:2px 8px;
                  border-radius:12px; font-size:0.8rem; font-weight:600; }
    .info-box { background:#EBF5FB; border-left:4px solid #2E86C1;
                padding:10px 14px; border-radius:4px; margin:10px 0; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/en/thumb/2/27/CBSE_Logo.png/200px-CBSE_Logo.png", width=80)
    st.title("⚙️ Settings")
    uploaded_file = st.file_uploader(
        "Upload CBSE Gazette (.txt)",
        type=["txt"],
        help="Upload the CBSE result gazette in standard 2-line per student format"
    )
    st.divider()
    st.markdown("**View Options**")
    show_raw = st.checkbox("Show raw parsed data", value=False)
    top_n = st.slider("Top/Bottom N students", 5, 20, 10)
    st.divider()
    st.markdown(
        "<small>Supports multiple 2nd language options:<br>"
        "Hindi · Telugu · Tamil · Sanskrit · etc.</small>",
        unsafe_allow_html=True
    )


# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("# 🎓 CBSE Class X Result Analysis")
st.markdown("*Comprehensive teacher dashboard — multi-language support*")

if not uploaded_file:
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
        <div class="info-box">
        <b>📂 Upload your CBSE result gazette</b><br>
        Supports the standard CBSE gazette format with subject codes.<br>
        Automatically detects multiple 2nd language choices (Hindi, Telugu, Tamil, etc.)
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""
        **Expected Format (2 lines per student):**
        ```
        20255617   M JOHN SHARMA       184  085  041  086  087   PASS
                                       075 B2 069 C2 068 B2 079 A2 083 B1
        20255618   F PRIYA REDDY       184  018  041  086  087   PASS
                                       092 A1 088 A2 095 A1 091 A1 089 A2
        ```
        """)
    with col2:
        st.markdown("""
        **Subject Codes:**
        | Code | Subject |
        |------|---------|
        | 184  | English |
        | 085  | Hindi   |
        | 018  | Telugu  |
        | 089  | Tamil   |
        | 041  | Maths   |
        | 086  | Science |
        | 087  | Soc.Sci |
        """)
    st.stop()


# ── Parse Data ────────────────────────────────────────────────────────────────
raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
lines = raw_text.splitlines()

with st.spinner("Parsing student records..."):
    df = sa.parse_student_data_from_lines(lines)

if df.empty:
    st.error("❌ No student records found. Please check the file format.")
    st.stop()

df = sa.add_total_marks(df)
mark_cols = sa.get_mark_cols(df)
lang2_groups = sa.get_lang2_breakdown(df)

# School name from file header
school_name = ""
for line in lines[:20]:
    m_school = __import__("re").search(r"SCHOOL\s*[:\-]\s*(.+)", line)
    if m_school:
        school_name = m_school.group(1).strip()
        break

if school_name:
    st.markdown(f"**🏫 School:** {school_name}")

# ── Top-level KPI Cards ───────────────────────────────────────────────────────
st.markdown('<div class="section-header">📌 Class Overview</div>', unsafe_allow_html=True)

total_students = len(df)
pass_count = (df.get("Result", pd.Series([])) == "PASS").sum() if "Result" in df.columns else 0
fail_count = (df.get("Result", pd.Series([])) == "FAIL").sum() if "Result" in df.columns else 0
pass_pct = (pass_count / total_students * 100) if total_students else 0
avg_total = df["Total"].mean()
highest = int(df["Total"].max())
male_count = (df.get("Gender", pd.Series([])) == "M").sum()
female_count = (df.get("Gender", pd.Series([])) == "F").sum()

kpi_cols = st.columns(6)
with kpi_cols[0]:
    st.markdown(f'<div class="metric-card"><div class="metric-val">{total_students}</div><div class="metric-lbl">Total Students</div></div>', unsafe_allow_html=True)
with kpi_cols[1]:
    st.markdown(f'<div class="metric-card" style="border-color:#27ae60"><div class="metric-val" style="color:#27ae60">{pass_count}</div><div class="metric-lbl">Passed ({pass_pct:.1f}%)</div></div>', unsafe_allow_html=True)
with kpi_cols[2]:
    st.markdown(f'<div class="metric-card" style="border-color:#e74c3c"><div class="metric-val" style="color:#e74c3c">{fail_count}</div><div class="metric-lbl">Failed / Compartment</div></div>', unsafe_allow_html=True)
with kpi_cols[3]:
    st.markdown(f'<div class="metric-card" style="border-color:#3498db"><div class="metric-val" style="color:#3498db">{avg_total:.1f}</div><div class="metric-lbl">Class Average (Total)</div></div>', unsafe_allow_html=True)
with kpi_cols[4]:
    st.markdown(f'<div class="metric-card" style="border-color:#9b59b6"><div class="metric-val" style="color:#9b59b6">{highest}</div><div class="metric-lbl">Highest Total</div></div>', unsafe_allow_html=True)
with kpi_cols[5]:
    st.markdown(f'<div class="metric-card" style="border-color:#f39c12"><div class="metric-val" style="color:#f39c12">{male_count}M / {female_count}F</div><div class="metric-lbl">Gender Split</div></div>', unsafe_allow_html=True)

# 2nd Language breakdown badge
if lang2_groups:
    langs = [f"**{lang}**: {len(sub)} students" for lang, sub in sorted(lang2_groups.items())]
    st.markdown(
        f'<div class="info-box">🌐 <b>2nd Language Breakdown:</b> ' +
        " &nbsp;|&nbsp; ".join(langs) + "</div>",
        unsafe_allow_html=True
    )

# ── Overview Charts ───────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📊 Class-wide Analysis</div>', unsafe_allow_html=True)

ov_col1, ov_col2, ov_col3 = st.columns(3)

with ov_col1:
    fig_pf = sa.fig_pass_fail_pie(df)
    if fig_pf:
        st.pyplot(fig_pf)
        plt.close(fig_pf)

with ov_col2:
    fig_subj = sa.fig_subject_comparison(df)
    if fig_subj:
        st.pyplot(fig_subj)
        plt.close(fig_subj)

with ov_col3:
    fig_gender = sa.fig_gender_comparison(df)
    if fig_gender:
        st.pyplot(fig_gender)
        plt.close(fig_gender)

# 2nd language comparison (only if multiple languages)
if len(lang2_groups) > 1:
    fig_lang2 = sa.fig_lang2_comparison(df)
    if fig_lang2:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.pyplot(fig_lang2)
            plt.close(fig_lang2)
        with c2:
            st.markdown("**Average Total by 2nd Language**")
            lang_table = []
            for lang, sub in sorted(lang2_groups.items()):
                lang_table.append({
                    "Language": lang,
                    "Students": len(sub),
                    "Avg Total": round(pd.to_numeric(sub.get("Total", pd.Series(dtype=float)), errors="coerce").mean(), 1),
                    "Pass %": f"{(sub.get('Result','') == 'PASS').sum() / len(sub) * 100:.1f}%" if "Result" in sub.columns else "N/A"
                })
            st.dataframe(pd.DataFrame(lang_table), hide_index=True)

# Correlation heatmap
fig_corr = sa.fig_correlation_heatmap(df)
if fig_corr:
    c1, c2 = st.columns([1, 2])
    with c1:
        st.markdown("**Subject Correlation Heatmap**")
        st.markdown("<small>Shows how performance in one subject relates to another</small>", unsafe_allow_html=True)
    with c2:
        st.pyplot(fig_corr)
        plt.close(fig_corr)

# ── Subject-wise Deep Dive ────────────────────────────────────────────────────
st.markdown('<div class="section-header">📘 Subject-wise Deep Dive</div>', unsafe_allow_html=True)

subject_options = [mc.replace("_Marks", "") for mc in mark_cols]
selected_subj = st.selectbox("Select Subject", subject_options, key="subject_select")
selected_mc = f"{selected_subj}_Marks"

if selected_mc in df.columns:
    # Filter by 2nd language if relevant
    lang_filter = "All Students"
    if selected_subj in [g for g in lang2_groups.keys()]:
        lang_options = ["All Students"] + sorted(lang2_groups.keys())
        lang_filter = st.selectbox("Filter by 2nd Language (for this subject only)", lang_options)

    plot_df = lang2_groups.get(lang_filter, df) if lang_filter != "All Students" else df

    s1, s2, s3, s4 = st.columns(4)
    subj_vals = pd.to_numeric(plot_df[selected_mc], errors="coerce").dropna()
    s1.metric("Average", f"{subj_vals.mean():.1f}")
    s2.metric("Highest", f"{subj_vals.max():.0f}")
    s3.metric("Lowest", f"{subj_vals.min():.0f}")
    s4.metric("Std Dev", f"{subj_vals.std():.1f}")

    col1, col2, col3 = st.columns(3)
    with col1:
        fig_hist = sa.fig_marks_histogram(plot_df, selected_mc)
        if fig_hist:
            st.pyplot(fig_hist)
            plt.close(fig_hist)
    with col2:
        fig_grade = sa.fig_grade_distribution(plot_df, selected_mc)
        if fig_grade:
            st.pyplot(fig_grade)
            plt.close(fig_grade)
    with col3:
        gc = sa.get_grade_col(df, selected_mc)
        if gc and gc in df.columns:
            grade_counts = plot_df[gc].value_counts().reindex(sa.GRADE_ORDER).dropna()
            grade_df = grade_counts.reset_index()
            grade_df.columns = ["Grade", "Count"]
            grade_df["% of Class"] = (grade_df["Count"] / len(plot_df) * 100).round(1).astype(str) + "%"
            st.markdown(f"**Grade Distribution Table**")
            st.dataframe(grade_df, hide_index=True)

    fig_tb = sa.fig_top_bottom(plot_df, selected_mc, n=top_n)
    if fig_tb:
        st.pyplot(fig_tb)
        plt.close(fig_tb)

    # Top / Bottom tables
    t1, t2 = st.columns(2)
    subj_df = plot_df[["Name", "Roll Number", "Gender", selected_mc] +
                      ([sa.get_grade_col(df, selected_mc)] if sa.get_grade_col(df, selected_mc) in df.columns else [])].copy()
    subj_df[selected_mc] = pd.to_numeric(subj_df[selected_mc], errors="coerce")
    subj_sorted = subj_df.sort_values(selected_mc, ascending=False).reset_index(drop=True)
    subj_sorted.index += 1

    with t1:
        st.markdown(f"**🏆 Top {top_n} in {selected_subj}**")
        st.dataframe(subj_sorted.head(top_n), use_container_width=True)
    with t2:
        st.markdown(f"**⚠️ Bottom {top_n} in {selected_subj}**")
        st.dataframe(subj_sorted.tail(top_n).iloc[::-1].reset_index(drop=True), use_container_width=True)

# ── Scatter Plot Explorer ─────────────────────────────────────────────────────
st.markdown('<div class="section-header">🔍 Subject Correlation Explorer</div>', unsafe_allow_html=True)
if len(mark_cols) >= 2:
    sc1, sc2 = st.columns(2)
    with sc1:
        subj_x = st.selectbox("X-axis subject", [mc.replace("_Marks","") for mc in mark_cols], index=0, key="sx")
    with sc2:
        subj_y = st.selectbox("Y-axis subject", [mc.replace("_Marks","") for mc in mark_cols], index=min(2, len(mark_cols)-1), key="sy")
    if subj_x != subj_y:
        fig_sc = sa.fig_scatter_two_subjects(df, f"{subj_x}_Marks", f"{subj_y}_Marks")
        if fig_sc:
            c1, c2 = st.columns([2, 1])
            with c1:
                st.pyplot(fig_sc)
                plt.close(fig_sc)
            with c2:
                x_vals = pd.to_numeric(df[f"{subj_x}_Marks"], errors="coerce")
                y_vals = pd.to_numeric(df[f"{subj_y}_Marks"], errors="coerce")
                mask = x_vals.notna() & y_vals.notna()
                corr_val = x_vals[mask].corr(y_vals[mask])
                st.metric("Pearson Correlation", f"{corr_val:.3f}")
                if abs(corr_val) > 0.7:
                    st.success("Strong correlation")
                elif abs(corr_val) > 0.4:
                    st.info("Moderate correlation")
                else:
                    st.warning("Weak correlation")

# ── Overall Rankings ──────────────────────────────────────────────────────────
st.markdown('<div class="section-header">🏆 Overall Rankings</div>', unsafe_allow_html=True)
rank_df = df[["Name", "Roll Number", "Gender"] + (["Result"] if "Result" in df.columns else []) +
             mark_cols + ["Total"]].copy()
rank_df = rank_df.sort_values("Total", ascending=False).reset_index(drop=True)
rank_df.index += 1

t1, t2 = st.columns(2)
with t1:
    st.markdown(f"**🥇 Top {top_n} Overall**")
    st.dataframe(rank_df.head(top_n).rename(columns={mc: mc.replace("_Marks","") for mc in mark_cols}),
                 use_container_width=True)
with t2:
    st.markdown(f"**⚠️ Bottom {top_n} Overall**")
    st.dataframe(rank_df.tail(top_n).rename(columns={mc: mc.replace("_Marks","") for mc in mark_cols}),
                 use_container_width=True)

# ── Raw Data ──────────────────────────────────────────────────────────────────
if show_raw:
    st.markdown('<div class="section-header">📄 Raw Parsed Data</div>', unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True)
    st.caption(f"{len(df)} students parsed")

# ── Grade Summary Table ───────────────────────────────────────────────────────
st.markdown('<div class="section-header">📈 Grade Distribution Summary</div>', unsafe_allow_html=True)
gs = sa.grade_summary(df)
if not gs.empty:
    # Keep only grades that appear
    gs_display = gs[[c for c in sa.GRADE_ORDER if c in gs.columns]]
    st.dataframe(gs_display.style.background_gradient(
        cmap="RdYlGn_r", axis=None, subset=[c for c in ["E1", "E2", "F"] if c in gs_display.columns]
    ).background_gradient(
        cmap="Greens", axis=None, subset=[c for c in ["A1", "A2"] if c in gs_display.columns]
    ), use_container_width=True)

# ── Excel Download ────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📥 Export to Excel</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
📊 The Excel report includes:<br>
• <b>All Students</b> — Complete data with colour-coded pass/fail rows<br>
• <b>Class Summary</b> — Subject statistics + 2nd language breakdown<br>
• <b>Grade Distribution</b> — Grade counts by subject<br>
• <b>Top Performers</b> — Medal rankings for top 20 students<br>
• <b>Needs Attention</b> — Students below 75% of class average or failed<br>
• <b>Per-subject Rank Lists</b> — Ranked lists for each subject
</div>
""", unsafe_allow_html=True)

if st.button("🔄 Generate Excel Report", type="primary"):
    with st.spinner("Building teacher-friendly Excel report..."):
        try:
            excel_bytes = sa.build_excel_report(df)
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_bytes,
                file_name=f"CBSE_Result_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.success("✅ Excel report ready! Click the button above to download.")
        except Exception as e:
            st.error(f"Error generating report: {e}")
            import traceback
            st.code(traceback.format_exc())
