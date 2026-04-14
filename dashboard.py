import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re
import student_analysis as sa

st.set_page_config(
    page_title="CBSE Result Analysis",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

GRADE_COLORS_HEX = {
    "A1":"#1a9850","A2":"#66bd63",
    "B1":"#3182bd","B2":"#6baed6",
    "C1":"#fdae61","C2":"#f46d43",
    "D1":"#d73027","D2":"#a50026",
}

def grade_palette(grades):
    return [GRADE_COLORS_HEX.get(g, "#aaaaaa") for g in grades]

SUBJECTS    = sa.SUBJECTS
SUBJ_LABELS = sa.SUBJ_LABELS

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🎓 CBSE Analyser")
    st.divider()
    uploaded = st.file_uploader("Upload Gazette (.txt)", type=["txt"])
    st.divider()
    st.subheader("Display Options")
    top_n    = st.slider("Top / Bottom N", 5, 20, 10)
    show_raw = st.checkbox("Show raw data table", False)
    st.divider()
    st.caption("Supports Hindi · Telugu · Painting as 2nd language options")

# ══════════════════════════════════════════════════════════════════════════════
# LANDING
# ══════════════════════════════════════════════════════════════════════════════
if not uploaded:
    st.title("🎓 CBSE Class X Result Analyser")
    st.info("Upload the CBSE gazette `.txt` file from the sidebar to get started.")
    with st.expander("📄 Expected file format"):
        st.code(
            "20255617   M STUDENT NAME                184    085    041    086    087   PASS\n"
            "                                         075 B2 069 C2 068 B2 079 A2 083 B1",
            language="text"
        )
        col1, col2 = st.columns(2)
        col1.markdown("**Subject Codes**\n| Code | Subject |\n|------|--------|\n| 184 | English |\n| 085 | Hindi |\n| 089 / 018 | Telugu |\n| 041 | Maths |\n| 086 | Science |\n| 087 | Social Sci |")
        col2.markdown("**Grades**\n| Grade | Range |\n|-------|-------|\n| A1 | 91–100 |\n| A2 | 81–90 |\n| B1 | 71–80 |\n| B2 | 61–70 |\n| C1 | 51–60 |\n| C2 | 41–50 |\n| D1 | 33–40 |")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PARSE
# ══════════════════════════════════════════════════════════════════════════════
raw   = uploaded.read().decode("utf-8", errors="ignore")
lines = raw.splitlines()

school_name = "School"
for ln in lines[:20]:
    m = re.search(r"SCHOOL\s*[:\-\s]+\d+\s+(.+)", ln)
    if m:
        school_name = m.group(1).strip()
        break

df = sa.parse(lines)
if df.empty:
    st.error("No student records found. Check the file format.")
    st.stop()

mark_cols   = [f"{s}_M" for s in SUBJECTS]
lang_groups = df.groupby("Lang2_Name")

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.title(f"🎓 {school_name}")
st.caption(f"CBSE Class X Results 2023 — {len(df)} students")

# ── KPI metrics ───────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total Students",  len(df))
k2.metric("Pass Rate",       "100%")
k3.metric("Class Average",   f"{df['Total'].mean():.1f} / 500")
k4.metric("Highest Score",   f"{int(df['Total'].max())} / 500")
k5.metric("Girls",           int((df.Gender == "F").sum()))
k6.metric("Boys",            int((df.Gender == "M").sum()))

# ── Language group summary ────────────────────────────────────────────────────
lang_summary = "  |  ".join(
    f"🌐 **{lang}**: {len(g)} students" for lang, g in lang_groups
)
st.markdown(lang_summary)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# OVERVIEW CHARTS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📊 Class Overview")

def make_fig(figsize=(5, 3.8)):
    fig, ax = plt.subplots(figsize=figsize)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.grid(axis="y", linewidth=0.7, alpha=0.4)
    return fig, ax

col1, col2, col3 = st.columns(3)

with col1:
    fig, ax = make_fig()
    avgs   = [pd.to_numeric(df[f"{s}_M"], errors="coerce").mean() for s in SUBJECTS]
    colors = ["#2E75B6","#1A7A4A","#E8A838","#6B2FBE","#D63384"]
    bars   = ax.bar([SUBJ_LABELS[s] for s in SUBJECTS], avgs, color=colors,
                    width=0.55, zorder=2, edgecolor="white", linewidth=0.8)
    for bar, v in zip(bars, avgs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                f"{v:.1f}", ha="center", va="bottom", fontsize=8, fontweight="600")
    ax.set_ylim(0, 105)
    ax.set_title("Average Marks by Subject", fontsize=10, fontweight="600", pad=8)
    ax.tick_params(axis="x", labelsize=7.5, rotation=15)
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

with col2:
    fig, ax = make_fig()
    x = np.arange(len(SUBJECTS))
    w = 0.36
    m_avgs = [pd.to_numeric(df[df.Gender=="M"][f"{s}_M"], errors="coerce").mean() for s in SUBJECTS]
    f_avgs = [pd.to_numeric(df[df.Gender=="F"][f"{s}_M"], errors="coerce").mean() for s in SUBJECTS]
    ax.bar(x - w/2, m_avgs, w, label="Male",   color="#3182BD", alpha=0.9, zorder=2)
    ax.bar(x + w/2, f_avgs, w, label="Female", color="#D63384", alpha=0.9, zorder=2)
    ax.set_xticks(x)
    ax.set_xticklabels([SUBJ_LABELS[s] for s in SUBJECTS], fontsize=7.5, rotation=15)
    ax.set_ylim(0, 105)
    ax.set_title("Male vs Female by Subject", fontsize=10, fontweight="600", pad=8)
    ax.legend(fontsize=8, frameon=False)
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

with col3:
    present_g = [g for g in sa.GRADE_ORDER
                 if any(df[f"{s}_G"].eq(g).any() for s in SUBJECTS)]
    mat = pd.DataFrame(
        {g: [df[f"{s}_G"].eq(g).sum() for s in SUBJECTS] for g in present_g},
        index=[SUBJ_LABELS[s] for s in SUBJECTS]
    )
    fig, ax = make_fig()
    im = ax.imshow(mat.values, cmap=plt.cm.RdYlGn, aspect="auto",
                   vmin=0, vmax=mat.values.max())
    ax.set_xticks(range(len(present_g)))
    ax.set_xticklabels(present_g, fontsize=8)
    ax.set_yticks(range(len(SUBJECTS)))
    ax.set_yticklabels(mat.index, fontsize=8)
    for i in range(len(SUBJECTS)):
        for j in range(len(present_g)):
            v = int(mat.values[i, j])
            if v > 0:
                ax.text(j, i, str(v), ha="center", va="center", fontsize=9, fontweight="600",
                        color="white" if mat.values[i, j] > mat.values.max() * 0.5 else "#333")
    ax.set_title("Grade Count Heatmap", fontsize=10, fontweight="600", pad=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

col4, col5 = st.columns([3, 2])

with col4:
    fig, ax = make_fig(figsize=(7, 3.5))
    vals = df["Total"].dropna()
    ax.hist(vals, bins=18, color="#2E75B6", edgecolor="white", linewidth=0.8, alpha=0.88, zorder=2)
    ax.axvline(vals.mean(),   color="#E8A838", lw=2, linestyle="--", label=f"Mean: {vals.mean():.1f}")
    ax.axvline(vals.median(), color="#1A7A4A", lw=2, linestyle="--", label=f"Median: {vals.median():.1f}")
    ax.set_title("Total Marks Distribution", fontsize=10, fontweight="600", pad=8)
    ax.set_xlabel("Total Marks (/500)", fontsize=8)
    ax.set_ylabel("Students", fontsize=8)
    ax.legend(fontsize=8, frameon=False)
    ax.tick_params(labelsize=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

with col5:
    lang_names = [ln for ln, _ in lang_groups]
    lang_avgs  = [grp["Total"].mean() for _, grp in lang_groups]
    lang_cnts  = [len(grp) for _, grp in lang_groups]
    fig, ax = make_fig(figsize=(4.5, 3.5))
    bar_colors = ["#2E75B6","#1A7A4A","#E8A838","#6B2FBE"]
    bars = ax.bar(lang_names, lang_avgs,
                  color=bar_colors[:len(lang_names)], width=0.5, zorder=2)
    for bar, avg, cnt in zip(bars, lang_avgs, lang_cnts):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f"{avg:.1f}\n(n={cnt})", ha="center", va="bottom", fontsize=8, fontweight="600")
    ax.set_ylim(0, max(lang_avgs) * 1.25)
    ax.set_title("Avg Total by 2nd Language", fontsize=10, fontweight="600", pad=8)
    ax.tick_params(labelsize=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SUBJECT DEEP DIVE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📘 Subject Deep Dive")

subj_sel = st.selectbox(
    "Select subject",
    options=SUBJECTS,
    format_func=lambda s: SUBJ_LABELS[s]
)
mc = f"{subj_sel}_M"
gc = f"{subj_sel}_G"
s  = pd.to_numeric(df[mc], errors="coerce").dropna()

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Average",  f"{s.mean():.1f}")
m2.metric("Median",   f"{s.median():.1f}")
m3.metric("Highest",  int(s.max()))
m4.metric("Lowest",   int(s.min()))
m5.metric("A1+A2 %",  f"{df[gc].isin(['A1','A2']).sum() / len(df) * 100:.0f}%")

ca, cb, cc = st.columns(3)

with ca:
    fig, ax = make_fig()
    ax.hist(s, bins=14, color="#2E75B6", edgecolor="white", linewidth=0.8, alpha=0.88, zorder=2)
    ax.axvline(s.mean(),   color="#E8A838", lw=2, linestyle="--", label=f"Mean {s.mean():.1f}")
    ax.axvline(s.median(), color="#1A7A4A", lw=2, linestyle="--", label=f"Median {s.median():.1f}")
    ax.set_title(f"{SUBJ_LABELS[subj_sel]} — Marks Distribution", fontsize=10, fontweight="600", pad=8)
    ax.set_xlabel("Marks", fontsize=8)
    ax.legend(fontsize=8, frameon=False)
    ax.tick_params(labelsize=8)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

with cb:
    grade_counts = df[gc].value_counts().reindex(sa.GRADE_ORDER).dropna()
    grade_counts = grade_counts[grade_counts > 0]
    fig, ax = make_fig()
    bars = ax.bar(grade_counts.index, grade_counts.values,
                  color=grade_palette(grade_counts.index),
                  edgecolor="white", linewidth=0.8, zorder=2)
    for bar, v in zip(bars, grade_counts.values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                str(int(v)), ha="center", va="bottom", fontsize=9, fontweight="600")
    ax.set_title(f"{SUBJ_LABELS[subj_sel]} — Grade Distribution", fontsize=10, fontweight="600", pad=8)
    ax.tick_params(labelsize=9)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

with cc:
    fig, ax = make_fig()
    m_vals = pd.to_numeric(df[df.Gender=="M"][mc], errors="coerce").dropna()
    f_vals = pd.to_numeric(df[df.Gender=="F"][mc], errors="coerce").dropna()
    bp = ax.boxplot([m_vals, f_vals], labels=["Male", "Female"],
                    patch_artist=True, widths=0.4,
                    medianprops=dict(color="white", linewidth=2))
    bp["boxes"][0].set_facecolor("#3182BD")
    bp["boxes"][1].set_facecolor("#D63384")
    for element in ["whiskers", "caps", "fliers"]:
        for item in bp[element]:
            item.set(color="#666666", linewidth=1.2)
    ax.set_title(f"{SUBJ_LABELS[subj_sel]} — Gender Boxplot", fontsize=10, fontweight="600", pad=8)
    ax.tick_params(labelsize=9)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

t1, t2 = st.columns(2)
df_subj = df[["Rank", "Name", "Roll", "Gender", mc, gc, "Total"]].copy()
df_subj[mc] = pd.to_numeric(df_subj[mc], errors="coerce")
df_subj = df_subj.sort_values(mc, ascending=False).reset_index(drop=True)
df_subj.index += 1

with t1:
    st.markdown(f"**🏆 Top {top_n} in {SUBJ_LABELS[subj_sel]}**")
    st.dataframe(
        df_subj.head(top_n).rename(columns={mc: "Marks", gc: "Grade"}),
        use_container_width=True, hide_index=False
    )
with t2:
    st.markdown(f"**⚠️ Bottom {top_n} in {SUBJ_LABELS[subj_sel]}**")
    st.dataframe(
        df_subj.tail(top_n).iloc[::-1].reset_index(drop=True).rename(columns={mc: "Marks", gc: "Grade"}),
        use_container_width=True, hide_index=True
    )

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# OVERALL RANKINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🏆 Overall Rankings")

df_rank = df.sort_values("Total", ascending=False).reset_index(drop=True)
df_rank.index += 1
disp_cols  = ["Rank", "Name", "Gender", "Lang2_Name",
              "English_M", "Lang2_M", "Maths_M", "Science_M", "Social_M", "Total"]
rename_map = {"Lang2_Name": "2nd Lang", "English_M": "English", "Lang2_M": "Lang2",
              "Maths_M": "Maths", "Science_M": "Science", "Social_M": "Social"}

r1, r2 = st.columns(2)
with r1:
    st.markdown(f"**🥇 Top {top_n} Students**")
    st.dataframe(
        df_rank[disp_cols].head(top_n).rename(columns=rename_map),
        use_container_width=True, hide_index=True
    )
with r2:
    st.markdown(f"**⬇️ Bottom {top_n} Students**")
    st.dataframe(
        df_rank[disp_cols].tail(top_n).iloc[::-1].reset_index(drop=True).rename(columns=rename_map),
        use_container_width=True, hide_index=True
    )

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# GRADE SUMMARY TABLE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📈 Grade Summary")

present_g = [g for g in sa.GRADE_ORDER
             if any(df[f"{s}_G"].eq(g).any() for s in SUBJECTS)]
gs_data = []
for subj in SUBJECTS:
    row = {"Subject": SUBJ_LABELS[subj]}
    for g in present_g:
        row[g] = int(df[f"{subj}_G"].eq(g).sum())
    a1a2 = row.get("A1", 0) + row.get("A2", 0)
    row["A1+A2 %"] = f"{a1a2 / len(df) * 100:.0f}%"
    gs_data.append(row)

gs_df = pd.DataFrame(gs_data).set_index("Subject")
st.dataframe(gs_df, use_container_width=True)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# NEEDS ATTENTION
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("⚠️ Students Needing Attention")

thr  = df["Total"].mean() - df["Total"].std()
flag = df[mark_cols].lt(60).any(axis=1) | df["Total"].lt(thr)
df_na = df[flag].sort_values("Total").reset_index(drop=True)
st.caption(f"{len(df_na)} students flagged — any subject below 60 marks, or total below {thr:.0f}")
st.dataframe(
    df_na[["Rank", "Name", "Gender", "English_M", "Lang2_M", "Maths_M", "Science_M", "Social_M", "Total"]]
    .rename(columns=rename_map),
    use_container_width=True, hide_index=True
)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# RAW DATA
# ══════════════════════════════════════════════════════════════════════════════
if show_raw:
    st.subheader("📄 Raw Parsed Data")
    st.dataframe(df, use_container_width=True)
    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📥 Export to Excel")
st.markdown("""
The Excel report contains **10 formatted sheets**:
All Students · Dashboard · Grade Distribution · Rank List ·
English · 2nd Language · Mathematics · Science · Social Science · Needs Attention
""")

if st.button("Generate Excel Report", type="primary"):
    with st.spinner("Building Excel report..."):
        try:
            data = sa.build_excel(df)
            st.download_button(
                "📥 Download CBSE_Result_Analysis.xlsx",
                data=data,
                file_name="CBSE_Result_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            import traceback
            st.error(f"Error: {e}")
            st.code(traceback.format_exc())
