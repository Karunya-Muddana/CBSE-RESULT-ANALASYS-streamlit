# dashboard.py
import streamlit as st
import pandas as pd
import student_analysis as sa
import matplotlib.pyplot as plt

st.set_page_config(page_title="Student Performance Dashboard", layout="wide")
st.title("🎓 Student Performance Dashboard")

st.sidebar.header("Upload & Settings")
uploaded_file = st.sidebar.file_uploader("Upload CBSE-style .txt file", type=["txt"])
map_names = st.sidebar.checkbox("Map Subject1..Subject5 → ENG, LANG II, MATH, SCI, SOC", value=True)

if uploaded_file:
    raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
    lines = raw_text.splitlines()
    df = sa.parse_student_data_from_lines(lines)

    if df.empty:
        st.error("No students parsed. Check file format. Use the debug prints or sample input.")
    else:
        if map_names:
            df = sa.map_subjects_to_names(df)

        # Add total marks column
        df = sa.add_total_marks(df)

        st.markdown("### 📊 Parsed Data (first 20 rows)")
        st.dataframe(df.head(20))

        # Subject selection
        mark_cols = sa.get_subject_mark_cols(df)
        if not mark_cols:
            st.warning("No subject mark columns detected. Check header mapping.")
        else:
            subject = st.selectbox("Select subject to analyze", mark_cols)

            # guess grade column
            grade_col_guess = None
            possible_grade = subject + " GRADE"
            if possible_grade in df.columns:
                grade_col_guess = possible_grade
            else:
                for c in df.columns:
                    if "GRADE" in c.upper() and subject.split()[0].lower() in c.lower():
                        grade_col_guess = c
                        break

            # ---- Subject Analysis ----
            st.subheader(f"📘 Analysis for {subject}")
            report = sa.analyze_subject(df, subject, grade_col_guess, produce_figs=True)

            st.write("**Top 10 scorers**")
            st.dataframe(report["top10"].reset_index(drop=True))

            st.write("**Bottom 10 scorers**")
            st.dataframe(report["bottom10"].reset_index(drop=True))

            st.write("**Stats**")
            st.json(report["stats"])

            st.write("**Grade counts**")
            if report["grade_counts"] is not None:
                st.bar_chart(report["grade_counts"])
            else:
                st.info("No grade column available for this subject.")

            if report.get("figs"):
                for k, fig in report["figs"].items():
                    st.write(f"**{k.replace('_',' ').title()}**")
                    st.pyplot(fig)

        # ---- Total Analysis ----
        st.subheader("📘 Overall Total Analysis")
        total_report = sa.analyze_total(df, produce_figs=True)

        st.write("**Top 10 by Total Marks**")
        st.dataframe(total_report["top10"].reset_index(drop=True))

        st.write("**Bottom 10 by Total Marks**")
        st.dataframe(total_report["bottom10"].reset_index(drop=True))

        st.write("**Stats (Total Marks)**")
        st.json(total_report["stats"])

        if total_report.get("figs"):
            for k, fig in total_report["figs"].items():
                st.write(f"**{k.replace('_',' ').title()}**")
                st.pyplot(fig)

        # ---- Excel Downloads ----
        st.subheader("📥 Download Reports")
        # Full consolidated report
        excel_bytes = sa.save_full_report_excel_bytes(df)
        st.download_button(
            "📥 Download Full Excel Report (All Subjects + Total)",
            data=excel_bytes,
            file_name="student_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Upload a CBSE-style two-line .txt file (roll/gender/name/subjectcodes then marks+grades).")
    st.markdown("""
    **Sample expected format (two lines per student):**
    ```
    1001 M John Sharma 041 042 043 044 045
    085 A2 071 B1 060 B2 093 A1 056 C1
    ```
    """)
