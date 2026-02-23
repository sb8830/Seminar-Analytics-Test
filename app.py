# ============================================================
# ENTERPRISE SEMINAR ANALYTICS DASHBOARD
# Version: Production Enterprise
# Author: Raj Analytics System
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Enterprise Seminar Analytics",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# CUSTOM CSS (Attractive UI)
# ============================================================

st.markdown("""
<style>

.main {
    background-color: #f8fbff;
}

.kpi-box {
    background: linear-gradient(135deg,#667eea,#764ba2);
    padding: 20px;
    border-radius: 12px;
    color: white;
    text-align: center;
}

.metric-card {
    background: white;
    padding: 15px;
    border-radius: 12px;
    box-shadow: 0px 4px 10px rgba(0,0,0,0.1);
}

h1, h2, h3 {
    color: #1f3b73;
}

</style>
""", unsafe_allow_html=True)

# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data():

    attendee = pd.read_excel("Offline Indepth Details_Attendees_New25-26.xlsx")
    report = pd.read_excel("Offline Seminar Report-Based On Attendee.xlsx")

    # Fix header issue
    report.columns = report.iloc[0]
    report = report.drop(0)

    # Standardize column names
    attendee.columns = attendee.columns.str.lower().str.strip()
    report.columns = report.columns.str.lower().str.strip()

    # Merge
    merged = pd.merge(
        attendee,
        report,
        how="left",
        on="phone",
        suffixes=("_attendee","_report")
    )

    # Create student_id
    merged["student_id"] = (
        merged["phone"].astype(str)
        + "_"
        + merged["student_name"].astype(str)
    )

    return merged

merged_df = load_data()

# ============================================================
# FILTER REQUIRED COURSES
# ============================================================

required_courses = [
    "Power Of Trading & Investing Combo Course",
    "Power Of Equity Market Strategy Course (Offline)"
]

merged_df = merged_df[
    merged_df["course_name"].isin(required_courses)
]

# ============================================================
# SIDEBAR FILTER
# ============================================================

st.sidebar.title("🎯 Enterprise Filters")

trainer_list = sorted(
    merged_df["trainer_name"].dropna().unique()
)

selected_trainer = st.sidebar.selectbox(
    "Select Trainer",
    trainer_list
)

trainer_df = merged_df[
    merged_df["trainer_name"] == selected_trainer
]

# ============================================================
# GLOBAL KPIs
# ============================================================

st.title("🚀 Enterprise Seminar Analytics Dashboard")

total_leads = trainer_df["student_id"].nunique()

total_attended = trainer_df[
    trainer_df["attendance_status"] == "Attended"
]["student_id"].nunique()

total_converted = trainer_df[
    trainer_df["conversion_status"] == "Converted"
]["student_id"].nunique()

conversion_rate = (
    round(total_converted / total_attended * 100,2)
    if total_attended > 0 else 0
)

total_seminars = trainer_df["seminar_name"].nunique()

hot_leads = trainer_df[
    (trainer_df["attendance_status"] == "Attended") &
    (trainer_df["conversion_status"] != "Converted")
]["student_id"].nunique()

# ============================================================
# KPI DISPLAY
# ============================================================

col1,col2,col3,col4,col5,col6 = st.columns(6)

col1.metric("Total Leads", total_leads)
col2.metric("Total Attended", total_attended)
col3.metric("Total Converted", total_converted)
col4.metric("Conversion %", f"{conversion_rate}%")
col5.metric("Total Seminars", total_seminars)
col6.metric("Hot Leads", hot_leads)

st.markdown("---")

# ============================================================
# SEMINAR LEVEL REPORT
# ============================================================

seminars = trainer_df["seminar_name"].unique()

for seminar in seminars:

    seminar_df = trainer_df[
        trainer_df["seminar_name"] == seminar
    ]

    st.subheader(f"📌 Seminar: {seminar}")

    leads = seminar_df["student_id"].nunique()

    attended = seminar_df[
        seminar_df["attendance_status"] == "Attended"
    ]["student_id"].nunique()

    converted = seminar_df[
        seminar_df["conversion_status"] == "Converted"
    ]["student_id"].nunique()

    conversion = (
        round(converted / attended * 100,2)
        if attended > 0 else 0
    )

    c1,c2,c3,c4 = st.columns(4)

    c1.metric("Leads", leads)
    c2.metric("Attended", attended)
    c3.metric("Converted", converted)
    c4.metric("Conversion %", f"{conversion}%")

    # ========================================================
    # STUDENT SUMMARY
    # ========================================================

    student_summary = seminar_df.groupby("student_id").agg(

        student_name=("student_name","first"),
        phone=("phone","first"),
        attended=("attendance_status",
                  lambda x: (x=="Attended").any()),

        converted=("conversion_status",
                   lambda x: (x=="Converted").any()),

        course=("course_name","first")

    ).reset_index()

    st.dataframe(student_summary, use_container_width=True)

    # ========================================================
    # FUNNEL CHART
    # ========================================================

    funnel = go.Figure(go.Funnel(
        y = ["Leads","Attended","Converted"],
        x = [leads, attended, converted]
    ))

    st.plotly_chart(funnel, use_container_width=True)

    # ========================================================
    # CONVERSION PIE
    # ========================================================

    pie_df = pd.DataFrame({

        "Status":["Converted","Not Converted"],
        "Count":[converted, attended-converted]

    })

    pie = px.pie(
        pie_df,
        names="Status",
        values="Count",
        title="Conversion Distribution"
    )

    st.plotly_chart(pie, use_container_width=True)

    # ========================================================
    # HOT LEADS TABLE
    # ========================================================

    hot = student_summary[
        (student_summary["attended"] == True) &
        (student_summary["converted"] == False)
    ]

    st.write("🔥 Hot Leads")
    st.dataframe(hot, use_container_width=True)

    st.markdown("---")

# ============================================================
# OVERALL ANALYTICS
# ============================================================

st.header("📊 Overall Performance Analytics")

seminar_summary = trainer_df.groupby("seminar_name").agg(

    Leads=("student_id","nunique"),

    Attended=("attendance_status",
              lambda x: (x=="Attended").sum()),

    Converted=("conversion_status",
               lambda x: (x=="Converted").sum())

).reset_index()

seminar_summary["Conversion %"] = (
    seminar_summary["Converted"] /
    seminar_summary["Attended"] * 100
).round(2)

bar = px.bar(
    seminar_summary,
    x="seminar_name",
    y="Conversion %",
    title="Seminar Conversion Comparison",
    color="Conversion %"
)

st.plotly_chart(bar, use_container_width=True)

# ============================================================
# DOWNLOAD REPORT
# ============================================================

st.download_button(

    "📥 Download Trainer Report",

    trainer_df.to_csv(index=False),

    file_name=f"{selected_trainer}_report.csv"

)

# ============================================================
# END
# ============================================================
