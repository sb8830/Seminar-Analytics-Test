import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Seminar Analytics Dashboard",
    layout="wide"
)

st.title("📊 Seminar Analytics Dashboard (Production Version)")

# =========================
# File Upload
# =========================

attendee_file = st.file_uploader(
    "Upload Indepth Attendee File",
    type=["xlsx", "csv"]
)

seminar_file = st.file_uploader(
    "Upload Seminar Report File",
    type=["xlsx", "csv"]
)

if attendee_file is None:
    st.warning("Please upload Indepth Attendee File")
    st.stop()

# =========================
# Load Data
# =========================

@st.cache_data
def load_file(file):

    if file.name.endswith("csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip().str.lower()

    return df

attendee_df = load_file(attendee_file)

if seminar_file:
    seminar_df = load_file(seminar_file)
else:
    seminar_df = pd.DataFrame()

# =========================
# Identify Columns Automatically
# =========================

def find_column(df, keywords):

    for col in df.columns:
        for key in keywords:
            if key in col:
                return col

    return None


phone_col = find_column(attendee_df, ["phone", "mobile", "contact"])
email_col = find_column(attendee_df, ["email", "mail"])
name_col = find_column(attendee_df, ["name", "student"])
payment_col = find_column(attendee_df, ["payment", "amount", "revenue", "paid"])
seminar_col = find_column(attendee_df, ["seminar", "event", "webinar", "session"])
trainer_col = find_column(attendee_df, ["trainer", "mentor", "faculty"])
date_col = find_column(attendee_df, ["date"])

# =========================
# Clean Data
# =========================

attendee_df[payment_col] = pd.to_numeric(
    attendee_df[payment_col],
    errors="coerce"
).fillna(0)

# =========================
# Create Unique Student ID
# =========================

attendee_df["student_id"] = (
    attendee_df[phone_col].astype(str)
    .replace("nan", "")
)

attendee_df.loc[
    attendee_df["student_id"] == "",
    "student_id"
] = attendee_df[email_col]

# =========================
# Create Conversion Flag
# =========================

attendee_df["converted"] = attendee_df[payment_col] > 0

# =========================
# Student Summary
# =========================

student_summary = attendee_df.groupby("student_id").agg(

    student_name=(name_col, "first"),
    phone=(phone_col, "first"),
    email=(email_col, "first"),
    total_payment=(payment_col, "sum"),
    converted=("converted", "max"),
    seminar=(seminar_col, "first"),
    trainer=(trainer_col, "first")

).reset_index()

# =========================
# KPIs
# =========================

total_students = student_summary["student_id"].nunique()

converted_students = student_summary["converted"].sum()

conversion_rate = (
    converted_students / total_students * 100
    if total_students else 0
)

total_revenue = student_summary["total_payment"].sum()

# =========================
# KPI Dashboard
# =========================

k1, k2, k3, k4 = st.columns(4)

k1.metric("Total Students", f"{total_students:,}")
k2.metric("Converted Students", f"{converted_students:,}")
k3.metric("Conversion Rate", f"{conversion_rate:.2f}%")
k4.metric("Total Revenue", f"₹{total_revenue:,.0f}")

st.divider()

# =========================
# Conversion Funnel
# =========================

funnel = go.Figure(go.Funnel(

    y=[
        "Total Students",
        "Converted Students"
    ],

    x=[
        total_students,
        converted_students
    ]

))

st.subheader("Conversion Funnel")

st.plotly_chart(
    funnel,
    use_container_width=True
)

# =========================
# Seminar Wise Analytics
# =========================

st.subheader("Seminar-wise Performance")

seminar_summary = attendee_df.groupby(seminar_col).agg(

    students=("student_id", "nunique"),
    converted=("converted", "sum"),
    revenue=(payment_col, "sum")

).reset_index()

seminar_summary["conversion_rate"] = (
    seminar_summary["converted"]
    / seminar_summary["students"]
    * 100
)

fig1 = px.bar(

    seminar_summary,
    x=seminar_col,
    y="revenue",
    title="Revenue by Seminar"

)

st.plotly_chart(fig1, use_container_width=True)

# =========================
# Trainer Wise Analytics
# =========================

if trainer_col:

    st.subheader("Trainer-wise Performance")

    trainer_summary = attendee_df.groupby(trainer_col).agg(

        students=("student_id", "nunique"),
        converted=("converted", "sum"),
        revenue=(payment_col, "sum")

    ).reset_index()

    trainer_summary["conversion_rate"] = (
        trainer_summary["converted"]
        / trainer_summary["students"]
        * 100
    )

    fig2 = px.bar(

        trainer_summary,
        x=trainer_col,
        y="revenue",
        title="Revenue by Trainer"

    )

    st.plotly_chart(fig2, use_container_width=True)

# =========================
# Daily Revenue Trend
# =========================

if date_col:

    st.subheader("Revenue Trend")

    attendee_df[date_col] = pd.to_datetime(
        attendee_df[date_col],
        errors="coerce"
    )

    trend = attendee_df.groupby(date_col)[payment_col].sum().reset_index()

    fig3 = px.line(

        trend,
        x=date_col,
        y=payment_col,
        title="Daily Revenue Trend"

    )

    st.plotly_chart(fig3, use_container_width=True)

# =========================
# Student Level Table
# =========================

st.subheader("Student Conversion Table")

st.dataframe(
    student_summary.sort_values(
        "total_payment",
        ascending=False
    ),
    use_container_width=True
)

# =========================
# Download Clean Report
# =========================

st.subheader("Download Report")

csv = student_summary.to_csv(index=False)

st.download_button(

    label="Download Student Conversion Report",

    data=csv,

    file_name="student_conversion_report.csv",

    mime="text/csv"

)
