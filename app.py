import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Advanced Seminar Analytics", layout="wide")

st.title("📊 Advanced Seminar Analytics Dashboard (Production Safe)")

# =============================
# Upload file
# =============================

file = st.file_uploader("Upload Attendee File", type=["xlsx", "csv"])

if file is None:
    st.stop()

# =============================
# Load data safely
# =============================

@st.cache_data
def load_data(file):

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip().str.lower()

    return df

df = load_data(file)

# =============================
# Smart column detection
# =============================

def find_col(keywords):

    for col in df.columns:
        for key in keywords:
            if key in col:
                return col
    return None


phone_col = find_col(["phone", "mobile", "contact"])
email_col = find_col(["email", "mail"])
name_col = find_col(["name", "student"])
payment_col = find_col(["payment", "amount", "paid", "fees"])
seminar_col = find_col(["seminar", "event"])
trainer_col = find_col(["trainer", "faculty", "mentor"])
date_col = find_col(["date"])
course_col = find_col(["course", "service"])
location_col = find_col(["location", "center", "branch", "city"])
source_col = find_col(["source", "lead", "campaign", "medium"])

# =============================
# Create safe columns
# =============================

if payment_col:
    df["payment"] = pd.to_numeric(df[payment_col], errors="coerce").fillna(0)
else:
    df["payment"] = 0

df["phone"] = df[phone_col].astype(str) if phone_col else ""
df["email"] = df[email_col].astype(str) if email_col else ""
df["student_name"] = df[name_col] if name_col else "Unknown"

df["student_id"] = df["phone"]
df.loc[df["student_id"] == "", "student_id"] = df["email"]

df["converted"] = df["payment"] > 0

# =============================
# Student summary SAFE
# =============================

student_summary = df.groupby("student_id").agg(

    student_name=("student_name", "first"),
    phone=("phone", "first"),
    email=("email", "first"),
    total_payment=("payment", "sum"),
    converted=("converted", "max")

).reset_index()

# =============================
# KPIs
# =============================

total_students = student_summary.shape[0]
converted_students = student_summary["converted"].sum()
revenue = student_summary["total_payment"].sum()

conversion_rate = (
    converted_students / total_students * 100
    if total_students else 0
)

# =============================
# KPI DISPLAY
# =============================

c1, c2, c3, c4 = st.columns(4)

c1.metric("Total Students", f"{total_students:,}")
c2.metric("Converted", f"{converted_students:,}")
c3.metric("Conversion Rate", f"{conversion_rate:.2f}%")
c4.metric("Revenue", f"₹{revenue:,.0f}")

st.divider()

# =============================
# Conversion Funnel
# =============================

st.subheader("Conversion Funnel")

fig = go.Figure(go.Funnel(

    y=["Leads", "Converted"],
    x=[total_students, converted_students]

))

st.plotly_chart(fig, use_container_width=True)

# =============================
# Course Analytics
# =============================

if course_col:

    st.subheader("Course Analytics")

    course = df.groupby(course_col).agg(

        students=("student_id", "nunique"),
        revenue=("payment", "sum")

    ).reset_index()

    fig = px.bar(course, x=course_col, y="revenue")

    st.plotly_chart(fig, use_container_width=True)

# =============================
# Trainer Analytics
# =============================

if trainer_col:

    st.subheader("Trainer Analytics")

    trainer = df.groupby(trainer_col).agg(

        students=("student_id", "nunique"),
        revenue=("payment", "sum")

    ).reset_index()

    fig = px.bar(trainer, x=trainer_col, y="revenue")

    st.plotly_chart(fig, use_container_width=True)

# =============================
# Location Analytics
# =============================

if location_col:

    st.subheader("Location Analytics")

    loc = df.groupby(location_col).agg(

        students=("student_id", "nunique"),
        revenue=("payment", "sum")

    ).reset_index()

    fig = px.bar(loc, x=location_col, y="revenue")

    st.plotly_chart(fig, use_container_width=True)

# =============================
# Lead Source Analytics
# =============================

if source_col:

    st.subheader("Lead Source Analytics")

    source = df.groupby(source_col).agg(

        students=("student_id", "nunique"),
        revenue=("payment", "sum")

    ).reset_index()

    fig = px.bar(source, x=source_col, y="revenue")

    st.plotly_chart(fig, use_container_width=True)

# =============================
# Revenue Trend
# =============================

if date_col:

    st.subheader("Revenue Trend")

    df["date"] = pd.to_datetime(df[date_col], errors="coerce")

    trend = df.groupby("date")["payment"].sum().reset_index()

    fig = px.line(trend, x="date", y="payment")

    st.plotly_chart(fig, use_container_width=True)

# =============================
# Student Table
# =============================

st.subheader("Student Table")

st.dataframe(student_summary, use_container_width=True)

# =============================
# Download
# =============================

csv = student_summary.to_csv(index=False)

st.download_button(

    "Download Student Report",
    csv,
    "student_report.csv",
    "text/csv"
)
