import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------
# CONFIG
# ---------------------------------------------------
st.set_page_config(
    page_title="Seminar Analytics PRO",
    page_icon="📊",
    layout="wide"
)

# ---------------------------------------------------
# CSS
# ---------------------------------------------------
st.markdown("""
<style>
.block-container { padding-top: 1rem; }
.stMetric {
    background: #f8fafc;
    padding: 12px;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------
# TITLE
# ---------------------------------------------------
st.title("📊 Seminar Analytics PRO Dashboard")
st.caption("Conversion • Revenue • ROI • Student Tracking")

# ---------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    indepth_file = st.file_uploader(
        "Upload Indepth Attendee File",
        type=["xlsx"],
        key="indepth"
    )

with col2:
    seminar_file = st.file_uploader(
        "Upload Seminar Report File",
        type=["xlsx"],
        key="seminar"
    )

if not indepth_file or not seminar_file:
    st.info("Upload both files to begin.")
    st.stop()


# ---------------------------------------------------
# LOAD FUNCTIONS (CACHED)
# ---------------------------------------------------

@st.cache_data(show_spinner=False)
def load_indepth(file):

    df = pd.read_excel(file, sheet_name=None)

    frames = []

    for name, sheet in df.items():

        sheet.columns = (
            sheet.columns
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

        if "student_name" in sheet.columns:
            frames.append(sheet)

    df = pd.concat(frames, ignore_index=True)

    numeric_cols = [
        "payment_received",
        "total_amount",
        "total_due",
        "total_gst"
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


@st.cache_data(show_spinner=False)
def load_seminar(file):

    df = pd.read_excel(file, header=1)

    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace("\n", "_")
        .str.replace(" ", "_")
    )

    numeric_cols = [
        "total_attended",
        "total_seat_booked_(in_seminar)",
        "actual_expenses",
        "actual_revenue(w/o_gst)_attendees",
        "surplus_or_deficit"
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(
                    df[col]
                    .astype(str)
                    .str.replace(",", ""),
                    errors="coerce"
                )
                .fillna(0)
            )

    df = df[pd.to_numeric(df["sr_no"], errors="coerce").notna()]

    return df


# ---------------------------------------------------
# LOAD DATA
# ---------------------------------------------------

indepth_df = load_indepth(indepth_file)
seminar_df = load_seminar(seminar_file)

# ---------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------

st.sidebar.header("Filters")

locations = seminar_df["location"].dropna().unique()

selected_locations = st.sidebar.multiselect(
    "Location",
    locations
)

profit_filter = st.sidebar.radio(
    "Profitability",
    ["All", "Profit", "Loss"]
)

# ---------------------------------------------------
# FILTER DATA
# ---------------------------------------------------

filtered = seminar_df

if selected_locations:
    filtered = filtered[
        filtered["location"].isin(selected_locations)
    ]

if profit_filter == "Profit":
    filtered = filtered[
        filtered["surplus_or_deficit"] > 0
    ]

if profit_filter == "Loss":
    filtered = filtered[
        filtered["surplus_or_deficit"] < 0
    ]

# ---------------------------------------------------
# KPI CALCULATIONS (FAST)
# ---------------------------------------------------

total_seminars = len(filtered)

total_attended = int(
    filtered["total_attended"].sum()
)

total_revenue = (
    filtered[
        "actual_revenue(w/o_gst)_attendees"
    ].sum()
)

total_expenses = filtered[
    "actual_expenses"
].sum()

profit = filtered[
    "surplus_or_deficit"
].sum()

# Conversion rate
total_indepth = indepth_df[
    "student_name"
].nunique()

conversion_rate = (
    total_indepth / total_attended * 100
    if total_attended > 0 else 0
)

# ---------------------------------------------------
# KPI DISPLAY
# ---------------------------------------------------

st.markdown("---")

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Seminars", total_seminars)

c2.metric("Attendees", f"{total_attended:,}")

c3.metric("Indepth Students", total_indepth)

c4.metric("Conversion %", f"{conversion_rate:.1f}%")

c5.metric("Revenue", f"₹{total_revenue:,.0f}")

c6.metric("Profit", f"₹{profit:,.0f}")

# ---------------------------------------------------
# CHARTS
# ---------------------------------------------------

st.markdown("---")

tab1, tab2, tab3 = st.tabs([
    "Revenue vs Expense",
    "Conversion Funnel",
    "Location Performance"
])


# Revenue chart
with tab1:

    chart = filtered.groupby("location", as_index=False).agg(
        revenue=("actual_revenue(w/o_gst)_attendees", "sum"),
        expense=("actual_expenses", "sum")
    )

    fig = px.bar(
        chart,
        x="location",
        y=["revenue", "expense"],
        barmode="group"
    )

    st.plotly_chart(fig, use_container_width=True)


# Funnel chart
with tab2:

    funnel = pd.DataFrame({

        "Stage": [
            "Seminar Attended",
            "Indepth Joined"
        ],

        "Count": [
            total_attended,
            total_indepth
        ]
    })

    fig = px.funnel(
        funnel,
        x="Count",
        y="Stage"
    )

    st.plotly_chart(fig, use_container_width=True)


# Location performance
with tab3:

    chart = filtered.groupby(
        "location",
        as_index=False
    ).agg({

        "total_attended": "sum",
        "surplus_or_deficit": "sum"

    })

    fig = px.bar(
        chart,
        x="location",
        y="surplus_or_deficit",
        color="surplus_or_deficit"
    )

    st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------
# STUDENT TABLE
# ---------------------------------------------------

st.markdown("---")

st.subheader("Student Conversion Data")

st.dataframe(
    indepth_df,
    use_container_width=True,
    height=400
)
