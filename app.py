import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta

st.set_page_config(page_title="Enterprise Seminar Dashboard", layout="wide")

st.title("🏢 Enterprise Seminar Analytics Dashboard")

# =========================================================
# Helpers
# =========================================================

def clean_numeric(val):

    if pd.isna(val):
        return 0

    if isinstance(val,(int,float)):
        return float(val)

    val=str(val).replace(",","").replace("₹","").replace("%","").strip()

    try:
        return float(val)
    except:
        return 0


def parse_pct(val):

    if pd.isna(val):
        return 0

    val=str(val).replace("%","")

    try:
        return float(val)
    except:
        return 0


# =========================================================
# Load Seminar Report
# =========================================================

@st.cache_data
def load_seminar(file):

    df=pd.read_excel(file,header=1)

    records=[]

    for _,row in df.iterrows():

        if pd.isna(row.iloc[0]):
            continue

        rec={

        "sr_no":int(clean_numeric(row.iloc[0])),

        "trainer":str(row.iloc[1]),

        "location":str(row.iloc[2]).upper(),

        "seminar_date":pd.to_datetime(row.iloc[3],errors="coerce"),

        "batch_date":pd.to_datetime(row.iloc[4],errors="coerce"),

        "targeted":clean_numeric(row.iloc[5]),

        "attended":clean_numeric(row.iloc[10]),

        "seat_booked":clean_numeric(row.iloc[15]),

        "expenses":clean_numeric(row.iloc[20]),

        "revenue":clean_numeric(row.iloc[22]),

        }

        rec["profit"]=rec["revenue"]-rec["expenses"]

        records.append(rec)

    return pd.DataFrame(records)


# =========================================================
# Load Conversion List
# =========================================================

@st.cache_data
def load_conversion(file):

    df=pd.read_excel(file)

    df.columns=df.columns.str.lower().str.replace(" ","_")

    if "batch_date" in df.columns:

        df["batch_date"]=pd.to_datetime(df["batch_date"],errors="coerce")

    if "payment_received" in df.columns:

        df["payment_received"]=df["payment_received"].apply(clean_numeric)

    if "total_amount" in df.columns:

        df["total_amount"]=df["total_amount"].apply(clean_numeric)

    return df


# =========================================================
# Match conversions
# =========================================================

def match_conversion(seminar,conv):

    matches=[]

    for _,c in conv.iterrows():

        if pd.isna(c.get("batch_date")):
            continue

        seminar["diff"]=(seminar["batch_date"]-c["batch_date"]).abs()

        closest=seminar[seminar["diff"]<=timedelta(days=7)]

        if len(closest)>0:

            best=closest.loc[closest["diff"].idxmin()]

            matches.append({

            "student":c.get("student_name",""),

            "phone":c.get("phone",""),

            "course":c.get("service_name",""),

            "revenue":c.get("total_amount",0),

            "received":c.get("payment_received",0),

            "location":best["location"],

            "trainer":best["trainer"]

            })

    return pd.DataFrame(matches)


# =========================================================
# Upload files
# =========================================================

col1,col2=st.columns(2)

with col1:

    seminar_file=st.file_uploader(
    "Upload Seminar Report",
    type=["xlsx"]
    )

with col2:

    conversion_file=st.file_uploader(
    "Upload Conversion List",
    type=["xlsx","xlsb"]
    )


if seminar_file is None:

    st.stop()


seminar_df=load_seminar(seminar_file)

# =========================================================
# Filters
# =========================================================

st.sidebar.header("Filters")

locations=st.sidebar.multiselect(
"Location",
seminar_df["location"].unique()
)

trainers=st.sidebar.multiselect(
"Trainer",
seminar_df["trainer"].unique()
)

filtered=seminar_df.copy()

if locations:
    filtered=filtered[filtered["location"].isin(locations)]

if trainers:
    filtered=filtered[filtered["trainer"].isin(trainers)]


# =========================================================
# KPIs
# =========================================================

total_seminars=len(filtered)

total_attended=filtered["attended"].sum()

total_revenue=filtered["revenue"].sum()

total_profit=filtered["profit"].sum()

k1,k2,k3,k4=st.columns(4)

k1.metric("Total Seminars",total_seminars)

k2.metric("Total Attended",f"{total_attended:,}")

k3.metric("Revenue",f"₹{total_revenue:,.0f}")

k4.metric("Profit",f"₹{total_profit:,.0f}")


# =========================================================
# Revenue vs Expense Chart
# =========================================================

fig=go.Figure()

fig.add_bar(
x=filtered["location"],
y=filtered["expenses"],
name="Expense"
)

fig.add_bar(
x=filtered["location"],
y=filtered["revenue"],
name="Revenue"
)

st.plotly_chart(fig,use_container_width=True)


# =========================================================
# Conversion analytics
# =========================================================

if conversion_file:

    conv_df=load_conversion(conversion_file)

    match_df=match_conversion(filtered,conv_df)

    if len(match_df)>0:

        st.header("Conversion Analytics")

        total_conv=len(match_df)

        conv_revenue=match_df["revenue"].sum()

        c1,c2=st.columns(2)

        c1.metric("Converted Students",total_conv)

        c2.metric("Conversion Revenue",f"₹{conv_revenue:,.0f}")


        # conversion by location

        loc_summary=match_df.groupby("location")["revenue"].sum().reset_index()

        fig=px.bar(
        loc_summary,
        x="location",
        y="revenue",
        title="Conversion Revenue by Location"
        )

        st.plotly_chart(fig,use_container_width=True)


        # course analysis

        course_summary=match_df.groupby("course")["revenue"].sum().reset_index()

        fig=px.pie(
        course_summary,
        names="course",
        values="revenue",
        title="Revenue by Course"
        )

        st.plotly_chart(fig,use_container_width=True)


        st.dataframe(match_df,use_container_width=True)

    else:

        st.warning("No conversions matched")


# =========================================================
# Seminar table
# =========================================================

st.header("Seminar Table")

st.dataframe(filtered,use_container_width=True)


# =========================================================
# Download
# =========================================================

st.download_button(
"Download Report",
filtered.to_csv(index=False),
"seminar_report.csv"
)
