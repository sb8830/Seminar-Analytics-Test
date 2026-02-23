# ============================================================
# FILE UPLOAD SECTION (Cloud Safe)
# ============================================================

st.sidebar.header("📂 Upload Data Files")

attendee_file = st.sidebar.file_uploader(
    "Upload Attendee File",
    type=["xlsx"]
)

report_file = st.sidebar.file_uploader(
    "Upload Seminar Report File",
    type=["xlsx"]
)

@st.cache_data
def load_data(attendee_file, report_file):

    attendee = pd.read_excel(attendee_file)
    report = pd.read_excel(report_file)

    # Fix header issue if exists
    report.columns = report.iloc[0]
    report = report.drop(0)

    attendee.columns = attendee.columns.str.lower().str.strip()
    report.columns = report.columns.str.lower().str.strip()

    merged = pd.merge(
        attendee,
        report,
        how="left",
        on="phone",
        suffixes=("_attendee","_report")
    )

    merged["student_id"] = (
        merged["phone"].astype(str)
        + "_"
        + merged["student_name"].astype(str)
    )

    return merged


if attendee_file and report_file:
    merged_df = load_data(attendee_file, report_file)
else:
    st.warning("Please upload both Excel files to continue.")
    st.stop()
