import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="Seminar Analytics", page_icon="📊", layout="wide")

# ── Custom CSS ──
st.markdown("""
<style>
    .block-container { padding-top: 1rem; }
    .stMetric { background: #f8f9fa; padding: 15px; border-radius: 10px; border: 1px solid #e9ecef; }
    h1 { color: #1a56db; }
    .report-card { background: #f0f4ff; border-left: 4px solid #1a56db; padding: 12px 16px; border-radius: 6px; margin: 8px 0; }
    .insight-box { background: #fff8e1; border: 1px solid #ffc107; padding: 10px 14px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Seminar Analytics Dashboard")
st.caption("Offline Seminar Performance • 2025-26")

# ── File Upload ──
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("Upload **Indepth Details (Attendees)**", type=["xlsx", "xls"], key="f1")
with col2:
    file2 = st.file_uploader("Upload **Seminar Report (Based on Attendee)**", type=["xlsx", "xls"], key="f2")

if not file1 or not file2:
    st.info("👆 Please upload both Excel files to begin analysis.")
    st.stop()

# ── Load Data ──
@st.cache_data
def load_attendee_data(file):
    sheets = pd.read_excel(file, sheet_name=None, header=0)
    frames = []
    for name, df in sheets.items():
        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
        if any(c in df.columns for c in ['student_name', 'studentname']):
            frames.append(df)
    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame()

@st.cache_data
def load_seminar_data(file):
    df = pd.read_excel(file, sheet_name=0, header=1)
    df.columns = df.columns.str.strip()
    return df

attendee_df = load_attendee_data(file1)
seminar_raw = load_seminar_data(file2)

# ── Parse Seminar Data ──
@st.cache_data
def parse_seminar(df):
    col_map = {}
    cols_lower = {c: c.strip().lower().replace('\n', ' ').replace('\r', ' ') for c in df.columns}
    for orig, low in cols_lower.items():
        if 'sr' in low and 'no' in low:
            col_map['sr_no'] = orig
        elif low in ['trainer']:
            col_map['trainer'] = orig
        elif low in ['location']:
            col_map['location'] = orig
        elif 'seminar' in low and 'date' in low:
            col_map['seminar_date'] = orig
        elif 'targeted' == low:
            col_map['targeted'] = orig
        elif 'total' in low and 'attended' in low and 'actual' not in low:
            col_map['total_attended'] = orig
        elif 'actual' in low and 'attended' in low:
            col_map['actual_attended'] = orig
        elif 'targeted' in low and 'attended' in low and '%' in low:
            col_map['targeted_to_attended_pct'] = orig
        elif 'total' in low and 'seat' in low and 'booked' in low:
            col_map['total_seat_booked'] = orig
        elif 'actual' in low and 'expense' in low:
            col_map['actual_expenses'] = orig
        elif 'expected' in low and 'revenue' in low:
            col_map['expected_revenue'] = orig
        elif 'actual' in low and 'revenue' in low and 'total' not in low:
            col_map['actual_revenue'] = orig
        elif 'total' in low and 'revenue' in low:
            col_map['total_revenue'] = orig
        elif 'surplus' in low or 'deficit' in low:
            col_map['surplus_deficit'] = orig
        elif low in ['er to ae']:
            col_map['er_to_ae'] = orig
        elif low in ['ar to ae']:
            col_map['ar_to_ae'] = orig
        elif 'attended' in low and 'seat' in low and 'booked' in low and '%' in low:
            col_map['attended_to_booked_pct'] = orig
        elif 'morning' in low and 'total' in low:
            col_map['morning_total'] = orig
        elif 'evening' in low and 'total' in low:
            col_map['evening_total'] = orig

    renamed = df.rename(columns={v: k for k, v in col_map.items()})
    if 'sr_no' in renamed.columns:
        renamed = renamed[pd.to_numeric(renamed['sr_no'], errors='coerce').notna()]

    numeric_cols = ['targeted', 'total_attended', 'actual_attended', 'total_seat_booked',
                    'actual_expenses', 'expected_revenue', 'actual_revenue', 'total_revenue',
                    'surplus_deficit', 'er_to_ae', 'ar_to_ae', 'morning_total', 'evening_total']
    for c in numeric_cols:
        if c in renamed.columns:
            renamed[c] = pd.to_numeric(renamed[c].astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce').fillna(0)
    return renamed

seminar_df = parse_seminar(seminar_raw)

# ── Parse Attendee Data ──
attendee_df.columns = attendee_df.columns.str.strip().str.lower().str.replace(' ', '_')
for c in ['payment_received', 'total_gst', 'total_amount', 'total_due', 'total_additional_charges']:
    if c in attendee_df.columns:
        attendee_df[c] = pd.to_numeric(attendee_df[c], errors='coerce').fillna(0)

# ── Sidebar Filters ──
st.sidebar.header("🔍 Filters")
locations = sorted(seminar_df['location'].dropna().unique()) if 'location' in seminar_df.columns else []
selected_locations = st.sidebar.multiselect("📍 Location", locations, default=[])

if 'trainer' in seminar_df.columns:
    all_trainers = set()
    for t in seminar_df['trainer'].dropna():
        for name in str(t).split(','):
            name = name.strip().split('\n')[0].strip()
            if name:
                all_trainers.add(name)
    all_trainers = sorted(all_trainers)
else:
    all_trainers = []
selected_trainers = st.sidebar.multiselect("👨‍🏫 Trainer", all_trainers, default=[])
profit_filter = st.sidebar.radio("💰 Profitability", ["All", "Profitable", "Loss-making"], horizontal=True)

# Date range filter
if 'seminar_date' in seminar_df.columns:
    seminar_df['seminar_date'] = pd.to_datetime(seminar_df['seminar_date'], errors='coerce')
    valid_dates = seminar_df['seminar_date'].dropna()
    if len(valid_dates) > 0:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        date_range = st.sidebar.date_input("📅 Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    else:
        date_range = None
else:
    date_range = None

# ── Apply Filters ──
filtered = seminar_df.copy()
if selected_locations:
    filtered = filtered[filtered['location'].isin(selected_locations)]
if selected_trainers:
    def has_trainer(trainer_str):
        names = [n.strip().split('\n')[0].strip() for n in str(trainer_str).split(',')]
        return any(n in selected_trainers for n in names)
    filtered = filtered[filtered['trainer'].apply(has_trainer)]
if profit_filter == "Profitable" and 'surplus_deficit' in filtered.columns:
    filtered = filtered[filtered['surplus_deficit'] > 0]
elif profit_filter == "Loss-making" and 'surplus_deficit' in filtered.columns:
    filtered = filtered[filtered['surplus_deficit'] < 0]
if date_range and len(date_range) == 2 and 'seminar_date' in filtered.columns:
    filtered = filtered[
        (filtered['seminar_date'].dt.date >= date_range[0]) &
        (filtered['seminar_date'].dt.date <= date_range[1])
    ]

# ── KPI Calculations ──
total_seminars = len(filtered)
total_attended = int(filtered['total_attended'].sum()) if 'total_attended' in filtered.columns else 0
total_revenue = filtered['actual_revenue'].sum() if 'actual_revenue' in filtered.columns else 0
total_expenses = filtered['actual_expenses'].sum() if 'actual_expenses' in filtered.columns else 0
net_surplus = total_revenue - total_expenses
with_exp = filtered[filtered['actual_expenses'] > 0] if 'actual_expenses' in filtered.columns else filtered
profitable_count = int((with_exp['surplus_deficit'] > 0).sum()) if 'surplus_deficit' in with_exp.columns else 0
avg_conversion = 0
if 'attended_to_booked_pct' in filtered.columns:
    avg_conversion = filtered['attended_to_booked_pct'].mean()
elif 'total_seat_booked' in filtered.columns and 'total_attended' in filtered.columns:
    total_att = filtered['total_attended'].sum()
    total_booked = filtered['total_seat_booked'].sum()
    avg_conversion = (total_booked / total_att * 100) if total_att > 0 else 0

# ── KPI Section ──
st.markdown("---")
k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
k1.metric("📋 Seminars", total_seminars)
k2.metric("👥 Attendees", f"{total_attended:,}")
k3.metric("💰 Revenue", f"₹{total_revenue/100000:.1f}L")
k4.metric("📤 Expenses", f"₹{total_expenses/100000:.1f}L")
k5.metric("📈 Net Surplus", f"₹{net_surplus/100000:.1f}L", delta=f"{'▲' if net_surplus >= 0 else '▼'}")
k6.metric("🎯 Avg Conversion", f"{avg_conversion:.1f}%")
k7.metric("✅ Profitable", f"{profitable_count}/{len(with_exp)}")

# ── Trainer Performance Helper ──
def build_trainer_summary(df):
    if 'trainer' not in df.columns:
        return pd.DataFrame()
    rows = []
    for _, row in df.iterrows():
        trainers = [t.strip().split('\n')[0].strip() for t in str(row.get('trainer', '')).split(',') if t.strip()]
        for t in trainers:
            if t:
                rows.append({
                    'trainer': t,
                    'actual_revenue': row.get('actual_revenue', 0),
                    'actual_expenses': row.get('actual_expenses', 0),
                    'surplus_deficit': row.get('surplus_deficit', 0),
                    'total_attended': row.get('total_attended', 0),
                    'total_seat_booked': row.get('total_seat_booked', 0),
                    'seminars': 1
                })
    if not rows:
        return pd.DataFrame()
    tdf = pd.DataFrame(rows)
    return tdf.groupby('trainer').agg(
        seminars=('seminars', 'sum'),
        total_attended=('total_attended', 'sum'),
        total_seat_booked=('total_seat_booked', 'sum'),
        actual_revenue=('actual_revenue', 'sum'),
        actual_expenses=('actual_expenses', 'sum'),
        surplus_deficit=('surplus_deficit', 'sum')
    ).reset_index().sort_values('actual_revenue', ascending=False)

trainer_summary = build_trainer_summary(filtered)

# ── Location Summary Helper ──
def build_location_summary(df):
    if 'location' not in df.columns:
        return pd.DataFrame()
    grp = df.groupby('location').agg(
        seminars=('sr_no', 'count') if 'sr_no' in df.columns else ('location', 'count'),
        total_attended=('total_attended', 'sum'),
        total_seat_booked=('total_seat_booked', 'sum'),
        actual_revenue=('actual_revenue', 'sum'),
        actual_expenses=('actual_expenses', 'sum'),
        surplus_deficit=('surplus_deficit', 'sum')
    ).reset_index().sort_values('actual_revenue', ascending=False)
    grp['roi_pct'] = grp.apply(lambda r: (r['surplus_deficit'] / r['actual_expenses'] * 100) if r['actual_expenses'] > 0 else 0, axis=1)
    grp['conversion_pct'] = grp.apply(lambda r: (r['total_seat_booked'] / r['total_attended'] * 100) if r['total_attended'] > 0 else 0, axis=1)
    return grp

location_summary = build_location_summary(filtered)

# ── Excel Report Generator ──
def generate_excel_report(filtered_df, attendee_df, trainer_sum, location_sum):
    wb = openpyxl.Workbook()

    # Styles
    hdr_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', start_color='1A56DB')
    sub_hdr_fill = PatternFill('solid', start_color='D6E4FF')
    sub_hdr_font = Font(name='Arial', bold=True, color='1A56DB', size=10)
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    title_font = Font(name='Arial', bold=True, size=14, color='1A56DB')
    pos_fill = PatternFill('solid', start_color='D1FAE5')
    neg_fill = PatternFill('solid', start_color='FEE2E2')
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header_row(ws, row_num, col_count):
        for col in range(1, col_count + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = center
            cell.border = border

    def style_data_row(ws, row_num, col_count, alt=False):
        alt_fill = PatternFill('solid', start_color='F8FAFF')
        for col in range(1, col_count + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.font = Font(name='Arial', size=10)
            cell.alignment = left
            cell.border = border
            if alt:
                cell.fill = alt_fill

    def add_title_block(ws, title, subtitle=''):
        ws.merge_cells('A1:H1')
        ws['A1'] = title
        ws['A1'].font = title_font
        ws['A1'].alignment = center
        if subtitle:
            ws.merge_cells('A2:H2')
            ws['A2'] = subtitle
            ws['A2'].font = Font(name='Arial', italic=True, size=10, color='555555')
            ws['A2'].alignment = center

    # ── Sheet 1: Executive Summary ──
    ws1 = wb.active
    ws1.title = "Executive Summary"
    add_title_block(ws1, "📊 SEMINAR ANALYTICS — EXECUTIVE SUMMARY", f"Generated: {datetime.now().strftime('%d %b %Y %I:%M %p')}")
    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 18

    ws1['A4'] = "KEY PERFORMANCE INDICATORS"
    ws1['A4'].font = sub_hdr_font
    ws1['A4'].fill = sub_hdr_fill

    kpi_data = [
        ("Total Seminars Conducted", total_seminars, ""),
        ("Total Attendees", total_attended, ""),
        ("Total Actual Revenue", f"₹{total_revenue:,.0f}", ""),
        ("Total Actual Expenses", f"₹{total_expenses:,.0f}", ""),
        ("Net Surplus / Deficit", f"₹{net_surplus:,.0f}", "▲ Profit" if net_surplus >= 0 else "▼ Loss"),
        ("Avg Seat Conversion %", f"{avg_conversion:.1f}%", ""),
        ("Profitable Seminars", f"{profitable_count} / {len(with_exp)}", f"{(profitable_count/len(with_exp)*100):.0f}%" if len(with_exp) > 0 else ""),
    ]
    ws1['A5'] = "Metric"
    ws1['B5'] = "Value"
    ws1['C5'] = "Note"
    style_header_row(ws1, 5, 3)
    for i, (metric, value, note) in enumerate(kpi_data, start=6):
        ws1[f'A{i}'] = metric
        ws1[f'B{i}'] = value
        ws1[f'C{i}'] = note
        style_data_row(ws1, i, 3, alt=(i % 2 == 0))

    ws1.column_dimensions['A'].width = 35
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 18

    # Top/Bottom performers
    if len(with_exp) > 0 and 'surplus_deficit' in with_exp.columns:
        r = 15
        ws1[f'A{r}'] = "TOP PERFORMERS"
        ws1[f'A{r}'].font = sub_hdr_font
        ws1[f'A{r}'].fill = sub_hdr_fill
        ws1.merge_cells(f'A{r}:C{r}')
        r += 1
        ws1[f'A{r}'] = "Location"
        ws1[f'B{r}'] = "Surplus/Deficit"
        ws1[f'C{r}'] = "Status"
        style_header_row(ws1, r, 3)
        r += 1
        top3 = with_exp.nlargest(3, 'surplus_deficit')
        for _, row in top3.iterrows():
            ws1[f'A{r}'] = row.get('location', 'N/A')
            ws1[f'B{r}'] = f"₹{int(row['surplus_deficit']):,}"
            ws1[f'C{r}'] = "✅ Profitable"
            ws1[f'A{r}'].fill = pos_fill
            ws1[f'B{r}'].fill = pos_fill
            ws1[f'C{r}'].fill = pos_fill
            for col in range(1, 4):
                ws1.cell(row=r, column=col).border = border
            r += 1
        r += 1
        ws1[f'A{r}'] = "NEEDS ATTENTION"
        ws1[f'A{r}'].font = sub_hdr_font
        ws1[f'A{r}'].fill = PatternFill('solid', start_color='FFE4E4')
        ws1.merge_cells(f'A{r}:C{r}')
        r += 1
        ws1[f'A{r}'] = "Location"
        ws1[f'B{r}'] = "Surplus/Deficit"
        ws1[f'C{r}'] = "Status"
        style_header_row(ws1, r, 3)
        r += 1
        bot3 = with_exp.nsmallest(3, 'surplus_deficit')
        for _, row in bot3.iterrows():
            ws1[f'A{r}'] = row.get('location', 'N/A')
            ws1[f'B{r}'] = f"₹{int(row['surplus_deficit']):,}"
            ws1[f'C{r}'] = "⚠️ Loss" if row['surplus_deficit'] < 0 else "Break-even"
            for col in range(1, 4):
                ws1.cell(row=r, column=col).fill = neg_fill if row['surplus_deficit'] < 0 else PatternFill('solid', start_color='FFF9E6')
                ws1.cell(row=r, column=col).border = border
            r += 1

    # ── Sheet 2: Financial Report ──
    ws2 = wb.create_sheet("Financial Report")
    add_title_block(ws2, "💰 FINANCIAL PERFORMANCE REPORT")
    ws2.row_dimensions[1].height = 28

    fin_cols = ['sr_no', 'location', 'trainer', 'seminar_date', 'actual_expenses',
                'expected_revenue', 'actual_revenue', 'surplus_deficit', 'er_to_ae', 'ar_to_ae']
    fin_cols = [c for c in fin_cols if c in filtered_df.columns]
    fin_headers = {
        'sr_no': 'Sr No', 'location': 'Location', 'trainer': 'Trainer',
        'seminar_date': 'Date', 'actual_expenses': 'Expenses (₹)',
        'expected_revenue': 'Expected Rev (₹)', 'actual_revenue': 'Actual Rev (₹)',
        'surplus_deficit': 'Surplus/Deficit (₹)', 'er_to_ae': 'ER:AE', 'ar_to_ae': 'AR:AE'
    }

    for ci, col in enumerate(fin_cols, 1):
        ws2.cell(row=4, column=ci).value = fin_headers.get(col, col.replace('_', ' ').title()
)
    style_header_row(ws2, 4, len(fin_cols))

    for ri, (_, row) in enumerate(filtered_df[fin_cols].iterrows(), start=5):
        for ci, col in enumerate(fin_cols, 1):
            val = row[col]
            if pd.isna(val):
                val = ''
            ws2.cell(row=ri, column=ci).value = val
        style_data_row(ws2, ri, len(fin_cols), alt=(ri % 2 == 0))
        # Color surplus/deficit cell
        if 'surplus_deficit' in fin_cols:
            sd_idx = fin_cols.index('surplus_deficit') + 1
            sd_val = row.get('surplus_deficit', 0)
            if sd_val > 0:
                ws2.cell(row=ri, column=sd_idx).fill = pos_fill
            elif sd_val < 0:
                ws2.cell(row=ri, column=sd_idx).fill = neg_fill

    # Totals row
    total_row = ri + 2
    ws2.cell(row=total_row, column=1).value = "TOTAL"
    ws2.cell(row=total_row, column=1).font = Font(name='Arial', bold=True)
    for ci, col in enumerate(fin_cols, 1):
        if col in ['actual_expenses', 'expected_revenue', 'actual_revenue', 'surplus_deficit']:
            ws2.cell(row=total_row, column=ci).value = f"=SUM({get_column_letter(ci)}5:{get_column_letter(ci)}{ri})"
            ws2.cell(row=total_row, column=ci).font = Font(name='Arial', bold=True)
            ws2.cell(row=total_row, column=ci).fill = PatternFill('solid', start_color='DBEAFE')

    for ci in range(1, len(fin_cols)+1):
        ws2.column_dimensions[get_column_letter(ci)].width = 18
    ws2.column_dimensions['B'].width = 22
    ws2.column_dimensions['C'].width = 24

    # ── Sheet 3: Attendance & Conversion ──
    ws3 = wb.create_sheet("Attendance & Conversion")
    add_title_block(ws3, "🎯 ATTENDANCE & CONVERSION REPORT")
    ws3.row_dimensions[1].height = 28

    att_cols = ['sr_no', 'location', 'seminar_date', 'targeted', 'total_attended',
                'actual_attended', 'total_seat_booked', 'targeted_to_attended_pct', 'morning_total', 'evening_total']
    att_cols = [c for c in att_cols if c in filtered_df.columns]
    att_headers = {
        'sr_no': 'Sr No', 'location': 'Location', 'seminar_date': 'Date',
        'targeted': 'Targeted', 'total_attended': 'Total Attended',
        'actual_attended': 'Actual Attended', 'total_seat_booked': 'Seats Booked',
        'targeted_to_attended_pct': 'Target→Attend %', 'morning_total': 'Morning',
        'evening_total': 'Evening'
    }
    for ci, col in enumerate(att_cols, 1):
        ws3.cell(row=4, column=ci).value = att_headers.get(col, col.replace('_', ' ').title())
    style_header_row(ws3, 4, len(att_cols))

    # Add computed conversion column
    extra_col = len(att_cols) + 1
    ws3.cell(row=4, column=extra_col).value = "Conversion %"
    ws3.cell(row=4, column=extra_col).font = hdr_font
    ws3.cell(row=4, column=extra_col).fill = hdr_fill
    ws3.cell(row=4, column=extra_col).alignment = center
    ws3.cell(row=4, column=extra_col).border = border

    for ri, (_, row) in enumerate(filtered_df[att_cols].iterrows(), start=5):
        for ci, col in enumerate(att_cols, 1):
            val = row[col]
            if pd.isna(val): val = ''
            ws3.cell(row=ri, column=ci).value = val
        style_data_row(ws3, ri, extra_col, alt=(ri % 2 == 0))
        # Conversion formula
        if 'total_seat_booked' in att_cols and 'total_attended' in att_cols:
            b_col = get_column_letter(att_cols.index('total_seat_booked') + 1)
            a_col = get_column_letter(att_cols.index('total_attended') + 1)
            ws3.cell(row=ri, column=extra_col).value = f"=IF({a_col}{ri}>0,{b_col}{ri}/{a_col}{ri}*100,0)"

    for ci in range(1, extra_col + 1):
        ws3.column_dimensions[get_column_letter(ci)].width = 16
    ws3.column_dimensions['B'].width = 22

    # ── Sheet 4: Trainer Performance ──
    ws4 = wb.create_sheet("Trainer Performance")
    add_title_block(ws4, "👨‍🏫 TRAINER PERFORMANCE REPORT")
    ws4.row_dimensions[1].height = 28

    if not trainer_sum.empty:
        t_headers = ['Trainer', 'Seminars', 'Total Attended', 'Seats Booked',
                     'Revenue (₹)', 'Expenses (₹)', 'Surplus/Deficit (₹)', 'Avg Rev/Seminar', 'Conversion %']
        for ci, h in enumerate(t_headers, 1):
            ws4.cell(row=4, column=ci).value = h
        style_header_row(ws4, 4, len(t_headers))

        for ri, (_, row) in enumerate(trainer_sum.iterrows(), start=5):
            ws4.cell(row=ri, column=1).value = row['trainer']
            ws4.cell(row=ri, column=2).value = int(row['seminars'])
            ws4.cell(row=ri, column=3).value = int(row['total_attended'])
            ws4.cell(row=ri, column=4).value = int(row['total_seat_booked'])
            ws4.cell(row=ri, column=5).value = round(row['actual_revenue'], 0)
            ws4.cell(row=ri, column=6).value = round(row['actual_expenses'], 0)
            ws4.cell(row=ri, column=7).value = round(row['surplus_deficit'], 0)
            ws4.cell(row=ri, column=8).value = f"=IF(B{ri}>0,E{ri}/B{ri},0)"
            ws4.cell(row=ri, column=9).value = f"=IF(C{ri}>0,D{ri}/C{ri}*100,0)"
            style_data_row(ws4, ri, len(t_headers), alt=(ri % 2 == 0))
            fill = pos_fill if row['surplus_deficit'] >= 0 else neg_fill
            ws4.cell(row=ri, column=7).fill = fill

        col_widths = [28, 12, 16, 14, 18, 16, 22, 20, 16]
        for ci, w in enumerate(col_widths, 1):
            ws4.column_dimensions[get_column_letter(ci)].width = w

    # ── Sheet 5: Location Summary ──
    ws5 = wb.create_sheet("Location Summary")
    add_title_block(ws5, "📍 LOCATION-WISE SUMMARY REPORT")
    ws5.row_dimensions[1].height = 28

    if not location_sum.empty:
        l_headers = ['Location', 'Seminars', 'Total Attended', 'Seats Booked',
                     'Revenue (₹)', 'Expenses (₹)', 'Surplus/Deficit (₹)', 'ROI %', 'Conversion %']
        for ci, h in enumerate(l_headers, 1):
            ws5.cell(row=4, column=ci).value = h
        style_header_row(ws5, 4, len(l_headers))

        for ri, (_, row) in enumerate(location_sum.iterrows(), start=5):
            ws5.cell(row=ri, column=1).value = row['location']
            ws5.cell(row=ri, column=2).value = int(row['seminars'])
            ws5.cell(row=ri, column=3).value = int(row['total_attended'])
            ws5.cell(row=ri, column=4).value = int(row['total_seat_booked'])
            ws5.cell(row=ri, column=5).value = round(row['actual_revenue'], 0)
            ws5.cell(row=ri, column=6).value = round(row['actual_expenses'], 0)
            ws5.cell(row=ri, column=7).value = round(row['surplus_deficit'], 0)
            ws5.cell(row=ri, column=8).value = round(row['roi_pct'], 1)
            ws5.cell(row=ri, column=9).value = round(row['conversion_pct'], 1)
            style_data_row(ws5, ri, len(l_headers), alt=(ri % 2 == 0))
            fill = pos_fill if row['surplus_deficit'] >= 0 else neg_fill
            ws5.cell(row=ri, column=7).fill = fill
            ws5.cell(row=ri, column=8).fill = fill

        for ci in range(1, 10):
            ws5.column_dimensions[get_column_letter(ci)].width = 20
        ws5.column_dimensions['A'].width = 28

    # ── Sheet 6: Student / Attendee Details ──
    ws6 = wb.create_sheet("Attendee Details")
    add_title_block(ws6, "👤 ATTENDEE DETAILS REPORT")
    ws6.row_dimensions[1].height = 28

    att_show_cols = [c for c in ['student_name', 'phone', 'email', 'service_name', 'batch_date',
                                  'payment_received', 'total_gst', 'total_amount', 'total_due', 'status']
                     if c in attendee_df.columns]
    att_show_headers = {
        'student_name': 'Student Name', 'phone': 'Phone', 'email': 'Email',
        'service_name': 'Service/Course', 'batch_date': 'Batch Date',
        'payment_received': 'Payment Received (₹)', 'total_gst': 'GST (₹)',
        'total_amount': 'Total Amount (₹)', 'total_due': 'Total Due (₹)', 'status': 'Status'
    }
    for ci, col in enumerate(att_show_cols, 1):
        ws6.cell(row=4, column=ci).value = att_show_headers.get(col, col.replace('_', ' ').title())
    style_header_row(ws6, 4, len(att_show_cols))

    for ri, (_, row) in enumerate(attendee_df[att_show_cols].iterrows(), start=5):
        for ci, col in enumerate(att_show_cols, 1):
            val = row[col]
            if pd.isna(val): val = ''
            ws6.cell(row=ri, column=ci).value = val
            ws6.cell(row=ri, column=ci).font = Font(name='Arial', size=9)
            ws6.cell(row=ri, column=ci).border = border
            if (ri % 2 == 0):
                ws6.cell(row=ri, column=ci).fill = PatternFill('solid', start_color='F8FAFF')

    for ci in range(1, len(att_show_cols) + 1):
        ws6.column_dimensions[get_column_letter(ci)].width = 20
    ws6.column_dimensions['A'].width = 28
    ws6.column_dimensions['C'].width = 30

    # ── Sheet 7: Course Revenue Breakdown ──
    ws7 = wb.create_sheet("Course Revenue")
    add_title_block(ws7, "📚 COURSE-WISE REVENUE BREAKDOWN")
    ws7.row_dimensions[1].height = 28

    if 'service_name' in attendee_df.columns and 'payment_received' in attendee_df.columns:
        course_data = attendee_df.groupby('service_name').agg(
            enrollments=('student_name', 'count') if 'student_name' in attendee_df.columns else ('service_name', 'count'),
            revenue=('payment_received', 'sum'),
            gst=('total_gst', 'sum') if 'total_gst' in attendee_df.columns else ('payment_received', 'count'),
            total_due=('total_due', 'sum') if 'total_due' in attendee_df.columns else ('payment_received', 'count')
        ).reset_index().sort_values('revenue', ascending=False)

        c_headers = ['Course / Service', 'Enrollments', 'Revenue Collected (₹)', 'GST (₹)', 'Total Due (₹)', 'Avg Revenue/Student']
        for ci, h in enumerate(c_headers, 1):
            ws7.cell(row=4, column=ci).value = h
        style_header_row(ws7, 4, len(c_headers))

        for ri, (_, row) in enumerate(course_data.iterrows(), start=5):
            ws7.cell(row=ri, column=1).value = row['service_name']
            ws7.cell(row=ri, column=2).value = int(row['enrollments'])
            ws7.cell(row=ri, column=3).value = round(row['revenue'], 0)
            ws7.cell(row=ri, column=4).value = round(row.get('gst', 0), 0)
            ws7.cell(row=ri, column=5).value = round(row.get('total_due', 0), 0)
            ws7.cell(row=ri, column=6).value = f"=IF(B{ri}>0,C{ri}/B{ri},0)"
            style_data_row(ws7, ri, len(c_headers), alt=(ri % 2 == 0))

        # Grand total
        tr = ri + 2
        ws7.cell(row=tr, column=1).value = "GRAND TOTAL"
        ws7.cell(row=tr, column=1).font = Font(name='Arial', bold=True, size=11)
        ws7.cell(row=tr, column=2).value = f"=SUM(B5:B{ri})"
        ws7.cell(row=tr, column=3).value = f"=SUM(C5:C{ri})"
        ws7.cell(row=tr, column=4).value = f"=SUM(D5:D{ri})"
        ws7.cell(row=tr, column=5).value = f"=SUM(E5:E{ri})"
        for ci in range(1, 7):
            ws7.cell(row=tr, column=ci).font = Font(name='Arial', bold=True)
            ws7.cell(row=tr, column=ci).fill = PatternFill('solid', start_color='DBEAFE')
            ws7.cell(row=tr, column=ci).border = border

        col_widths_7 = [35, 16, 24, 16, 18, 22]
        for ci, w in enumerate(col_widths_7, 1):
            ws7.column_dimensions[get_column_letter(ci)].width = w

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ── Main Dashboard Tabs ──
st.markdown("---")
main_tab1, main_tab2, main_tab3, main_tab4, main_tab5, main_tab6 = st.tabs([
    "📊 Charts & Analytics", "📍 Location Report", "👨‍🏫 Trainer Report",
    "📋 Data Tables", "💡 Insights", "📥 Download Reports"
])

# ── Tab 1: Charts ──
with main_tab1:
    chart_tab1, chart_tab2, chart_tab3, chart_tab4, chart_tab5 = st.tabs([
        "Revenue vs Expense", "Surplus/Deficit", "Attendance Funnel", "Student Status", "Revenue Breakdown"
    ])

    with chart_tab1:
        if 'actual_revenue' in filtered.columns and 'actual_expenses' in filtered.columns:
            chart_data = filtered[filtered['actual_expenses'] > 0][['location', 'actual_revenue', 'actual_expenses']].copy()
            fig = go.Figure()
            fig.add_trace(go.Bar(name='Revenue', x=chart_data['location'], y=chart_data['actual_revenue'], marker_color='#1a56db'))
            fig.add_trace(go.Bar(name='Expense', x=chart_data['location'], y=chart_data['actual_expenses'], marker_color='#f59e0b'))
            fig.update_layout(barmode='group', height=450, title="Revenue vs Expenses by Location",
                              xaxis_tickangle=-45, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

    with chart_tab2:
        if 'surplus_deficit' in filtered.columns:
            chart_data = filtered[filtered['actual_expenses'] > 0][['location', 'surplus_deficit']].copy()
            colors = ['#10b981' if v >= 0 else '#ef4444' for v in chart_data['surplus_deficit']]
            fig = px.bar(chart_data, x='location', y='surplus_deficit', title="Surplus / Deficit by Location")
            fig.update_traces(marker_color=colors)
            fig.add_hline(y=0, line_dash="dash", line_color="gray")
            fig.update_layout(height=450, xaxis_tickangle=-45,
                              plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

    with chart_tab3:
        if all(c in filtered.columns for c in ['targeted', 'total_attended', 'total_seat_booked']):
            fig = go.Figure()
            fig.add_trace(go.Bar(name='Targeted', x=filtered['location'], y=filtered['targeted'], marker_color='#9ca3af', opacity=0.5))
            fig.add_trace(go.Bar(name='Attended', x=filtered['location'], y=filtered['total_attended'], marker_color='#1a56db'))
            fig.add_trace(go.Bar(name='Booked', x=filtered['location'], y=filtered['total_seat_booked'], marker_color='#10b981'))
            fig.update_layout(barmode='group', height=450, title="Attendance Funnel by Location",
                              xaxis_tickangle=-45, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

    with chart_tab4:
        if 'status' in attendee_df.columns:
            status_counts = attendee_df['status'].value_counts()
            fig = px.pie(values=status_counts.values, names=status_counts.index, hole=0.45,
                         title="Student Status Distribution",
                         color_discrete_sequence=['#10b981', '#ef4444', '#f59e0b', '#6366f1'])
            fig.update_layout(height=420)
            st.plotly_chart(fig, use_container_width=True)

    with chart_tab5:
        if 'service_name' in attendee_df.columns:
            course_rev = attendee_df.groupby('service_name')['payment_received'].sum().sort_values(ascending=False).head(10)
            fig = px.bar(x=course_rev.values, y=course_rev.index, orientation='h',
                         title="Top 10 Courses by Revenue", color_discrete_sequence=['#1a56db'])
            fig.update_layout(height=450, yaxis_title='', xaxis_title='Revenue (₹)',
                              plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

# ── Tab 2: Location Report ──
with main_tab2:
    st.subheader("📍 Location-wise Performance Summary")
    if not location_summary.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(location_summary.head(10), x='location', y=['actual_revenue', 'actual_expenses'],
                         barmode='group', title="Revenue vs Expenses by Location",
                         color_discrete_sequence=['#1a56db', '#f59e0b'])
            fig.update_layout(height=380, xaxis_tickangle=-30,
                              plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.scatter(location_summary, x='total_attended', y='actual_revenue',
                             size='surplus_deficit', color='roi_pct',
                             hover_name='location', title="Attendance vs Revenue Bubble Chart",
                             color_continuous_scale='RdYlGn', size_max=40)
            fig.update_layout(height=380, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            location_summary.style.format({
                'actual_revenue': '₹{:,.0f}', 'actual_expenses': '₹{:,.0f}',
                'surplus_deficit': '₹{:,.0f}', 'roi_pct': '{:.1f}%', 'conversion_pct': '{:.1f}%'
            }).background_gradient(subset=['surplus_deficit'], cmap='RdYlGn'),
            use_container_width=True, height=350
        )
    else:
        st.warning("Location data not available.")

# ── Tab 3: Trainer Report ──
with main_tab3:
    st.subheader("👨‍🏫 Trainer Performance Summary")
    if not trainer_summary.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(trainer_summary.head(10), x='trainer', y='actual_revenue',
                         title="Revenue by Trainer", color='surplus_deficit',
                         color_continuous_scale='RdYlGn')
            fig.update_layout(height=380, xaxis_tickangle=-30,
                              plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(trainer_summary.head(10), x='trainer', y='seminars',
                         title="Seminars Conducted per Trainer", color_discrete_sequence=['#6366f1'])
            fig.update_layout(height=380, xaxis_tickangle=-30,
                              plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)

        trainer_display = trainer_summary.copy()
        trainer_display['avg_rev_per_seminar'] = (trainer_display['actual_revenue'] / trainer_display['seminars']).round(0)
        trainer_display['conversion_pct'] = (trainer_display['total_seat_booked'] / trainer_display['total_attended'].replace(0, 1) * 100).round(1)

        st.dataframe(
            trainer_display.style.format({
                'actual_revenue': '₹{:,.0f}', 'actual_expenses': '₹{:,.0f}',
                'surplus_deficit': '₹{:,.0f}', 'avg_rev_per_seminar': '₹{:,.0f}',
                'conversion_pct': '{:.1f}%'
            }).background_gradient(subset=['surplus_deficit'], cmap='RdYlGn'),
            use_container_width=True, height=350
        )
    else:
        st.warning("Trainer data not available.")

# ── Tab 4: Data Tables ──
with main_tab4:
    st.subheader("📋 Seminar Performance Data")
    display_cols = [c for c in ['sr_no', 'location', 'trainer', 'seminar_date', 'targeted',
                                 'total_attended', 'total_seat_booked', 'actual_expenses',
                                 'actual_revenue', 'surplus_deficit', 'er_to_ae'] if c in filtered.columns]
    st.dataframe(
        filtered[display_cols].reset_index(drop=True),
        use_container_width=True, height=380,
        column_config={
            "actual_expenses": st.column_config.NumberColumn("Expenses", format="₹%d"),
            "actual_revenue": st.column_config.NumberColumn("Revenue", format="₹%d"),
            "surplus_deficit": st.column_config.NumberColumn("Surplus/Deficit", format="₹%d"),
        }
    )

    st.markdown("---")
    st.subheader("👤 Attendee Details")
    att_display = [c for c in ['student_name', 'phone', 'email', 'service_name', 'batch_date',
                                'payment_received', 'total_gst', 'status', 'total_amount', 'total_due']
                   if c in attendee_df.columns]
    if 'status' in attendee_df.columns:
        status_filter = st.multiselect("Filter by Status", attendee_df['status'].dropna().unique().tolist())
        display_att = attendee_df[attendee_df['status'].isin(status_filter)] if status_filter else attendee_df
    else:
        display_att = attendee_df
    st.dataframe(
        display_att[att_display].head(500).reset_index(drop=True),
        use_container_width=True, height=380,
        column_config={
            "payment_received": st.column_config.NumberColumn("Payment", format="₹%d"),
            "total_amount": st.column_config.NumberColumn("Total Amt", format="₹%d"),
        }
    )
    st.caption(f"Showing {min(500, len(display_att))} of {len(display_att)} records")

# ── Tab 5: Insights ──
with main_tab5:
    st.subheader("💡 Automated Insights & Recommendations")

    if len(with_exp) > 0 and 'surplus_deficit' in with_exp.columns:
        best = with_exp.loc[with_exp['surplus_deficit'].idxmax()]
        worst = with_exp.loc[with_exp['surplus_deficit'].idxmin()]
        loss_count = int((with_exp['surplus_deficit'] < 0).sum())
        profit_rate = profitable_count / len(with_exp) * 100 if len(with_exp) > 0 else 0

        col_a, col_b = st.columns(2)
        with col_a:
            st.success(f"**🏆 Best ROI:** {best.get('location', 'N/A')} — Surplus ₹{int(best.get('surplus_deficit', 0)):,}")
            if 'total_attended' in filtered.columns and len(filtered) > 0:
                top_att = filtered.loc[filtered['total_attended'].idxmax()]
                st.info(f"**👥 Highest Attendance:** {top_att.get('location', 'N/A')} — {int(top_att['total_attended']):,} attendees")
            if not trainer_summary.empty:
                top_trainer = trainer_summary.iloc[0]
                st.info(f"**🌟 Top Trainer by Revenue:** {top_trainer['trainer']} — ₹{int(top_trainer['actual_revenue']):,}")
        with col_b:
            st.error(f"**⚠️ Worst ROI:** {worst.get('location', 'N/A')} — Deficit ₹{abs(int(worst.get('surplus_deficit', 0))):,}")
            st.warning(f"**📉 Loss-making seminars:** {loss_count} out of {len(with_exp)} ran at a loss")
            if profit_rate >= 75:
                st.success(f"**✅ Profitability Rate: {profit_rate:.0f}%** — Excellent performance!")
            elif profit_rate >= 50:
                st.warning(f"**⚠️ Profitability Rate: {profit_rate:.0f}%** — Room for improvement.")
            else:
                st.error(f"**❌ Profitability Rate: {profit_rate:.0f}%** — Immediate review needed.")

        st.markdown("---")
        st.markdown("#### 📌 Recommendations")
        recs = []
        if loss_count > 0:
            recs.append(f"🔴 **{loss_count} seminars ran at a loss.** Review cost structure and pricing for these locations.")
        if avg_conversion < 20:
            recs.append(f"🟡 **Low conversion rate ({avg_conversion:.1f}%).** Consider improving follow-up process and seminar content.")
        elif avg_conversion >= 35:
            recs.append(f"🟢 **Strong conversion rate ({avg_conversion:.1f}%).** Replicate this model across underperforming locations.")
        if not location_summary.empty and len(location_summary) > 1:
            top_loc = location_summary.iloc[0]['location']
            recs.append(f"🏆 **{top_loc}** is your highest revenue location. Increase seminar frequency here.")
        if not trainer_summary.empty and len(trainer_summary) > 1:
            top_t = trainer_summary.iloc[0]['trainer']
            recs.append(f"⭐ **Trainer {top_t}** generates the most revenue. Assign to high-potential locations.")
        if total_revenue > 0:
            expense_ratio = total_expenses / total_revenue * 100
            if expense_ratio > 60:
                recs.append(f"⚠️ **Expense ratio is {expense_ratio:.1f}% of revenue.** Target below 50% for healthy margins.")
            else:
                recs.append(f"✅ **Expense ratio is {expense_ratio:.1f}%.** Within acceptable range.")
        for r in recs:
            st.markdown(f"> {r}")

# ── Tab 6: Download Reports ──
with main_tab6:
    st.subheader("📥 Download Comprehensive Reports")
    st.markdown("""
    The report package includes **7 professional sheets**:

    | Sheet | Contents |
    |---|---|
    | 📊 Executive Summary | KPIs, Top/Bottom performers, financial snapshot |
    | 💰 Financial Report | Full revenue, expense, surplus/deficit per seminar |
    | 🎯 Attendance & Conversion | Targeted vs attended vs booked with conversion % |
    | 👨‍🏫 Trainer Performance | Revenue, attendance, ROI per trainer |
    | 📍 Location Summary | Location-wise aggregated performance & ROI |
    | 👤 Attendee Details | Complete student/attendee data |
    | 📚 Course Revenue | Course-wise enrollment & revenue breakdown |
    """)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("🔄 Generate Full Report", use_container_width=True, type="primary"):
            with st.spinner("Building report... please wait"):
                excel_data = generate_excel_report(filtered, attendee_df, trainer_summary, location_summary)
            st.success("✅ Report ready! Click below to download.")
            st.download_button(
                label="📥 Download Excel Report (.xlsx)",
                data=excel_data,
                file_name=f"Seminar_Analytics_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("---")
    st.markdown("#### 📊 Quick CSV Downloads")
    csv_col1, csv_col2, csv_col3 = st.columns(3)
    with csv_col1:
        if not filtered.empty:
            st.download_button("📋 Seminar Data (CSV)", filtered.to_csv(index=False).encode(),
                               "seminar_data.csv", "text/csv", use_container_width=True)
    with csv_col2:
        if not trainer_summary.empty:
            st.download_button("👨‍🏫 Trainer Report (CSV)", trainer_summary.to_csv(index=False).encode(),
                               "trainer_report.csv", "text/csv", use_container_width=True)
    with csv_col3:
        if not location_summary.empty:
            st.download_button("📍 Location Report (CSV)", location_summary.to_csv(index=False).encode(),
                               "location_report.csv", "text/csv", use_container_width=True)
