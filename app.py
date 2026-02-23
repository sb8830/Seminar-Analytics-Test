import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re

st.set_page_config(page_title="Seminar Analytics Dashboard", layout="wide", page_icon="📊")

# ── Dark theme styling ──
st.markdown("""
<style>
    .stMetric { background: #1a1a2e; padding: 15px; border-radius: 10px; border: 1px solid #2d2d44; }
    .stMetric label { color: #8b8ba3 !important; font-size: 0.75rem !important; text-transform: uppercase; letter-spacing: 1px; }
    .stMetric [data-testid="stMetricValue"] { color: #e0e0ff !important; font-size: 1.8rem !important; }
    .insight-card { background: #1a1a2e; border: 1px solid #2d2d44; border-radius: 10px; padding: 15px; margin: 5px 0; }
    h1, h2, h3 { color: #e0e0ff !important; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ──
def clean_numeric(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).replace(',', '').replace('₹', '').replace('%', '').strip()
    try: return float(s)
    except: return 0

def parse_pct(val):
    if pd.isna(val): return 0
    s = str(val).replace('%', '').strip()
    try: return float(s)
    except: return 0

# ── Load Seminar Report ──
@st.cache_data
def load_seminar_report(file):
    df = pd.read_excel(file, sheet_name=0, header=1)
    # Auto-detect columns by position (handles messy headers)
    cols = df.columns.tolist()
    records = []
    for _, row in df.iterrows():
        sr = clean_numeric(row.iloc[0])
        if sr == 0 or pd.isna(row.iloc[0]): continue
        rec = {
            'sr_no': int(sr),
            'trainer': str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else '',
            'location': str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else '',
            'seminar_date': str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else '',
            'batch_date': str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else '',
            'targeted': int(clean_numeric(row.iloc[5])),
            'morning_total': int(clean_numeric(row.iloc[6])),
            'evening_total': int(clean_numeric(row.iloc[8])),
            'total_attended': int(clean_numeric(row.iloc[10])),
            'actual_attended': int(clean_numeric(row.iloc[11])) if clean_numeric(row.iloc[11]) > 0 else int(clean_numeric(row.iloc[10])),
            'targeted_to_attended_pct': parse_pct(row.iloc[12]),
            'total_seat_booked': int(clean_numeric(row.iloc[15])),
            'non_webinar': int(clean_numeric(row.iloc[16])),
            'attended_to_seat_booked_pct': parse_pct(row.iloc[17]),
            'targeted_to_seat_booked_pct': parse_pct(row.iloc[19]),
            'actual_expenses': clean_numeric(row.iloc[20]),
            'expected_revenue': clean_numeric(row.iloc[21]),
            'actual_revenue': clean_numeric(row.iloc[22]),
            'total_revenue': clean_numeric(row.iloc[23]),
            'surplus_deficit': clean_numeric(row.iloc[24]),
            'surplus_to_expense_pct': parse_pct(row.iloc[25]),
            'er_to_ae': clean_numeric(row.iloc[26]),
            'ar_to_ae': clean_numeric(row.iloc[27]),
        }
        records.append(rec)
    return pd.DataFrame(records)

# ── Load Location Revenue from Sheet 2 ──
@st.cache_data
def load_location_revenue(file):
    try:
        df = pd.read_excel(file, sheet_name=1, header=1)
        records = []
        for _, row in df.iterrows():
            place = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            if not place or place.lower() in ['place', 'nan', '']: continue
            rec = {
                'location': place,
                'pti_students': int(clean_numeric(row.iloc[1])),
                'series_students': int(clean_numeric(row.iloc[2])),
                'other_series': int(clean_numeric(row.iloc[3])),
                'other_course': int(clean_numeric(row.iloc[4])),
                'active': int(clean_numeric(row.iloc[5])),
                'inactive': int(clean_numeric(row.iloc[6])),
            }
            records.append(rec)
        # Revenue data from right side of same sheet
        df2 = pd.read_excel(file, sheet_name=1, header=1)
        rev_map = {}
        for _, row in df2.iterrows():
            place = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ''
            if not place or place.lower() in ['place', 'nan', '']: continue
            rev_map[place] = {
                'pti_revenue': clean_numeric(row.iloc[8]),
                'series_revenue': clean_numeric(row.iloc[9]),
                'others_revenue': clean_numeric(row.iloc[10]),
                'total_revenue': clean_numeric(row.iloc[11]),
            }
        result = pd.DataFrame(records)
        for col in ['pti_revenue', 'series_revenue', 'others_revenue', 'total_revenue']:
            result[col] = result['location'].map(lambda x: rev_map.get(x, {}).get(col, 0))
        return result
    except:
        return pd.DataFrame()

# ── Load Conversion List ──
@st.cache_data
def load_conversion_list(file):
    ext = file.name.split('.')[-1].lower()
    if ext == 'xlsb':
        df = pd.read_excel(file, engine='pyxlsb')
    else:
        df = pd.read_excel(file)

    # Normalize column names
    df.columns = [str(c).strip().lower().replace(' ', '_') for c in df.columns]

    # Parse key columns
    if 'total_amount' in df.columns:
        df['total_amount'] = df['total_amount'].apply(clean_numeric)
    if 'payment_received' in df.columns:
        df['payment_received'] = df['payment_received'].apply(clean_numeric)
    if 'total_due' in df.columns:
        df['total_due'] = df['total_due'].apply(clean_numeric)
    if 'order_date' in df.columns:
        df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce')
    if 'batch_date' in df.columns:
        df['batch_date_parsed'] = pd.to_datetime(df['batch_date'], errors='coerce', dayfirst=True)

    return df

# ── Match Conversions to Seminar Locations ──
def match_conversions_to_seminars(seminar_df, conv_df):
    """Cross-reference conversion list with seminar locations using batch date proximity"""
    # Parse seminar batch dates
    seminar_dates = []
    for _, row in seminar_df.iterrows():
        bd = row['batch_date']
        try:
            parsed = pd.to_datetime(bd, dayfirst=True, errors='coerce')
            seminar_dates.append({
                'location': row['location'],
                'batch_date': parsed,
                'seminar_date': row['seminar_date'],
                'sr_no': row['sr_no']
            })
        except:
            seminar_dates.append({
                'location': row['location'],
                'batch_date': pd.NaT,
                'seminar_date': row['seminar_date'],
                'sr_no': row['sr_no']
            })

    sem_dates_df = pd.DataFrame(seminar_dates)

    # For each conversion, find the closest matching seminar by batch date
    matched = []
    if 'batch_date_parsed' not in conv_df.columns:
        return pd.DataFrame()

    for _, conv in conv_df.iterrows():
        conv_batch = conv.get('batch_date_parsed')
        if pd.isna(conv_batch): continue

        # Find seminars with batch date within 7 days
        valid = sem_dates_df[sem_dates_df['batch_date'].notna()].copy()
        valid['date_diff'] = (valid['batch_date'] - conv_batch).abs()
        close = valid[valid['date_diff'] <= timedelta(days=7)]

        if len(close) > 0:
            best = close.loc[close['date_diff'].idxmin()]
            matched.append({
                'student_name': conv.get('student_name', ''),
                'phone': conv.get('phone', ''),
                'service_name': conv.get('service_name', ''),
                'total_amount': conv.get('total_amount', 0),
                'payment_received': conv.get('payment_received', 0),
                'status': conv.get('status', ''),
                'order_date': conv.get('order_date'),
                'batch_date': conv.get('batch_date'),
                'matched_location': best['location'],
                'matched_seminar_sr': best['sr_no'],
                'sales_rep_name': conv.get('sales_rep_name', ''),
            })

    return pd.DataFrame(matched)

# ────────────── APP ──────────────
st.title("📊 Seminar Analytics Dashboard")
st.caption("Upload Seminar Report & Conversion List for cross-referenced insights")

col_up1, col_up2 = st.columns(2)
with col_up1:
    seminar_file = st.file_uploader("📄 Upload Seminar Report (.xlsx)", type=['xlsx', 'xls'])
with col_up2:
    conv_file = st.file_uploader("📄 Upload Conversion List (.xlsx / .xlsb)", type=['xlsx', 'xls', 'xlsb'])

if seminar_file:
    seminar_df = load_seminar_report(seminar_file)
    loc_rev_df = load_location_revenue(seminar_file)

    # ── Sidebar Filters ──
    st.sidebar.header("🎛️ Filters")
    all_locations = sorted(seminar_df['location'].unique())
    all_trainers = sorted(set(t.strip() for trainers in seminar_df['trainer'] for t in trainers.split(',')))

    sel_locations = st.sidebar.multiselect("Locations", all_locations)
    sel_trainers = st.sidebar.multiselect("Trainers", all_trainers)
    profit_filter = st.sidebar.radio("Profitability", ["All", "Profitable", "Loss"], horizontal=True)

    # Apply filters
    filtered = seminar_df.copy()
    if sel_locations:
        filtered = filtered[filtered['location'].isin(sel_locations)]
    if sel_trainers:
        filtered = filtered[filtered['trainer'].apply(
            lambda x: any(t.strip() in sel_trainers for t in x.split(','))
        )]
    if profit_filter == "Profitable":
        filtered = filtered[filtered['surplus_deficit'] > 0]
    elif profit_filter == "Loss":
        filtered = filtered[filtered['surplus_deficit'] < 0]

    # ── KPIs ──
    st.markdown("---")
    with_exp = filtered[filtered['actual_expenses'] > 0]
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Seminars", len(filtered))
    k2.metric("Total Attendees", f"{filtered['total_attended'].sum():,}")
    k3.metric("Total Revenue", f"₹{filtered['actual_revenue'].sum()/100000:.1f}L")
    k4.metric("Total Expenses", f"₹{filtered['actual_expenses'].sum()/100000:.1f}L")
    avg_conv = filtered['attended_to_seat_booked_pct'].mean() if len(filtered) > 0 else 0
    k5.metric("Avg Conversion", f"{avg_conv:.1f}%")
    profitable_count = len(with_exp[with_exp['surplus_deficit'] > 0])
    k6.metric("Profitable", f"{profitable_count}/{len(with_exp)}")

    # ── Charts ──
    tab1, tab2, tab3, tab4 = st.tabs(["💰 Revenue vs Expense", "📊 Surplus/Deficit", "👥 Attendance", "📈 Revenue Breakdown"])

    with tab1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Expense', x=filtered['location'], y=filtered['actual_expenses'], marker_color='#ff6b6b'))
        fig.add_trace(go.Bar(name='Revenue', x=filtered['actual_revenue'], y=filtered['actual_revenue'], marker_color='#51cf66'))
        fig.update_layout(barmode='group', template='plotly_dark', height=400,
                         xaxis_tickangle=-45, title='Revenue vs Expense by Location')
        # Fix: use location for x-axis on both traces
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Expense', x=filtered['location'], y=filtered['actual_expenses'], marker_color='#ff6b6b'))
        fig.add_trace(go.Bar(name='Revenue', x=filtered['location'], y=filtered['actual_revenue'], marker_color='#51cf66'))
        fig.update_layout(barmode='group', template='plotly_dark', height=400, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        colors = ['#51cf66' if v >= 0 else '#ff6b6b' for v in filtered['surplus_deficit']]
        fig = go.Figure(go.Bar(x=filtered['location'], y=filtered['surplus_deficit'], marker_color=colors))
        fig.update_layout(template='plotly_dark', height=400, xaxis_tickangle=-45, title='Surplus / Deficit')
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Targeted', x=filtered['location'], y=filtered['targeted'], marker_color='#748ffc'))
        fig.add_trace(go.Bar(name='Attended', x=filtered['location'], y=filtered['total_attended'], marker_color='#51cf66'))
        fig.add_trace(go.Bar(name='Booked', x=filtered['location'], y=filtered['total_seat_booked'], marker_color='#ffd43b'))
        fig.update_layout(barmode='group', template='plotly_dark', height=400, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    with tab4:
        if len(loc_rev_df) > 0:
            top15 = loc_rev_df.nlargest(15, 'total_revenue')
            fig = go.Figure()
            fig.add_trace(go.Bar(name='PTI', x=top15['location'], y=top15['pti_revenue'], marker_color='#748ffc'))
            fig.add_trace(go.Bar(name='Series', x=top15['location'], y=top15['series_revenue'], marker_color='#51cf66'))
            fig.add_trace(go.Bar(name='Others', x=top15['location'], y=top15['others_revenue'], marker_color='#ffd43b'))
            fig.update_layout(barmode='stack', template='plotly_dark', height=400, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Revenue breakdown not available in this file format")

    # ── Conversion List Cross-Reference ──
    if conv_file:
        st.markdown("---")
        st.header("🔗 Conversion List Cross-Reference")

        conv_df = load_conversion_list(conv_file)
        matched_df = match_conversions_to_seminars(seminar_df, conv_df)

        if len(matched_df) > 0:
            # Summary by location
            conv_summary = matched_df.groupby('matched_location').agg(
                total_conversions=('student_name', 'count'),
                total_amount=('total_amount', 'sum'),
                total_received=('payment_received', 'sum'),
                active_count=('status', lambda x: (x == 'Active').sum()),
            ).reset_index()

            # Merge with seminar data
            merged = seminar_df.merge(conv_summary, left_on='location', right_on='matched_location', how='left')
            merged['total_conversions'] = merged['total_conversions'].fillna(0).astype(int)
            merged['conversion_revenue'] = merged['total_amount'].fillna(0)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Matched Students", f"{len(matched_df):,}")
            c2.metric("Conversion Revenue", f"₹{matched_df['total_amount'].sum()/100000:.1f}L")
            c3.metric("Active Students", f"{(matched_df['status'] == 'Active').sum()}")
            c4.metric("Locations Covered", f"{matched_df['matched_location'].nunique()}")

            tab_c1, tab_c2, tab_c3 = st.tabs(["📊 Conversions by Location", "📋 Student Details", "🔍 Top Courses"])

            with tab_c1:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    name='Conversions',
                    x=conv_summary['matched_location'],
                    y=conv_summary['total_conversions'],
                    marker_color='#748ffc'
                ))
                fig.update_layout(template='plotly_dark', height=400, xaxis_tickangle=-45,
                                 title='Number of Conversions per Seminar Location')
                st.plotly_chart(fig, use_container_width=True)

                # Revenue from conversions vs seminar expense
                fig2 = go.Figure()
                for _, row in conv_summary.iterrows():
                    sem_row = seminar_df[seminar_df['location'] == row['matched_location']]
                    exp = sem_row['actual_expenses'].values[0] if len(sem_row) > 0 else 0
                    fig2.add_trace(go.Bar(
                        name=row['matched_location'],
                        x=['Expense', 'Conv Revenue'],
                        y=[exp, row['total_amount']],
                        showlegend=True
                    ))
                fig2.update_layout(template='plotly_dark', height=400,
                                  title='Seminar Expense vs Conversion Revenue')
                st.plotly_chart(fig2, use_container_width=True)

            with tab_c2:
                st.dataframe(
                    matched_df[['student_name', 'phone', 'service_name', 'total_amount',
                               'status', 'matched_location', 'sales_rep_name']].sort_values('matched_location'),
                    use_container_width=True, height=400
                )

            with tab_c3:
                course_summary = matched_df.groupby('service_name').agg(
                    count=('student_name', 'count'),
                    revenue=('total_amount', 'sum')
                ).sort_values('revenue', ascending=False).reset_index()
                fig = px.pie(course_summary.head(10), values='revenue', names='service_name',
                            title='Top Courses by Revenue', template='plotly_dark',
                            color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No conversions matched to seminar dates. Check batch date formats.")

    # ── Key Insights ──
    st.markdown("---")
    st.header("💡 Key Insights")

    if len(with_exp) > 0:
        best_roi = with_exp.loc[with_exp['surplus_to_expense_pct'].idxmax()]
        worst_roi = with_exp.loc[with_exp['surplus_to_expense_pct'].idxmin()]
        highest_att = filtered.loc[filtered['total_attended'].idxmax()]
        best_conv = filtered.loc[filtered['attended_to_seat_booked_pct'].idxmax()]
        avg_expense = with_exp['actual_expenses'].mean()
        loss_count = len(with_exp[with_exp['surplus_deficit'] < 0])

        i1, i2 = st.columns(2)
        with i1:
            st.success(f"🏆 **Best ROI:** {best_roi['location']} — {best_roi['surplus_to_expense_pct']}% return, ₹{best_roi['surplus_deficit']:,.0f} surplus")
            st.info(f"👥 **Highest Attendance:** {highest_att['location']} — {highest_att['total_attended']} attendees ({highest_att['targeted_to_attended_pct']}% show rate)")
            st.info(f"💰 **Cost Efficiency:** Avg cost ₹{avg_expense:,.0f}. {profitable_count}/{len(with_exp)} seminars profitable")
        with i2:
            st.warning(f"⚠️ **Lowest ROI:** {worst_roi['location']} — {worst_roi['surplus_to_expense_pct']}% return. {loss_count} seminars at loss")
            st.info(f"🎯 **Best Conversion:** {best_conv['location']} — {best_conv['attended_to_seat_booked_pct']}% of attendees booked seats")

    # ── Student Status ──
    if len(loc_rev_df) > 0:
        st.markdown("---")
        col_pie1, col_pie2 = st.columns(2)
        with col_pie1:
            total_active = loc_rev_df['active'].sum()
            total_inactive = loc_rev_df['inactive'].sum()
            fig = px.pie(values=[total_active, total_inactive], names=['Active', 'Inactive/Closed'],
                        title='Student Status Distribution', template='plotly_dark',
                        color_discrete_map={'Active': '#51cf66', 'Inactive/Closed': '#ff6b6b'})
            st.plotly_chart(fig, use_container_width=True)

        with col_pie2:
            total_pti = loc_rev_df['pti_students'].sum()
            total_series = loc_rev_df['series_students'].sum()
            total_other = loc_rev_df['other_series'].sum() + loc_rev_df['other_course'].sum()
            fig = px.pie(values=[total_pti, total_series, total_other],
                        names=['PTI', 'Series 10', 'Others'],
                        title='Student Course Distribution', template='plotly_dark',
                        color_discrete_sequence=['#748ffc', '#51cf66', '#ffd43b'])
            st.plotly_chart(fig, use_container_width=True)

    # ── Data Table ──
    st.markdown("---")
    st.header("📋 Seminar Performance Table")
    display_cols = ['sr_no', 'location', 'trainer', 'seminar_date', 'targeted', 'total_attended',
                   'targeted_to_attended_pct', 'total_seat_booked', 'actual_expenses',
                   'actual_revenue', 'surplus_deficit', 'surplus_to_expense_pct']
    st.dataframe(filtered[display_cols].sort_values('sr_no'), use_container_width=True, height=500)

    # ── Download ──
    csv = filtered.to_csv(index=False)
    st.download_button("📥 Download Filtered Data (CSV)", csv, "seminar_report.csv", "text/csv")

else:
    st.info("👆 Upload the Seminar Report Excel to get started. Optionally add the Conversion List for cross-referencing.")
