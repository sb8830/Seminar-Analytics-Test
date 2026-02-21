"""
Enterprise Seminar Analytics Dashboard
======================================
A comprehensive analytics solution for offline seminar performance tracking.

Version: 2.0.0
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
from typing import Dict, List, Optional, Any
import logging

# ── Configuration & Setup ──
st.set_page_config(
    page_title="Enterprise Seminar Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://support.example.com',
        'Report a bug': 'https://bugs.example.com',
        'About': '# Enterprise Seminar Analytics v2.0'
    }
)

# ── Logging ──
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Custom CSS - Enterprise Grade ──
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 98%; }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px; border-radius: 12px; color: white;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1); transition: transform 0.2s ease;
    }
    .metric-card:hover { transform: translateY(-2px); }
    .metric-card.success { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
    .metric-card.warning { background: linear-gradient(135deg, #f2994a 0%, #f2c94c 100%); }
    .metric-card.danger { background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%); }
    h1 { color: #1a1a2e; font-weight: 700; letter-spacing: -0.5px; }
    h2, h3 { color: #16213e; font-weight: 600; }
    .custom-card { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .sidebar-section { background: #f8f9fa; padding: 15px; border-radius: 10px; margin-bottom: 15px; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
    .animate-fade { animation: fadeIn 0.5s ease-out; }
</style>
""", unsafe_allow_html=True)

# ── Session State ──
def init_session_state():
    defaults = {
        'data_loaded': False, 'last_refresh': None, 'view_mode': 'dashboard',
        'selected_seminar': None, 'date_range': None, 'custom_filters': {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# ── Data Loading & Processing ──
class DataProcessor:
    @staticmethod
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_attendee_data(file) -> pd.DataFrame:
        try:
            sheets = pd.read_excel(file, sheet_name=None, header=0)
            frames = []
            for name, df in sheets.items():
                df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace(r'[^\w]', '_', regex=True)
                student_cols = ['student_name', 'studentname', 'name']
                if any(col in df.columns for col in student_cols):
                    df['source_sheet'] = name
                    frames.append(df)
            if frames:
                result = pd.concat(frames, ignore_index=True)
                return DataProcessor._clean_attendee_data(result)
            return pd.DataFrame()
        except Exception as e:
            logger.error(f"Error loading attendee data: {e}")
            st.error(f"Failed to load attendee data: {str(e)}")
            return pd.DataFrame()

    @staticmethod
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_seminar_data(file) -> pd.DataFrame:
        try:
            df = pd.read_excel(file, sheet_name=0, header=1)
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            logger.error(f"Error loading seminar data: {e}")
            st.error(f"Failed to load seminar data: {str(e)}")
            return pd.DataFrame()

    @staticmethod
    def _clean_attendee_data(df: pd.DataFrame) -> pd.DataFrame:
        col_mapping = {
            'student_name': 'student_name', 'studentname': 'student_name', 'name': 'student_name',
            'phone_no': 'phone', 'contact': 'phone', 'email_id': 'email',
            'service': 'service_name', 'course': 'service_name', 'batch': 'batch_date',
            'payment': 'payment_received', 'amount_paid': 'payment_received',
            'gst_amount': 'total_gst', 'total': 'total_amount', 'due_amount': 'total_due',
            'balance': 'total_due', 'student_status': 'status', 'enrollment_status': 'status'
        }
        df = df.rename(columns={k: v for k, v in col_mapping.items() if k in df.columns})
        
        numeric_cols = ['payment_received', 'total_gst', 'total_amount', 'total_due', 'total_additional_charges']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', '').str.replace('₹', '').str.replace(' ', ''),
                    errors='coerce'
                ).fillna(0)
        
        text_cols = ['student_name', 'email', 'service_name', 'status']
        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().replace('nan', '')
        return df

    @staticmethod
    @st.cache_data(ttl=3600, show_spinner=False)
    def parse_seminar_data(df: pd.DataFrame) -> pd.DataFrame:
        try:
            col_map = {}
            cols_lower = {c: c.strip().lower().replace('\n', ' ').replace('\r', ' ') for c in df.columns}
            
            patterns = {
                'sr_no': ['sr no', 'sr.no', 'serial', 's.no', 'no'],
                'trainer': ['trainer', 'faculty', 'speaker', 'mentor'],
                'location': ['location', 'venue', 'city', 'branch'],
                'seminar_date': ['seminar date', 'event date', 'date'],
                'targeted': ['targeted', 'target', 'expected'],
                'total_attended': ['total attended', 'attendance', 'attended'],
                'actual_attended': ['actual attended', 'actual attendance'],
                'total_seat_booked': ['total seat booked', 'booked seats', 'seats booked'],
                'actual_expenses': ['actual expense', 'expense', 'cost', 'actual cost'],
                'expected_revenue': ['expected revenue', 'expected', 'projected revenue'],
                'actual_revenue': ['actual revenue', 'revenue', 'collection'],
                'total_revenue': ['total revenue', 'total collection'],
                'surplus_deficit': ['surplus', 'deficit', 'profit loss', 'p/l'],
                'er_to_ae': ['er to ae', 'er/ae'],
                'ar_to_ae': ['ar to ae', 'ar/ae'],
                'attended_to_booked_pct': ['attended to booked', 'conversion %'],
                'morning_total': ['morning total', 'morning'],
                'evening_total': ['evening total', 'evening']
            }
            
            for target, keywords in patterns.items():
                for orig, low in cols_lower.items():
                    if any(kw in low for kw in keywords):
                        col_map[target] = orig
                        break
            
            renamed = df.rename(columns={v: k for k, v in col_map.items() if v in df.columns})
            
            if 'sr_no' in renamed.columns:
                renamed = renamed[pd.to_numeric(renamed['sr_no'], errors='coerce').notna()]
                renamed['sr_no'] = renamed['sr_no'].astype(int)
            
            numeric_cols = [
                'targeted', 'total_attended', 'actual_attended', 'total_seat_booked',
                'actual_expenses', 'expected_revenue', 'actual_revenue', 'total_revenue',
                'surplus_deficit', 'er_to_ae', 'ar_to_ae', 'morning_total', 'evening_total'
            ]
            
            for col in numeric_cols:
                if col in renamed.columns:
                    renamed[col] = pd.to_numeric(
                        renamed[col].astype(str).str.replace(',', '').str.replace('%', '').str.replace('₹', ''),
                        errors='coerce'
                    ).fillna(0)
            
            if 'seminar_date' in renamed.columns:
                renamed['seminar_date'] = pd.to_datetime(renamed['seminar_date'], errors='coerce')
                renamed['month'] = renamed['seminar_date'].dt.to_period('M')
                renamed['quarter'] = renamed['seminar_date'].dt.to_period('Q')
            
            return renamed
        except Exception as e:
            logger.error(f"Error parsing seminar data: {e}")
            return df

# ── Analytics Engine ──
class AnalyticsEngine:
    @staticmethod
    def calculate_kpis(df: pd.DataFrame) -> Dict[str, Any]:
        if df.empty:
            return {}
        
        kpis = {}
        kpis['total_seminars'] = len(df)
        
        if 'total_attended' in df.columns:
            kpis['total_attended'] = int(df['total_attended'].sum())
            kpis['avg_attendance'] = round(df['total_attended'].mean(), 1)
            kpis['max_attendance'] = int(df['total_attended'].max())
        
        if 'actual_revenue' in df.columns:
            kpis['total_revenue'] = float(df['actual_revenue'].sum())
            kpis['avg_revenue'] = round(df['actual_revenue'].mean(), 2)
            kpis['max_revenue'] = float(df['actual_revenue'].max())
        
        if 'actual_expenses' in df.columns:
            kpis['total_expenses'] = float(df['actual_expenses'].sum())
            kpis['avg_expense'] = round(df['actual_expenses'].mean(), 2)
        
        if 'surplus_deficit' in df.columns:
            kpis['total_profit'] = float(df['surplus_deficit'].sum())
            kpis['profitable_seminars'] = int((df['surplus_deficit'] > 0).sum())
            kpis['loss_seminars'] = int((df['surplus_deficit'] < 0).sum())
            if kpis['total_revenue'] > 0:
                kpis['profit_margin'] = round(kpis['total_profit'] / kpis['total_revenue'] * 100, 2)
            else:
                kpis['profit_margin'] = 0
        
        if 'total_seat_booked' in df.columns and 'total_attended' in df.columns:
            total_booked = df['total_seat_booked'].sum()
            total_attended = df['total_attended'].sum()
            kpis['total_booked'] = int(total_booked)
            kpis['conversion_rate'] = round((total_attended / total_booked * 100) if total_booked > 0 else 0, 2)
        
        if 'targeted' in df.columns and 'total_attended' in df.columns:
            kpis['target_achievement'] = round(
                (df['total_attended'].sum() / df['targeted'].sum() * 100) if df['targeted'].sum() > 0 else 0, 2
            )
        
        return kpis
    
    @staticmethod
    def calculate_location_stats(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or 'location' not in df.columns:
            return pd.DataFrame()
        
        stats = df.groupby('location').agg({
            'sr_no': 'count',
            'total_attended': 'sum',
            'actual_revenue': 'sum',
            'actual_expenses': 'sum',
            'surplus_deficit': 'sum'
        }).rename(columns={'sr_no': 'seminar_count'})
        
        stats['profit_margin'] = (stats['surplus_deficit'] / stats['actual_revenue'] * 100).round(2)
        return stats.sort_values('surplus_deficit', ascending=False)
    
    @staticmethod
    def calculate_trainer_stats(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or 'trainer' not in df.columns:
            return pd.DataFrame()
        
        all_trainers = []
        for t in df['trainer'].dropna():
            for name in str(t).split(','):
                name = name.strip().split('\n')[0].strip()
                if name:
                    all_trainers.append(name)
        
        if not all_trainers:
            return pd.DataFrame()
        
        trainer_df = pd.DataFrame({'trainer': all_trainers})
        stats = trainer_df.groupby('trainer').size().reset_index(name='seminars')
        return stats.sort_values('seminars', ascending=False)

# ── Main Application ──
def main():
    st.title("📊 Enterprise Seminar Analytics Dashboard")
    st.caption("Offline Seminar Performance • 2025-26 | Real-time Analytics")
    
    # ── File Upload Section ──
    with st.container():
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            file1 = st.file_uploader("📁 Upload **Attendee Details**", type=["xlsx", "xls"], key="f1")
        with col2:
            file2 = st.file_uploader("📁 Upload **Seminar Report**", type=["xlsx", "xls"], key="f2")
        with col3:
            st.write("")
            st.write("")
            if st.button("🔄 Refresh Data", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
    
    if not file1 or not file2:
        st.info("👆 Please upload both Excel files to begin analysis.")
        st.stop()
    
    # ── Load Data ──
    with st.spinner('Loading and processing data...'):
        attendee_df = DataProcessor.load_attendee_data(file1)
        seminar_raw = DataProcessor.load_seminar_data(file2)
        seminar_df = DataProcessor.parse_seminar_data(seminar_raw)
    
    if seminar_df.empty:
        st.error("❌ No valid seminar data found. Please check your file format.")
        st.stop()
    
    # ── Sidebar Filters ──
    st.sidebar.header("🔍 Advanced Filters")
    st.sidebar.markdown("---")
    
    # Location filter
    locations = sorted(seminar_df['location'].dropna().unique()) if 'location' in seminar_df.columns else []
    selected_locations = st.sidebar.multiselect("📍 Location", locations, default=locations, help="Select one or more locations")
    
    # Trainer filter
    all_trainers = set()
    if 'trainer' in seminar_df.columns:
        for t in seminar_df['trainer'].dropna():
            for name in str(t).split(','):
                name = name.strip().split('\n')[0].strip()
                if name:
                    all_trainers.add(name)
    all_trainers = sorted(all_trainers)
    
    selected_trainers = st.sidebar.multiselect("👨‍🏫 Trainer", all_trainers, default=[], help="Filter by trainer")
    
    # Date Range Filter
    if 'seminar_date' in seminar_df.columns:
        min_date = seminar_df['seminar_date'].min()
        max_date = seminar_df['seminar_date'].max()
        if pd.notna(min_date) and pd.notna(max_date):
            date_range = st.sidebar.date_input(
                "📅 Date Range", 
                value=(min_date, max_date),
                help="Filter by seminar date"
            )
        else:
            date_range = None
    else:
        date_range = None
    
    # Profitability filter
    st.sidebar.markdown("### 💰 Profitability")
    profit_filter = st.sidebar.radio("Filter", ["All", "Profitable", "Loss-making"], horizontal=True)
    
    # Attendance Range
    if 'total_attended' in seminar_df.columns:
        min_att, max_att = int(seminar_df['total_attended'].min()), int(seminar_df['total_attended'].max())
        attendance_range = st.sidebar.slider(
            "👥 Attendance Range",
            min_value=min_att, max_value=max_att,
            value=(min_att, max_att)
        )
    else:
        attendance_range = None
    
    # ── Apply Filters ──
    filtered = seminar_df.copy()
    
    if selected_locations:
        filtered = filtered[filtered['location'].isin(selected_locations)]
    
    if selected_trainers:
        def has_trainer(trainer_str):
            names = [n.strip().split('\n')[0].strip() for n in str(trainer_str).split(',')]
            return any(n in selected_trainers for n in names)
        filtered = filtered[filtered['trainer'].apply(has_trainer)]
    
    if date_range and len(date_range) == 2:
        filtered = filtered[(filtered['seminar_date'] >= pd.to_datetime(date_range[0])) &
