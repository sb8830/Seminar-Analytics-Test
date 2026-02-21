"""
Enterprise Seminar Analytics Dashboard
======================================
A comprehensive analytics solution for offline seminar performance tracking,
featuring advanced analytics, interactive visualizations, and enterprise-grade features.

Version: 2.0.0
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple
import logging
import io

# ── Configuration & Setup ──
st.set_page_config(
    page_title="Enterprise Seminar Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://support.example.com',
        'Report a bug': 'https://bugs.example.com',
        'About': '# Enterprise Seminar Analytics v2.0\nBuilt with ❤️ for data-driven decisions'
    }
)

# ── Logging Configuration ──
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Custom CSS - Enterprise Grade ──
st.markdown("""
<style>
    /* Base Styles */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }
    
    /* Headers */
    h1 {
        color: #1a1a2e;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    h2, h3 {
        color: #16213e;
        font-weight: 600;
    }
    
    /* Custom Cards */
    .custom-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border: 1px solid #e8e8e8;
    }
    
    /* Sidebar Section */
    .sidebar-section {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 15px;
    }
    
    /* Status Colors */
    .status-positive { color: #10b981; font-weight: 600; }
    .status-negative { color: #ef4444; font-weight: 600; }
    .status-neutral { color: #6b7280; font-weight: 600; }
    
    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .animate-fade {
        animation: fadeIn 0.5s ease-out;
    }
    
    /* Metric Styling */
    div[data-testid="stMetric"] {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e9ecef;
    }
    div[data-testid="stMetric"] label {
        color: #6b7280;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #1a1a2e;
        font-weight: 700;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 12px 24px;
        border-radius: 8px;
        font-weight: 500;
    }
    
    /* DataFrame */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }
    
    /* File Uploader */
    .stFileUploader {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border: 2px dashed #d1d5db;
    }
</style>
""", unsafe_allow_html=True)

# ── Session State Management ──
def init_session_state():
    """Initialize session state variables."""
    defaults = {
        'data_loaded': False,
        'last_refresh': None,
        'view_mode': 'dashboard',
        'selected_seminar': None,
        'custom_filters': {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# ── Data Loading & Processing ──
class DataProcessor:
    """Enterprise-grade data processing class."""
    
    @staticmethod
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_attendee_data(file) -> pd.DataFrame:
        """Load and process attendee data from Excel file."""
        try:
            logger.info("Loading attendee data...")
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
                result = DataProcessor._clean_attendee_data(result)
                logger.info(f"Loaded {len(result)} attendee records")
                return result
            return pd.DataFrame()
        except Exception as e:
            logger.error(f"Error loading attendee data: {e}")
            st.error(f"Failed to load attendee data: {str(e)}")
            return pd.DataFrame()
    
    @staticmethod
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_seminar_data(file) -> pd.DataFrame:
        """Load and process seminar report data."""
        try:
            logger.info("Loading seminar data...")
            df = pd.read_excel(file, sheet_name=0, header=1)
            df.columns = df.columns.str.strip()
            logger.info(f"Loaded {len(df)} seminar records")
            return df
        except Exception as e:
            logger.error(f"Error loading seminar data: {e}")
            st.error(f"Failed to load seminar data: {str(e)}")
            return pd.DataFrame()
    
    @staticmethod
    def _clean_attendee_data(df: pd.DataFrame) -> pd.DataFrame:
        """Clean and standardize attendee data."""
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
        """Parse and standardize seminar data with intelligent column mapping."""
        try:
            logger.info("Parsing seminar data...")
            
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
            
            logger.info(f"Parsed {len(renamed)} seminar records")
            return renamed
            
        except Exception as e:
            logger.error(f"Error parsing seminar data: {e}")
            return df

# ── Analytics Engine ──
class AnalyticsEngine:
    """Enterprise analytics computation engine."""
    
    @staticmethod
    def calculate_kpis(df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate comprehensive KPIs."""
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
            if kpis.get('total_revenue', 0) > 0:
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
        """Calculate statistics by location."""
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
        """Calculate statistics by trainer."""
        if df.empty or 'trainer' not in df.columns:
            return pd.DataFrame()
        
        all_trainers = []
        for t in df['trainer'].dropna():
            for name in str(t).split(','):
                name = name.strip().split('\n')[0].strip()
                if name:
                    all_trainers.append({'trainer': name, 'location': t})
        
        if not all_trainers:
            return pd.DataFrame()
        
        trainer_df = pd.DataFrame(all_trainers)
        stats = trainer_df.groupby('trainer').size().reset_index(name='seminars')
        return stats.sort_values('seminars', ascending=False)
    
    @staticmethod
    def get_top_performers(df: pd.DataFrame, metric: str = 'surplus_deficit', n: int = 5) -> pd.DataFrame:
        """Get top performing seminars by metric."""
        if df.empty or metric not in df.columns:
            return pd.DataFrame()
        
        return df.nlargest(n, metric)[['location', 'trainer', 'seminar_date', metric, 'total_attended']]
    
    @staticmethod
    def get_bottom_performers(df: pd.DataFrame, metric: str = 'surplus_deficit', n: int = 5) -> pd.DataFrame:
        """Get bottom performing seminars by metric."""
        if df.empty or metric not in df.columns:
            return pd.DataFrame()
        
        return df.nsmallest(n, metric)[['location', 'trainer', 'seminar_date', metric, 'total_attended']]

# ── Visualization Functions ──
def create_revenue_expense_chart(df: pd.DataFrame) -> go.Figure:
    """Create revenue vs expense comparison chart."""
    if df.empty or 'actual_revenue' not in df.columns or 'actual_expenses' not in df.columns:
        return go.Figure()
    
    chart_data = df[df['actual_expenses'] > 0][['location', 'actual_revenue', 'actual_expenses']].copy()
    if chart_data.empty:
        return go.Figure()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Revenue', 
        x=chart_data['location'], 
        y=chart_data['actual_revenue'], 
        marker_color='#1a56db',
        hovertemplate='₹%{y:,.0f}<extra>Revenue</extra>'
    ))
    fig.add_trace(go.Bar(
        name='Expense', 
        x=chart_data['location'], 
        y=chart_data['actual_expenses'], 
        marker_color='#f59e0b',
        hovertemplate='₹%{y:,.0f}<extra>Expense</extra>'
    ))
    fig.update_layout(
        barmode='group',
        height=450,
        xaxis_tickangle=-45,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(l=20, r=20, t=80, b=80)
    )
    return fig

def create_profit_chart(df: pd.DataFrame) -> go.Figure:
    """Create surplus/deficit chart."""
    if df.empty or 'surplus_deficit' not in df.columns:
        return go.Figure()
    
    chart_data = df[df['actual_expenses'] > 0][['location', 'surplus_deficit']].copy()
    if chart_data.empty:
        return go.Figure()
    
    colors = ['#10b981' if v >= 0 else '#ef4444' for v in chart_data['surplus_deficit']]
    
    fig = px.bar(
        chart_data, 
        x='location', 
        y='surplus_deficit',
        color_discrete_sequence=['#10b981']
    )
    fig.update_traces(marker_color=colors)
    fig.update_layout(
        height=450,
        xaxis_tickangle=-45,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=80, b=80)
    )
    return fig

def create_attendance_funnel_chart(df: pd.DataFrame) -> go.Figure:
    """Create attendance funnel chart."""
    if df.empty or not all(c in df.columns for c in ['targeted', 'total_attended', 'total_seat_booked']):
        return go.Figure()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Targeted', 
        x=df['location'], 
        y=df['targeted'], 
        marker_color='#9ca3af',
        opacity=0.5
    ))
    fig.add_trace(go.Bar(
        name='Attended', 
        x=df['location'], 
        y=df['total_attended'], 
        marker_color='#1a56db'
    ))
    fig.add_trace(go.Bar(
        name='Booked', 
        x=df['location'], 
        y=df['total_seat_booked'], 
        marker_color='#10b981'
    ))
    fig.update_layout(
        barmode='group',
        height=450,
        xaxis_tickangle=-45,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(l=20, r=20, t=80, b=80)
    )
    return fig

def create_status_pie_chart(attendee_df: pd.DataFrame) -> go.Figure:
    """Create student status pie chart."""
    if 'status' not in attendee_df.columns or attendee_df.empty:
        return go.Figure()
    
    status_counts = attendee_df['status'].value_counts()
    if status_counts.empty:
        return go.Figure()
    
    fig = px.pie(
        values=status_counts.values, 
        names=status_counts.index, 
        hole=0.45,
        color_discrete_sequence=['#10b981', '#ef4444', '#f59e0b', '#6366f1', '#8b5cf6']
    )
    fig.update_layout(height=400)
    return fig

def create_revenue_breakdown_chart(attendee_df: pd.DataFrame) -> go.Figure:
    """Create revenue breakdown by course chart."""
    if 'service_name' not in attendee_df.columns or 'payment_received' not in attendee_df.columns:
        return go.Figure()
    
    course_rev = attendee_df.groupby('service_name')['payment_received'].sum().sort_values(ascending=False).head(10)
    if course_rev.empty:
        return go.Figure()
    
    fig = px.bar(
        x=course_rev.values, 
        y=course_rev.index, 
        orientation='h',
        color_discrete_sequence=['#1a56db']
    )
    fig.update_layout(
        height=450,
        yaxis_title='',
        xaxis_title='Revenue (₹)',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=80, b=80)
    )
    return fig

def create_trend_chart(df: pd.DataFrame) -> go.Figure:
    """Create trend analysis chart."""
    if df.empty or 'seminar_date' not in df.columns:
        return go.Figure()
    
    df_sorted = df.sort_values('seminar_date')
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    if 'total_attended' in df.columns:
        fig.add_trace(
            go.Bar(
                x=df_sorted['seminar_date'],
                y=df_sorted['total_attended'],
                name='Attendance',
                marker_color='#1a56db',
                opacity=0.7
            ),
            secondary_y=False
        )
    
    if 'surplus_deficit' in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df_sorted['seminar_date'],
                y=df_sorted['surplus_deficit'],
                name='Profit/Loss',
                mode='lines+markers',
                line=dict(color='#10b981', width=3),
                marker=dict(size=8)
            ),
            secondary_y=True
        )
    
    fig.update_layout(
        height=450,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(l=20, r=20, t=80, b=80)
    )
    fig.update_yaxes(title_text="Attendance", secondary_y=False)
    fig.update_yaxes(title_text="Profit/Loss (₹)", secondary_y=True)
    
    return fig

def create_location_heatmap(df: pd.DataFrame) -> go.Figure:
    """Create location performance heatmap."""
    if df.empty or 'location' not in df.columns:
        return go.Figure()
    
    location_stats = AnalyticsEngine.calculate_location_stats(df)
    if location_stats.empty:
        return go.Figure()
    
    # Normalize data for heatmap
    normalized = location_stats.copy()
    for col in ['seminar_count', 'total_attended', 'actual_revenue', 'surplus_deficit']:
        if col in normalized.columns:
            max_val = normalized[col].max()
            if max_val > 0:
                normalized[col] = normalized[col] / max_val * 100
    
    fig = px.imshow(
        normalized[['seminar_count', 'total_attended', 'actual_revenue', 'surplus_deficit']].T,
        labels=dict(x="Location", y="Metric", color="Performance (%)"),
        x=normalized.index,
        y=['Seminars', 'Attendance', 'Revenue', 'Profit'],
        color_continuous_scale='RdYlGn'
    )
    fig.update_layout(height=400)
    return fig

# ── Helper Functions ──
def format_currency(value: float) -> str:
    """Format value as currency."""
    if abs(value) >= 100000:
        return f"₹{value/100000:.1f}L"
    elif abs(value) >= 1000:
        return f"₹{value/1000:.1f}K"
    else:
        return f"₹{value:,.0f}"

def format_number(value: int) -> str:
    """Format large numbers."""
    if abs(value) >= 1000000:
        return f"{value/1000000:.1f}M"
    elif abs(value) >= 1000:
        return f"{value/1000:.1f}K"
    else:
        return f"{value:,}"

# ── Main Application ──
def main():
    """Main application function."""
    
    # Header
    st.title("📊 Enterprise Seminar Analytics Dashboard")
    st.caption("Offline Seminar Performance • 2025-26 | Real-time Analytics")
    
    # ── File Upload Section ──
    st.markdown("### 📁 Data Import")
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        file1 = st.file_uploader("Upload **Attendee Details**", type=["xlsx", "xls"], key="f1")
    with col2:
        file2 = st.file_uploader("Upload **Seminar Report**", type=["xlsx", "xls"], key="f2")
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
    
    st.session_state['data_loaded'] = True
    st.session_state['last_refresh'] = datetime.now()
    
    # ── Sidebar Filters ──
    st.sidebar.header("🔍 Advanced Filters")
    st.sidebar.markdown("---")
    
    # Location filter
    locations = sorted(seminar_df['location'].dropna().unique()) if 'location' in seminar_df.columns else []
    selected_locations = st.sidebar.multiselect(
        "📍 Location", 
        locations, 
        default=locations if locations else None,
        help="Select one or more locations"
    )
    
    # Trainer filter
    all_trainers = set()
    if 'trainer' in seminar_df.columns:
        for t in seminar_df['trainer'].dropna():
            for name in str(t).split(','):
                name = name.strip().split('\n')[0].strip()
                if name:
                    all_trainers.add(name)
    all_trainers = sorted(all_trainers)
    
    selected_trainers = st.sidebar.multiselect(
        "👨‍🏫 Trainer", 
        all_trainers, 
        default=[],
        help="Filter by trainer"
    )
    
    # Date Range Filter
    if 'seminar_date' in seminar_df.columns:
        valid_dates = seminar_df[seminar_df['seminar_date'].notna()]
        if not valid_dates.empty:
            min_date = valid_dates['seminar_date'].min()
            max_date = valid_dates['seminar_date'].max()
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
    else:
        date_range = None
    
    # Profitability filter
    st.sidebar.markdown("### 💰 Profitability")
    profit_filter = st.sidebar.radio("Filter", ["All", "Profitable", "Loss-making"], horizontal=True, label_visibility="collapsed")
    
    # Attendance Range
    if 'total_attended' in seminar_df.columns:
        min_att = int(seminar_df['total_attended'].min())
        max_att = int(seminar_df['total_attended'].max())
        attendance_range = st.sidebar.slider(
            "👥 Attendance Range",
            min_value=min_att, 
            max_value=max_att,
            value=(min_att, max_att),
            help="Filter by attendance count"
        )
    else:
        attendance_range = None
    
    # Revenue Range
    if 'actual_revenue' in seminar_df.columns:
        min_rev = int(seminar_df['actual_revenue'].min())
        max_rev = int(seminar_df['actual_revenue'].max())
        revenue_range = st.sidebar.slider(
            "💰 Revenue Range (₹)",
            min_value=min_rev,
            max_value=max_rev,
            value=(min_rev, max_rev),
            help="Filter by revenue"
        )
    else:
        revenue_range = None
    
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
        filtered = filtered[
            (filtered['seminar_date'] >= pd.to_datetime(date_range[0])) & 
            (filtered['seminar_date'] <= pd.to_datetime(date_range[1]))
        ]
    
    if profit_filter == "Profitable" and 'surplus_deficit' in filtered.columns:
        filtered = filtered[filtered['surplus_deficit'] > 0]
    elif profit_filter == "Loss-making" and 'surplus_deficit' in filtered.columns:
        filtered = filtered[filtered['surplus_deficit'] < 0]
    
    if attendance_range and 'total_attended' in filtered.columns:
        filtered = filtered[
            (filtered['total_attended'] >= attendance_range[0]) & 
            (filtered['total_attended'] <= attendance_range[1])
        ]
    
    if revenue_range and 'actual_revenue' in filtered.columns:
        filtered = filtered[
            (filtered['actual_revenue'] >= revenue_range[0]) & 
            (filtered['actual_revenue'] <= revenue_range[1])
        ]
    
    # Filter attendee data based on filtered seminars
    if 'location' in filtered.columns:
        filtered_locations = filtered['location'].unique()
        if not attendee_df.empty and 'service_name' in attendee_df.columns:
            # Keep all attendees if no specific location filter
            pass
    
    # ── KPI Section ──
    st.markdown("---")
    st.subheader("📈 Key Performance Indicators")
    
    kpis = AnalyticsEngine.calculate_kpis(filtered)
    
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    
    k1.metric("📋 Seminars", kpis.get('total_seminars', 0))
    k2.metric("👥 Attendees", format_number(kpis.get('total_attended', 0)))
    k3.metric("💰 Revenue", format_currency(kpis.get('total_revenue', 0)))
    k4.metric("📤 Expenses", format_currency(kpis.get('total_expenses', 0)))
    k5.metric("🎯 Conversion", f"{kpis.get('conversion_rate', 0)}%")
    k6.metric("✅ Profitable", f"{kpis.get('profitable_seminars', 0)}/{kpis.get('total_seminars', 0)}")
    
    # Secondary KPIs
    s1, s2, s3, s4, s5, s6 = st.columns(6)
    s1.metric("📊 Avg Attendance", kpis.get('avg_attendance', 0))
    s2.metric("💵 Avg Revenue", format_currency(kpis.get('avg_revenue', 0)))
    s3.metric("📉 Total Profit", format_currency(kpis.get('total_profit', 0)))
    s4.metric("📈 Profit Margin", f"{kpis.get('profit_margin', 0)}%")
    s5.metric("🎯 Target Achievement", f"{kpis.get('target_achievement', 0)}%")
    s6.metric("📍 Loss-making", kpis.get('loss_seminars', 0))
    
    # ── Charts Section ──
    st.markdown("---")
    st.subheader("📊 Analytics Dashboard")
    
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "💰 Revenue vs Expense", 
        "📈 Profit Analysis", 
        "🎯 Attendance Funnel",
        "📅 Trend Analysis",
        "🗺️ Location Heatmap",
        "👥 Student Status"
    ])
    
    with tab1:
        st.markdown("#### Revenue vs Expense by Location")
        fig_rev = create_revenue_expense_chart(filtered)
        if fig_rev:
            st.plotly_chart(fig_rev, use_container_width=True)
        else:
            st.warning("Insufficient data for revenue vs expense chart")
    
    with tab2:
        st.markdown("#### Surplus/Deficit by Location")
        fig_profit = create_profit_chart(filtered)
        if fig_profit:
            st.plotly_chart(fig_profit, use_container_width=True)
        else:
            st.warning("Insufficient data for profit chart")
    
    with tab3:
        st.markdown("#### Attendance Funnel Analysis")
        fig_funnel = create_attendance_funnel_chart(filtered)
        if fig_funnel:
            st.plotly_chart(fig_funnel, use_container_width=True)
        else:
            st.warning("Insufficient data for attendance funnel")
    
    with tab4:
        st.markdown("#### Performance Trends Over Time")
        fig_trend = create_trend_chart(filtered)
        if fig_trend:
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.warning("Insufficient data for trend analysis")
    
    with tab5:
        st.markdown("#### Location Performance Heatmap")
        fig_heatmap = create_location_heatmap(filtered)
        if fig_heatmap:
            st.plotly_chart(fig_heatmap, use_container_width=True)
        else:
            st.warning("Insufficient data for heatmap")
    
    with tab6:
        st.markdown("#### Student Status Distribution")
        fig_status = create_status_pie_chart(attendee_df)
        if fig_status:
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.warning("Insufficient data for status chart")
    
    # ── Course Revenue Breakdown ──
    st.markdown("---")
    st.subheader("💰 Revenue Breakdown by Course")
    
    fig_course = create_revenue_breakdown_chart(attendee_df)
    if fig_course:
        st.plotly_chart(fig_course, use_container_width=True)
    else:
        st.warning("Insufficient data for course revenue breakdown")
    
    # ── Key Insights ──
    st.markdown("---")
    st.subheader("💡 Key Insights & Recommendations")
    
    with_exp = filtered[filtered['actual_expenses'] > 0] if 'actual_expenses' in filtered.columns else filtered
    
    if not with_exp.empty and 'surplus_deficit' in with_exp.columns:
        col_a, col_b = st.columns(2)
        
        # Best and worst performers
        best = with_exp.loc[with_exp['surplus_deficit'].idxmax()]
        worst = with_exp.loc[with_exp['surplus_deficit'].idxmin()]
        
        with col_a:
            st.success(f"**🏆 Best ROI:** {best.get('location', 'N/A')} — Surplus of ₹{int(best.get('surplus_deficit', 0)):,}")
            if 'total_attended' in filtered.columns:
                top_att = filtered.loc[filtered['total_attended'].idxmax()]
                st.info(f"**👥 Highest Attendance:** {top_att.get('location', 'N/A')} — {int(top_att['total_attended']):,} attendees")
        
        with col_b:
            st.error(f"**⚠️ Worst ROI:** {worst.get('location', 'N/A')} — Deficit of ₹{abs(int(worst.get('surplus_deficit', 0))):,}")
            loss_count = int((with_exp['surplus_deficit'] < 0).sum())
            st.warning(f"**📉 Loss-making seminars:** {loss_count} out of {len(with_exp)} ran at a loss")
        
        # Recommendations
        st.markdown("#### 📋 Recommendations")
        
        loss_rate = (loss_count / len(with_exp) * 100) if len(with_exp) > 0 else 0
        if loss_rate > 30:
            st.markdown("🔴 **High Priority:** More than 30% of seminars are running at a loss. Review pricing strategy and cost control measures.")
        
        avg_conversion = kpis.get('conversion_rate', 0)
        if avg_conversion < 50:
            st.markdown("🟡 **Medium Priority:** Conversion rate below 50%. Consider improving follow-up processes and engagement strategies.")
        
        if kpis.get('profit_margin', 0) < 10:
            st.markdown("🟡 **Medium Priority:** Profit margins below 10%. Analyze expense structure and explore revenue optimization.")
    
    # ── Top/Bottom Performers ──
    st.markdown("---")
    st.subheader("🏆 Top & Bottom Performers")
    
    col_top, col_bottom = st.columns(2)
    
    with col_top:
        st.markdown("#### 🏆 Top 5 by Profit")
        top_performers = AnalyticsEngine.get_top_performers(filtered, 'surplus_deficit', 5)
        if not top_performers.empty:
            st.dataframe(
                top_performers.style.format({
                    'surplus_deficit': '₹{:,.0f}',
                    'total_attended': '{:,}'
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No data available")
    
    with col_bottom:
        st.markdown("#### 📉 Bottom 5 by Profit")
        bottom_performers = AnalyticsEngine.get_bottom_performers(filtered, 'surplus_deficit', 5)
        if not bottom_performers.empty:
            st.dataframe(
                bottom_performers.style.format({
                    'surplus_deficit': '₹{:,.0f}',
                    'total_attended': '{:,}'
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No data available")
    
    # ── Location Statistics ──
    st.markdown("---")
    st.subheader("📍 Location-wise Statistics")
    
    location_stats = AnalyticsEngine.calculate_location_stats(filtered)
    if not location_stats.empty:
        st.dataframe(
            location_stats.style.format({
                'total_attended': '{:,}',
                'actual_revenue': '₹{:,.0f}',
                'actual_expenses': '₹{:,.0f}',
                'surplus_deficit': '₹{:,.0f}',
                'profit_margin': '{:.2f}%'
            }),
            use_container_width=True
        )
    else:
        st.info("No location data available")
    
    # ── Seminar Data Table ──
    st.markdown("---")
    st.subheader("📋 Seminar Performance Data")
    
    display_cols = [c for c in [
        'sr_no', 'location', 'trainer', 'seminar_date', 'targeted',
        'total_attended', 'total_seat_booked', 'actual_expenses',
        'actual_revenue', 'surplus_deficit', 'er_to_ae'
    ] if c in filtered.columns]
    
    if display_cols:
        st.dataframe(
            filtered[display_cols].reset_index(drop=True),
            use_container_width=True,
            height=400,
            column_config={
                "actual_expenses": st.column_config.NumberColumn("Expenses", format="₹%d"),
                "actual_revenue": st.column_config.NumberColumn("Revenue", format="₹%d"),
                "surplus_deficit": st.column_config.NumberColumn("Surplus/Deficit", format="₹%d"),
                "seminar_date": st.column_config.DateColumn("Date", format="DD/MM/YYYY")
            }
        )
    
    # ── Attendee Details ──
    st.markdown("---")
    st.subheader("👤 Attendee Details")
    
    att_display = [c for c in [
        'student_name', 'phone', 'email', 'service_name', 'batch_date',
        'payment_received', 'total_gst', 'status', 'total_amount', 'total_due'
    ] if c in attendee_df.columns]
    
    if 'status' in attendee_df.columns:
        status_filter = st.multiselect(
            "Filter by Status", 
            attendee_df['status'].dropna().unique().tolist()
        )
        display_att = attendee_df[attendee_df['status'].isin(status_filter)] if status_filter else attendee_df
    else:
        display_att = attendee_df
    
    if att_display:
        st.dataframe(
            display_att[att_display].head(500).reset_index(drop=True),
            use_container_width=True,
            height=400,
            column_config={
                "payment_received": st.column_config.NumberColumn("Payment", format="₹%d"),
                "total_amount": st.column_config.NumberColumn("Total Amt", format="₹%d"),
                "total_due": st.column_config.NumberColumn("Due", format="₹%d")
            }
        )
        st.caption(f"Showing {min(500, len(display_att))} of {len(display_att)} records")
    
    # ── Export Options ──
    st.markdown("---")
    st.subheader("📤 Export Data")
    
    col_exp1, col_exp2, col_exp3 = st.columns(3)
    
    with col_exp1:
        # Export filtered seminar data
        if not filtered.empty:
            csv_seminar = filtered.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Seminar Data (CSV)",
                data=csv_seminar,
                file_name="seminar_analytics.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with col_exp2:
        # Export location stats
        if not location_stats.empty:
            csv_location = location_stats.to_csv().encode('utf-8')
            st.download_button(
                label="📥 Download Location Stats (CSV)",
                data=csv_location,
                file_name="location_stats.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with col_exp3:
        # Export attendee data
        if not attendee_df.empty:
            csv_attendee = attendee_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Attendee Data (CSV)",
                data=csv_attendee,
                file_name="attendee_data.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    # ── Footer ──
    st.markdown("---")
    st.caption(
        f"""
        **Enterprise Seminar Analytics v2.0** | Last Updated: {st.session_state.get('last_refresh', 'N/A')}
        
        📊 Built with Streamlit & Plotly | Data processed securely in-memory
        """
    )

# ── Entry Point ──
if __name__ == "__main__":
    main()
