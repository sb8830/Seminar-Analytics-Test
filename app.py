"""
Enterprise Seminar Analytics Dashboard
======================================
A comprehensive analytics solution for offline seminar performance tracking,
featuring advanced analytics, interactive visualizations, and enterprise-grade features.

Author: Analytics Team
Version: 2.0.0
Last Updated: 2025-06-15
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any
import io
import logging
from functools import wraps
import time

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
    
    /* Metric Cards */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        transition: transform 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .metric-card.success {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    .metric-card.warning {
        background: linear-gradient(135deg, #f2994a 0%, #f2c94c 100%);
    }
    .metric-card.danger {
        background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
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
    
    /* Cards */
    .custom-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border: 1px solid #e8e8e8;
    }
    
    /* Sidebar */
    .sidebar-section {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 15px;
    }
    
    /* DataFrame Styling */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 12px 24px;
        border-radius: 8px;
        font-weight: 500;
    }
    
    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .animate-fade {
        animation: fadeIn 0.5s ease-out;
    }
    
    /* Status Badges */
    .status-badge {
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .status-active { background: #d4edda; color: #155724; }
    .status-inactive { background: #f8d7da; color: #721c24; }
    .status-pending { background: #fff3cd; color: #856404; }
</style>
""", unsafe_allow_html=True)

# ── Session State Management ──
def init_session_state():
    """Initialize session state variables for enterprise features."""
    defaults = {
        'data_loaded': False,
        'last_refresh': None,
        'view_mode': 'dashboard',
        'selected_seminar': None,
        'date_range': None,
        'expanded_sections': [],
        'custom_filters': {},
        'comparison_mode': False,
        'comparison_period': None
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# ── Performance Decorators ──
def cache_with_logging(ttl: int = 3600):
    """Decorator for caching with logging."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            cache_key = f"{func.__name__}_{str(args)}_{str(kwargs)}"
            if cache_key in st.session_state:
                return st.session_state[cache_key]
            result = func(*args, **kwargs)
            st.session_state[cache_key] = result
            return result
        return wrapper
    return decorator

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
                # Standardize column names
                df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace(r'[^\w]', '_', regex=True)
                
                # Check for student identifier columns
                student_cols = ['student_name', 'studentname', 'student_name', 'name']
                if any(col in df.columns for col in student_cols):
                    df['source_sheet'] = name
                    frames.append(df)
            
            if frames:
                result = pd.concat(frames, ignore_index=True)
                # Clean data
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
        # Standardize column names
        col_mapping = {
            'student_name': 'student_name',
            'studentname': 'student_name',
            'name': 'student_name',
            'phone_no': 'phone',
            'contact': 'phone',
            'email_id': 'email',
            'service': 'service_name',
            'course': 'service_name',
            'batch': 'batch_date',
            'date': 'batch_date',
            'payment': 'payment_received',
            'amount_paid': 'payment_received',
            'gst_amount': 'total_gst',
            'total': 'total_amount',
            'due_amount': 'total_due',
            'balance': 'total_due',
            'student_status': 'status',
            'enrollment_status': 'status'
        }
        
        df = df.rename(columns={k: v for k, v in col_mapping.items() if k in df.columns})
        
        # Convert numeric columns
        numeric_cols = ['payment_received', 'total_gst', 'total_amount', 'total_due', 'total_additional_charges']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', '').str.replace('₹', '').str.replace(' ', ''),
                    errors='coerce'
                ).fillna(0)
        
        # Clean text columns
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
            
            # Create column mapping based on patterns
            col_map = {}
            cols_lower = {c: c.strip().lower().replace('\n', ' ').replace('\r', ' ') for c in df.columns}
            
            # Pattern matching for columns
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
            
            # Rename columns
            renamed = df.rename(columns={v: k for k, v in col_map.items() if v in df.columns})
            
            # Filter valid rows
            if 'sr_no' in renamed.columns:
                renamed = renamed[pd.to_numeric(renamed['sr_no'], errors='coerce').notna()]
                renamed['sr_no'] = renamed['sr_no'].astype(int)
            
            # Convert numeric columns
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
            
            # Parse dates
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
        
        # Basic counts
        kpis['total_seminars'] = len(df)
        
        # Attendance metrics
        if 'total_attended' in df.columns:
            kpis['total_attended'] = int(df['total_attended'].sum())
            kpis['avg_attendance'] = round(df['total_attended'].mean(), 1)
            kpis['max_attendance'] = int(df['total_attended'].max())
        
        # Financial metrics
        if 'actual_revenue' in df.columns:
            kpis['total_revenue'] = float(df['actual_revenue'].sum())
            kpis['avg_revenue'] = round(df['actual_revenue'].mean(), 2)
            kpis['max_revenue'] = float(df['actual_revenue'].max())
        
        if 'actual_expenses' in df.columns:
            kpis['total_expenses'] = float(df['actual_expenses'].sum())
            kpis['avg_expense'] = round(df['actual_expenses'].mean(), 2)
        
        # Profitability
        if 'surplus_deficit' in df.columns:
            kpis['total_profit'] = float(df['surplus_deficit'].sum())
            kpis['profitable_seminars'] = int((df['surplus_deficit'] > 0).sum())
            kpis['loss_seminars'] = int((df['surplus_deficit'] < 0).sum())
            kpis['profit_margin'] = round(
                (kpis['total_profit'] / kpis['total_revenue'] * 100) if kpis['total_revenue'] > 0 else 0, 2
            )
        
        # Conversion metrics
        if 'total_seat_booked' in df.columns and 'total_attended' in df.columns:
            total_booked = df['total_seat_booked'].sum()
            total_attended = df['total_attended'].sum()
            kpis['total_booked'] = int(total_booked)
            kpis['conversion_rate'] = round((total_attended / total_booked * 100) if total_booked > 0 else 0, 2)
        
        # Target achievement
        if 'targeted' in df.columns and 'total_attended' in df.columns:
            kpis['target_achievement'] = round(
                (df['total_attended'].sum() / df['targeted'].sum() * 100) if df['targeted'].sum() > 0 else 0, 2
            )
        
        return kpis
    
    @staticmethod
    def calculate_trends(df: pd.DataFrame, period: str = 'month') -> pd.DataFrame:
        """Calculate time-based trends."""
        if df.empty or 'seminar_date' not in df.columns:
            return pd.DataFrame()
        
        if period == 'month':
            grouped = df.groupby(df['seminar_date'].dt.to_period('M'))
        elif period == 'quarter':
            grouped = df.groupby(df['seminar_date'].dt.to_period('Q'))
        else:
            grouped = df.groupby(df['seminar_date'].dt.date)
        
        trends = grouped.agg({
            'total_attended': 'sum',
            'actual_revenue': 'sum',
            'actual_expenses': 'sum',
            'surplus_deficit': 'sum',
            'sr_no': 'count'
        }).rename(columns={'sr_no': 'seminars'})
        
        return trends.reset_index()
    
    @staticmethod
    def get_top_performers(df: pd.DataFrame, metric: str = 'surplus_deficit', n: int = 5) -> pd.DataFrame:
        """Get top performing seminars by metric."""
        if df.empty or metric not in df.columns:
            return pd.DataFrame()
        
        return df.nlargest(n, metric)[['location', 'trainer', 'seminar_date', metric, 'total_attended']]
    
    @staticmethod
    def calculate_location_stats(df: pd.DataFrame
