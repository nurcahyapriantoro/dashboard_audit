import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import os
import utils
import time
import sys

# Get resource path for bundled app
base_path = utils.get_base_path() if hasattr(utils, 'get_base_path') else os.path.dirname(os.path.abspath(__file__))
LOGO_PERTAMINA_PATH = os.path.join(base_path, "logo_pertamina.png")

st.set_page_config(
    page_title="Dashboard Audit - Pertamina Corporate",
    page_icon=LOGO_PERTAMINA_PATH if os.path.exists(LOGO_PERTAMINA_PATH) else "🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

def inject_custom_css():
    
    if "theme" not in st.session_state:
        st.session_state.theme = "light"

    is_dark = st.session_state.theme == "dark"

    if is_dark:
        bg_color = "#0E1117"
        bg_gradient = "linear-gradient(180deg, #0E1117 0%, #161B22 100%)"
        text_primary = "#FFFFFF"
        text_secondary = "#A3AED0"
        card_bg = "#262730"
        sidebar_bg = "#161B22"
        border_color = "#30363D"
        table_bg = "#262730"
    else:
        bg_color = "#E6F3FF"
        bg_gradient = "linear-gradient(180deg, #D6E4F0 0%, #FFFFFF 100%)"
        text_primary = "#2B3674"
        text_secondary = "#A3AED0"
        card_bg = "#FFFFFF"
        sidebar_bg = "#FFFFFF"
        border_color = "#E0E5F2"
        table_bg = "#FFFFFF"

    st.markdown(f"""
    <style>
    :root {{
        --primary-blue: #005DAA;
        --primary-red: #ED1C24;
        --primary-green: #69BE28;
        --bg-light: {bg_color};          
        --text-dark: {text_primary};         
        --text-grey: {text_secondary};         
        --white: {card_bg};
        --card-shadow: 0px 18px 40px rgba(112, 144, 176, 0.12);
        --border-color: {border_color};
    }}
    
    .stApp {{
        background: {bg_gradient};
        color: var(--text-dark);
        font-family: 'Segoe UI', sans-serif;
    }}
    
    h1, h2, h3 {{
        color: var(--text-dark);
        font-weight: 700;
    }}
    
    [data-testid="stSidebar"] {{
        background: {sidebar_bg};
        border-right: 1px solid var(--border-color);
        height: 100vh;
        overflow: hidden !important;
        box-shadow: 2px 0 10px rgba(0,0,0,0.05);
    }}

    [data-testid="stSidebar"] > div:first-child {{
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }}
    
    [data-testid="stSidebarUserContent"] {{
        padding-top: 0px !important;
        display: flex;
        flex-direction: column;
        gap: 0px; 
    }}

    [data-testid="stSidebar"] * {{
        color: var(--text-dark) !important;
    }}

    [data-testid="stSidebar"] hr {{
        margin: 1.5rem 0 !important;
        border-top: 1px solid var(--border-color) !important;
        opacity: 0.6;
    }}
    
    [data-testid="stSidebar"] .stRadio > div {{
        gap: 8px;
    }}

    [data-testid="stSidebar"] .stRadio label {{
        color: var(--text-dark) !important;
        cursor: pointer;
        padding: 12px 16px !important;
        border-radius: 12px;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        margin-bottom: 6px !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        border: 1px solid transparent;
        display: flex !important;
        align-items: center !important;
    }}

    [data-testid="stSidebar"] .stRadio label:hover {{
        background-color: rgba(0, 93, 170, 0.05);
        border-color: rgba(0, 93, 170, 0.2);
        transform: translateX(3px);
    }}

    [data-testid="stSidebar"] [data-baseweb="radio"] div[aria-checked="true"] {{
        background: linear-gradient(135deg, var(--primary-blue) 0%, #004B8D 100%) !important;
        box-shadow: 0 4px 15px rgba(0, 93, 170, 0.3);
    }}

    [data-testid="stSidebar"] [data-baseweb="radio"] div[aria-checked="true"] p {{
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }}
    
    .main-header {{
        background: transparent;
        color: var(--text-dark);
        padding: 0px 0px 20px 0px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid var(--border-color);
        margin-bottom: 20px;
    }}
    
    .breadcrumb {{
        font-size: 12px;
        color: var(--text-grey);
        margin-bottom: 5px;
    }}
    
    .header-title {{
        font-size: 24px;
        font-weight: 700;
        margin: 0;
        color: var(--text-dark);
    }}
    
    .kpi-card {{
        background: var(--white);
        border-radius: 20px;
        padding: 20px;
        box-shadow: var(--card-shadow);
        transition: transform 0.2s;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        border: 1px solid var(--border-color);
    }}
    
    .kpi-card:hover {{
        transform: translateY(-5px);
        cursor: pointer;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
    }}
    
    .kpi-icon {{
        width: 45px;
        height: 45px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        margin-bottom: 10px;
    }}
    
    .kpi-value {{
        font-size: 32px;
        font-weight: 700;
        color: var(--text-dark);
        margin: 5px 0;
    }}
    
    .kpi-label {{
        color: var(--text-grey);
        font-size: 14px;
    }}
    
    .trend-badge {{
        font-size: 12px;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 5px;
    }}
    
    .trend-up {{ color: var(--primary-green); }}
    .trend-down {{ color: var(--primary-red); }}
    
    .stButton > button {{
        border-radius: 12px;
        font-weight: 600;
        border: none;
        padding: 0.5rem 1rem;
        transition: all 0.2s;
    }}
    
    .stButton > button[kind="primary"] {{
        background-color: var(--primary-blue);
        color: #FFFFFF;
    }}

    .stButton > button:hover {{
        opacity: 0.9;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 20px;
        background-color: transparent;
        padding-bottom: 10px;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        background: transparent;
        border: none;
        color: var(--text-grey);
        font-weight: 600;
        padding: 10px 0;
    }}
    
    .stTabs [aria-selected="true"] {{
        color: var(--primary-blue) !important;
        border-bottom: 3px solid var(--primary-blue) !important;
    }}
    
    [data-testid="stDataFrame"] {{
        background: var(--white);
        padding: 15px;
        border-radius: 20px;
        box-shadow: var(--card-shadow);
        border: 1px solid var(--border-color);
        width: 100%;
    }}
    
    .stDataFrame > div {{
        max-height: 80vh;
        overflow-y: auto;
    }}

    .status-badge {{
        padding: 4px 12px;
        border-radius: 50px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        text-align: center;
    }}
    
    </style>
    """, unsafe_allow_html=True)

def _get_initial_menu(menu_options):
    query_menu = st.query_params.get("menu", None)
    if isinstance(query_menu, list):
        query_menu = query_menu[0] if query_menu else None

    if "last_menu" not in st.session_state:
        st.session_state["last_menu"] = query_menu if query_menu in menu_options else menu_options[0]

    if st.session_state["last_menu"] not in menu_options:
        st.session_state["last_menu"] = menu_options[0]

    return st.session_state["last_menu"]

def render_sidebar():
    menu_options = [
        "📊 Dashboard Audit",
        "doc_control", # Use a key internally if needed, but display text is strings
        "📈 Analytics & Insights",
        "📧 Email Reminders",
        "⚙️ Master Data"
    ]
    # Display names
    display_options = [
        "📊 Dashboard Audit",
        "📂 Document Control",
        "📈 Analytics & Insights",
        "📧 Email Reminders",
        "⚙️ Master Data"
    ]
    
    initial_menu = _get_initial_menu(display_options)

    with st.sidebar:
        col_logo, col_title = st.columns([0.8, 3.2])
        with col_logo:
             if os.path.exists(LOGO_PERTAMINA_PATH):
                st.image(LOGO_PERTAMINA_PATH, width=40)
        with col_title:
             st.markdown("""
            <div style='display: flex; flex-direction: column; justify-content: center; height: 40px;'>
                 <h1 style='color: var(--text-dark); font-size: 18px; margin: 0; line-height: 1.2;'>DASHBOARD AUDIT</h1>
                 <p style='color: var(--text-grey); font-size: 10px; margin: 0; line-height: 1;'>Pertamina Corporate</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.write("") # Spacer
        
        menu = st.radio(
            "",
            options=display_options,
            index=display_options.index(initial_menu),
            key="sidebar_menu",
            label_visibility="collapsed"
        )


        st.session_state["last_menu"] = menu
        st.query_params["menu"] = menu
        
        st.markdown("<div class='sidebar-divider'></div>", unsafe_allow_html=True)
        
        def toggle_theme():
            st.session_state.theme = "dark" if st.session_state.theme_toggle else "light"

        if "theme_toggle" not in st.session_state:
            st.session_state.theme_toggle = st.session_state.get("theme") == "dark"

        st.toggle("Dark Mode", key="theme_toggle", on_change=toggle_theme)

        st.markdown("<div class='sidebar-divider'></div>", unsafe_allow_html=True)
        
        stats = utils.get_document_statistics()
        st.markdown("<p style='color: var(--text-grey); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px;'>⚡ Quick Stats</p>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="quick-stats-card">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <span style="font-size: 11px; color: #666;">Total Documents</span>
                <span style="font-size: 14px; font-weight: bold; color: var(--primary-blue);">{stats['total']}</span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <span style="font-size: 11px; color: #666;">Completion</span>
                <span style="font-size: 14px; font-weight: bold; color: var(--primary-green);">{stats['done']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<div class='sidebar-divider'></div>", unsafe_allow_html=True)
        
        st.markdown("""
        <div class='footer' style='margin-top: 10px;'>
            <p style='color: var(--text-grey); font-size: 10px; margin-bottom: 2px;'>© 2026 Dashboard Audit</p>
            <p style='color: var(--text-grey); font-size: 9px;'>Corporate Edition v1.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        return menu

@st.cache_data(ttl=60)
def load_data_cached():
    return utils.load_documents()

def clear_cache():
    load_data_cached.clear()

def render_dashboard():
    col_header, col_search = st.columns([3, 2])
    with col_header:
        st.markdown(f"""
        <div style="padding-top:10px;">
            <p class='breadcrumb'>Home > Dashboard Audit</p>
            <h1 class='header-title'>Executive Overview</h1>
            <p style='color: var(--text-grey); font-size: 14px; margin-top:5px;'>Real-time monitoring & insights</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_search:
        st.markdown('<div style="height: 25px;"></div>', unsafe_allow_html=True)
        global_search = st.text_input("🔍 Global Search", placeholder="Search document ID, unit, or status...", label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Use cached data
    df = load_data_cached()
    
    stats = utils.get_document_statistics() # This should ideally use the df passed to it or be cached too
    # Recalculate stats from df to use cache
    if not df.empty:
        stats = {
            'total': len(df),
            'outstanding': len(df[df['Document Status'] == 'Outstanding']),
            'need_review': len(df[df['Document Status'] == 'Need to Review']),
            'done': len(df[df['Document Status'] == 'Done']),
            'overdue': len(df[df['Days To Go'] < 0]) if 'Days To Go' in df.columns else 0,
            'completion_rate': round((len(df[df['Document Status'] == 'Done']) / len(df) * 100), 1)
        }
    else:
        stats = {'total': 0, 'outstanding': 0, 'need_review': 0, 'done': 0, 'overdue': 0, 'completion_rate': 0}

    if global_search and not df.empty:
        mask = df.astype(str).apply(lambda x: x.str.contains(global_search, case=False)).any(axis=1)
        search_results = df[mask]
        if not search_results.empty:
            st.info(f"🔍 Found {len(search_results)} documents matching '{global_search}'")
            st.dataframe(
                search_results,
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "Document Status": st.column_config.SelectboxColumn(
                        "Status",
                        options=["Outstanding", "Need to Review", "Done"],
                        required=True,
                    )
                }
            )
            st.markdown("---")
        else:
            st.warning(f"No documents found matching '{global_search}'")
    
    with st.expander("⚙️ **Dashboard Filters**", expanded=False):
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            filter_fungsi = st.multiselect("Filter by Divisi", options=utils.get_fungsi_list())
        with f_col2:
            filter_status = st.multiselect("Filter by Status", options=utils.DOC_STATUS_OPTIONS)
        
        filtered_df = df.copy()
        if filter_fungsi:
            filtered_df = filtered_df[filtered_df['Fungsi'].isin(filter_fungsi)]
        if filter_status:
            filtered_df = filtered_df[filtered_df['Document Status'].isin(filter_status)]
            
        if not filtered_df.equals(df):
            stats = {
                'total': len(filtered_df),
                'outstanding': len(filtered_df[filtered_df['Document Status'] == 'Outstanding']),
                'need_review': len(filtered_df[filtered_df['Document Status'] == 'Need to Review']),
                'done': len(filtered_df[filtered_df['Document Status'] == 'Done']),
                'overdue': len(filtered_df[df['Days To Go'] < 0]) if 'Days To Go' in filtered_df.columns else 0,
                 'completion_rate': round((len(filtered_df[filtered_df['Document Status'] == 'Done']) / len(filtered_df) * 100) if len(filtered_df) > 0 else 0, 1)
            }

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="kpi-card">
            <div>
                <div class="kpi-icon" style="background-color: #E6F7FF; color: #0091FF;">📄</div>
                <div class="kpi-label">Total Documents</div>
                <div class="kpi-value">{stats['total']}</div>
            </div>
            <div class="trend-badge trend-up">
                <span>↑ 12%</span> <span style="color: var(--text-grey); font-weight:400;">vs last month</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi-card" style="border-left: 4px solid var(--primary-red);">
            <div>
                <div class="kpi-icon" style="background-color: #FFF1F0; color: #F5222D;">🔥</div>
                <div class="kpi-label">Outstanding</div>
                <div class="kpi-value">{stats['outstanding']}</div>
            </div>
            <div class="trend-badge trend-down">
                <span>↓ 5%</span> <span style="color: var(--text-grey); font-weight:400;">getting better</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi-card" style="border-left: 4px solid #FAAD14;">
            <div>
                <div class="kpi-icon" style="background-color: #FFFBE6; color: #FAAD14;">⚖️</div>
                <div class="kpi-label">Need Review</div>
                <div class="kpi-value">{stats['need_review']}</div>
            </div>
            <div class="trend-badge" style="color: #FAAD14;">
                <span>• Stable</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="kpi-card" style="border-left: 4px solid var(--primary-green);">
            <div>
                <div class="kpi-icon" style="background-color: #F6FFED; color: #52C41A;">✅</div>
                <div class="kpi-label">Completed</div>
                <div class="kpi-value">{stats['done']}</div>
            </div>
            <div class="trend-badge trend-up">
                <span>↑ 8%</span> <span style="color: var(--text-grey); font-weight:400;">vs last month</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.container():
        st.markdown("### 📊 Analytics & Performance")
        
        if not filtered_df.empty:
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("##### 🏢 Documents by Unit/Divisi")
                status_fungsi = filtered_df.groupby(['Fungsi', 'Document Status']).size().reset_index(name='Count')
                
                if not status_fungsi.empty:
                    fig_bar = px.bar(
                        status_fungsi,
                        x='Fungsi',
                        y='Count',
                        color='Document Status',
                        barmode='group',
                        color_discrete_map={
                            'Outstanding': '#ED1C24',
                            'Need to Review': '#FFC107',
                            'Done': '#69BE28'
                        }
                    )
                    fig_bar.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Segoe UI"),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        margin=dict(l=20, r=20, t=20, b=20),
                        height=350,
                        xaxis_tickangle=-45
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("No data available for selected filters.")
            
            with chart_col2:
                st.markdown("##### 🍩 Completion Rate")
                status_counts = filtered_df['Document Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                
                if not status_counts.empty:
                    fig_donut = px.pie(
                        status_counts,
                        values='Count',
                        names='Status',
                        hole=0.7,
                        color='Status',
                        color_discrete_map={
                            'Outstanding': '#ED1C24',
                            'Need to Review': '#FFC107',
                            'Done': '#69BE28'
                        }
                    )
                    fig_donut.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Segoe UI"),
                        showlegend=True,
                        legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5),
                        margin=dict(l=20, r=20, t=20, b=20),
                        height=350,
                        annotations=[dict(
                            text=f'{stats["completion_rate"]}%',
                            x=0.5, y=0.5,
                            font_size=28,
                            showarrow=False,
                            font=dict(color='#005DAA', family="Segoe UI", weight='bold')
                        )]
                    )
                    st.plotly_chart(fig_donut, use_container_width=True)
                else:
                    st.info("No data available.")

    st.markdown("### 🕒 Recent Updates")
    if not filtered_df.empty:
        st.dataframe(
            filtered_df.sort_values(by="Doc. ID Number", ascending=False).head(5),
            use_container_width=True,
            hide_index=True,
            column_order=["Doc. ID Number", "Document Description", "Document Status", "Doc. Submission Deadline", "Days To Go", "Fungsi"],
            column_config={
                 "Document Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=utils.DOC_STATUS_OPTIONS,
                    width="medium",
                    required=True,
                ),
                 "Days To Go": st.column_config.ProgressColumn(
                    "Urgency",
                    format="%d days",
                    min_value=-10,
                    max_value=30,
                ),
            }
        )
    else:
        st.info("No data to display.")

def render_document_control():
    
    col_header, _ = st.columns([3, 1])
    with col_header:
        st.markdown(f"""
        <div style="padding-top:10px;">
            <p class='breadcrumb'>Home > Document Control Center</p>
            <h1 class='header-title'>Document Management</h1>
            <p style='color: var(--text-grey); font-size: 14px; margin-top:5px;'>Create, track, and manage audit documents.</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.expander("📥 Export & Reports", expanded=False):
        st.markdown("Download comprehensive report including executive summary and detailed data.")
        
        report_type = st.radio("Filter By:", ["All Data", "Yearly", "Monthly", "Date Range"], horizontal=True)
        
        df_export = utils.load_documents()
        
        # Pastikan kolom Doc. Request Date bertipe datetime
        try:
            df_export['Doc. Request Date'] = pd.to_datetime(df_export['Doc. Request Date'])
        except:
             pass

        if report_type == "All Data":
            pass # No filtering

        elif report_type == "Yearly":
            if 'Doc. Request Date' in df_export.columns:
                years = sorted(df_export['Doc. Request Date'].dt.year.dropna().astype(int).unique(), reverse=True)
                if not years:
                    years = [datetime.now().year]
                selected_year = st.selectbox("Select Year", options=years)
                # Filter by year
                df_export = df_export[df_export['Doc. Request Date'].dt.year == selected_year]
            else:
                 st.error("Column 'Doc. Request Date' missing")
            
        elif report_type == "Monthly":
            col_year, col_month = st.columns(2)
            if 'Doc. Request Date' in df_export.columns:
                years = sorted(df_export['Doc. Request Date'].dt.year.dropna().astype(int).unique(), reverse=True)
                if not years:
                    years = [datetime.now().year]
                with col_year:
                    selected_year = st.selectbox("Select Year", options=years, key="monthly_year")
                with col_month:
                    months = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June", 
                            7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"}
                    selected_month_name = st.selectbox("Select Month", options=list(months.values()))
                    selected_month = [k for k, v in months.items() if v == selected_month_name][0]
                
                # Filter by Month and Year
                df_export = df_export[
                    (df_export['Doc. Request Date'].dt.year == selected_year) & 
                    (df_export['Doc. Request Date'].dt.month == selected_month)
                ]
            else:
                 st.error("Column 'Doc. Request Date' missing")
            
        elif report_type == "Date Range":
            col_start, col_end = st.columns(2)
            with col_start:
                start_date = st.date_input("Start Date", value=date.today().replace(day=1))
            with col_end:
                end_date = st.date_input("End Date", value=date.today())
            
            # Filter range
            if 'Doc. Request Date' in df_export.columns:
                 # Ensure comparison is date to date
                 mask = (df_export['Doc. Request Date'].dt.date >= start_date) & (df_export['Doc. Request Date'].dt.date <= end_date)
                 df_export = df_export[mask]
        
        st.caption(f"Generating report for {len(df_export)} documents...")
        
        if not df_export.empty:
            report_data = utils.generate_spectacular_report(df_export)
            st.download_button(
                label=f"Download Report ({len(df_export)} docs)",
                data=report_data,
                file_name=f"Audit_Report_{report_type.replace(' ', '')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_report_btn",
                help="Generate and download a professional Excel report"
            )
        else:
            st.warning("⚠️ No data found for the selected period.")

    tab_view, tab_add, tab_update, tab_audit, tab_delete, tab_upload = st.tabs([
        "📋 View & Filter", "➕ Add Document", "✏️ Update Status", "🕒 Audit Trail", "🗑️ Delete", "📤 Import Excel"
    ])
    
    with tab_view:
        
        df = utils.load_documents()
        
        if not df.empty:
            col_search, col_filter = st.columns([2, 1])
            with col_search:
                search_term = st.text_input("🔍 Quick Search", placeholder="Type to search...", label_visibility="collapsed")
            with col_filter:
                status_filter = st.multiselect("Filter Status", utils.DOC_STATUS_OPTIONS, default=None, placeholder="All Status", label_visibility="collapsed")
            
            filtered_df = df.copy()
            if search_term:
                mask = filtered_df.astype(str).apply(lambda row: search_term.lower() in row.str.lower().values, axis=1)
                filtered_df = filtered_df[mask]
            if status_filter:
                filtered_df = filtered_df[filtered_df['Document Status'].isin(status_filter)]
            
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                height=500,
                column_config={
                    "Doc. ID Number": st.column_config.TextColumn("Doc. ID", width="small", help="Unique Identifier"),
                    "Document Description": st.column_config.TextColumn("Description", width="large"),
                    "Fungsi": st.column_config.TextColumn("Division", width="medium"),
                    "Doc. Request Date": st.column_config.DateColumn("Requested", format="DD MMM YYYY"),
                    "Doc. Submission Deadline": st.column_config.DateColumn("Deadline", format="DD MMM YYYY"),
                    "Days To Go": st.column_config.ProgressColumn(
                        "Time Remaining", 
                        help="Days until deadline",
                        format="%d days",
                        min_value=-10,
                        max_value=30,
                    ),
                    "Document Status": st.column_config.SelectboxColumn(
                        "Status", 
                        options=utils.DOC_STATUS_OPTIONS,
                        width="medium",
                        required=True
                    ),
                    "Remarks": st.column_config.LinkColumn("Form Link", display_text="Open Form"),
                    "Status Reminder": st.column_config.TextColumn("Last Reminder", width="small")
                }
            )
            st.caption(f"Showing {len(filtered_df)} documents")
        else:
            st.info("📭 No documents found. Add a new document to get started.")
    
    with tab_add:
        st.markdown("##### ➕ Create New Document")
        
        with st.form("add_document_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Document Details**")
                new_doc_id = st.text_input(
                    "Doc. ID Number *",
                    placeholder="e.g., Fin-2026-001",
                    help="Auto-generated ID is recommended"
                )
                if not new_doc_id:
                    suggested_id = utils.generate_doc_id("AUD")
                    st.caption(f"💡 Suggested ID: `{suggested_id}`")
                
                new_desc = st.text_area(
                    "Description *",
                    placeholder="Enter document description...",
                    height=100
                )
                
                fungsi_list = utils.get_fungsi_list()
                new_fungsi = st.selectbox("Division/Unit *", options=fungsi_list)
            
            with col2:
                st.markdown("**Timeline & Contact**")
                new_request_date = st.date_input("Request Date *", value=date.today())
                new_deadline = st.date_input("Submission Deadline *", value=date.today() + timedelta(days=7))
                new_email_to = st.text_input("Auditee Email (To) *", placeholder="auditee@pertamina.com")
                new_email_cc = st.text_input("CC Recipients", placeholder="manager@pertamina.com")
            
            st.markdown("---")
            col_stat, col_link = st.columns(2)
            with col_stat:
                 new_status = st.selectbox("Initial Status *", options=["Outstanding", "Need to Review", "Done"])
            with col_link:
                 new_remarks = st.text_input("Remarks / Form Link", placeholder="Optional: Add notes or form link here...")
            
            submitted = st.form_submit_button("💾 Save Document", type="primary", use_container_width=True)
            
            if submitted:
                if not new_desc or not new_email_to:
                    st.error("❌ Please fill in all required fields marked with (*)")
                else:
                    final_id = new_doc_id if new_doc_id else suggested_id
                    
                    doc_data = {
                        "Doc. ID Number": final_id,
                        "Document Description": new_desc,
                        "Fungsi": new_fungsi,
                        "Doc. Request Date": new_request_date,
                        "Doc. Submission Deadline": new_deadline,
                        "Email - Auditee1": new_email_to,
                        "Recepient - cc": new_email_cc,
                        "Document Status": new_status,
                        "Remarks": new_remarks,
                        "Status Reminder": None
                    }
                    success, message = utils.add_document(doc_data)
                    if success:
                        st.success(f"✅ Document {final_id} created successfully!")
                        utils.log_audit_event("Current User", "Create", final_id, "Document created manually")
                    else:
                        st.error(message)
    
    with tab_update:
        st.markdown("##### ✏️ Update Document Details")
        
        df = utils.load_documents()
        
        if not df.empty:
            doc_ids = df['Doc. ID Number'].dropna().tolist()
            selected_doc = st.selectbox("Select Document ID:", options=doc_ids)
            
            if selected_doc:
                doc_row = df[df['Doc. ID Number'] == selected_doc].iloc[0]
                
                with st.form("update_document_form"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("**📝 General Information**")
                        
                        current_desc = doc_row['Document Description'] if pd.notna(doc_row['Document Description']) else ""
                        update_desc = st.text_area("Description", value=current_desc, height=100)
                        
                        fungsi_list = utils.get_fungsi_list()
                        current_fungsi = doc_row['Fungsi'] if pd.notna(doc_row['Fungsi']) and doc_row['Fungsi'] in fungsi_list else (fungsi_list[0] if fungsi_list else None)
                        update_fungsi = st.selectbox("Division/Unit", options=fungsi_list, index=fungsi_list.index(current_fungsi) if current_fungsi else 0)

                    with col2:
                        st.markdown("**📅 Timeline & Contact**")
                        
                        try:
                            default_req_date = pd.to_datetime(doc_row['Doc. Request Date']).date()
                        except:
                            default_req_date = date.today()
                            
                        try:
                            default_deadline = pd.to_datetime(doc_row['Doc. Submission Deadline']).date()
                        except:
                            default_deadline = date.today()

                        update_req_date = st.date_input("Request Date", value=default_req_date)
                        update_deadline = st.date_input("Submission Deadline", value=default_deadline)
                        
                        current_email_to = doc_row['Email - Auditee1'] if pd.notna(doc_row['Email - Auditee1']) else ""
                        update_email_to = st.text_input("Auditee Email (To)", value=current_email_to)
                        
                        current_email_cc = doc_row['Recepient - cc'] if pd.notna(doc_row['Recepient - cc']) else ""
                        update_email_cc = st.text_input("CC Recipients", value=current_email_cc)

                    st.markdown("---")
                    col3, col4 = st.columns(2)
                    
                    with col3:
                        current_status = doc_row['Document Status']
                        status_index = utils.DOC_STATUS_OPTIONS.index(current_status) if current_status in utils.DOC_STATUS_OPTIONS else 0
                        update_status = st.selectbox("Status", options=utils.DOC_STATUS_OPTIONS, index=status_index)
                    
                    with col4:
                        current_remarks = doc_row['Remarks'] if pd.notna(doc_row['Remarks']) else ""
                        update_remarks = st.text_input("Form Link / Remarks", value=current_remarks)
                    
                    update_submitted = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                    
                    if update_submitted:
                        updates = {
                            "Document Description": update_desc,
                            "Fungsi": update_fungsi,
                            "Doc. Request Date": update_req_date,
                            "Doc. Submission Deadline": update_deadline,
                            "Email - Auditee1": update_email_to,
                            "Recepient - cc": update_email_cc,
                            "Document Status": update_status,
                            "Remarks": update_remarks
                        }
                        success, message = utils.update_document(selected_doc, updates)
                        if success:
                            st.success(message)
                            utils.log_audit_event("Current User", "Update", selected_doc, f"Updated details for {selected_doc}")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(message)
        else:
            st.info("📭 No documents available to update.")
            
    with tab_audit:
        st.markdown("##### 🕒 Audit History Log")
        audit_df = utils.load_audit_logs()
        
        if not audit_df.empty:
             st.dataframe(
                audit_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Timestamp": st.column_config.DatetimeColumn("Time", format="DD/MM/YYYY HH:mm:ss"),
                    "User": st.column_config.TextColumn("User", width="small"),
                    "Action": st.column_config.TextColumn("Action", width="small"),
                    "Doc ID": st.column_config.TextColumn("Doc ID", width="small"),
                    "Details": st.column_config.TextColumn("Details", width="large"),
                }
            )
        else:
            st.info("No audit logs found.")

    with tab_delete:
        st.markdown("##### 🗑️ Delete Document")
        st.warning("⚠️ **Warning:** This action cannot be undone.")
        
        df = utils.load_documents()
        
        if not df.empty:
            doc_ids = df['Doc. ID Number'].dropna().tolist()
            delete_doc = st.selectbox("Select Document to Delete:", options=doc_ids, key="delete_select")
            
            if delete_doc:
                doc_row = df[df['Doc. ID Number'] == delete_doc].iloc[0]
                
                st.markdown(f"""
                <div style="background: #FFF1F0; padding: 10px; border-radius: 8px; border: 1px solid #FFA39E;">
                    <strong>{delete_doc}</strong><br>
                    {doc_row['Document Description']}
                </div>
                """, unsafe_allow_html=True)
                
                confirm_delete = st.checkbox("✅ I confirm I want to delete this document")
                
                if st.button("🗑️ Delete Permanently", type="primary", disabled=not confirm_delete):
                    success, message = utils.delete_document(delete_doc)
                    if success:
                        st.success(message)
                        utils.log_audit_event("Current User", "Delete", delete_doc, "Document deleted")
                        st.rerun()
                    else:
                        st.error(message)
        else:
            st.info("📭 No documents to delete.")
            
    with tab_upload:
        st.markdown("##### 📤 Import Data from Excel")
        st.info("Upload file Excel untuk menambahkan banyak dokumen sekaligus.")
        
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
        
        if uploaded_file is not None:
            # Preview first 5 rows
            try:
                df_preview = pd.read_excel(uploaded_file, nrows=5)
                st.write("Preview Data:")
                st.dataframe(df_preview)
                
                if st.button("🚀 Process Import", type="primary"):
                    success, msg, count = utils.process_bulk_upload(uploaded_file)
                    if success:
                        st.success(msg)
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error(msg)
            except Exception as e:
                st.error(f"Error reading file: {e}")
        
        with st.expander("ℹ️ Format Kolom Excel yang Didukung"):
            st.markdown("""
            Pastikan header kolom di Excel Anda mengandung nama berikut (case insensitive):
            - **Document Description** (Wajib): Deskripsi dokumen
            - **Fungsi** / **Divisi** (Wajib): Nama unit/divisi
            - **Doc. Submission Deadline** (Wajib): Tanggal batas waktu (Format Date Excel)
            - **Doc. Request Date**: Tanggal permintaan (Opsional, Default: Hari ini)
            - **Email - Auditee1**: Email penerima (Opsional)
            - **Recepient - cc**: Email CC (Opsional)
            - **Document Status**: Status awal (Opsional, Default: Outstanding)
            - **Remarks**: Catatan (Opsional)
            """)

def render_email_automation():
    
    st.markdown("""
    <div class='main-header'>
        <h1>📧 Email Reminders</h1>
        <p>Generate reminder emails via Microsoft Outlook</p>
    </div>
    """, unsafe_allow_html=True)
    
    reminder_df = utils.get_documents_for_reminder()
    
    st.caption("Menampilkan dokumen berstatus Outstanding / Need to Review untuk dikirim reminder.")
    
    if not reminder_df.empty:
        col_left, col_right = st.columns([1, 2])
        
        with col_left:
            st.markdown("##### 📋 Dokumen untuk Reminder")
            st.caption(f"Total: {len(reminder_df)} dokumen")
            
            for idx, row in reminder_df.iterrows():
                days = row.get('Days To Go', 0)
                if days is not None and days < 0:
                    badge_class = "status-outstanding"
                    badge_icon = "🔴"
                elif days is not None and days <= 3:
                    badge_class = "status-review"
                    badge_icon = "🟡"
                else:
                    badge_class = "status-done"
                    badge_icon = "🟢"
                
                with st.container():
                    st.markdown(f"""
                    <div style='background: white; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid {"#DA291C" if days and days < 0 else "#FFC107"};'>
                        <strong>{row['Doc. ID Number']}</strong><br>
                        <small style='color: #666;'>{row['Document Description'][:50]}...</small><br>
                        <span class='status-badge {badge_class}'>{badge_icon} {row['Document Status']}</span>
                    </div>
                    """, unsafe_allow_html=True)
        
        with col_right:
            st.markdown("##### ✉️ Compose Email")
            
            doc_options = reminder_df['Doc. ID Number'].tolist()
            selected_doc_id = st.selectbox(
                "Pilih No. Dokumen untuk Reminder:",
                options=doc_options,
                help="Pilih dokumen yang akan dikirim reminder"
            )
            
            if selected_doc_id:
                doc_data = reminder_df[reminder_df['Doc. ID Number'] == selected_doc_id].iloc[0].to_dict()
                
                st.markdown("---")
                
                with st.form("email_compose_form"):
                    default_to = doc_data.get('Email - Auditee1', '') or ''
                    email_to = st.text_input(
                        "📧 To (Penerima Utama)",
                        value=default_to,
                        help="Email utama penerima"
                    )
                    
                    default_cc = doc_data.get('Recepient - cc', '') or ''
                    email_cc = st.text_input(
                        "📋 CC (Tembusan)",
                        value=default_cc,
                        help="Email CC, pisahkan dengan koma"
                    )
                    
                    default_subject = utils.generate_email_subject(doc_data)
                    email_subject = st.text_input(
                        "📝 Subject",
                        value=default_subject
                    )
                    
                    st.markdown("##### 📄 Template Email (Editable)")
                    default_body = utils.generate_email_body(doc_data)
                    email_body = st.text_area(
                        "Body",
                        value=default_body,
                        height=300,
                        label_visibility="collapsed",
                        help="Template profesional yang dapat diedit sebelum membuka draft email"
                    )
                    
                    show_preview = st.checkbox("👁️ Preview Email", value=False)
                    
                    if show_preview:
                        st.markdown("##### Preview:")
                        st.markdown(f"**To:** {email_to}")
                        st.markdown(f"**CC:** {email_cc}")
                        st.markdown(f"**Subject:** {email_subject}")
                        st.markdown("---")
                        st.markdown(email_body, unsafe_allow_html=True)
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        generate_btn = st.form_submit_button(
                            "📤 Generate Outlook Email",
                            type="primary",
                            use_container_width=True
                        )
                    
                    if generate_btn:
                        if not email_to:
                            st.error("❌ Email penerima (To) wajib diisi!")
                        else:
                            success, message = utils.generate_outlook_email_new(
                                to=email_to,
                                cc=email_cc,
                                subject=email_subject,
                                body_html=email_body,
                                doc_id=selected_doc_id
                            )
                            if success:
                                st.success(message)
                                st.caption("Silakan review isi draft, lalu kirim dari Outlook.")
                            else:
                                st.error(message)
                                st.caption("Tips: set New Outlook sebagai default mail app di Windows agar tombol membuka jendela compose langsung.")
    
    else:
        st.success("🎉 Tidak ada dokumen yang perlu di-remind! Semua dokumen sudah **Done**.")
        
        st.markdown("""
        <div style='text-align: center; padding: 50px;'>
            <h2>✅ All Clear!</h2>
            <p style='color: #666;'>Semua dokumen audit sudah selesai diproses.</p>
        </div>
        """, unsafe_allow_html=True)

def render_master_data():
    
    st.markdown("""
    <div class='main-header'>
        <h1>⚙️ Master Data Management</h1>
        <p>Manage divisions and user profiles</p>
    </div>
    """, unsafe_allow_html=True)
    
    tab_fungsi, tab_users = st.tabs(["🏢 Master Fungsi/Divisi", "👥 User Management"])
    
    with tab_fungsi:
        st.markdown("##### 🏢 Kelola Divisi/Fungsi")
        st.info("📝 Edit langsung di tabel di bawah. Klik tombol **Simpan** setelah melakukan perubahan.")
        
        df_fungsi = utils.load_fungsi()
        
        edited_fungsi = st.data_editor(
            df_fungsi,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Nama Fungsi": st.column_config.TextColumn(
                    "Nama Divisi/Fungsi",
                    help="Nama divisi yang akan tampil di dropdown",
                    width="large",
                    required=True
                )
            },
            hide_index=True
        )
        
        if st.button("💾 Simpan Perubahan Fungsi", type="primary"):
            edited_fungsi = edited_fungsi.dropna(subset=['Nama Fungsi'])
            success, message = utils.save_fungsi(edited_fungsi)
            if success:
                st.success(message)
            else:
                st.error(message)
    
    with tab_users:
        st.markdown("##### 👥 Kelola User")
        st.info("📝 Kelola user yang dapat mengakses sistem (untuk pengembangan fitur login)")
        
        df_users = utils.load_users()
        
        fungsi_list = utils.get_fungsi_list()
        
        edited_users = st.data_editor(
            df_users,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Nama": st.column_config.TextColumn(
                    "Nama Lengkap",
                    help="Nama user",
                    width="medium",
                    required=True
                ),
                "Email": st.column_config.TextColumn(
                    "Email",
                    help="Alamat email user",
                    width="medium",
                    required=True
                ),
                "Role": st.column_config.SelectboxColumn(
                    "Role",
                    help="Hak akses user",
                    options=utils.ROLE_OPTIONS,
                    width="small",
                    required=True
                ),
                "Fungsi": st.column_config.SelectboxColumn(
                    "Divisi",
                    help="Divisi user",
                    options=fungsi_list,
                    width="medium"
                )
            },
            hide_index=True
        )
        
        if st.button("💾 Simpan Perubahan User", type="primary"):
            edited_users = edited_users.dropna(subset=['Nama', 'Email'])
            success, message = utils.save_users(edited_users)
            if success:
                st.success(message)
            else:
                st.error(message)

def render_analytics():
    st.markdown("""
    <div class='main-header'>
        <h1>📈 Analytics & Insights</h1>
        <p>Deep dive into your audit performance metrics</p>
    </div>
    """, unsafe_allow_html=True)

    df = load_data_cached()
    
    if df.empty:
        st.info("No data available for analysis.")
        return

    # Filter Section
    with st.expander("🔎 Advanced Filters", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
             selected_fungsi = st.multiselect("Filter Division", options=utils.get_fungsi_list())
        with col2:
             selected_status = st.multiselect("Filter Status", options=utils.DOC_STATUS_OPTIONS)
        with col3:
             date_range = st.date_input("Date Range (Submission Deadline)", value=[], help="Select start and end date")

    # Apply Filters
    df_filtered = df.copy()
    if selected_fungsi:
        df_filtered = df_filtered[df_filtered['Fungsi'].isin(selected_fungsi)]
    if selected_status:
        df_filtered = df_filtered[df_filtered['Document Status'].isin(selected_status)]
    if len(date_range) == 2:
        start_date, end_date = date_range
        # Ensure datetime conversion for comparison
        df_filtered['Doc. Submission Deadline'] = pd.to_datetime(df_filtered['Doc. Submission Deadline'])
        df_filtered = df_filtered[
            (df_filtered['Doc. Submission Deadline'].dt.date >= start_date) & 
            (df_filtered['Doc. Submission Deadline'].dt.date <= end_date)
        ]

    st.markdown("### 📊 Status Distribution Analysis")
    col_pie, col_bar = st.columns([1, 2])
    
    with col_pie:
        status_counts = df_filtered['Document Status'].value_counts()
        fig_pie = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title="Status Composition",
            hole=0.4,
            color=status_counts.index,
            color_discrete_map={'Outstanding': '#ED1C24', 'Need to Review': '#FFC107', 'Done': '#69BE28'}
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        # Group by Fungsi and Status
        fungsi_status = df_filtered.groupby(['Fungsi', 'Document Status']).size().reset_index(name='Count')
        fig_bar = px.bar(
            fungsi_status,
            x='Fungsi',
            y='Count',
            color='Document Status',
            title="Status by Division",
            barmode='stack',
            color_discrete_map={'Outstanding': '#ED1C24', 'Need to Review': '#FFC107', 'Done': '#69BE28'}
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("### 📉 Timeline Analysis")
    # Convert dates if not already
    df_filtered['Doc. Submission Deadline'] = pd.to_datetime(df_filtered['Doc. Submission Deadline'])
    
    # Trend of Deadlines
    timeline_data = df_filtered.resample('M', on='Doc. Submission Deadline').size().reset_index(name='Documents Due')
    fig_line = px.line(
        timeline_data,
        x='Doc. Submission Deadline',
        y='Documents Due',
        title="Document Submission Deadlines Over Time",
        markers=True,
        line_shape='spline'
    )
    fig_line.update_traces(line_color='#005DAA', line_width=3)
    st.plotly_chart(fig_line, use_container_width=True)
    
    # Overdue Analysis
    st.markdown("### ⚠️ Overdue Aging Analysis")
    overdue_df = df_filtered[ (df_filtered['Days To Go'] < 0) & (df_filtered['Document Status'] != 'Done') ].copy()
    
    if not overdue_df.empty:
        overdue_df['Overdue Days'] = abs(overdue_df['Days To Go'])
        fig_scatter = px.scatter(
            overdue_df,
            x='Overdue Days',
            y='Fungsi',
            size='Overdue Days',
            color='Document Status',
            hover_data=['Doc. ID Number', 'Document Description'],
            title="Overdue Severity by Division (Bubble Size = Days Late)",
            color_discrete_map={'Outstanding': '#ED1C24', 'Need to Review': '#FFC107'}
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.success("🎉 No overdue documents found in the current selection!")

def main():
    
    inject_custom_css()
    
    selected_menu = render_sidebar()
    
    # Map menu selection to function
    # Note: Using 'in' check to handle potentially decorated strings
    if "Dashboard" in selected_menu:
        render_dashboard()
    elif "Document Control" in selected_menu:
        render_document_control()
    elif "Analytics" in selected_menu:
        render_analytics()
    elif "Email" in selected_menu:
        render_email_automation()
    elif "Master Data" in selected_menu:
        render_master_data()

if __name__ == "__main__":
    main()
