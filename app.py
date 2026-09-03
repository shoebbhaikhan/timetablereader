import streamlit as st
import pandas as pd
import datetime

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="UID Industrial Design | Timetable & Faculty Dispatch",
    page_icon="📅",
    layout="wide"
)

# Custom CSS for clean, high-contrast aesthetics
st.markdown("""
<style>
    .metric-box {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
    }
    .badge-free {
        background-color: #dcfce7;
        color: #15803d;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-busy {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_and_parse_data(file_path="Timetable-ODD Sem 2026-27.xlsx"):
    xls = pd.ExcelFile(file_path)
    df_dash = pd.read_excel(xls, sheet_name='Dashboard', header=None)
    
    # 1. Parse Calendar Date Columns from Dashboard
    months = df_dash.iloc[1].ffill()
    dates = df_dash.iloc[2]
    weeks = df_dash.iloc[3].ffill()

    calendar_days = []
    month_map = {'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}

    for col_idx in range(2, df_dash.shape[1]):
        d_val = dates[col_idx]
        if pd.notna(d_val) and str(d_val).strip() not in ['', 'nan', 'Date']:
            try:
                day_int = int(float(d_val))
                m_str = str(months[col_idx]).strip()
                w_str = str(weeks[col_idx]).strip()
                m_int = month_map.get(m_str, 7)
                
                # Standard academic year mapping
                date_obj = datetime.date(2026, m_int, day_int)
                calendar_days.append({
                    'col_idx': col_idx,
                    'date': date_obj,
                    'month': m_str,
                    'day': day_int,
                    'week': w_str,
                    'formatted': f"{date_obj.strftime('%a, %d %b %Y')} ({w_str})"
                })
            except Exception:
                continue

    # 2. Extract Faculty Rows & Assignments
    faculty_blocks = []
    for r in range(5, df_dash.shape[0]):
        c1 = df_dash.iloc[r, 1]
        c0 = df_dash.iloc[r, 0]
        if pd.notna(c1) and str(c1).strip() not in ['', 'nan', 'Month', 'Date', 'Week', 'No of Days']:
            fac_name = str(c1).strip()
            faculty_blocks.append((r, fac_name))

    faculty_schedule = {} # Key: fac_name -> { date_str: module_string }
    all_faculties = sorted(list(set([name for _, name in faculty_blocks])))

    for i, (r_start, fac_name) in enumerate(faculty_blocks):
        r_end = faculty_blocks[i+1][0] if i+1 < len(faculty_blocks) else min(r_start + 4, df_dash.shape[0])
        if fac_name not in faculty_schedule:
            faculty_schedule[fac_name] = {}

        for c_day in calendar_days:
            col = c_day['col_idx']
            assigned_modules = []
            for sub_r in range(r_start, r_end):
                cell_val = str(df_dash.iloc[sub_r, col]).strip()
                if cell_val not in ['', 'nan', 'None', '-']:
                    if not cell_val.replace('.', '', 1).isdigit():
                        clean_name = " ".join(cell_val.split())
                        if clean_name not in assigned_modules:
                            assigned_modules.append(clean_name)
            
            if assigned_modules:
                faculty_schedule[fac_name][c_day['date']] = ", ".join(assigned_modules)

    return calendar_days, all_faculties, faculty_schedule


# --- DATA INITIALIZATION ---
try:
    calendar_days, all_faculties, faculty_schedule = load_and_parse_data()
except Exception as e:
    st.error(f"Error reading timetable workbook: {e}")
    st.info("Make sure 'Timetable-ODD Sem 2026-27.xlsx' is present in the repository root directory.")
    st.stop()

# --- TOP NAVIGATION & HEADER ---
st.title("UID Department of Industrial Design")
st.caption("Master Academic Schedule & Dynamic Faculty Dispatch System")

tab_free, tab_personal, tab_matrix = st.tabs([
    "🔍 Faculty Availability on a Date", 
    "👤 Individual Faculty Timeline", 
    "📊 Master Contact Matrix"
])


# ==========================================================
# TAB 1: FACULTY AVAILABILITY (WHO IS FREE TODAY?)
# ==========================================================
with tab_free:
    st.subheader("Faculty Availability Checker")
    
    date_options = {c['formatted']: c['date'] for c in calendar_days}
    selected_label = st.selectbox("Select Academic Day:", options=list(date_options.keys()))
    selected_date = date_options[selected_label]

    free_faculty = []
    busy_faculty = []

    for fac in all_faculties:
        assigned = faculty_schedule.get(fac, {}).get(selected_date, None)
        if assigned:
            busy_faculty.append({'Faculty': fac, 'Scheduled Module': assigned})
        else:
            free_faculty.append(fac)

    # Summary Metrics
    c1, c2, c3 = st.columns(3)
    c1.metric("Selected Date", selected_date.strftime('%d %B %Y'))
    c2.metric("Available / Free Faculty", len(free_faculty))
    c3.metric("Scheduled / Busy Faculty", len(busy_faculty))

    st.markdown("---")
    col_free, col_busy = st.columns([1, 1])

    with col_free:
        st.markdown(f"#### ✅ Available Faculty ({len(free_faculty)})")
        st.caption("Instructors with zero teaching modules assigned on this date:")
        if free_faculty:
            # Display free faculty in a dual-column list
            f_col1, f_col2 = st.columns(2)
            for idx, fac in enumerate(free_faculty):
                target = f_col1 if idx % 2 == 0 else f_col2
                target.markdown(f"• **{fac}** &nbsp; <span class='badge-free'>Free</span>", unsafe_allow_html=True)
        else:
            st.warning("All faculty members have active sessions scheduled today.")

    with col_busy:
        st.markdown(f"#### 🔒 Scheduled Faculty ({len(busy_faculty)})")
        st.caption("Instructors actively scheduled in studios/lectures:")
        if busy_faculty:
            df_busy = pd.DataFrame(busy_faculty)
            st.dataframe(df_busy, use_container_width=True, hide_index=True)
        else:
            st.info("No active studio sessions found on this day (Academic holiday or review day).")


# ==========================================================
# TAB 2: INDIVIDUAL FACULTY TIMELINE
# ==========================================================
with tab_personal:
    st.subheader("Individual Faculty Semester Calendar")
    selected_fac = st.selectbox("Select Faculty Member:", options=all_faculties)

    fac_days = faculty_schedule.get(selected_fac, {})
    if fac_days:
        timeline_list = []
        for c in calendar_days:
            mod = fac_days.get(c['date'], None)
            if mod:
                timeline_list.append({
                    'Date': c['date'].strftime('%d-%b-%Y'),
                    'Week': c['week'],
                    'Assigned Module': mod
                })
        
        df_fac = pd.DataFrame(timeline_list)
        st.markdown(f"**Total Instructional Days Scheduled:** `{len(df_fac)} days`")
        st.dataframe(df_fac, use_container_width=True, hide_index=True)
    else:
        st.info(f"No specific module timeline recorded in Dashboard for {selected_fac}.")


# ==========================================================
# TAB 3: MASTER SCHEDULE MATRIX
# ==========================================================
with tab_matrix:
    st.subheader("Master Daily Matrix (Faculty vs Calendar Days)")
    filter_week = st.selectbox(
        "Filter by Academic Week:", 
        options=["All Weeks"] + sorted(list(set([c['week'] for c in calendar_days])))
    )

    # Build matrix dataframe
    active_cols = calendar_days if filter_week == "All Weeks" else [c for c in calendar_days if c['week'] == filter_week]
    matrix_data = {'Faculty': all_faculties}

    for c in active_cols:
        col_name = f"{c['date'].strftime('%d-%b')} ({c['week']})"
        matrix_data[col_name] = [faculty_schedule.get(f, {}).get(c['date'], "-") for f in all_faculties]

    df_matrix = pd.DataFrame(matrix_data)
    st.dataframe(df_matrix, use_container_width=True, hide_index=True)