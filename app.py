import streamlit as st
import pandas as pd
import datetime
import os

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="UID ID Timetable & Faculty Dispatch",
    page_icon="📅",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .badge-free {
        background-color: #dcfce7;
        color: #15803d;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-busy {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_data():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if not files:
        st.error("No timetable `.xlsx` spreadsheet found in the repository root directory.")
        st.stop()

    file_path = files[0]
    xls = pd.ExcelFile(file_path)

    # Prefer 'Dashboard' sheet for clean faculty-day mappings
    sheet = 'Dashboard' if 'Dashboard' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    # 1. Parse Calendar Date Columns (Row 1: Month, Row 2: Date, Row 3: Week)
    months = df.iloc[1].ffill()
    dates = df.iloc[2]
    weeks = df.iloc[3].ffill()

    month_map = {'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
    calendar_days = []

    for col in range(2, df.shape[1]):
        d_val = dates[col]
        if pd.notna(d_val) and str(d_val).strip() not in ['', 'nan', 'Date']:
            try:
                day_int = int(float(d_val))
                m_str = str(months[col]).strip()
                w_str = str(weeks[col]).strip()
                month_int = month_map.get(m_str, 7)

                date_obj = datetime.date(2026, month_int, day_int)
                calendar_days.append({
                    'col': col,
                    'date': date_obj,
                    'week': w_str,
                    'label': f"{date_obj.strftime('%a, %d %b %Y')} ({w_str})"
                })
            except Exception:
                continue

    # 2. Extract Faculty Roster
    faculty_blocks = []
    for r in range(5, df.shape[0]):
        fac_name = df.iloc[r, 1]
        if pd.notna(fac_name) and str(fac_name).strip() not in ['', 'nan', 'Month', 'Date', 'Week', 'No of Days']:
            name_clean = " ".join(str(fac_name).strip().split())
            faculty_blocks.append((r, name_clean))

    all_faculties = sorted(list(set([name for _, name in faculty_blocks])))

    # 3. Build Schedule Dictionary: faculty -> { date: module_string }
    faculty_schedule = {f: {} for f in all_faculties}

    for i, (r_start, fac_name) in enumerate(faculty_blocks):
        r_end = faculty_blocks[i + 1][0] if i + 1 < len(faculty_blocks) else min(r_start + 4, df.shape[0])

        for c_info in calendar_days:
            col = c_info['col']
            date_key = c_info['date']
            modules_found = []

            for sub_r in range(r_start, r_end):
                val = str(df.iloc[sub_r, col]).strip()
                if val and val not in ['nan', 'None', '-']:
                    # Exclude standalone digit strings representing contact hours/days
                    if not val.replace('.', '', 1).isdigit():
                        clean_val = " ".join(val.split())
                        if clean_val not in modules_found:
                            modules_found.append(clean_val)

            if modules_found:
                faculty_schedule[fac_name][date_key] = ", ".join(modules_found)

    return calendar_days, all_faculties, faculty_schedule


# --- DATA LOAD ---
try:
    calendar_days, all_faculties, faculty_schedule = load_data()
except Exception as e:
    st.error(f"Error parsing workbook: {e}")
    st.stop()

# --- TOP HEADER ---
st.title("UID Department of Industrial Design")
st.caption("Master Academic Timetable & Faculty Availability Dispatch")

tab_free, tab_personal, tab_matrix = st.tabs([
    "🔍 Faculty Availability on a Date", 
    "👤 Individual Faculty Timeline", 
    "📊 Master Contact Matrix"
])

# ==========================================================
# TAB 1: WHO IS FREE TODAY?
# ==========================================================
with tab_free:
    st.subheader("Daily Availability Checker")

    date_options = {c['label']: c['date'] for c in calendar_days}
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

    # Metric summary row
    c1, c2, c3 = st.columns(3)
    c1.metric("Selected Date", selected_date.strftime('%d %B %Y'))
    c2.metric("Available / Free Faculty", len(free_faculty))
    c3.metric("Scheduled / Busy Faculty", len(busy_faculty))

    st.markdown("---")
    col_free, col_busy = st.columns([1, 1])

    with col_free:
        st.markdown(f"#### ✅ Available Faculty ({len(free_faculty)})")
        st.caption("Instructors with zero teaching modules on this date:")
        if free_faculty:
            f_col1, f_col2 = st.columns(2)
            for idx, fac in enumerate(free_faculty):
                target = f_col1 if idx % 2 == 0 else f_col2
                target.markdown(f"• **{fac}** &nbsp; <span class='badge-free'>Free</span>", unsafe_allow_html=True)
        else:
            st.warning("All faculty members are scheduled today.")

    with col_busy:
        st.markdown(f"#### 🔒 Scheduled Faculty ({len(busy_faculty)})")
        st.caption("Instructors scheduled in studios/lectures:")
        if busy_faculty:
            df_busy = pd.DataFrame(busy_faculty)
            st.dataframe(df_busy, use_container_width=True, hide_index=True)
        else:
            st.info("No teaching modules scheduled on this day.")

# ==========================================================
# TAB 2: INDIVIDUAL FACULTY TIMELINE
# ==========================================================
with tab_personal:
    st.subheader("Individual Faculty Schedule")
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
        st.info(f"No specific module timeline recorded for {selected_fac}.")

# ==========================================================
# TAB 3: MASTER SCHEDULE MATRIX
# ==========================================================
with tab_matrix:
    st.subheader("Master Daily Matrix (Faculty vs Calendar Days)")
    filter_week = st.selectbox(
        "Filter by Academic Week:",
        options=["All Weeks"] + sorted(list(set([c['week'] for c in calendar_days if c['week']])))
    )

    active_cols = calendar_days if filter_week == "All Weeks" else [c for c in calendar_days if c['week'] == filter_week]
    matrix_data = {'Faculty': all_faculties}

    for c in active_cols:
        col_name = f"{c['date'].strftime('%d-%b')} ({c['week']})"
        matrix_data[col_name] = [faculty_schedule.get(f, {}).get(c['date'], "-") for f in all_faculties]

    df_matrix = pd.DataFrame(matrix_data)
    st.dataframe(df_matrix, use_container_width=True, hide_index=True)
