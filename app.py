import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(page_title="UID ID Timetable", page_icon="📅", layout="wide")

@st.cache_data
def load_unmerged_data():
    # Target unmerged timetable file
    files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if not files:
        st.error("No timetable spreadsheet found.")
        st.stop()
        
    df = pd.read_excel(files[0], sheet_name='Dashboard', header=None)

    # 1. Parse Calendar Header Rows
    # Row 1: Months | Row 2: Dates | Row 3: Weeks
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
                d_obj = datetime.date(2026, month_map.get(m_str, 7), day_int)
                calendar_days.append({
                    'col': col,
                    'date': d_obj,
                    'week': w_str,
                    'label': f"{d_obj.strftime('%a, %d %b %Y')} ({w_str})"
                })
            except Exception:
                continue

    # 2. Extract Faculty Roster & Direct Column Lookups
    faculty_roster = []
    faculty_rows = {}

    for r in range(5, df.shape[0]):
        fac_name = df.iloc[r, 1]
        if pd.notna(fac_name) and str(fac_name).strip() not in ['', 'nan', 'Month', 'Date', 'Week', 'No of Days']:
            name_clean = str(fac_name).strip()
            faculty_roster.append(name_clean)
            faculty_rows[name_clean] = r

    # 3. Direct Matrix Population (No unmerging algorithms needed)
    faculty_schedule = {f: {} for f in faculty_roster}

    for fac, r_idx in faculty_rows.items():
        # Check faculty row and the 2 descriptor rows beneath it
        for c_info in calendar_days:
            col = c_info['col']
            entries = []
            for sub_r in range(r_idx, min(r_idx + 4, df.shape[0])):
                val = str(df.iloc[sub_r, col]).strip()
                if val and val not in ['nan', 'None', '-'] and not val.replace('.', '', 1).isdigit():
                    if val not in entries:
                        entries.append(" ".join(val.split()))
            
            if entries:
                faculty_schedule[fac][c_info['date']] = ", ".join(entries)

    return calendar_days, sorted(list(set(faculty_roster))), faculty_schedule


calendar_days, faculties, schedule = load_unmerged_data()

st.title("UID Industrial Design — Faculty Dispatch")
tab_free, tab_all = st.tabs(["🔍 Faculty Availability", "📊 Master Timetable View"])

with tab_free:
    date_map = {c['label']: c['date'] for c in calendar_days}
    selected_label = st.selectbox("Select Date:", list(date_map.keys()))
    target_date = date_map[selected_label]

    free_list = []
    busy_list = []

    for f in faculties:
        assigned = schedule[f].get(target_date)
        if assigned:
            busy_list.append({"Faculty": f, "Assigned Module": assigned})
        else:
            free_list.append(f)

    c1, c2 = st.columns(2)
    c1.metric("Available / Free", len(free_list))
    c2.metric("Scheduled / Busy", len(busy_list))

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Available Faculty")
        for faculty in free_list:
            st.success(faculty)

    with col_b:
        st.subheader("Busy Faculty")
        if busy_list:
            st.dataframe(pd.DataFrame(busy_list), use_container_width=True, hide_index=True)
        else:
            st.info("No sessions scheduled.")

with tab_all:
    st.subheader("Master Allocation Matrix")
    matrix_records = []
    for f in faculties:
        row = {"Faculty": f}
        for c in calendar_days:
            row[c['date'].strftime('%d-%b')] = schedule[f].get(c['date'], "-")
        matrix_records.append(row)
    st.dataframe(pd.DataFrame(matrix_records), use_container_width=True, hide_index=True)import streamlit as st
import pandas as pd
import openpyxl
import datetime

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="UID Industrial Design | Timetable & Faculty Dispatch",
    page_icon="📅",
    layout="wide"
)

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
def load_and_parse_timetable(file_name):
    # Load workbook directly with openpyxl to resolve merged cells
    wb = openpyxl.load_workbook(file_name, data_only=True)
    ws = wb['Dashboard']

    # 1. Build a lookup map of all merged ranges -> propagate top-left value
    merged_map = {}
    for rng in ws.merged_cells.ranges:
        top_left_val = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_map[(r, c)] = top_left_val

    def get_val(r, c):
        if (r, c) in merged_map:
            return merged_map[(r, c)]
        return ws.cell(r, c).value

    # 2. Parse Months, Dates, and Weeks across columns
    month_map = {'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
    calendar_days = []
    
    current_month_str = "Jul"
    for col in range(2, ws.max_column + 1):
        m_cand = get_val(2, col)
        if m_cand in ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
            current_month_str = m_cand

        d_val = get_val(3, col)
        w_val = get_val(4, col)

        if d_val is not None and str(d_val).strip() not in ['', 'None', 'Date']:
            try:
                day_int = int(float(d_val))
                month_int = month_map.get(current_month_str, 7)
                week_str = str(w_val).strip() if w_val else "Week"
                
                date_obj = datetime.date(2026, month_int, day_int)
                calendar_days.append({
                    'col': col,
                    'date': date_obj,
                    'week': week_str,
                    'formatted': f"{date_obj.strftime('%a, %d %b %Y')} ({week_str})"
                })
            except Exception:
                continue

    # 3. Locate all Faculty blocks
    faculty_list = []
    faculty_row_indices = []

    for r in range(6, ws.max_row + 1):
        c1 = ws.cell(r, 2).value
        if c1 is not None and str(c1).strip() not in ['', 'None', 'Month', 'Date', 'Week', 'No of Days']:
            fac_name = str(c1).strip()
            faculty_list.append(fac_name)
            faculty_row_indices.append((r, fac_name))

    # 4. Extract schedule per faculty across all unmerged cells
    faculty_schedule = {fac: {} for fac in faculty_list}

    for i, (r_start, fac_name) in enumerate(faculty_row_indices):
        r_end = faculty_row_indices[i + 1][0] if i + 1 < len(faculty_row_indices) else min(r_start + 4, ws.max_row + 1)
        
        for c_day in calendar_days:
            col = c_day['col']
            date_key = c_day['date']
            modules_found = []

            for r in range(r_start, r_end):
                val = get_val(r, col)
                if val is not None and str(val).strip() not in ['', 'None', '-', 'nan']:
                    s = str(val).strip()
                    # Filter out stray day-count numbers like '10', '9'
                    if not s.replace('.', '', 1).isdigit():
                        clean_mod = " ".join(s.split())
                        if clean_mod not in modules_found:
                            modules_found.append(clean_mod)

            if modules_found:
                faculty_schedule[fac_name][date_key] = ", ".join(modules_found)

    return calendar_days, sorted(list(set(faculty_list))), faculty_schedule


# --- RUN LOADER ---
try:
    import os
    # Automatically locate the Excel file in current directory
    excel_candidates = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if not excel_candidates:
        st.error("No `.xlsx` file found in the app directory.")
        st.stop()
    
    file_path = excel_candidates[0]
    calendar_days, all_faculties, faculty_schedule = load_and_parse_timetable(file_path)
except Exception as e:
    st.error(f"Error reading timetable workbook: {e}")
    st.stop()

# --- TOP INTERFACE ---
st.title("UID Department of Industrial Design")
st.caption("Master Academic Schedule & Dynamic Faculty Dispatch System")

tab_free, tab_personal, tab_matrix = st.tabs([
    "🔍 Faculty Availability on a Date", 
    "👤 Individual Faculty Timeline", 
    "📊 Master Contact Matrix"
])

# ==========================================================
# TAB 1: WHO IS FREE TODAY?
# ==========================================================
with tab_free:
    st.subheader("Faculty Availability Checker")

    date_options = {c['formatted']: c['date'] for c in calendar_days}
    selected_label = st.selectbox("Select Academic Day:", options=list(date_options.keys()))
    selected_date = date_options[selected_label]

    free_faculty = []
    busy_faculty = []

    for fac in all_faculties:
        mod = faculty_schedule.get(fac, {}).get(selected_date, None)
        if mod:
            busy_faculty.append({'Faculty': fac, 'Scheduled Module': mod})
        else:
            free_faculty.append(fac)

    # Metrics
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
            f_col1, f_col2 = st.columns(2)
            for idx, fac in enumerate(free_faculty):
                target = f_col1 if idx % 2 == 0 else f_col2
                target.markdown(f"• **{fac}** &nbsp; <span class='badge-free'>Free</span>", unsafe_allow_html=True)
        else:
            st.warning("All faculty members are scheduled today.")

    with col_busy:
        st.markdown(f"#### 🔒 Scheduled Faculty ({len(busy_faculty)})")
        st.caption("Instructors actively scheduled in studios/lectures:")
        if busy_faculty:
            df_busy = pd.DataFrame(busy_faculty)
            st.dataframe(df_busy, use_container_width=True, hide_index=True)
        else:
            st.info("No teaching modules scheduled on this day.")

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
