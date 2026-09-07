import streamlit as st
import pandas as pd
import openpyxl
import datetime
import os
import re
import json

# --- ADMIN PASSWORD CONFIGURATION ---
ADMIN_PASSWORD = "ArsenalFC"  # Change this to your preferred password
EXCLUSION_FILE = "excluded_faculty.json"

st.set_page_config(
    page_title="UID Industrial Design | Timetable",
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
        display: inline-block;
        margin: 2px 0;
    }
    .badge-busy {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
        margin: 2px 0;
    }
    .faculty-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)


# --- PERSISTENT EXCLUSION HELPERS ---
def get_excluded_faculty():
    if os.path.exists(EXCLUSION_FILE):
        try:
            with open(EXCLUSION_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_excluded_faculty(excluded_list):
    with open(EXCLUSION_FILE, "w") as f:
        json.dump(excluded_list, f)


@st.cache_data
def load_and_parse():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
    if not files:
        st.error("No timetable `.xlsx` file found in the repository.")
        st.stop()
    
    file_path = files[0]
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    # 1. Faculty Roster
    df_fw = pd.read_excel(file_path, sheet_name='Faculty Work Load ')
    raw_faculties = df_fw['Faculty Name '].dropna().unique().tolist()
    
    faculty_list = []
    first_name_to_full = {}
    
    for f in raw_faculties:
        full = " ".join(str(f).split())
        tokens = [t.lower().strip("().,") for t in full.split()]
        tokens = [t for t in tokens if t not in ['dr', 'mr', 'ms', 'prof']]
        if tokens:
            fn = tokens[0]
            first_name_to_full[fn] = full
            faculty_list.append(full)
            
    faculty_list = sorted(list(set(faculty_list)))
    
    # 2. Daily Timetable Sheet
    sheet_name = 'Morning_Afternoon Updated_Timet' if 'Morning_Afternoon Updated_Timet' in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    month_map = {'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
    calendar_cols = []

    for col in range(2, ws.max_column + 1):
        m_val = ws.cell(3, col).value
        d_val = ws.cell(4, col).value
        w_val = ws.cell(5, col).value

        if d_val is not None and str(d_val).strip() not in ['', 'None', 'Date']:
            try:
                day_int = int(float(d_val))
                m_str = str(m_val).strip()
                if m_str in month_map:
                    w_str = str(w_val).strip() if w_val else "Week"
                    d_obj = datetime.date(2026, month_map[m_str], day_int)
                    calendar_cols.append({
                        'col': col,
                        'date': d_obj,
                        'week': w_str,
                        'label': f"{d_obj.strftime('%a, %d %b %Y')} ({w_str})"
                    })
            except Exception:
                continue

    # 3. Active Session Coordinates
    teaching_rows = [
        (14, 'UG-Sem 3', 'Sec A', 'Morning', 8),
        (15, 'UG-Sem 3', 'Sec A', 'Afternoon', 8),
        (16, 'UG-Sem 3', 'Sec B', 'Morning', 8),
        (17, 'UG-Sem 3', 'Sec B', 'Afternoon', 8),
        (18, 'UG-Sem 3', 'Sec C', 'Morning', 8),
        (19, 'UG-Sem 3', 'Sec C', 'Afternoon', 8),
        (20, 'UG-Sem 3', 'Sec D', 'Morning', 8),
        (21, 'UG-Sem 3', 'Sec D', 'Afternoon', 8),
        (22, 'UG-Sem 3', 'Sec E', 'Morning', 8),
        (23, 'UG-Sem 3', 'Sec E', 'Afternoon', 8),
        (33, 'UG-Sem 5', 'Sec A', 'Morning', 30),
        (34, 'UG-Sem 5', 'Sec A', 'Afternoon', 30),
        (35, 'UG-Sem 5', 'Sec B', 'Morning', 30),
        (36, 'UG-Sem 5', 'Sec B', 'Afternoon', 30),
        (37, 'UG-Sem 5', 'Sec C', 'Morning', 30),
        (38, 'UG-Sem 5', 'Sec C', 'Afternoon', 30),
        (39, 'UG-Sem 5', 'Sec D', 'Morning', 30),
        (40, 'UG-Sem 5', 'Sec D', 'Afternoon', 30),
        (51, 'UG-Sem 7', 'Sec A', 'Lead', 48),
        (52, 'UG-Sem 7', 'Sec A', 'Assisting', 48),
        (53, 'UG-Sem 7', 'Sec B', 'Lead', 48),
        (54, 'UG-Sem 7', 'Sec B', 'Assisting', 48),
        (55, 'UG-Sem 7', 'Sec C', 'Lead', 48),
        (56, 'UG-Sem 7', 'Sec C', 'Assisting', 48),
        (65, 'PG-Sem 1', 'Main', 'Full Day', 62),
        (66, 'PG-Sem 1', 'Master Class', 'Session', 62),
        (74, 'PG-Sem 3', 'Main', 'Full Day', 71),
        (75, 'PG-Sem 3', 'Master Class', 'Session', 71),
    ]

    faculty_day_schedule = {f: {} for f in faculty_list}

    for c_info in calendar_cols:
        col = c_info['col']
        dt = c_info['date']

        for r, cohort, sec, slot, mod_r in teaching_rows:
            cell_val = ws.cell(r, col).value
            if cell_val is not None and str(cell_val).strip() not in ['', 'None', '-']:
                txt = str(cell_val).strip()
                mod_name = str(ws.cell(mod_r, col).value or "Studio").strip()
                
                cell_tokens = [re.sub(r'[^a-zA-Z]', '', w).lower() for w in txt.split()]
                matched_faculty = None
                for tok in cell_tokens:
                    if tok in first_name_to_full:
                        matched_faculty = first_name_to_full[tok]
                        break

                if matched_faculty:
                    detail = f"{cohort} | {sec} ({slot}) - {mod_name}"
                    if dt not in faculty_day_schedule[matched_faculty]:
                        faculty_day_schedule[matched_faculty][dt] = []
                    if detail not in faculty_day_schedule[matched_faculty][dt]:
                        faculty_day_schedule[matched_faculty][dt].append(detail)

    return calendar_cols, faculty_list, faculty_day_schedule


try:
    calendar_cols, faculty_list, faculty_schedule = load_and_parse()
except Exception as e:
    st.error(f"Error initializing timetable engine: {e}")
    st.stop()

st.title("UID Department of Industrial Design")
st.caption("Master Academic Schedule & Faculty Dispatch")

tab1, tab2, tab3 = st.tabs([
    "🔍 Faculty Availability by Date", 
    "👤 Individual Faculty Schedule",
    "🔒 Admin: Exclude Faculty"
])

# ==========================================================
# TAB 1: AVAILABILITY ENGINE
# ==========================================================
with tab1:
    excluded_faculties = get_excluded_faculty()

    date_to_week = {c['date']: c['week'] for c in calendar_cols}
    valid_dates = [c['date'] for c in calendar_cols]
    min_date = valid_dates[0]
    max_date = valid_dates[-1]

    default_date = datetime.date(2026, 9, 7) if min_date <= datetime.date(2026, 9, 7) <= max_date else min_date

    selected_date = st.date_input(
        "Select Academic Date:",
        value=default_date,
        min_value=min_date,
        max_value=max_date
    )

    current_week = date_to_week.get(selected_date, "Non-Instructional / Holiday")

    busy_members = []
    free_members = []

    for fac in faculty_list:
        classes = faculty_schedule.get(fac, {}).get(selected_date, [])
        if classes:
            busy_members.append((fac, classes))
        else:
            # Only add to free list if not marked as permanently excluded
            if fac not in excluded_faculties:
                free_members.append(fac)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Week", current_week)
    m2.metric("Date", selected_date.strftime('%d %B %Y'))
    m3.metric("Available / Free", len(free_members))
    m4.metric("Scheduled in Class", len(busy_members))

    st.markdown("---")
    col_free, col_busy = st.columns([1, 1.2])

    with col_free:
        st.subheader(f"✅ Free Faculty ({len(free_members)})")
        st.caption("Available for substitution/jury duties:")
        f1, f2 = st.columns(2)
        for idx, fac in enumerate(free_members):
            target = f1 if idx % 2 == 0 else f2
            target.markdown(f"• **{fac}** &nbsp; <span class='badge-free'>Available</span>", unsafe_allow_html=True)

    with col_busy:
        st.subheader(f"🔒 Scheduled Faculty ({len(busy_members)})")
        st.caption("Instructors with active class assignments:")
        for fac, sessions in busy_members:
            details_html = "<br>".join([f"&nbsp;&nbsp;↳ {s}" for s in sessions])
            st.markdown(f"""
            <div class="faculty-card">
                <b>{fac}</b> &nbsp; <span class="badge-busy">In Class</span><br>
                <small style="color: #475569;">{details_html}</small>
            </div>
            """, unsafe_allow_html=True)

# ==========================================================
# TAB 2: INDIVIDUAL LOOKUP
# ==========================================================
with tab2:
    selected_fac = st.selectbox("Select Faculty Member:", options=faculty_list)
    sched = faculty_schedule.get(selected_fac, {})

    if sched:
        st.markdown(f"**Total Teaching Days:** `{len(sched)} days`")
        records = []
        for c in calendar_cols:
            if c['date'] in sched:
                records.append({
                    'Date': c['date'].strftime('%d-%b-%Y'),
                    'Week': c['week'],
                    'Assigned Sessions': "; ".join(sched[c['date']])
                })
        st.table(pd.DataFrame(records))
    else:
        st.info(f"No active teaching assignments found for {selected_fac}.")

# ==========================================================
# TAB 3: ADMIN FACULTY EXCLUSION (PASSWORD PROTECTED)
# ==========================================================
with tab3:
    st.subheader("Admin Control: Faculty Availability Visibility")
    st.caption("Manage faculty members who should never appear in the available/free list (e.g., Deans, Leadership, or Non-teaching staff).")

    # Injected clean CSS styling for table borders
    st.markdown("""
    <style>
        .roster-header {
            border: 1px solid #cbd5e1;
            border-bottom: 2px solid #94a3b8;
            border-radius: 8px 8px 0 0;
            background-color: #f1f5f9;
            padding: 10px 16px;
            font-weight: 700;
            color: #0f172a;
            font-size: 0.95rem;
        }
        .roster-row-even {
            border-left: 1px solid #cbd5e1;
            border-right: 1px solid #cbd5e1;
            border-bottom: 1px solid #cbd5e1;
            background-color: #ffffff;
            padding: 4px 16px;
        }
        .roster-row-odd {
            border-left: 1px solid #cbd5e1;
            border-right: 1px solid #cbd5e1;
            border-bottom: 1px solid #cbd5e1;
            background-color: #f8fafc;
            padding: 4px 16px;
        }
        .roster-last {
            border-radius: 0 0 8px 8px;
        }
    </style>
    """, unsafe_allow_html=True)

    password_attempt = st.text_input("Enter Admin Password to Unlock:", type="password")

    if password_attempt:
        if password_attempt == ADMIN_PASSWORD:
            st.success("Admin authenticated.")

            current_exclusions = set(get_excluded_faculty())

            st.write("### Faculty Visibility Roster")
            st.caption("Check the box next to any faculty member you want to **exclude** from the Free list. Click **Save Changes** below when done.")

            # Header
            st.markdown("""
            <div class="roster-header">
                <div style="display: flex; justify-content: space-between;">
                    <span style="flex: 3;">Faculty Name</span>
                    <span style="flex: 1; text-align: right; padding-right: 12px;">Exclude from Free List</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            new_exclusions = []
            total_fac = len(faculty_list)

            # Rows
            for idx, fac in enumerate(faculty_list):
                row_cls = "roster-row-even" if idx % 2 == 0 else "roster-row-odd"
                if idx == total_fac - 1:
                    row_cls += " roster-last"

                with st.container():
                    c1, c2 = st.columns([3, 1])
                    c1.markdown(f"<div class='{row_cls}' style='border-right: none; height: 100%; display: flex; align-items: center; padding-top: 10px; font-weight: 600; color: #1e293b;'>{fac}</div>", unsafe_allow_html=True)
                    
                    with c2:
                        st.markdown(f"<div class='{row_cls}' style='border-left: none; text-align: right; padding-top: 6px;'>", unsafe_allow_html=True)
                        is_excluded = st.checkbox(
                            "Exclude",
                            value=(fac in current_exclusions),
                            key=f"excl_{idx}",
                            label_visibility="collapsed"
                        )
                        st.markdown("</div>", unsafe_allow_html=True)

                    if is_excluded:
                        new_exclusions.append(fac)

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Save Changes", type="primary"):
                save_excluded_faculty(new_exclusions)
                st.success("Roster visibility updated successfully!")
                st.rerun()
        else:
            st.error("Incorrect password. Access denied.")
