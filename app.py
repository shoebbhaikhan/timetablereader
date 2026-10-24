import datetime
import os
import re
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="UID ID Timetable & Generator",
    page_icon="📅",
    layout="wide"
)

st.markdown(
    """
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
    .wt-table {
        width: 100%;
        border-collapse: collapse;
        font-family: inherit;
        margin-top: 14px;
    }
    .wt-th {
        background-color: #0f172a;
        color: #ffffff;
        padding: 10px 8px;
        text-align: center;
        font-size: 0.88rem;
        border: 1px solid #334155;
    }
    .wt-td-label {
        background-color: #f8fafc;
        color: #0f172a;
        font-weight: 700;
        padding: 8px 10px;
        border: 1px solid #cbd5e1;
        font-size: 0.85rem;
        vertical-align: middle;
        white-space: nowrap;
    }
    .wt-td {
        border: 1px solid #cbd5e1;
        padding: 6px 4px;
        vertical-align: top;
        background-color: #ffffff;
        width: 18%;
    }
    .s-pill {
        border-radius: 4px;
        padding: 3px 6px;
        margin-bottom: 4px;
        font-size: 0.74rem;
        line-height: 1.25;
        display: block;
        border-left: 3px solid transparent;
    }
    .s1 { background: #dbeafe; color: #1e40af; border-color: #2563eb; }
    .s2 { background: #ccfbf1; color: #0f766e; border-color: #0d9488; }
    .s3 { background: #fef3c7; color: #92400e; border-color: #d97706; }
    .s4 { background: #f3e8ff; color: #6b21a8; border-color: #9333ea; }
    .s-off { background: #f1f5f9; color: #94a3b8; font-style: italic; }

    /* Completion Status Badges */
    .status-badge {
        padding: 4px 12px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
        text-align: center;
        letter-spacing: 0.02em;
    }
    .status-completed {
        background-color: #dcfce7;
        color: #166534;
        border: 1px solid #bbf7d0;
    }
    .status-active {
        background-color: #ffedd5;
        color: #9a3412;
        border: 1px solid #fed7aa;
    }
    .status-upcoming {
        background-color: #cffafe;
        color: #155e75;
        border: 1px solid #a5f3fc;
    }
</style>
""",
    unsafe_allow_html=True,
)

# =============================================================
# SUPABASE CLIENT INITIALIZATION (ROBUST SECRETS LOADER)
# =============================================================
@st.cache_resource
def init_supabase():
    url = None
    key = None

    # 1. Try Streamlit native secrets
    try:
        if "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
            url = st.secrets["SUPABASE_URL"]
            key = st.secrets["SUPABASE_KEY"]
    except Exception:
        pass

    # 2. Direct file fallback if st.secrets fails
    if not url or not key:
        secrets_path = os.path.join(".streamlit", "secrets.toml")
        if os.path.exists(secrets_path):
            with open(secrets_path, "r") as f:
                for line in f:
                    if "SUPABASE_URL" in line and "=" in line:
                        url = line.split("=")[1].strip().strip('"').strip("'")
                    if "SUPABASE_KEY" in line and "=" in line:
                        key = line.split("=")[1].strip().strip('"').strip("'")

    # 3. Hardcoded project fallback for your deployment
    if not url or not key:
        url = "https://fpxuovievrepkseceszo.supabase.co"
        key = "sb_publishable_CR7h1w2d9oJUlpjyklV1fg_7fujOFuB"

    if url and key:
        try:
            from supabase import create_client
            return create_client(url, key)
        except Exception as e:
            st.error(f"Error connecting to Supabase: {e}")
            return None
    return None

supabase_client = init_supabase()

# =============================================================
# DYNAMIC FACULTY DATABASE LOADER
# =============================================================
@st.cache_data(ttl=300)
def get_faculty_roster():
    if supabase_client:
        try:
            res = (
                supabase_client.table("faculties")
                .select("faculty_code, name, first_name_alias, active_until, is_active, is_available_for_substitution")
                .eq("is_active", True)
                .order("name")
                .execute()
            )
            if res.data:
                db_list = {"-- None / Leave Empty --": "", "✏️ [Manual / Custom Entry]": "CUSTOM"}
                for f in res.data:
                    db_list[f"{f['name']} ({f['faculty_code']})"] = f["faculty_code"]
                return db_list, res.data
        except Exception:
            pass

    fallback_list = {
        "-- None / Leave Empty --": "",
        "✏️ [Manual / Custom Entry]": "CUSTOM",
        "Sharad Shekar Shetty (15044)": "15044",
        "Kishan Chavda (15177)": "15177",
        "Umang Shah (15069)": "15069",
        "Shoeb Iqbal Khan (15255)": "15255",
        "Abhishek Karmakar (15254)": "15254",
        "Malekulashter (15342)": "15342",
        "Arshkirat Singh Gill (15279)": "15279",
        "Aditya Lingam (15067)": "15067",
        "Sreya Acharyya (15461)": "15461",
        "Dhanush Kumar (15497)": "15497",
        "Vipul Vinayak Jadhav (15500)": "15500",
        "Ashuj Chawda (15506)": "15506",
        "DA Siddharth (15445)": "15445",
        "Sree Hari B Lal (15546)": "15546",
        "Mithil Suresh (15552)": "15552",
        "Bhargav Manchalla (15556)": "15556",
        "Atul Kedia (15477)": "15477",
        "Bhargav Mistry (15478)": "15478",
    }
    raw_fallback = [
        {
            "faculty_code": code,
            "name": label.split(" (")[0],
            "active_until": "2026-10-10" if "Mithil" in label else None,
            "is_available_for_substitution": False if "Atul" in label else True
        }
        for label, code in fallback_list.items() if code not in ["", "CUSTOM"]
    ]
    return fallback_list, raw_fallback

DEFAULT_FACULTY_LIST, RAW_DB_FACULTIES = get_faculty_roster()
DATABASE_FACULTY_NAMES = [f["name"] for f in RAW_DB_FACULTIES]
def is_faculty_active_on(fac_entry, target_date):
    """
    Checks if a faculty member is active on target_date based on their active_until date.
    Permanent faculty (active_until = None) are always active.
    """
    if not isinstance(fac_entry, dict):
        return True
    last_date = fac_entry.get("active_until")
    if not last_date:
        return True
    if isinstance(last_date, str):
        try:
            last_date = datetime.date.fromisoformat(last_date)
        except ValueError:
            return True
    return target_date <= last_date

# =============================================================
# DYNAMIC COURSES DATABASE LOADER
# =============================================================
@st.cache_data(ttl=300)
def get_courses_roster():
    if supabase_client:
        try:
            res = (
                supabase_client.table("courses")
                .select("course_code, title, semester, semester_type, program_code, classification, hours_needed, days_needed")
                .order("semester")
                .order("title")
                .execute()
            )
            if res.data:
                c_list = []
                c_options = {}
                for c in res.data:
                    c_dict = {
                        "code": c["course_code"],
                        "title": c["title"],
                        "sem": c["semester"],
                        "sem_type": c.get("semester_type", "ODD"),
                        "prog_code": c.get("program_code", "1019"),
                        "batch": f"{c['semester']} ID",
                        "classification": c.get("classification", "PRACTICAL"),
                        "hours_needed": float(c.get("hours_needed") or 0),
                        "days_needed": float(c.get("days_needed") or 0)
                    }
                    c_list.append(c_dict)
                    c_options[f"{c_dict['title']}  |  [{c_dict['code']}]  ({c_dict['sem']})"] = c_dict
                c_options["✏️ [Custom / Manual Course Entry]"] = {
                    "code": "", "title": "Custom", "sem": "I", "prog_code": "1019", "batch": "Custom"
                }
                return c_list, c_options
        except Exception:
            pass

    fallback_courses = [
        {"code": "UC012030001", "title": "Personality Development", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203001203", "title": "Design Research", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203006211", "title": "Product Visualization", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203006212", "title": "Form, Aesthetic and Emotion", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203006207", "title": "Studio- Human Centric Design", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203006213", "title": "Design Articulation with AI", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "31203006214", "title": "Indian Design System", "sem": "Sem 3", "prog_code": "1019", "batch": "3rd Sem B.Des PD 2025-29"},
        {"code": "UC013050001", "title": "CAID", "sem": "Sem 5", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
        {"code": "31305001330", "title": "The Art of Delightful Design", "sem": "Sem 5", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
        {"code": "31305006322", "title": "Human Factors", "sem": "Sem 5", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
        {"code": "31305006324", "title": "Studio- Humanizing Technology", "sem": "Sem 5", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
        {"code": "31305006329", "title": "Portfolio Design with Voice Agents", "sem": "Sem 5", "prog_code": "1019", "batch": "5th Sem B.Des PD 2024-28"},
        {"code": "31407001400", "title": "Design Management", "sem": "Sem 7", "prog_code": "1019", "batch": "7th Sem B.Des PD 2023-27"},
        {"code": "31407006405", "title": "System Analysis and Design with Voice Agents", "sem": "Sem 7", "prog_code": "1019", "batch": "7th Sem B.Des PD 2023-27"},
    ]
    fb_opts = {f"{c['title']}  |  [{c['code']}]  ({c['batch']})": c for c in fallback_courses}
    fb_opts["✏️ [Custom / Manual Course Entry]"] = {"code": "", "title": "Custom", "sem": "I", "prog_code": "1019", "batch": "Custom"}
    return fallback_courses, fb_opts

COURSES_DATA, COURSE_OPTIONS = get_courses_roster()

MODULE_ALIASES = {
    "system design": ("31407006405", "System Analysis and Design with Voice Agents"),
    "system analysis and design with voice agents": ("31407006405", "System Analysis and Design with Voice Agents"),
    "system analysis and design": ("31407006405", "System Analysis and Design with Voice Agents"),
    "dm": ("31407001400", "Design Management"),
    "design management": ("31407001400", "Design Management"),
    "studio-humanizing technology": ("31305006324", "Studio- Humanizing Technology"),
    "studio- humanizing technology": ("31305006324", "Studio- Humanizing Technology"),
    "humanizing technology": ("31305006324", "Studio- Humanizing Technology"),
    "studio: design and technology": ("32203002610", "Studio- Design and Technology"),
    "studio- design and technology": ("32203002610", "Studio- Design and Technology"),
    "design and technology": ("32203002610", "Studio- Design and Technology"),
    "ds 1": ("26UDP01020", "Design Studio I"),
    "design studio i": ("26UDP01020", "Design Studio I"),
    "design studio 1": ("26UDP01020", "Design Studio I"),
    "uid common- design research": ("31203001203", "Design Research"),
    "design research": ("31203001203", "Design Research"),
    "personality development": ("UC012030001", "Personality Development"),
    "art of delightful design": ("31305001330", "The Art of Delightful Design"),
    "the art of delightful design": ("31305001330", "The Art of Delightful Design"),
    "portfolio with ai": ("31305006329", "Portfolio Design with Voice Agents"),
    "human factors": ("31305006322", "Human Factors"),
    "caid": ("UC013050001", "CAID"),
    "product visualization": ("31203006211", "Product Visualization"),
    "form, aesthetic and emotion": ("31203006212", "Form, Aesthetic and Emotion"),
    "studio- human centric design": ("31203006207", "Studio- Human Centric Design"),
    "design articulation with ai": ("31203006213", "Design Articulation with AI"),
    "indian design system": ("31203006214", "Indian Design System"),
    "design foundation": ("32101002501", "Design Foundation"),
    "form studies": ("32101002502", "Form Studies"),
    "caid & visualization with ai": ("32101002507", "CAID & Visualization with AI"),
    "emergent technology": ("32101002506", "Emergent Technology"),
    "frugal innovation": ("32101002505", "Frugal Innovation"),
    "design prototyping": ("32101002504", "Design Prototyping"),
    "professional communication": ("32101001500", "Professional Communication"),
    "research methodology": ("32203001604", "Research Methodology"),
    "user experience design": ("32203002609", "User Experience Design"),
    "entrepreneurship": ("32203001603", "Entrepreneurship"),
    "lighting design": ("32203002611", "Lighting Design"),
    "craft and technology": ("32203002612", "Craft and Technology"),
    "experience design": ("31305006325", "Experience Design"),
    "packaging design": ("31305006326", "Packaging Design"),
    "speculative design": ("31305006327", "Speculative design"),
}

def resolve_module(raw_name):
    raw_lower = raw_name.strip().lower()
    clean_words = " ".join(re.sub(r"[^a-zA-Z0-9\s]", " ", raw_lower).split())
    if raw_lower in MODULE_ALIASES:
        return MODULE_ALIASES[raw_lower]
    for k, (code, title) in MODULE_ALIASES.items():
        if k in raw_lower or k in clean_words:
            return code, title
    for c in COURSES_DATA:
        c_title_clean = " ".join(re.sub(r"[^a-zA-Z0-9\s]", " ", c["title"].lower()).split())
        if c_title_clean in clean_words or clean_words in c_title_clean:
            return c["code"], c["title"]
    return "—", raw_name

SEMESTERS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
DAYS_MAP = {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5}
SLOT_TIMINGS_MAP = {
    1: "09:30 TO 11:00",
    2: "11:20 TO 13:10",
    3: "14:05 TO 15:40",
    4: "16:00 TO 17:30",
}

@st.cache_data
def load_master_data():
    files = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~")]
    if not files:
        st.error("No `.xlsx` timetable spreadsheet found in the directory.")
        st.stop()

    file_path = files[0]
    wb = openpyxl.load_workbook(file_path, data_only=True)

    df_fw = pd.read_excel(file_path, sheet_name="Faculty Work Load ")
    raw_faculties = df_fw["Faculty Name "].dropna().unique().tolist()

    official_faculty_list = []
    first_name_to_full = {}
    first_name_to_code = {}

    for f in raw_faculties:
        full = " ".join(str(f).split())
        tokens = [t.lower().strip("().,") for t in full.split()]
        tokens = [t for t in tokens if t not in ["dr", "mr", "ms", "prof"]]
        if tokens:
            fn = tokens[0]
            first_name_to_full[fn] = full
            official_faculty_list.append(full)
            for k_erp, code_erp in DEFAULT_FACULTY_LIST.items():
                if fn in k_erp.lower():
                    first_name_to_code[fn] = code_erp
                    break

    official_faculty_list = sorted(list(set(official_faculty_list)))

    target_sheet = "Morning_Afternoon Updated_Timet"
    matched_sheet_name = next(
        (s for s in wb.sheetnames if target_sheet.lower() in s.lower().strip()),
        wb.sheetnames[1] if len(wb.sheetnames) > 1 else wb.sheetnames[0],
    )
    ws = wb[matched_sheet_name]

    month_map = {
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }
    calendar_cols = []

    for col in range(1, ws.max_column + 1):
        m_val = ws.cell(3, col).value
        d_val = ws.cell(4, col).value
        w_val = ws.cell(5, col).value

        if d_val is not None:
            d_str = str(d_val).strip().split(".")[0]
            if d_str.isdigit():
                day_int = int(d_str)
                m_str = str(m_val or "").strip().lower()
                matched_month = next((month_map[k] for k in month_map if k in m_str), None)
                if matched_month:
                    w_str = str(w_val).strip() if w_val else "Week"
                    try:
                        d_obj = datetime.date(2026, matched_month, day_int)
                        calendar_cols.append({
                            "col": col,
                            "date": d_obj,
                            "week": w_str,
                            "label": f"{d_obj.strftime('%a, %d %b %Y')} ({w_str})",
                        })
                    except ValueError:
                        continue

    teaching_rows = [
        (14, "UG-Sem 3", "A", "Morning", 8),
        (15, "UG-Sem 3", "A", "Afternoon", 8),
        (16, "UG-Sem 3", "B", "Morning", 8),
        (17, "UG-Sem 3", "B", "Afternoon", 8),
        (18, "UG-Sem 3", "C", "Morning", 8),
        (19, "UG-Sem 3", "C", "Afternoon", 8),
        (20, "UG-Sem 3", "D", "Morning", 8),
        (21, "UG-Sem 3", "D", "Afternoon", 8),
        (22, "UG-Sem 3", "E", "Morning", 8),
        (23, "UG-Sem 3", "E", "Afternoon", 8),
        (33, "UG-Sem 5", "A", "Morning", 30),
        (34, "UG-Sem 5", "A", "Afternoon", 30),
        (35, "UG-Sem 5", "B", "Morning", 30),
        (36, "UG-Sem 5", "B", "Afternoon", 30),
        (37, "UG-Sem 5", "C", "Morning", 30),
        (38, "UG-Sem 5", "C", "Afternoon", 30),
        (39, "UG-Sem 5", "D", "Morning", 30),
        (40, "UG-Sem 5", "D", "Afternoon", 30),
        (51, "UG-Sem 7", "A", "Lead", 48),
        (52, "UG-Sem 7", "A", "Assisting", 48),
        (53, "UG-Sem 7", "B", "Lead", 48),
        (54, "UG-Sem 7", "B", "Assisting", 48),
        (55, "UG-Sem 7", "C", "Lead", 48),
        (56, "UG-Sem 7", "C", "Assisting", 48),
        (65, "PG-Sem 1", "Combined", "Full Day", 62),
        (74, "PG-Sem 3", "Combined", "Full Day", 71),
    ]

    faculty_day_schedule = {f: {} for f in official_faculty_list}
    cohort_day_schedule = {}

    for c_info in calendar_cols:
        col = c_info["col"]
        dt = c_info["date"]
        if dt not in cohort_day_schedule:
            cohort_day_schedule[dt] = {}

        for r, cohort, sec, slot, mod_r in teaching_rows:
            cell_val = ws.cell(r, col).value
            if cell_val is not None and str(cell_val).strip() not in ["", "None", "-"]:
                txt = str(cell_val).strip()
                mod_name = str(ws.cell(mod_r, col).value or "").strip()
                clean_mod = " ".join(mod_name.split())

                cell_tokens = [re.sub(r"[^a-zA-Z]", "", w).lower() for w in txt.split()]
                matched_faculty = None
                faculty_code = None

                for tok in cell_tokens:
                    if tok in first_name_to_full:
                        matched_faculty = first_name_to_full[tok]
                        faculty_code = first_name_to_code.get(tok, "")
                        break

                if not matched_faculty:
                    matched_faculty = txt
                    faculty_code = "VISITING"

                if matched_faculty not in faculty_day_schedule:
                    faculty_day_schedule[matched_faculty] = {}

                detail = f"{cohort} | Sec {sec} ({slot}) - {clean_mod}"
                if dt not in faculty_day_schedule[matched_faculty]:
                    faculty_day_schedule[matched_faculty][dt] = []
                if detail not in faculty_day_schedule[matched_faculty][dt]:
                    faculty_day_schedule[matched_faculty][dt].append(detail)

                if cohort not in cohort_day_schedule[dt]:
                    cohort_day_schedule[dt][cohort] = {}
                if sec not in cohort_day_schedule[dt][cohort]:
                    cohort_day_schedule[dt][cohort][sec] = {
                        "module": clean_mod,
                        "morning": None,
                        "afternoon": None,
                        "faculty_name_m": None,
                        "faculty_name_a": None,
                        "faculties_all_day": [],
                    }

                if slot in ["Morning", "Full Day", "Lead"]:
                    cohort_day_schedule[dt][cohort][sec]["morning"] = faculty_code
                    cohort_day_schedule[dt][cohort][sec]["faculty_name_m"] = matched_faculty
                if slot in ["Afternoon", "Full Day", "Assisting"]:
                    cohort_day_schedule[dt][cohort][sec]["afternoon"] = faculty_code
                    cohort_day_schedule[dt][cohort][sec]["faculty_name_a"] = matched_faculty

                if (
                    cohort == "UG-Sem 7"
                    and matched_faculty not in cohort_day_schedule[dt][cohort][sec]["faculties_all_day"]
                ):
                    cohort_day_schedule[dt][cohort][sec]["faculties_all_day"].append(matched_faculty)

    return calendar_cols, official_faculty_list, faculty_day_schedule, cohort_day_schedule, ws

try:
    calendar_cols, official_faculty_list, faculty_schedule, cohort_day_schedule, ws = load_master_data()
except Exception as e:
    st.error(f"Error loading master dataset: {e}")
    st.stop()


# ==========================================================
# HELPER: RANGE AVAILABILITY & FIRST CLASH DETECTOR
# ==========================================================
def check_faculty_availability_in_range(fac_name, start_date, end_date, session_type):
    """
    Checks faculty availability across a date range for a given session.
    Returns:
      is_fully_free (bool)
      free_until (date or None): The last date they are free before their first class
      first_clash_date (date or None): The first date they are busy
      first_clash_detail (str): Details of the first conflicting class
    """
    if not fac_name or fac_name in ["-- None / Leave Empty --", "✏️ [Manual / Custom Entry]"]:
        return True, None, None, ""

    fac_record = next((f for f in RAW_DB_FACULTIES if f["name"] == fac_name), {})
    alias = (fac_record.get("first_name_alias") or "").lower()
    fac_tokens = [t for t in re.findall(r"[a-zA-Z]+", fac_name.lower()) if t not in ["dr", "mr", "ms", "prof"]]

    cur = start_date
    free_days = []
    first_clash_date = None
    first_clash_detail = ""

    while cur <= end_date:
        if cur.weekday() < 5:  # Monday to Friday
            # 1. Check if active in department (departure date)
            if not is_faculty_active_on(fac_record, cur):
                first_clash_date = cur
                first_clash_detail = f"Contract / tenure ended on {fac_record.get('active_until')}"
                break

            # 2. Check scheduled classes on this date
            day_busy = False
            day_detail = ""
            for raw_f, dates in faculty_schedule.items():
                classes = dates.get(cur, [])
                if not classes:
                    continue
                
                raw_tokens = [t for t in re.findall(r"[a-zA-Z]+", raw_f.lower()) if t not in ["dr", "mr", "ms", "prof", "sir"]]
                matches = (alias and alias in raw_tokens) or any(len(t) > 2 and t in raw_tokens for t in fac_tokens) or (fac_name == raw_f)

                if matches:
                    if "Morning" in session_type:
                        conflicts = [c for c in classes if "(Morning)" in c or "PG-" in c or "UG-Sem 7" in c]
                    elif "Afternoon" in session_type:
                        conflicts = [c for c in classes if "(Afternoon)" in c or "PG-" in c or "UG-Sem 7" in c]
                    else:
                        conflicts = classes

                    if conflicts:
                        day_busy = True
                        day_detail = conflicts[0]
                        break

            if day_busy:
                first_clash_date = cur
                first_clash_detail = day_detail
                break
            else:
                free_days.append(cur)

        cur += datetime.timedelta(days=1)

    if first_clash_date is None:
        return True, end_date, None, ""

    free_until = free_days[-1] if free_days else None
    return False, free_until, first_clash_date, first_clash_detail


st.title("UID Department of Industrial Design")
st.caption("Master Academic Schedule, Cloud Timetable Management & Availability Engine")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "🔍 Faculty Availability",
    "👤 Individual Schedule",
    "📋 Faculty Course Workload",
    "⚡ Live Cloud Timetable Editor",
    "🗓️ Weekly Timetable",
    "🏁 Course Completion Dates",
    "📈 Subject Hours & Progress",
    "⚡ Auto Timetable Export",
    "🛠️ Custom Generator Form",
])

# ==========================================================
# TAB 1: AVAILABILITY ENGINE
# ==========================================================
with tab1:
    date_to_week = {c["date"]: c["week"] for c in calendar_cols}
    valid_dates = [c["date"] for c in calendar_cols]
    min_date = valid_dates[0]
    max_date = valid_dates[-1]

    today = datetime.date.today()
    default_date = today if min_date <= today <= max_date else (min_date if today < min_date else max_date)

    selected_date = st.date_input(
        "Select Academic Date:",
        value=default_date,
        min_value=min_date,
        max_value=max_date,
        key="avail_date_picker",
    )

    current_week = date_to_week.get(selected_date, "Non-Instructional / Holiday")

    # 1. Identify busy instructors and extract clean tokens
    busy_members = []
    busy_tokens = set()

    for fac_raw, dates in faculty_schedule.items():
        classes = dates.get(selected_date, [])
        if classes and not any(skip in fac_raw.lower() for skip in ["mid term", "no faculty", "workshop", "tours", "holiday"]):
            busy_members.append((fac_raw, classes))
            for token in re.findall(r"[a-zA-Z]+", fac_raw.lower()):
                if token not in ["dr", "mr", "ms", "prof", "sir", "updated"]:
                    busy_tokens.add(token)

    # 2. Filter free faculty strictly from official database roster
    free_members = []
    active_roster = DATABASE_FACULTY_NAMES if DATABASE_FACULTY_NAMES else official_faculty_list

    for fac in active_roster:
        fac_record = next((f for f in RAW_DB_FACULTIES if f["name"] == fac), {})
        is_active = is_faculty_active_on(fac_record, selected_date)
        
        # Check substitution toggle (defaults to True if column not yet added)
        can_substitute = fac_record.get("is_available_for_substitution", True)
        if can_substitute is None:
            can_substitute = True

        # Never show Atul Kedia or anyone with can_substitute = False
        if not can_substitute or "atul kedia" in fac.lower():
            continue

        alias = (fac_record.get("first_name_alias") or "").lower()
        fac_tokens = [t for t in re.findall(r"[a-zA-Z]+", fac.lower()) if t not in ["dr", "mr", "ms", "prof"]]

        # Matches if alias or any substantive name token (like 'siddharth') is in busy_tokens
        is_busy = (
            (alias and alias in busy_tokens)
            or any((len(t) > 2 and t in busy_tokens) for t in fac_tokens)
            or (fac in [b[0] for b in busy_members])
        )

        if is_active and not is_busy:
            free_members.append(fac)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Week", current_week)
    m2.metric("Date", selected_date.strftime("%d %B %Y"))
    m3.metric("Available / Free", len(free_members))
    m4.metric("Scheduled in Class", len(busy_members))

    st.markdown("---")
    col_free, col_busy = st.columns([1, 1.2])

    with col_free:
        st.subheader(f"✅ Free Faculty ({len(free_members)})")
        st.caption("Official department faculty available for substitution:")
        f1, f2 = st.columns(2)
        for idx, fac in enumerate(free_members):
            target = f1 if idx % 2 == 0 else f2
            target.markdown(f"• **{fac}** &nbsp; <span class='badge-free'>Available</span>", unsafe_allow_html=True)

    with col_busy:
        st.subheader(f"🔒 Scheduled Faculty ({len(busy_members)})")
        st.caption("Instructors with active class assignments:")
        for fac, sessions in busy_members:
            details_html = "<br>".join([f"&nbsp;&nbsp;↳ {s}" for s in sessions])
            st.markdown(
                f"""
                <div class="faculty-card">
                    <b>{fac}</b> &nbsp; <span class="badge-busy">In Class</span><br>
                    <small style="color: #475569;">{details_html}</small>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ==========================================================
# TAB 2: INDIVIDUAL LOOKUP & CONTACT HOURS
# ==========================================================
with tab2:
    faculty_dropdown_options = DATABASE_FACULTY_NAMES if DATABASE_FACULTY_NAMES else official_faculty_list
    selected_fac = st.selectbox("Select Faculty Member:", options=faculty_dropdown_options, key="indiv_fac_select")

    sched = faculty_schedule.get(selected_fac, {})

    if sched:
        total_contact_hours = 0.0
        records = []
        course_stats = {}

        for c in calendar_cols:
            dt = c["date"]
            if dt in sched:
                sessions = sched[dt]

                if dt < datetime.date(2026, 9, 1):
                    full_day_hrs = 7.0
                    half_day_hrs = 3.5
                else:
                    full_day_hrs = 6.5
                    half_day_hrs = 3.25

                is_thursday = dt.weekday() == 3
                thu_ug_off = is_thursday and (dt >= datetime.date(2026, 8, 1))

                has_pg = any("PG-" in s for s in sessions)
                has_ug_full = any("UG-Sem 7" in s for s in sessions)
                has_morning = any("(Morning)" in s for s in sessions)
                has_afternoon = any("(Afternoon)" in s for s in sessions)

                if has_pg:
                    daily_hours = full_day_hrs
                elif has_ug_full:
                    daily_hours = half_day_hrs if thu_ug_off else full_day_hrs
                else:
                    daily_hours = 0.0
                    if has_morning:
                        daily_hours += half_day_hrs
                    if has_afternoon and not thu_ug_off:
                        daily_hours += half_day_hrs
                    daily_hours = min(daily_hours, full_day_hrs)

                total_contact_hours += daily_hours

                day_modules = set()
                for s in sessions:
                    parts = s.split(" - ", 1)
                    raw_mod = parts[1].strip() if len(parts) > 1 else s.strip()
                    code, canonical_title = resolve_module(raw_mod)
                    day_modules.add((code, canonical_title))

                split_hours = daily_hours / len(day_modules) if day_modules else daily_hours
                for code, title in day_modules:
                    key = (code, title)
                    if key not in course_stats:
                        course_stats[key] = {"days": set(), "hours": 0.0}
                    course_stats[key]["days"].add(dt)
                    course_stats[key]["hours"] += split_hours

                records.append({
                    "Date": dt.strftime("%d-%b-%Y"),
                    "Week": c["week"],
                    "Day": dt.strftime("%A"),
                    "Daily Contact Hours": f"{daily_hours:.2f} hrs",
                    "Assigned Sessions": "; ".join(sessions),
                })

        clean_fac_name = re.sub(r"[^a-zA-Z0-9]", "_", selected_fac)

        st.markdown(f"#### 📊 Workload Summary for **{selected_fac}**")
        summary_data = [
            {"Metric": "Total Teaching Days", "Value": f"{len(records)} days"},
            {"Metric": "Total Contact / Teaching Hours", "Value": f"{total_contact_hours:.2f} hrs"},
        ]
        df_summary = pd.DataFrame(summary_data)
        st.table(df_summary)

        csv_summary = df_summary.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Workload Summary (.csv)",
            data=csv_summary,
            file_name=f"Summary_{clean_fac_name}_ODD_2026.csv",
            mime="text/csv",
            key="summary_download_btn",
        )

        st.markdown("---")
        st.markdown("#### 📚 Course-wise Workload Breakdown")
        course_breakdown_rows = []
        for (code, title), data in sorted(course_stats.items(), key=lambda x: len(x[1]["days"]), reverse=True):
            course_breakdown_rows.append({
                "Course Code": code,
                "Course Name": title,
                "Number of Days": f"{len(data['days'])} days",
                "Teaching Hours": f"{data['hours']:.2f} hrs",
            })

        df_courses = pd.DataFrame(course_breakdown_rows)
        st.table(df_courses)

        csv_courses = df_courses.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Course Breakdown (.csv)",
            data=csv_courses,
            file_name=f"Course_Breakdown_{clean_fac_name}_ODD_2026.csv",
            mime="text/csv",
            key="course_download_btn",
        )

        st.markdown("---")
        st.markdown("#### 📅 Teaching Session Breakdown")
        df_indiv = pd.DataFrame(records)
        st.table(df_indiv)

        csv_indiv = df_indiv.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Session Breakdown (.csv)",
            data=csv_indiv,
            file_name=f"Sessions_{clean_fac_name}_ODD_2026.csv",
            mime="text/csv",
            key="indiv_download_btn",
        )
    else:
        st.info(f"No active teaching assignments found for {selected_fac}.")

# ==========================================================
# TAB 3: FACULTY COURSE WORKLOAD & GRAND TOTALS
# ==========================================================
with tab3:
    st.subheader("📋 Faculty-Wise Subject Allocations & Totals")
    st.caption("Itemized list of all subjects taught by faculty members with total teaching days, contact hours, and grand totals.")

    view_mode = st.radio(
        "Select View Mode:",
        ["Individual Faculty Breakdown", "All Faculty Consolidated Master View"],
        horizontal=True,
        key="fac_wl_view_mode",
    )

    def compute_faculty_course_stats(fac_name):
        sched_data = faculty_schedule.get(fac_name, {})
        stats = {}

        for c_info in calendar_cols:
            dt = c_info["date"]
            if dt in sched_data:
                sessions = sched_data[dt]

                if dt < datetime.date(2026, 9, 1):
                    full_day_hrs = 7.0
                    half_day_hrs = 3.5
                else:
                    full_day_hrs = 6.5
                    half_day_hrs = 3.25

                is_thu = dt.weekday() == 3
                thu_ug_off = is_thu and (dt >= datetime.date(2026, 8, 1))

                has_pg = any("PG-" in s for s in sessions)
                has_ug_full = any("UG-Sem 7" in s for s in sessions)
                has_m = any("(Morning)" in s for s in sessions)
                has_a = any("(Afternoon)" in s for s in sessions)

                if has_pg:
                    daily_hrs = full_day_hrs
                elif has_ug_full:
                    daily_hrs = half_day_hrs if thu_ug_off else full_day_hrs
                else:
                    daily_hrs = 0.0
                    if has_m:
                        daily_hrs += half_day_hrs
                    if has_a and not thu_ug_off:
                        daily_hrs += half_day_hrs
                    daily_hrs = min(daily_hrs, full_day_hrs)

                day_courses = set()
                for s in sessions:
                    parts = s.split(" - ", 1)
                    raw_mod = parts[1].strip() if len(parts) > 1 else s.strip()
                    code, title = resolve_module(raw_mod)
                    day_courses.add((code, title))

                split_hrs = daily_hrs / len(day_courses) if day_courses else daily_hrs
                for code, title in day_courses:
                    key = (code, title)
                    if key not in stats:
                        stats[key] = {"days": set(), "hours": 0.0}
                    stats[key]["days"].add(dt)
                    stats[key]["hours"] += split_hrs

        return stats

    if view_mode == "Individual Faculty Breakdown":
        selected_fac_wl = st.selectbox(
            "Select Faculty Member:",
            options=official_faculty_list,
            key="fac_wl_dropdown",
        )
        course_stats = compute_faculty_course_stats(selected_fac_wl)

        if course_stats:
            rows = []
            all_unique_days = set()
            total_hours = 0.0

            for (code, title), data in sorted(course_stats.items(), key=lambda x: len(x[1]["days"]), reverse=True):
                n_days = len(data["days"])
                hrs = data["hours"]
                all_unique_days.update(data["days"])
                total_hours += hrs

                rows.append({
                    "Course Code": code,
                    "Course Name": title,
                    "Teaching Days": f"{n_days} days",
                    "Teaching Hours": f"{hrs:.2f} hrs",
                })

            rows.append({
                "Course Code": "—",
                "Course Name": "TOTAL",
                "Teaching Days": f"{len(all_unique_days)} days",
                "Teaching Hours": f"{total_hours:.2f} hrs",
            })

            df_single_wl = pd.DataFrame(rows)
            st.table(df_single_wl)

            clean_name = re.sub(r"[^a-zA-Z0-9]", "_", selected_fac_wl)
            csv_single_wl = df_single_wl.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"📥 Download Workload Table for {selected_fac_wl} (.csv)",
                data=csv_single_wl,
                file_name=f"Faculty_Courses_{clean_name}.csv",
                mime="text/csv",
                key="download_single_fac_wl",
            )
        else:
            st.info(f"No scheduled courses found for {selected_fac_wl}.")
    else:
        all_rows = []
        for fac in official_faculty_list:
            c_stats = compute_faculty_course_stats(fac)
            if not c_stats:
                continue
            fac_unique_days = set()
            fac_total_hrs = 0.0

            for (code, title), data in sorted(c_stats.items(), key=lambda x: len(x[1]["days"]), reverse=True):
                fac_unique_days.update(data["days"])
                fac_total_hrs += data["hours"]
                all_rows.append({
                    "Faculty Name": fac,
                    "Course Code": code,
                    "Course Name": title,
                    "Teaching Days": len(data["days"]),
                    "Teaching Hours": round(data["hours"], 2),
                })

            all_rows.append({
                "Faculty Name": f"TOTAL: {fac}",
                "Course Code": "—",
                "Course Name": "—",
                "Teaching Days": len(fac_unique_days),
                "Teaching Hours": round(fac_total_hrs, 2),
            })

        if all_rows:
            df_all_wl = pd.DataFrame(all_rows)
            st.dataframe(df_all_wl, use_container_width=True, height=600)

            csv_all_wl = df_all_wl.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Department-Wide Master Allocation (.csv)",
                data=csv_all_wl,
                file_name="Department_Faculty_Course_Workload_Master.csv",
                mime="text/csv",
                key="download_all_fac_wl",
            )

# ==========================================================
# TAB 4: LIVE CLOUD TIMETABLE & SUBJECT/FACULTY EDITOR (SUPABASE)
# ==========================================================
with tab4:
    st.subheader("⚡ Live Timetable Module Scheduler")
    st.caption("Assign courses and instructors across dates and sections with live cloud synchronization.")

    if not supabase_client:
        st.warning("⚠️ Supabase is not connected. Configure `SUPABASE_URL` and `SUPABASE_KEY` in `.streamlit/secrets.toml`.")
    else:
        # Administrative quick-add drawers
        with st.expander("⚙️ Roster & Course Quick Tools (Add New)"):
            c_adm1, c_adm2 = st.columns(2)
            with c_adm1:
                st.markdown("**➕ Add Faculty Member**")
                n_f_code = st.text_input("Faculty ERP Code (e.g., 15999)", key="tb4_f_code")
                n_f_name = st.text_input("Full Name", key="tb4_f_name")
                n_f_end = st.date_input("Last Working Date (Optional)", value=None, key="tb4_f_end")
                if st.button("Save Faculty to Supabase", key="tb4_btn_save_f"):
                    if n_f_code and n_f_name:
                        supabase_client.table("faculties").upsert({
                            "faculty_code": n_f_code.strip(),
                            "name": n_f_name.strip(),
                            "first_name_alias": n_f_name.split()[0].lower(),
                            "active_until": str(n_f_end) if n_f_end else None,
                            "is_active": True,
                            "is_available_for_substitution": True
                        }).execute()
                        st.success(f"✓ Saved {n_f_name}!")
                        st.cache_data.clear()
                        st.rerun()

            with c_adm2:
                st.markdown("**📚 Add New Subject**")
                n_c_code = st.text_input("Course Code (e.g., 31204006220)", key="tb4_c_code")
                n_c_title = st.text_input("Course Title", key="tb4_c_title")
                n_c_sem = st.selectbox("Semester", ["Sem 2", "Sem 4", "Sem 6", "Sem 3", "Sem 5", "Sem 7", "PG Sem 1", "PG Sem 3"], key="tb4_c_sem")
                if st.button("Save Subject to Supabase", key="tb4_btn_save_c"):
                    if n_c_code and n_c_title:
                        supabase_client.table("courses").upsert({
                            "course_code": n_c_code.strip(),
                            "title": n_c_title.strip(),
                            "semester": n_c_sem,
                            "semester_type": "EVEN" if n_c_sem in ["Sem 2", "Sem 4", "Sem 6"] else "ODD",
                            "program_code": "1019" if not n_c_sem.startswith("PG") else "1025"
                        }).execute()
                        st.success(f"✓ Saved {n_c_title}!")
                        st.cache_data.clear()
                        st.rerun()

        st.markdown("---")

        # Global Session Selection Checkboxes & Target Cohort
        st.markdown("<p style='font-size: 0.88rem; font-weight: 700; color: #1e293b; margin-bottom: 6px;'>SELECT ACTIVE SESSIONS TO CONFIGURE</p>", unsafe_allow_html=True)
        top_c1, top_c2, top_c3 = st.columns([1.2, 1.2, 1.2])
        with top_c1:
            plan_morning = st.checkbox("🌅 Morning Session (S1 + S2)", value=True, key="chk_plan_m")
        with top_c2:
            plan_afternoon = st.checkbox("🌇 Afternoon Session (S3 + S4)", value=False, key="chk_plan_a")
        with top_c3:
            target_cohort = st.selectbox(
                "Target Cohort",
                ["UG-Sem 3", "UG-Sem 5", "UG-Sem 7", "PG-Sem 1", "PG-Sem 3", "UG-Sem 2", "UG-Sem 4", "UG-Sem 6"],
                key="wf_target_cohort"
            )

        if not plan_morning and not plan_afternoon:
            st.info("💡 Please check at least one session (Morning and/or Afternoon) to plan.")

        # Common Dates & Sections Row
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        d_col1, d_col2, d_col3 = st.columns([1.2, 1.2, 0.8])
        with d_col1:
            st.markdown("<p style='font-size: 0.82rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 4px;'>START DATE</p>", unsafe_allow_html=True)
            wf_start_date = st.date_input("START DATE", value=datetime.date.today(), label_visibility="collapsed", key="wf_start_date")
        with d_col2:
            st.markdown("<p style='font-size: 0.82rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 4px;'>END DATE</p>", unsafe_allow_html=True)
            wf_end_date = st.date_input("END DATE", value=datetime.date.today() + datetime.timedelta(days=4), label_visibility="collapsed", key="wf_end_date")
        with d_col3:
            st.markdown("<p style='font-size: 0.82rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 4px;'>NUMBER OF SECTIONS</p>", unsafe_allow_html=True)
            wf_num_sections = st.number_input("NUMBER OF SECTIONS", min_value=1, max_value=5, value=4, step=1, label_visibility="collapsed", key="wf_num_sections")

        # Red text button styling for Check Availability triggers
        st.markdown(
            """
            <style>
                div[data-testid*="chk_avail_btn_"] button {
                    color: #dc2626 !important;
                    background: transparent !important;
                    border: none !important;
                    padding: 0px !important;
                    font-size: 0.82rem !important;
                    font-weight: 700 !important;
                    text-decoration: underline !important;
                    cursor: pointer !important;
                    height: auto !important;
                    margin-top: -6px !important;
                    box-shadow: none !important;
                }
                div[data-testid*="chk_avail_btn_"] button:hover {
                    color: #991b1b !important;
                    background: transparent !important;
                }
                div[data-testid*="chk_avail_btn_"] button:focus:not(:focus-visible) {
                    box-shadow: none !important;
                }
            </style>
            """,
            unsafe_allow_html=True
        )

        all_section_letters = ["A", "B", "C", "D", "E"]
        active_sections = all_section_letters[:int(wf_num_sections)]
        course_keys = list(COURSE_OPTIONS.keys())
        faculty_opts = list(DEFAULT_FACULTY_LIST.keys())

        def render_section_slot_box(sec, prefix, session_label):
            st.markdown(f"<p style='font-size: 0.86rem; font-weight: 600; color: #1e293b; margin-bottom: 4px;'>Faculty Section {sec}</p>", unsafe_allow_html=True)
            f_choice = st.selectbox(
                f"Faculty Section {sec} ({prefix})",
                options=faculty_opts,
                key=f"wf_{prefix}_fac_sec_{sec}",
                label_visibility="collapsed"
            )
            selected_full_name = f_choice.split(" (")[0].strip()

            chk_clicked = st.button("Check Availability", key=f"chk_avail_btn_{prefix}_{sec}")
            if chk_clicked:
                if f_choice in ["-- None / Leave Empty --", "✏️ [Manual / Custom Entry]"]:
                    st.caption("ℹ️ Please select an official faculty member to check availability.")
                else:
                    is_free, free_until, clash_date, clash_detail = check_faculty_availability_in_range(
                        selected_full_name, wf_start_date, wf_end_date, session_label
                    )

                    if is_free:
                        st.markdown(
                            f"<span style='color: #16a34a; font-weight: 600; font-size: 0.82rem;'>✓ Fully Free ({wf_start_date.strftime('%d %b')} – {wf_end_date.strftime('%d %b')})</span>",
                            unsafe_allow_html=True
                        )
                    else:
                        if free_until:
                            st.markdown(
                                f"<div style='background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 6px 10px; margin-top: 4px; font-size: 0.82rem; color: #991b1b;'>"
                                f"<b>⚠️ Not Available full duration:</b><br>"
                                f"• Free until: <b>{free_until.strftime('%d %b %Y')}</b><br>"
                                f"• Busy on: <b>{clash_date.strftime('%d %b %Y')}</b><br>"
                                f"<small style='color: #7f1d1d;'>({clash_detail})</small>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                        else:
                            st.markdown(
                                f"<div style='background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 6px 10px; margin-top: 4px; font-size: 0.82rem; color: #991b1b;'>"
                                f"<b>⚠️ Busy from Day 1 ({clash_date.strftime('%d %b %Y')}):</b><br>"
                                f"<small style='color: #7f1d1d;'>{clash_detail}</small>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
            return DEFAULT_FACULTY_LIST[f_choice]

        # ------------------------------------------------------
        # 1. MORNING SESSION CARD (S1 + S2)
        # ------------------------------------------------------
        m_course_info = None
        fac_per_section_m = {}

        if plan_morning:
            st.markdown("<hr style='margin: 22px 0; border: none; border-top: 1px solid #cbd5e1;'>", unsafe_allow_html=True)
            st.markdown(
                """
                <div style="background-color: #f8fafc; border: 1px solid #94a3b8; border-left: 6px solid #2563eb; border-radius: 10px; padding: 18px 24px; margin-bottom: 14px;">
                    <h3 style="margin: 0; color: #0f172a; font-size: 1.5rem; font-weight: 800;">Subject (Morning Session: S1 + S2)</h3>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.markdown("<p style='font-size: 0.82rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 4px;'>SUBJECT NAME (MORNING)</p>", unsafe_allow_html=True)
            m_course_key = st.selectbox("SUBJECT NAME (MORNING)", options=course_keys, index=0, label_visibility="collapsed", key="wf_m_course_select")
            m_course_info = COURSE_OPTIONS[m_course_key]

            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

            sec_r1 = active_sections[:3]
            if sec_r1:
                cols_r1 = st.columns(3)
                for idx, sec in enumerate(sec_r1):
                    with cols_r1[idx]:
                        fac_per_section_m[sec] = render_section_slot_box(sec, "m", "Morning Session (S1 + S2)")

            sec_r2 = active_sections[3:]
            if sec_r2:
                st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                cols_r2 = st.columns(3)
                for idx, sec in enumerate(sec_r2):
                    with cols_r2[idx]:
                        fac_per_section_m[sec] = render_section_slot_box(sec, "m", "Morning Session (S1 + S2)")

        # ------------------------------------------------------
        # 2. AFTERNOON SESSION CARD (S3 + S4)
        # ------------------------------------------------------
        a_course_info = None
        fac_per_section_a = {}

        if plan_afternoon:
            st.markdown("<hr style='margin: 22px 0; border: none; border-top: 1px solid #cbd5e1;'>", unsafe_allow_html=True)
            st.markdown(
                """
                <div style="background-color: #f8fafc; border: 1px solid #94a3b8; border-left: 6px solid #d97706; border-radius: 10px; padding: 18px 24px; margin-bottom: 14px;">
                    <h3 style="margin: 0; color: #0f172a; font-size: 1.5rem; font-weight: 800;">Subject (Afternoon Session: S3 + S4)</h3>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.markdown("<p style='font-size: 0.82rem; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 4px;'>SUBJECT NAME (AFTERNOON)</p>", unsafe_allow_html=True)
            def_a_idx = 1 if len(course_keys) > 1 else 0
            a_course_key = st.selectbox("SUBJECT NAME (AFTERNOON)", options=course_keys, index=def_a_idx, label_visibility="collapsed", key="wf_a_course_select")
            a_course_info = COURSE_OPTIONS[a_course_key]

            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

            sec_r1 = active_sections[:3]
            if sec_r1:
                cols_r1 = st.columns(3)
                for idx, sec in enumerate(sec_r1):
                    with cols_r1[idx]:
                        fac_per_section_a[sec] = render_section_slot_box(sec, "a", "Afternoon Session (S3 + S4)")

            sec_r2 = active_sections[3:]
            if sec_r2:
                st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                cols_r2 = st.columns(3)
                for idx, sec in enumerate(sec_r2):
                    with cols_r2[idx]:
                        fac_per_section_a[sec] = render_section_slot_box(sec, "a", "Afternoon Session (S3 + S4)")

        st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

        # Optional Settings Drawer
        with st.expander("🛠️ Advanced Options (Room Allocation & Thursday Half-Day)"):
            opt_c1, opt_c2, opt_c3 = st.columns(3)
            with opt_c1:
                wf_block = opt_c1.text_input("Academic Block", value="F", key="wf_block_val")
            with opt_c2:
                wf_room = opt_c2.text_input("Default Room", placeholder="e.g. F2, E10", key="wf_room_val")
            with opt_c3:
                wf_skip_thu_afternoon = opt_c3.checkbox("Thursday Afternoon Off (S3 & S4 Off for UG)", value=True, key="wf_thu_off_chk")

        # ------------------------------------------------------
        # SUBMIT & SYNC TO SUPABASE
        # ------------------------------------------------------
        if st.button("🚀 Push Schedule to Supabase Cloud", type="primary", use_container_width=True, key="wf_btn_sync"):
            if not plan_morning and not plan_afternoon:
                st.error("Please select at least one session (Morning and/or Afternoon).")
            elif wf_start_date > wf_end_date:
                st.error("Start Date must be before or equal to End Date.")
            elif plan_morning and not m_course_info.get("code"):
                st.error("Please select a valid Morning subject.")
            elif plan_afternoon and not a_course_info.get("code"):
                st.error("Please select a valid Afternoon subject.")
            else:
                records_to_insert = []
                cur = wf_start_date

                while cur <= wf_end_date:
                    if cur.weekday() < 5:  # Monday to Friday
                        date_str = str(cur)
                        is_thu = cur.weekday() == 3

                        # Process Morning Slots (S1, S2)
                        if plan_morning:
                            for s_num in [1, 2]:
                                for sec in active_sections:
                                    f_code = fac_per_section_m.get(sec)
                                    records_to_insert.append({
                                        "date": date_str,
                                        "cohort": target_cohort,
                                        "section": sec,
                                        "slot_num": s_num,
                                        "slot_time": SLOT_TIMINGS_MAP[s_num],
                                        "shift_timing": "09:30 TO 17:30",
                                        "course_code": m_course_info["code"],
                                        "faculty_code": f_code if f_code else None,
                                        "academic_block": wf_block,
                                        "room": wf_room if wf_room else None,
                                        "classification": "PRACTICAL",
                                        "is_off": False
                                    })

                        # Process Afternoon Slots (S3, S4)
                        if plan_afternoon:
                            for s_num in [3, 4]:
                                is_off = (is_thu and wf_skip_thu_afternoon and not target_cohort.startswith("PG"))
                                for sec in active_sections:
                                    f_code = fac_per_section_a.get(sec)
                                    records_to_insert.append({
                                        "date": date_str,
                                        "cohort": target_cohort,
                                        "section": sec,
                                        "slot_num": s_num,
                                        "slot_time": SLOT_TIMINGS_MAP[s_num],
                                        "shift_timing": "09:30 TO 17:30",
                                        "course_code": a_course_info["code"],
                                        "faculty_code": None if is_off else (f_code if f_code else None),
                                        "academic_block": wf_block,
                                        "room": wf_room if wf_room else None,
                                        "classification": "THEORY",
                                        "is_off": is_off
                                    })

                    cur += datetime.timedelta(days=1)

                if records_to_insert:
                    try:
                        supabase_client.table("timetable_slots").upsert(
                            records_to_insert,
                            on_conflict="date,cohort,section,slot_num"
                        ).execute()
                        st.success(f"✓ Successfully scheduled {len(records_to_insert)} slots for {target_cohort} in Supabase!")
                        st.cache_data.clear()
                    except Exception as ex:
                        st.error(f"Error updating Supabase database: {ex}")
                else:
                    st.warning("No instructional days found in selected date range.")
                    
# ==========================================================
# TAB 5: WEEKLY CALENDAR TIMETABLE VIEW
# ==========================================================
with tab5:
    st.subheader("🗓️ Weekly Department Timetable")
    st.caption("Comprehensive weekly grid displaying all sections with Sessions 1–4 color-coded.")

    weeks_dict = {}
    for c_info in calendar_cols:
        w = c_info["week"]
        if w not in weeks_dict:
            weeks_dict[w] = []
        weeks_dict[w].append(c_info)

    week_options = []
    week_data_map = {}
    def_idx = 0
    today = datetime.date.today()

    for idx, (w, days) in enumerate(weeks_dict.items()):
        d_start = days[0]["date"].strftime("%d %b")
        d_end = days[-1]["date"].strftime("%d %b %Y")
        lbl = f"{w}  ({d_start} - {d_end})"
        week_options.append(lbl)
        week_data_map[lbl] = (w, days)
        if any(d["date"] == today for d in days):
            def_idx = idx

    if "selected_cohort" not in st.session_state:
        st.session_state["selected_cohort"] = "UG-Sem 3"

    col_week, col_cohorts = st.columns([1.1, 2.4])

    with col_week:
        selected_week_lbl = st.selectbox("Select Week:", options=week_options, index=def_idx, key="weekly_select")
        selected_w_name, week_days = week_data_map[selected_week_lbl]

    with col_cohorts:
        st.markdown("<p style='font-size: 0.88rem; font-weight: 500; margin-bottom: 6px;'>Select Cohort:</p>", unsafe_allow_html=True)
        cohort_buttons = ["UG-Sem 3", "UG-Sem 5", "UG-Sem 7", "PG-Sem 1", "PG-Sem 3", "All"]
        btn_cols = st.columns(len(cohort_buttons))

        for b_col, opt in zip(btn_cols, cohort_buttons):
            is_active = st.session_state["selected_cohort"] == opt
            if b_col.button(opt, type="primary" if is_active else "secondary", use_container_width=True, key=f"btn_ch_{opt}"):
                st.session_state["selected_cohort"] = opt
                st.rerun()

    cohort_filter = "All Cohorts" if st.session_state["selected_cohort"] == "All" else st.session_state["selected_cohort"]
    display_days = [d for d in week_days if d["date"].weekday() < 5]

    cohort_sections = {
        "UG-Sem 3": ["A", "B", "C", "D", "E"],
        "UG-Sem 5": ["A", "B", "C", "D"],
        "UG-Sem 7": ["A", "B", "C"],
        "PG-Sem 1": ["Combined"],
        "PG-Sem 3": ["Combined"],
    }

    active_cohorts = [cohort_filter] if cohort_filter != "All Cohorts" else list(cohort_sections.keys())

    html = ['<table class="wt-table">']
    html.append("<thead><tr>")
    html.append('<th class="wt-th" style="width: 12%;">Cohort / Section</th>')
    for d in display_days:
        d_name = d["date"].strftime("%a")
        d_num = d["date"].strftime("%d %b")
        html.append(f'<th class="wt-th">{d_name}<br><small>{d_num}</small></th>')
    html.append("</tr></thead><tbody>")

    for ch in active_cohorts:
        for sec in cohort_sections.get(ch, []):
            row_lbl = f"{ch}<br><span style='color: #0284c7;'>Sec {sec}</span>"
            html.append("<tr>")
            html.append(f'<td class="wt-td-label">{row_lbl}</td>')

            for d in display_days:
                dt = d["date"]
                day_ch_data = cohort_day_schedule.get(dt, {}).get(ch, {})
                sec_info = day_ch_data.get(sec, None)

                html.append('<td class="wt-td">')
                if sec_info:
                    mod_name = sec_info.get("module", "")
                    mod_short = (mod_name[:24] + "...") if len(mod_name) > 24 else mod_name
                    is_thursday_off = dt.weekday() == 3 and not ch.startswith("PG")

                    if ch == "UG-Sem 7" and sec_info.get("faculties_all_day"):
                        fac_display = " & ".join(sec_info["faculties_all_day"])
                        html.append(f'<div class="s-pill s1"><b>S1:</b> {fac_display}<br><small>{mod_short}</small></div>')
                        html.append(f'<div class="s-pill s2"><b>S2:</b> {fac_display}<br><small>{mod_short}</small></div>')
                        if is_thursday_off:
                            html.append('<div class="s-pill s-off">S3 & S4: Off</div>')
                        else:
                            html.append(f'<div class="s-pill s3"><b>S3:</b> {fac_display}<br><small>{mod_short}</small></div>')
                            html.append(f'<div class="s-pill s4"><b>S4:</b> {fac_display}<br><small>{mod_short}</small></div>')
                    else:
                        fac_m = sec_info.get("faculty_name_m") or "—"
                        fac_a = sec_info.get("faculty_name_a") or "—"
                        html.append(f'<div class="s-pill s1"><b>S1:</b> {fac_m}<br><small>{mod_short}</small></div>')
                        html.append(f'<div class="s-pill s2"><b>S2:</b> {fac_m}<br><small>{mod_short}</small></div>')
                        if is_thursday_off:
                            html.append('<div class="s-pill s-off">S3 & S4: Off</div>')
                        else:
                            html.append(f'<div class="s-pill s3"><b>S3:</b> {fac_a}<br><small>{mod_short}</small></div>')
                            html.append(f'<div class="s-pill s4"><b>S4:</b> {fac_a}<br><small>{mod_short}</small></div>')
                else:
                    html.append('<div class="s-pill s-off" style="text-align: center; padding: 22px 0;">No Sessions</div>')
                html.append("</td>")
            html.append("</tr>")

    html.append("</tbody></table>")
    st.markdown("".join(html), unsafe_allow_html=True)

# ==========================================================
# TAB 6: COURSE COMPLETION DATES
# ==========================================================
with tab6:
    st.subheader("🏁 Curriculum Module Completion Tracker")
    st.caption("Detailed overview of course completion dates.")

    COURSE_COMPLETION_DATA = [
        {"Cohort": "Sem 3", "Subject": "Personality Development", "Start Date": "06 Jul 2026", "Completion Date": "17 Jul 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 3", "Subject": "Indian Design System", "Start Date": "06 Jul 2026", "Completion Date": "17 Jul 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 3", "Subject": "Product Visualization", "Start Date": "20 Jul 2026", "Completion Date": "14 Aug 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 3", "Subject": "Form, Aesthetic and Emotion", "Start Date": "17 Aug 2026", "Completion Date": "03 Sep 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 3", "Subject": "Design Research", "Start Date": "07 Sep 2026", "Completion Date": "22 Sep 2026", "RawStatus": "Active"},
        {"Cohort": "Sem 3", "Subject": "Studio- Human Centric Design", "Start Date": "28 Sep 2026", "Completion Date": "19 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "Sem 3", "Subject": "Design Articulation with AI", "Start Date": "26 Oct 2026", "Completion Date": "04 Nov 2026", "RawStatus": "Upcoming"},
        {"Cohort": "Sem 5", "Subject": "CAID", "Start Date": "06 Jul 2026", "Completion Date": "10 Jul 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 5", "Subject": "Human Factors", "Start Date": "13 Jul 2026", "Completion Date": "03 Aug 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 5", "Subject": "The Art of Delightful Design", "Start Date": "17 Aug 2026", "Completion Date": "27 Aug 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 5", "Subject": "Studio- Humanizing Technology", "Start Date": "04 Aug 2026", "Completion Date": "08 Sep 2026", "RawStatus": "Completed"},
        {"Cohort": "Sem 5", "Subject": "Portfolio with AI", "Start Date": "09 Sep 2026", "Completion Date": "23 Oct 2026", "RawStatus": "Active"},
        {"Cohort": "Sem 5", "Subject": "Experience Design", "Start Date": "28 Sep 2026", "Completion Date": "15 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "Sem 5", "Subject": "Packaging Design", "Start Date": "28 Sep 2026", "Completion Date": "15 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "Sem 5", "Subject": "Speculative design", "Start Date": "28 Sep 2026", "Completion Date": "15 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "Sem 7", "Subject": "Internship", "Start Date": "Pre-Semester", "Completion Date": "Self-paced", "RawStatus": "Completed"},
        {"Cohort": "Sem 7", "Subject": "Design Management", "Start Date": "15 Sep 2026", "Completion Date": "28 Sep 2026", "RawStatus": "Active"},
        {"Cohort": "Sem 7", "Subject": "Studio: System Analysis and Design", "Start Date": "29 Sep 2026", "Completion Date": "03 Nov 2026", "RawStatus": "Upcoming"},
        {"Cohort": "PG Sem 1", "Subject": "Professional Communication", "Start Date": "Pre-Semester", "Completion Date": "Orientation", "RawStatus": "Completed"},
        {"Cohort": "PG Sem 1", "Subject": "Design Foundation", "Start Date": "03 Aug 2026", "Completion Date": "13 Aug 2026", "RawStatus": "Completed"},
        {"Cohort": "PG Sem 1", "Subject": "Form Studies", "Start Date": "14 Aug 2026", "Completion Date": "03 Sep 2026", "RawStatus": "Completed"},
        {"Cohort": "PG Sem 1", "Subject": "Design Studio I", "Start Date": "07 Sep 2026", "Completion Date": "01 Oct 2026", "RawStatus": "Active"},
        {"Cohort": "PG Sem 1", "Subject": "CAID & Visualization with AI", "Start Date": "05 Oct 2026", "Completion Date": "09 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "PG Sem 1", "Subject": "Emergent Technology", "Start Date": "12 Oct 2026", "Completion Date": "27 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "PG Sem 1", "Subject": "Design Prototyping", "Start Date": "28 Oct 2026", "Completion Date": "03 Nov 2026", "RawStatus": "Upcoming"},
        {"Cohort": "PG Sem 3", "Subject": "Internship", "Start Date": "Pre-Semester", "Completion Date": "Self-paced", "RawStatus": "Completed"},
        {"Cohort": "PG Sem 3", "Subject": "Research Methodology", "Start Date": "17 Aug 2026", "Completion Date": "21 Aug 2026", "Status": "Completed"},
        {"Cohort": "PG Sem 3", "Subject": "Studio- Design and Technology", "Start Date": "24 Aug 2026", "Completion Date": "16 Sep 2026", "RawStatus": "Completed"},
        {"Cohort": "PG Sem 3", "Subject": "User Experience Design", "Start Date": "17 Sep 2026", "Completion Date": "01 Oct 2026", "RawStatus": "Active"},
        {"Cohort": "PG Sem 3", "Subject": "Craft and Technology", "Start Date": "08 Oct 2026", "Completion Date": "27 Oct 2026", "RawStatus": "Upcoming"},
        {"Cohort": "PG Sem 3", "Subject": "Entrepreneurship", "Start Date": "28 Oct 2026", "Completion Date": "05 Nov 2026", "RawStatus": "Upcoming"},
    ]

    status_map = {
        "Completed": '<span class="status-badge status-completed">Completed</span>',
        "Active": '<span class="status-badge status-active">Active</span>',
        "Upcoming": '<span class="status-badge status-upcoming">Upcoming</span>',
    }

    for item in COURSE_COMPLETION_DATA:
        raw = item.get("RawStatus", "Upcoming")
        item["Status"] = status_map.get(raw, f'<span class="status-badge status-upcoming">{raw}</span>')

    df_comp = pd.DataFrame(COURSE_COMPLETION_DATA)
    c_list = ["All Cohorts", "Sem 3", "Sem 5", "Sem 7", "PG Sem 1", "PG Sem 3"]
    selected_comp_cohort = st.segmented_control("Filter by Cohort:", options=c_list, default="All Cohorts", key="comp_cohort_filter")

    if selected_comp_cohort != "All Cohorts":
        df_filtered = df_comp[df_comp["Cohort"] == selected_comp_cohort]
    else:
        df_filtered = df_comp

    display_df = df_filtered[["Cohort", "Subject", "Start Date", "Completion Date", "Status"]]
    st.write(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

# ==========================================================
# TAB 7: SUBJECT TEACHING HOURS & PROGRESS TRACKER
# ==========================================================
with tab7:
    st.subheader("📈 Subject Hours: Completed vs Needed in Semester")
    st.caption("Tracks total teaching hours delivered to date against the total prescribed hours per semester from the curriculum credit structure.")

    SUBJECT_HOURS_MASTER = [
        {"Cohort": "Sem 3", "Code": "UC012030001", "Subject": "Personality Development", "Hours Needed": 30.0, "Keywords": ["personality development"], "HalfDay": True},
        {"Cohort": "Sem 3", "Code": "31203006214", "Subject": "Indian Design System", "Hours Needed": 30.0, "Keywords": ["indian design system"], "HalfDay": True},
        {"Cohort": "Sem 3", "Code": "31203006211", "Subject": "Product Visualization", "Hours Needed": 120.0, "Keywords": ["product visualization"], "HalfDay": False},
        {"Cohort": "Sem 3", "Code": "31203006212", "Subject": "Form, Aesthetic and Emotion", "Hours Needed": 75.0, "Keywords": ["form, aesthetic and emotion"], "HalfDay": False},
        {"Cohort": "Sem 3", "Code": "31203001203", "Subject": "Design Research", "Hours Needed": 60.0, "Keywords": ["design research"], "HalfDay": False},
        {"Cohort": "Sem 3", "Code": "31203006207", "Subject": "Studio- Human Centric Design", "Hours Needed": 90.0, "Keywords": ["studio- human centric design", "human centric design"], "HalfDay": False},
        {"Cohort": "Sem 3", "Code": "31203006213", "Subject": "Design Articulation with AI", "Hours Needed": 45.0, "Keywords": ["design articulation with ai", "design articulation"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "UC013050001", "Subject": "CAID", "Hours Needed": 30.0, "Keywords": ["caid"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305001330", "Subject": "The Art of Delightful Design", "Hours Needed": 60.0, "Keywords": ["art of delightful design", "delightful design"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006322", "Subject": "Human Factors", "Hours Needed": 90.0, "Keywords": ["human factors"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006324", "Subject": "Studio- Humanizing Technology", "Hours Needed": 105.0, "Keywords": ["humanizing technology"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006329", "Subject": "Portfolio Design with Voice Agents / Portfolio with AI", "Hours Needed": 75.0, "Keywords": ["portfolio with ai", "portfolio design"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006325", "Subject": "Experience Design", "Hours Needed": 90.0, "Keywords": ["experience design"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006326", "Subject": "Packaging Design", "Hours Needed": 90.0, "Keywords": ["packaging design"], "HalfDay": False},
        {"Cohort": "Sem 5", "Code": "31305006327", "Subject": "Speculative design", "Hours Needed": 90.0, "Keywords": ["speculative design"], "HalfDay": False},
        {"Cohort": "Sem 7", "Code": "31407001400", "Subject": "Design Management", "Hours Needed": 45.0, "Keywords": ["design management", "dm"], "HalfDay": False},
        {"Cohort": "Sem 7", "Code": "31407006403", "Subject": "Internship", "Hours Needed": 240.0, "Keywords": ["internship"], "HalfDay": False},
        {"Cohort": "Sem 7", "Code": "31407006405", "Subject": "System Analysis and Design with Voice Agents", "Hours Needed": 165.0, "Keywords": ["system analysis and design", "system design"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101001500", "Subject": "Professional Communication", "Hours Needed": 30.0, "Keywords": ["professional communication", "communication"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002501", "Subject": "Design Foundation", "Hours Needed": 60.0, "Keywords": ["design foundation"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002502", "Subject": "Form Studies", "Hours Needed": 90.0, "Keywords": ["form studies"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002503", "Subject": "Design Studio I", "Hours Needed": 120.0, "Keywords": ["design studio i", "ds 1", "design studio 1"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002507", "Subject": "CAID & Visualization with AI", "Hours Needed": 45.0, "Keywords": ["caid & visualization with ai", "visualization with ai"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002505", "Subject": "Frugal Innovation", "Hours Needed": 75.0, "Keywords": ["frugal innovation"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002506", "Subject": "Emergent Technology", "Hours Needed": 75.0, "Keywords": ["emergent technology"], "HalfDay": False},
        {"Cohort": "PG Sem 1", "Code": "32101002504", "Subject": "Design Prototyping", "Hours Needed": 30.0, "Keywords": ["design prototyping"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203001602", "Subject": "Internship", "Hours Needed": 120.0, "Keywords": ["internship"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203001604", "Subject": "Research Methodology", "Hours Needed": 30.0, "Keywords": ["research methodology"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203002610", "Subject": "Studio- Design and Technology", "Hours Needed": 90.0, "Keywords": ["studio- design and technology", "studio: design and technology", "design and technology"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203002609", "Subject": "User Experience Design", "Hours Needed": 60.0, "Keywords": ["user experience design", "ux design"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203002611", "Subject": "Lighting Design", "Hours Needed": 90.0, "Keywords": ["lighting design"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203002612", "Subject": "Craft and Technology", "Hours Needed": 90.0, "Keywords": ["craft and technology"], "HalfDay": False},
        {"Cohort": "PG Sem 3", "Code": "32203001603", "Subject": "Entrepreneurship", "Hours Needed": 60.0, "Keywords": ["entrepreneurship"], "HalfDay": False},
    ]

    cohort_row_map = {
        "Sem 3": [8, 9, 10],
        "Sem 5": [30],
        "Sem 7": [48],
        "PG Sem 1": [62],
        "PG Sem 3": [71],
    }

    current_date = datetime.date.today()
    subject_progress_records = []

    for item in SUBJECT_HOURS_MASTER:
        cohort = item["Cohort"]
        kws = item["Keywords"]
        needed_hrs = item["Hours Needed"]
        is_half_day = item["HalfDay"]
        rows_to_check = cohort_row_map.get(cohort, [8])

        teaching_days_done = 0
        hours_done = 0.0
        total_scheduled_hrs = 0.0

        for c_info in calendar_cols:
            col_idx = c_info["col"]
            dt = c_info["date"]

            matched = False
            for r in rows_to_check:
                val = str(ws.cell(r, col_idx).value or "").lower()
                if any(k in val for k in kws):
                    matched = True
                    break

            if matched:
                if dt < datetime.date(2026, 9, 1):
                    f_hrs, h_hrs = 7.0, 3.5
                else:
                    f_hrs, h_hrs = 6.5, 3.25

                is_thu = dt.weekday() == 3
                thu_ug_off = is_thu and (dt >= datetime.date(2026, 8, 1)) and not cohort.startswith("PG")

                if cohort.startswith("PG"):
                    day_hrs = f_hrs
                elif is_half_day:
                    day_hrs = h_hrs
                elif thu_ug_off:
                    day_hrs = h_hrs
                else:
                    day_hrs = f_hrs

                total_scheduled_hrs += day_hrs
                if dt <= current_date:
                    hours_done += day_hrs
                    teaching_days_done += 1

        remaining_hrs = max(0.0, needed_hrs - hours_done)
        pct_done = min(100.0, (hours_done / needed_hrs * 100)) if needed_hrs > 0 else 0.0

        if pct_done >= 100.0:
            status_html = '<span class="status-badge status-completed">Completed</span>'
        elif pct_done > 0.0:
            status_html = '<span class="status-badge status-active">In Progress</span>'
        else:
            status_html = '<span class="status-badge status-upcoming">Pending</span>'

        subject_progress_records.append({
            "Cohort": cohort,
            "Course Code": item["Code"],
            "Course Name": item["Subject"],
            "Hours Done": f"{hours_done:.1f} hrs",
            "Hours Needed": f"{needed_hrs:.1f} hrs",
            "Remaining": f"{remaining_hrs:.1f} hrs",
            "Progress": f"{pct_done:.1f}%",
            "Status": status_html,
            "_raw_cohort": cohort,
        })

    df_prog = pd.DataFrame(subject_progress_records)
    cohort_options = ["All Cohorts", "Sem 3", "Sem 5", "Sem 7", "PG Sem 1", "PG Sem 3"]
    selected_cohort_filter = st.segmented_control("Filter by Semester:", options=cohort_options, default="All Cohorts", key="prog_cohort_filter")

    if selected_cohort_filter != "All Cohorts":
        df_filtered_prog = df_prog[df_prog["_raw_cohort"] == selected_cohort_filter]
    else:
        df_filtered_prog = df_prog

    cols_to_show = ["Cohort", "Course Code", "Course Name", "Hours Done", "Hours Needed", "Remaining", "Progress", "Status"]
    st.write(df_filtered_prog[cols_to_show].to_html(escape=False, index=False), unsafe_allow_html=True)

    csv_prog = df_filtered_prog[["Cohort", "Course Code", "Course Name", "Hours Done", "Hours Needed", "Remaining", "Progress"]].to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Subject Hours Progress (.csv)",
        data=csv_prog,
        file_name=f"Subject_Hours_Progress_{current_date}.csv",
        mime="text/csv",
        key="prog_download_btn",
    )

# ==========================================================
# TAB 8: WORKFLOW A — AUTO TIMETABLE EXPORT
# ==========================================================
with tab8:
    st.subheader("⚡ Auto-Generate Timetable CSV from Master Sheet")
    st.caption("Inspects master data, extracts assigned courses, and exports CSV in ERP format.")

    today = datetime.date.today()
    a_col1, a_col2, a_col3 = st.columns(3)
    with a_col1:
        cohort_choice = st.selectbox("Select Cohort:", ["UG-Sem 3", "UG-Sem 5", "UG-Sem 7", "PG-Sem 1", "PG-Sem 3"])
    with a_col2:
        auto_start = st.date_input("Start Date", today, key="auto_start")
    with a_col3:
        auto_end = st.date_input("End Date", today + datetime.timedelta(days=(4 - today.weekday()) if today.weekday() < 5 else 4), key="auto_end")

    auto_thu_half = st.checkbox("Thursday Afternoon Off (Slots 3 & 4 off for UG)", value=True, key="auto_thu")
    default_block = st.text_input("Academic Block", value="F", key="auto_block")

    cohort_erp_defaults = {
        "UG-Sem 3": {"prog": "1019", "sem": "III"},
        "UG-Sem 5": {"prog": "1019", "sem": "V"},
        "UG-Sem 7": {"prog": "1019", "sem": "VII"},
        "PG-Sem 1": {"prog": "1025", "sem": "I"},
        "PG-Sem 3": {"prog": "1025", "sem": "III"},
    }

    if st.button("Auto-Extract & Build Timetable", type="primary", key="btn_auto_build"):
        if auto_start > auto_end:
            st.error("Start Date must be before or equal to End Date.")
        else:
            extracted_rows = []
            cur_date = auto_start

            while cur_date <= auto_end:
                if cur_date.weekday() < 5:
                    date_str = cur_date.strftime("%Y-%m-%d")
                    day_cohorts = cohort_day_schedule.get(cur_date, {})
                    cohort_data = day_cohorts.get(cohort_choice, {})

                    is_thursday = cur_date.weekday() == 3
                    is_thu_off = is_thursday and auto_thu_half and not cohort_choice.startswith("PG")
                    active_slots = [1, 2] if is_thu_off else [1, 2, 3, 4]

                    for sec_code, sec_info in cohort_data.items():
                        mod_title = sec_info["module"]
                        fac_morning = sec_info["morning"]
                        fac_afternoon = sec_info["afternoon"]

                        c_code, _ = resolve_module(mod_title)
                        if c_code == "—":
                            matched_course = next((c for c in COURSES_DATA if c["title"].lower() in mod_title.lower() or mod_title.lower() in c["title"].lower()), None)
                            c_code = matched_course["code"] if matched_course else "3120300"

                        c_type = "MANDATORY"
                        erp_info = cohort_erp_defaults.get(cohort_choice, {"prog": "1019", "sem": "III"})

                        for s_num in active_slots:
                            assigned_faculty = fac_morning if s_num in [1, 2] else fac_afternoon
                            slot_classification = "PRACTICAL" if s_num in [1, 2] else "THEORY"

                            extracted_rows.append({
                                "Date": date_str,
                                "Program Code": erp_info["prog"],
                                "Semester Code": erp_info["sem"],
                                "Year": None,
                                "Course Classification": slot_classification,
                                "Course Code": c_code,
                                "Course Type": c_type,
                                "Faculty Code": assigned_faculty,
                                "Shift Timing": "09:30 TO 17:30",
                                "Slot Time": SLOT_TIMINGS_MAP[s_num],
                                "Section Code": sec_code,
                                "Group": None,
                                "Academic Block": default_block,
                                "Room Allocation": None,
                                "Combined Class": None,
                                "Mode of Class": "OFFLINE",
                                "Time Table Type": None,
                            })

                cur_date += datetime.timedelta(days=1)

            if extracted_rows:
                df_auto = pd.DataFrame(extracted_rows)
                st.success(f"Extracted {len(df_auto)} rows directly from the spreadsheet!")
                st.table(df_auto)

                csv_auto = df_auto.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Auto-Generated CSV",
                    data=csv_auto,
                    file_name=f"Auto_Timetable_{cohort_choice}_{auto_start}.csv",
                    mime="text/csv",
                )
            else:
                st.warning(f"No scheduled sessions found for {cohort_choice} between {auto_start} and {auto_end}.")

# ==========================================================
# TAB 9: WORKFLOW B — MANUAL / FORM-BASED GENERATOR
# ==========================================================
with tab9:
    st.subheader("🛠️ Custom / Manual Timetable Generator")
    st.caption("Build and customize a timetable schedule using manual overrides, custom rooms, and configurable slots.")

    def render_course_form(label_prefix, key_suffix, default_index=0, default_classif="PRACTICAL"):
        st.markdown(f"**{label_prefix}**")
        selected_key = st.selectbox(f"Subject for {label_prefix}", options=list(COURSE_OPTIONS.keys()), index=default_index, key=f"course_select_{key_suffix}")
        course_info = COURSE_OPTIONS[selected_key]

        if selected_key == "✏️ [Custom / Manual Course Entry]":
            c1, c2 = st.columns(2)
            code = c1.text_input("Course Code", key=f"custom_code_{key_suffix}")
            ctype = c2.selectbox("Course Type", ["MANDATORY", "ELECTIVE"], index=0, key=f"ctype_{key_suffix}")
            c3, c4, c5 = st.columns(3)
            sem = c3.selectbox("Semester", SEMESTERS, index=0, key=f"sem_{key_suffix}")
            prog = c4.text_input("Program Code", value="1019", key=f"prog_{key_suffix}")
            classif = c5.selectbox("Classification", ["PRACTICAL", "THEORY"], index=0 if default_classif == "PRACTICAL" else 1, key=f"classif_{key_suffix}")
        else:
            code = course_info["code"]
            default_sem_idx = SEMESTERS.index(course_info["sem"]) if course_info["sem"] in SEMESTERS else 0
            c1, c2, c3, c4 = st.columns([1.5, 1, 1, 1.2])
            ctype = c1.selectbox("Course Type", ["MANDATORY", "ELECTIVE"], index=0, key=f"ctype_{key_suffix}")
            sem = c2.selectbox("Semester", SEMESTERS, index=default_sem_idx, key=f"sem_{key_suffix}")
            prog = c3.text_input("Prog Code", value=course_info["prog_code"], key=f"prog_{key_suffix}")
            classif = c4.selectbox("Classification", ["PRACTICAL", "THEORY"], index=0 if default_classif == "PRACTICAL" else 1, key=f"classif_{key_suffix}")
            st.caption(f"📌 Mapped: **Code:** `{code}` | **Sem:** `{sem}` | **Prog:** `{prog}`")

        return {"code": code, "type": ctype, "sem": sem, "prog": prog, "classification": classif}

    def render_faculty_dropdown(label, key_name):
        fac_label = st.selectbox(label, options=list(DEFAULT_FACULTY_LIST.keys()), key=f"fac_select_{key_name}")
        if fac_label == "✏️ [Manual / Custom Entry]":
            fac_code = st.text_input(f"Type Faculty Code ({label})", key=f"custom_fac_{key_name}")
        elif fac_label == "-- None / Leave Empty --":
            fac_code = ""
        else:
            fac_code = DEFAULT_FACULTY_LIST[fac_label]
        return fac_code.strip() if fac_code else None

    st.markdown("#### 1. Schedule & Slot Timings")
    t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 1])

    with t_col1:
        today = datetime.date.today()
        c_d1, c_d2 = st.columns(2)
        start_date = c_d1.date_input("Start Date", today, key="man_start")
        end_date = c_d2.date_input("End Date", today + datetime.timedelta(days=4), key="man_end")
        thursday_half_day = st.checkbox("Thursday Afternoon Off (Half Day)", value=True, key="man_thu")

    with t_col2:
        selected_days = st.multiselect("Active Days of Week", options=list(DAYS_MAP.keys()), default=["Mon", "Tue", "Wed", "Thu", "Fri"], key="man_days")

    with t_col3:
        academic_block = st.text_input("Academic Block", value="F", key="man_block")

    st.markdown("---")
    st.markdown("#### 2. Course Selection")
    c_top1, c_top2 = st.columns([2, 1])
    with c_top1:
        sync_morning = st.checkbox("🔄 Auto-fill Morning Slot 2 from Slot 1", value=True, key="man_sync_m")
        sync_afternoon = st.checkbox("🔄 Auto-fill Afternoon Slot 4 from Slot 3", value=True, key="man_sync_a")
    with c_top2:
        num_sections = st.slider("Number of Sections in Batch", min_value=1, max_value=5, value=4, key="man_sec_count")

    m_col, a_col = st.columns(2, gap="large")
    with m_col:
        st.markdown("### 🌅 Morning Slots (Default: PRACTICAL)")
        slot1_cfg = render_course_form("Slot 1 (09:30 - 11:00)", "man_s1", default_index=0, default_classif="PRACTICAL")
        if sync_morning:
            s2_classif = st.selectbox("Slot 2 Classification", ["PRACTICAL", "THEORY"], index=0, key="man_s2_sync")
            slot2_cfg = {**slot1_cfg, "classification": s2_classif}
        else:
            slot2_cfg = render_course_form("Slot 2 (11:20 - 13:10)", "man_s2", default_index=0, default_classif="PRACTICAL")

    with a_col:
        st.markdown("### 🌇 Afternoon Slots (Default: THEORY)")
        slot3_cfg = render_course_form("Slot 3 (14:05 - 15:40)", "man_s3", default_index=1, default_classif="THEORY")
        if sync_afternoon:
            s4_classif = st.selectbox("Slot 4 Classification", ["THEORY", "PRACTICAL"], index=0, key="man_s4_sync")
            slot4_cfg = {**slot3_cfg, "classification": s4_classif}
        else:
            slot4_cfg = render_course_form("Slot 4 (16:00 - 17:30)", "man_s4", default_index=1, default_classif="THEORY")

    slots_configuration = {1: slot1_cfg, 2: slot2_cfg, 3: slot3_cfg, 4: slot4_cfg}

    st.markdown("---")
    st.markdown("#### 3. Section Allocations")
    section_labels = ["A", "B", "C", "D", "E"][:num_sections]
    section_allocations = []

    for sec in section_labels:
        st.markdown(f"**Section {sec}**")
        if sync_morning and sync_afternoon:
            col_m, col_a, col_r = st.columns([2.5, 2.5, 1.5])
            with col_m:
                fac_m = render_faculty_dropdown("Morning Faculty (Slots 1 & 2)", f"man_m_{sec}")
                fac_s1, fac_s2 = fac_m, fac_m
            with col_a:
                fac_a = render_faculty_dropdown("Afternoon Faculty (Slots 3 & 4)", f"man_a_{sec}")
                fac_s3, fac_s4 = fac_a, fac_a
            with col_r:
                room = st.text_input("Studio Room", key=f"man_room_{sec}", placeholder="e.g. F2, E10")
        else:
            col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 2, 1.5])
            with col1:
                fac_s1 = render_faculty_dropdown("Slot 1 Fac", f"man_s1_{sec}")
            with col2:
                fac_s2 = fac_s1 if sync_morning else render_faculty_dropdown("Slot 2 Fac", f"man_s2_{sec}")
            with col3:
                fac_s3 = render_faculty_dropdown("Slot 3 Fac", f"man_s3_{sec}")
            with col4:
                fac_s4 = fac_s3 if sync_afternoon else render_faculty_dropdown("Slot 4 Fac", f"man_s4_{sec}")
            with col5:
                room = st.text_input("Studio Room", key=f"man_room_{sec}", placeholder="e.g. F2, E10")

        section_allocations.append({
            "section": sec,
            "fac_slot_1": fac_s1,
            "fac_slot_2": fac_s2,
            "fac_slot_3": fac_s3,
            "fac_slot_4": fac_s4,
            "room": room.strip() if room else None,
        })

    st.markdown("---")
    if st.button("Generate Manual Timetable", type="primary", use_container_width=True, key="man_gen_btn"):
        empty_courses = [f"Slot {k}" for k, v in slots_configuration.items() if not v["code"]]
        if empty_courses:
            st.error(f"Missing Course Code for: {', '.join(empty_courses)}")
        elif start_date > end_date:
            st.error("Start Date must be before or equal to End Date.")
        else:
            active_day_ints = [DAYS_MAP[d] for d in selected_days]
            rows = []
            cur_date = start_date

            while cur_date <= end_date:
                if cur_date.weekday() in active_day_ints:
                    is_thursday = cur_date.weekday() == 3
                    date_str = cur_date.strftime("%Y-%m-%d")
                    shift_timing = "09:30 TO 17:30"
                    active_slots = [1, 2] if (is_thursday and thursday_half_day) else [1, 2, 3, 4]

                    for slot_num in active_slots:
                        s_cfg = slots_configuration[slot_num]
                        for sec in section_allocations:
                            rows.append({
                                "Date": date_str,
                                "Program Code": s_cfg["prog"],
                                "Semester Code": s_cfg["sem"],
                                "Year": None,
                                "Course Classification": s_cfg["classification"],
                                "Course Code": s_cfg["code"],
                                "Course Type": s_cfg["type"],
                                "Faculty Code": sec[f"fac_slot_{slot_num}"],
                                "Shift Timing": shift_timing,
                                "Slot Time": SLOT_TIMINGS_MAP[slot_num],
                                "Section Code": sec["section"],
                                "Group": None,
                                "Academic Block": academic_block,
                                "Room Allocation": sec["room"],
                                "Combined Class": None,
                                "Mode of Class": "OFFLINE",
                                "Time Table Type": None,
                            })
                cur_date += datetime.timedelta(days=1)

            df_result = pd.DataFrame(rows)
            st.success(f"Generated {len(df_result)} schedule rows successfully!")
            st.table(df_result)

            csv_bytes = df_result.to_csv(index=False).encode("utf-8")
            sample_prog = slot1_cfg["prog"]
            sample_sem = slot1_cfg["sem"]

            st.download_button(
                label="📥 Download Timetable CSV (.csv)",
                data=csv_bytes,
                file_name=f"UID_Timetable_{sample_prog}_Sem{sample_sem}_{start_date}.csv",
                mime="text/csv",
                key="man_download_btn",
            )
