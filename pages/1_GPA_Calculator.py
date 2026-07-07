import streamlit as st

st.set_page_config(
    page_title="GPA Calculator · Humanize AI",
    layout="wide",
    page_icon="🎓",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# SHARED DESIGN SYSTEM (mirrors app.py)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&family=Inter:wght@400;500;600;700;800&display=swap');

[data-testid="stSidebarNav"] { display: none !important; }

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Inter', 'Noto Sans KR', sans-serif !important; }

.stApp { background: #FAFAF8; }
.block-container { padding: 2.5rem 2.5rem 5rem !important; max-width: 1100px !important; }

section[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #EBEBEA !important;
}
section[data-testid="stSidebar"] .block-container { padding: 2rem 1.25rem !important; }

/* inputs */
div[data-testid="stTextInput"] input {
    background: #FFFFFF !important; color: #1A1915 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 10px !important;
    font-size: 14px !important; padding: 10px 12px !important;
    transition: all 0.18s !important;
}
div[data-testid="stTextInput"] input:focus {
    background: #fff !important; border-color: #4F46E5 !important;
    box-shadow: 0 0 0 3px rgba(79,70,229,0.10) !important;
}
div[data-testid="stTextInput"] label {
    font-size: 12px !important; font-weight: 700 !important;
    color: #5C5A55 !important; text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
}

/* number input */
div[data-testid="stNumberInput"] input {
    background: #FFFFFF !important; color: #1A1915 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 10px !important;
    font-size: 14px !important; padding: 10px 12px !important;
}
div[data-testid="stNumberInput"] label {
    font-size: 12px !important; font-weight: 700 !important;
    color: #5C5A55 !important; text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
}

/* selectbox */
div[data-testid="stSelectbox"] > div > div {
    background: #FFFFFF !important; border: 1.5px solid #DDDCDA !important;
    border-radius: 10px !important; color: #1A1915 !important;
    font-size: 14px !important;
}
div[data-testid="stSelectbox"] label {
    font-size: 12px !important; font-weight: 700 !important;
    color: #5C5A55 !important; text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
}

/* primary button */
.stButton > button[kind="primary"] {
    background: #1A1915 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 12px !important;
    padding: 13px 24px !important; font-size: 14px !important;
    font-weight: 700 !important; letter-spacing: -0.2px !important;
    transition: all 0.18s !important; height: auto !important;
}
.stButton > button[kind="primary"]:hover {
    background: #2D2C28 !important; transform: translateY(-1px) !important;
}

/* secondary button */
.stButton > button[kind="secondary"] {
    background: #FFFFFF !important; color: #5C5A55 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 10px !important;
    font-size: 13px !important; font-weight: 600 !important;
    padding: 9px 16px !important; transition: all 0.18s !important;
    height: auto !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #1A1915 !important; color: #1A1915 !important;
    background: #F5F4F1 !important;
}

/* radio */
div[data-testid="stRadio"] label { font-size: 14px !important; color: #1A1915 !important; }
div[data-testid="stRadio"] > div { gap: 8px !important; }

/* scrollbar */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #DDDCDA; border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: #B8B6B2; }

/* alert */
div[data-testid="stAlert"] { border-radius: 12px !important; font-size: 14px !important; font-weight: 500 !important; }

/* remove streamlit top padding */
div[data-testid="stToolbar"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# LANGUAGE
# ══════════════════════════════════════════════════════════════════════════════
if "lang" not in st.session_state:
    st.session_state.lang = "ko"

GL = {
    "ko": {
        "brand_sub": "Student Tools Suite", "nav_tools": "도구",
        "nav_humanizer": "AI 인간화", "nav_gpa": "GPA 계산기", "nav_active": "현재", "card_h_desc": "자소서·CV·에세이를 AI 흔적 없이 자연스럽게 재작성", "card_g_desc": "4.0 / 4.3 스케일 · 무제한 과목 · 목표 GPA 계산기",
        "page_title": "GPA 계산기", "page_sub": "4.0 & 4.3 스케일 · 무제한 과목 · 실시간 계산",
        "scale_label": "GPA 스케일", "scale_40": "4.0 스케일", "scale_43": "4.3 스케일",
        "guide_title": "등급 가이드", "formula_title": "계산식",
        "cur_gpa": "현재 GPA", "total_cred": "총 학점", "grade_pts": "학점 포인트",
        "courses_lbl": "과목 수", "enrolled": "수강 중",
        "credit_hours": "학점",  "total_points": "총 포인트",
        "pct_max": "% of maximum",
        "courses_title": "과목 목록", "add_course": "+ 과목 추가", "reset_all": "↺ 전체 초기화",
        "col_name": "과목명", "col_credits": "학점", "col_grade": "등급", "col_points": "포인트",
        "course_num": "과목", "no_courses": "아직 과목이 없습니다", "no_courses_sub": "<strong>+ 과목 추가</strong> 버튼을 눌러 시작하세요",
        "default_name": "과목",
        "invalid_warn": "개 과목의 데이터가 불완전하여 GPA 계산에서 제외됩니다.",
        "dist_title": "등급 분포", "breakdown_title": "계산 내역",
        "col_course": "과목", "col_gp": "등급 포인트", "col_earned": "취득 포인트",
        "total_row": "합계",
        "goal_title": "🎯 목표 GPA 계산기 — 몇 학점 더 필요할까?",
        "target_gpa": "목표 GPA", "future_grade": "앞으로 수강할 과목의 예상 등급",
        "goal_label": "필요 예상 학점", "goal_sub_a": "점의 성적으로", "goal_sub_b": "부터 목표 GPA",
        "goal_sub_c": "를 달성하기 위해",
        "goal_already": "이미 목표 GPA를 달성했습니다! 현재:",
        "goal_impossible": "해당 등급으로는 목표 GPA에 도달할 수 없습니다. 더 높은 예상 등급을 시도해보세요.",
        "goal_no_course": "목표 계산기를 사용하려면 과목을 하나 이상 추가하세요.",
        "course_placeholder": "과목명...",
    },
    "en": {
        "brand_sub": "Student Tools Suite", "nav_tools": "Tools",
        "nav_humanizer": "AI Humanizer", "nav_gpa": "GPA Calculator", "nav_active": "Active", "card_h_desc": "Rewrite cover letters, CVs & essays to bypass AI detection", "card_g_desc": "4.0 / 4.3 Scale · Unlimited Courses · Goal GPA Estimator",
        "page_title": "GPA Calculator", "page_sub": "4.0 & 4.3 Scale · Unlimited Courses · Live Calculation",
        "scale_label": "GPA Scale", "scale_40": "4.0 Scale", "scale_43": "4.3 Scale",
        "guide_title": "Scale Guide", "formula_title": "Formula",
        "cur_gpa": "Current GPA", "total_cred": "Total Credits", "grade_pts": "Grade Points",
        "courses_lbl": "Courses", "enrolled": "enrolled",
        "credit_hours": "credit hours", "total_points": "total points",
        "pct_max": "% of maximum",
        "courses_title": "Courses", "add_course": "+ Add Course", "reset_all": "↺ Reset All",
        "col_name": "Course Name", "col_credits": "Credits", "col_grade": "Grade", "col_points": "Points",
        "course_num": "Course", "no_courses": "No courses yet", "no_courses_sub": "Click <strong>+ Add Course</strong> to get started",
        "default_name": "Course",
        "invalid_warn": " course(s) have incomplete data and are excluded from the GPA calculation.",
        "dist_title": "Grade Distribution", "breakdown_title": "Calculation Breakdown",
        "col_course": "Course", "col_gp": "Grade Points", "col_earned": "Earned",
        "total_row": "TOTAL",
        "goal_title": "🎯 GPA Goal Estimator — How many more credits do I need?",
        "target_gpa": "Target GPA", "future_grade": "Expected Grade in Future Courses",
        "goal_label": "Estimated Credits Needed", "goal_sub_a": "of all-", "goal_sub_b": " work to reach a ", "goal_sub_c": " GPA",
        "goal_already": "You've already surpassed your target GPA! Current:",
        "goal_impossible": "With that grade, you can't reach your target GPA — try a higher expected grade.",
        "goal_no_course": "Add at least one course to use the goal estimator.",
        "course_placeholder": "Course name...",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
# GPA SCALE DATA
# ══════════════════════════════════════════════════════════════════════════════
GRADE_POINTS_40 = {
    "A+": 4.0, "A": 4.0, "A-": 3.7,
    "B+": 3.3, "B": 3.0, "B-": 2.7,
    "C+": 2.3, "C": 2.0, "C-": 1.7,
    "D+": 1.3, "D": 1.0,
    "F": 0.0,
}

GRADE_POINTS_43 = {
    "A+": 4.3, "A": 4.0, "A-": 3.7,
    "B+": 3.3, "B": 3.0, "B-": 2.7,
    "C+": 2.3, "C": 2.0, "C-": 1.7,
    "D+": 1.3, "D": 1.0,
    "F": 0.0,
}

GRADE_OPTIONS = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "F"]

GRADE_COLOR = {
    "A+": "#059669", "A": "#059669", "A-": "#10B981",
    "B+": "#0891B2", "B": "#0891B2", "B-": "#06B6D4",
    "C+": "#F59E0B", "C": "#F59E0B", "C-": "#FBBF24",
    "D+": "#EF4444", "D": "#EF4444",
    "F":  "#DC2626",
}


def gpa_color(gpa: float, scale: float) -> str:
    ratio = gpa / scale
    if ratio >= 0.9:   return "#059669"
    elif ratio >= 0.75: return "#0891B2"
    elif ratio >= 0.6:  return "#F59E0B"
    else:               return "#EF4444"


def calculate_gpa(courses: list, scale: str) -> dict:
    table = GRADE_POINTS_40 if scale == "4.0" else GRADE_POINTS_43
    max_gpa = 4.0 if scale == "4.0" else 4.3
    total_credits = 0
    total_points = 0.0
    valid = [c for c in courses if c.get("grade") and c.get("credits", 0) > 0]
    for c in valid:
        pts = table.get(c["grade"], 0.0)
        total_credits += c["credits"]
        total_points += pts * c["credits"]
    gpa = round(total_points / total_credits, 2) if total_credits > 0 else 0.0
    return {
        "gpa": gpa, "max_gpa": max_gpa,
        "total_credits": total_credits,
        "total_points": round(total_points, 2),
        "course_count": len(valid),
    }


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INIT
# ══════════════════════════════════════════════════════════════════════════════
if "gpa_courses" not in st.session_state:
    st.session_state.gpa_courses = [
        {"name": "Calculus",    "credits": 3, "grade": "A-"},
        {"name": "English",     "credits": 2, "grade": "B+"},
        {"name": "History",     "credits": 4, "grade": "A"},
    ]

if "gpa_scale" not in st.session_state:
    st.session_state.gpa_scale = "4.0"

if "gpa_next_id" not in st.session_state:
    st.session_state.gpa_next_id = len(st.session_state.gpa_courses)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    G = GL[st.session_state.lang]

    st.markdown(f"""
<div style="display:flex;align-items:center;gap:10px;padding-bottom:16px;
  border-bottom:1px solid #EBEBEA;margin-bottom:16px">
  <div style="width:36px;height:36px;background:#1A1915;border-radius:10px;
    display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:17px;font-weight:800;flex-shrink:0">✦</div>
  <div>
    <div style="font-size:15px;font-weight:800;color:#1A1915;letter-spacing:-0.3px">Humanize AI</div>
    <div style="font-size:11px;color:#9B9790;margin-top:1px">{G['brand_sub']}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # Language toggle
    lc1, lc2 = st.columns(2, gap="small")
    with lc1:
        if st.button("🇰🇷  한국어", key="lang_ko", use_container_width=True,
                     type="primary" if st.session_state.lang == "ko" else "secondary"):
            st.session_state.lang = "ko"; st.rerun()
    with lc2:
        if st.button("🇺🇸  English", key="lang_en", use_container_width=True,
                     type="primary" if st.session_state.lang == "en" else "secondary"):
            st.session_state.lang = "en"; st.rerun()

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # Navigation
    st.markdown(f"""
<div style="margin-bottom:20px;padding-bottom:20px;border-bottom:1px solid #EBEBEA">
  <p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;
    letter-spacing:0.6px;margin-bottom:10px">{G['nav_tools']}</p>
  <div style="display:flex;flex-direction:column;gap:10px">
    <a href="/" target="_self" style="text-decoration:none">
      <div style="padding:10px 12px;border-radius:9px;border:1px solid #E2E8F0;
        background:#FFFFFF;transition:all 0.15s;border-left:3px solid #1B3A6B"
        onmouseenter="this.style.background='#EEF2FA'"
        onmouseleave="this.style.background='#FFFFFF'">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
          <span style="font-size:14px">✦</span>
          <span style="font-size:13px;font-weight:600;color:#1E293B">{G['nav_humanizer']}</span>
          <span style="margin-left:auto;font-size:12px;color:#94A3B8">→</span>
        </div>
        <div style="font-size:11px;color:#6B7A99;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
          {G['card_h_desc']}
        </div>
      </div>
    </a>
    <div style="padding:10px 12px;border-radius:9px;background:#EEF2FA;
      border-left:3px solid #047857">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
        <span style="font-size:14px">🎓</span>
        <span style="font-size:13px;font-weight:700;color:#047857">{G['nav_gpa']}</span>
        <span style="margin-left:auto;font-size:10px;background:#047857;color:#fff;
          padding:2px 6px;border-radius:4px;font-weight:700">{G['nav_active'].upper()}</span>
      </div>
      <div style="font-size:11px;color:#6B7A99;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
        {G['card_g_desc']}
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown(f'<p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:10px">{G["scale_label"]}</p>', unsafe_allow_html=True)
    scale_choice = st.radio(
        "Scale", [G["scale_40"], G["scale_43"]],
        index=0 if st.session_state.gpa_scale == "4.0" else 1,
        label_visibility="collapsed",
    )
    st.session_state.gpa_scale = "4.0" if scale_choice == G["scale_40"] else "4.3"

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown(f"""
<div style="background:#F7F6F3;border-radius:12px;padding:14px 16px">
  <p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px;margin:0 0 10px">{G['guide_title']}</p>
  <div style="display:flex;flex-direction:column;gap:6px">
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">A+ / A</span><span style="color:#059669;font-weight:700">4.0 / 4.3</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">A-</span><span style="color:#10B981;font-weight:700">3.7</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">B+</span><span style="color:#0891B2;font-weight:700">3.3</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">B</span><span style="color:#0891B2;font-weight:700">3.0</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">B-</span><span style="color:#06B6D4;font-weight:700">2.7</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">C+ / C / C-</span><span style="color:#F59E0B;font-weight:700">2.3 – 1.7</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">D+ / D</span><span style="color:#EF4444;font-weight:700">1.3 – 1.0</span></div>
    <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#5C5A55;font-weight:600">F</span><span style="color:#DC2626;font-weight:700">0.0</span></div>
  </div>
</div>
""", unsafe_allow_html=True)




# ══════════════════════════════════════════════════════════════════════════════
# REFRESH G after sidebar
# ══════════════════════════════════════════════════════════════════════════════
G = GL[st.session_state.lang]

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="margin-bottom:32px">
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px">
    <div style="width:44px;height:44px;background:linear-gradient(135deg,#4F46E5,#7C3AED);
      border-radius:14px;display:flex;align-items:center;justify-content:center;
      font-size:22px;flex-shrink:0">🎓</div>
    <div>
      <h1 style="font-size:30px;font-weight:800;color:#1A1915;letter-spacing:-0.8px;margin:0">
        {G['page_title']}
      </h1>
      <p style="font-size:14px;color:#9B9790;margin:2px 0 0;font-weight:400">{G['page_sub']}</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# LIVE GPA SUMMARY CARDS
# ══════════════════════════════════════════════════════════════════════════════
result = calculate_gpa(st.session_state.gpa_courses, st.session_state.gpa_scale)
gpa_val  = result["gpa"]
max_gpa  = result["max_gpa"]
tot_cred = result["total_credits"]
tot_pts  = result["total_points"]
n_courses= result["course_count"]
gpa_c    = gpa_color(gpa_val, max_gpa)
bar_pct  = int((gpa_val / max_gpa) * 100) if max_gpa else 0

def gpa_label(gpa, scale):
    ratio = gpa / scale if scale else 0
    if ratio >= 0.925:   return "A+"
    elif ratio >= 0.875: return "A"
    elif ratio >= 0.825: return "A-"
    elif ratio >= 0.775: return "B+"
    elif ratio >= 0.725: return "B"
    elif ratio >= 0.675: return "B-"
    elif ratio >= 0.575: return "C+"
    elif ratio >= 0.5:   return "C"
    else:                return "—"

label = gpa_label(gpa_val, max_gpa)

st.markdown(f"""
<div style="display:grid;grid-template-columns:2fr 1fr 1fr 1fr;gap:14px;margin-bottom:28px">
  <div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:18px;padding:24px 28px;position:relative;overflow:hidden">
    <div style="position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,{gpa_c},{gpa_c}88);"></div>
    <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px">{G['cur_gpa']}</div>
    <div style="display:flex;align-items:baseline;gap:6px">
      <div style="font-size:42px;font-weight:800;color:{gpa_c};letter-spacing:-2px;line-height:1">{gpa_val:.2f}</div>
      <div style="font-size:16px;font-weight:600;color:#B8B6B2">/ {max_gpa}</div>
      <div style="margin-left:6px;background:{gpa_c}18;color:{gpa_c};font-size:13px;font-weight:800;padding:3px 8px;border-radius:6px">{label}</div>
    </div>
    <div style="margin-top:14px;background:#F0EFEC;border-radius:99px;height:6px;overflow:hidden">
      <div style="width:{bar_pct}%;height:100%;background:{gpa_c};border-radius:99px;transition:width 0.4s ease"></div>
    </div>
    <div style="font-size:11px;color:#B8B6B2;margin-top:6px">{bar_pct}{G['pct_max']}</div>
  </div>
  <div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:18px;padding:24px 20px">
    <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px">{G['total_cred']}</div>
    <div style="font-size:34px;font-weight:800;color:#1A1915;letter-spacing:-1.5px;line-height:1">{tot_cred}</div>
    <div style="font-size:12px;color:#B8B6B2;margin-top:8px">{G['credit_hours']}</div>
  </div>
  <div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:18px;padding:24px 20px">
    <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px">{G['grade_pts']}</div>
    <div style="font-size:34px;font-weight:800;color:#1A1915;letter-spacing:-1.5px;line-height:1">{tot_pts:.1f}</div>
    <div style="font-size:12px;color:#B8B6B2;margin-top:8px">{G['total_points']}</div>
  </div>
  <div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:18px;padding:24px 20px">
    <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px">{G['courses_lbl']}</div>
    <div style="font-size:34px;font-weight:800;color:#1A1915;letter-spacing:-1.5px;line-height:1">{n_courses}</div>
    <div style="font-size:12px;color:#B8B6B2;margin-top:8px">{G['enrolled']}</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# COURSE TABLE HEADER + ACTION BUTTONS
# ══════════════════════════════════════════════════════════════════════════════
hc1, hc2 = st.columns([1, 1])
with hc1:
    st.markdown(f'<h2 style="font-size:18px;font-weight:700;color:#1A1915;margin:0">{G["courses_title"]}</h2>', unsafe_allow_html=True)
with hc2:
    btn_cols = st.columns([1, 1], gap="small")
    with btn_cols[0]:
        if st.button(G["add_course"], type="primary", use_container_width=True, key="add_course"):
            st.session_state.gpa_courses.append({
                "name": f"{G['default_name']} {st.session_state.gpa_next_id + 1}",
                "credits": 3, "grade": "A",
            })
            st.session_state.gpa_next_id += 1
            st.rerun()
    with btn_cols[1]:
        if st.button(G["reset_all"], use_container_width=True, key="reset_courses"):
            st.session_state.gpa_courses = []
            st.session_state.gpa_next_id = 0
            st.rerun()

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# COURSE ROWS
# ══════════════════════════════════════════════════════════════════════════════
if not st.session_state.gpa_courses:
    st.markdown(f"""
<div style="background:#FFFFFF;border:1.5px dashed #DDDCDA;border-radius:16px;
  padding:48px;text-align:center">
  <div style="font-size:36px;margin-bottom:12px">📚</div>
  <div style="font-size:16px;font-weight:600;color:#1A1915;margin-bottom:6px">{G['no_courses']}</div>
  <div style="font-size:14px;color:#9B9790">{G['no_courses_sub']}</div>
</div>
""", unsafe_allow_html=True)
else:
    st.markdown(f"""
<div style="display:grid;grid-template-columns:1fr 110px 110px 80px 44px;gap:10px;
  padding:0 12px 8px;margin-bottom:2px">
  <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px">{G['col_name']}</div>
  <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px">{G['col_credits']}</div>
  <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px">{G['col_grade']}</div>
  <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px">{G['col_points']}</div>
  <div></div>
</div>
""", unsafe_allow_html=True)

    delete_idx = None
    table = GRADE_POINTS_40 if st.session_state.gpa_scale == "4.0" else GRADE_POINTS_43

    for i, course in enumerate(st.session_state.gpa_courses):
        grade_pts = table.get(course.get("grade", "A"), 0.0)
        earned = grade_pts * course.get("credits", 0)
        g_color = GRADE_COLOR.get(course.get("grade", "A"), "#5C5A55")

        st.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:14px;
  padding:14px 16px 4px;margin-bottom:10px;transition:border-color 0.18s;"
  onmouseenter="this.style.borderColor='#4F46E5'" onmouseleave="this.style.borderColor='#EBEBEA'">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
    <div style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px">{G['course_num']} {i+1}</div>
    <div style="background:{g_color}18;color:{g_color};font-size:12px;font-weight:800;padding:3px 8px;border-radius:6px">{course.get('grade','?')} · {earned:.1f} pts</div>
  </div>
</div>
""", unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns([4, 1.6, 1.6, 0.8], gap="small")
        with c1:
            new_name = st.text_input(
                "name", value=course.get("name", ""),
                placeholder=G["course_placeholder"],
                key=f"name_{i}", label_visibility="collapsed",
            )
            if new_name != course.get("name"):
                st.session_state.gpa_courses[i]["name"] = new_name
        with c2:
            new_cred = st.number_input(
                "credits", value=float(course.get("credits", 3)),
                min_value=0.5, max_value=20.0, step=0.5,
                key=f"cred_{i}", label_visibility="collapsed", format="%.1f",
            )
            if new_cred != course.get("credits"):
                st.session_state.gpa_courses[i]["credits"] = new_cred
        with c3:
            cur_grade = course.get("grade", "A")
            grade_idx = GRADE_OPTIONS.index(cur_grade) if cur_grade in GRADE_OPTIONS else 1
            new_grade = st.selectbox(
                "grade", GRADE_OPTIONS, index=grade_idx,
                key=f"grade_{i}", label_visibility="collapsed",
            )
            if new_grade != course.get("grade"):
                st.session_state.gpa_courses[i]["grade"] = new_grade
        with c4:
            if st.button("✕", key=f"del_{i}", use_container_width=True):
                delete_idx = i

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    if delete_idx is not None:
        st.session_state.gpa_courses.pop(delete_idx)
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# VALIDATION WARNINGS
# ══════════════════════════════════════════════════════════════════════════════
invalid = [
    c for c in st.session_state.gpa_courses
    if not c.get("name", "").strip() or c.get("credits", 0) <= 0 or c.get("grade") not in GRADE_OPTIONS
]
if invalid:
    st.warning(f"⚠ {len(invalid)}{G['invalid_warn']}")


# ══════════════════════════════════════════════════════════════════════════════
# GRADE DISTRIBUTION
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.gpa_courses:
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown(f'<h2 style="font-size:18px;font-weight:700;color:#1A1915;margin:0 0 16px">{G["dist_title"]}</h2>', unsafe_allow_html=True)

    dist: dict[str, int] = {}
    for c in st.session_state.gpa_courses:
        g = c.get("grade", "")
        if g in GRADE_OPTIONS:
            dist[g] = dist.get(g, 0) + 1

    total_c = sum(dist.values())
    bars_html = ""
    for g in GRADE_OPTIONS:
        count = dist.get(g, 0)
        if count == 0: continue
        pct = int(count / total_c * 100)
        g_color = GRADE_COLOR.get(g, "#9B9790")
        course_word = "과목" if st.session_state.lang == "ko" else ("course" if count == 1 else "courses")
        bars_html += f"""
<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px">
  <div style="width:32px;text-align:center;font-size:13px;font-weight:700;color:{g_color}">{g}</div>
  <div style="flex:1;background:#F0EFEC;border-radius:99px;height:10px;overflow:hidden">
    <div style="width:{pct}%;height:100%;background:{g_color};border-radius:99px;transition:width 0.4s"></div>
  </div>
  <div style="width:70px;text-align:right;font-size:12px;color:#5C5A55;font-weight:600">{count} {course_word} · {pct}%</div>
</div>"""

    st.markdown(f'<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:16px;padding:20px 24px">{bars_html}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CALCULATION BREAKDOWN TABLE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.gpa_courses and n_courses > 0:
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown(f'<h2 style="font-size:18px;font-weight:700;color:#1A1915;margin:0 0 16px">{G["breakdown_title"]}</h2>', unsafe_allow_html=True)

    table_rows = ""
    tbl = GRADE_POINTS_40 if st.session_state.gpa_scale == "4.0" else GRADE_POINTS_43
    for c in st.session_state.gpa_courses:
        if not c.get("grade") or c.get("credits", 0) <= 0: continue
        pts = tbl.get(c["grade"], 0.0)
        earned = pts * c["credits"]
        g_color = GRADE_COLOR.get(c["grade"], "#9B9790")
        table_rows += f"""<tr>
  <td style="padding:10px 16px;font-size:14px;color:#1A1915;font-weight:500">{c.get('name','—')}</td>
  <td style="padding:10px 16px;font-size:14px;color:#5C5A55;text-align:center">{c['credits']:.1f}</td>
  <td style="padding:10px 16px;text-align:center"><span style="background:{g_color}18;color:{g_color};font-size:13px;font-weight:700;padding:3px 10px;border-radius:6px">{c['grade']}</span></td>
  <td style="padding:10px 16px;font-size:13px;color:#5C5A55;text-align:center">{pts:.1f}</td>
  <td style="padding:10px 16px;font-size:14px;font-weight:700;color:#1A1915;text-align:center">{earned:.2f}</td>
</tr>"""

    th = lambda t: f'<th style="padding:10px 16px;font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px;text-align:center">{t}</th>'
    st.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:16px;overflow:hidden">
<table style="width:100%;border-collapse:collapse">
  <thead><tr style="background:#F7F6F3;border-bottom:1.5px solid #EBEBEA">
    <th style="padding:10px 16px;font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px;text-align:left">{G['col_course']}</th>
    {th(G['col_credits'])}{th(G['col_grade'])}{th(G['col_gp'])}{th(G['col_earned'])}
  </tr></thead>
  <tbody>{table_rows}
    <tr style="background:#F7F6F3;border-top:1.5px solid #EBEBEA">
      <td style="padding:12px 16px;font-size:13px;font-weight:700;color:#1A1915">{G['total_row']}</td>
      <td style="padding:12px 16px;font-size:13px;font-weight:700;color:#1A1915;text-align:center">{tot_cred}</td>
      <td></td><td></td>
      <td style="padding:12px 16px;font-size:14px;font-weight:800;color:{gpa_c};text-align:center">GPA: {gpa_val:.2f}</td>
    </tr>
  </tbody>
</table></div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# GPA GOAL ESTIMATOR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

with st.expander(G["goal_title"], expanded=False):
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    gc1, gc2 = st.columns(2, gap="medium")
    with gc1:
        target_gpa = st.number_input(
            G["target_gpa"], min_value=0.0, max_value=float(max_gpa),
            value=min(3.5, float(max_gpa)), step=0.05, format="%.2f", key="target_gpa",
        )
    with gc2:
        future_grade = st.selectbox(G["future_grade"], GRADE_OPTIONS, index=1, key="future_grade")

    if tot_cred > 0:
        future_pts = (GRADE_POINTS_40 if st.session_state.gpa_scale == "4.0" else GRADE_POINTS_43).get(future_grade, 4.0)
        denom = future_pts - target_gpa
        if abs(denom) < 0.001:
            if abs(gpa_val - target_gpa) < 0.01:
                st.success(f"{G['goal_already']} {gpa_val:.2f}")
            else:
                st.info(G["goal_impossible"])
        else:
            needed_credits = (target_gpa * tot_cred - tot_pts) / denom
            if needed_credits <= 0:
                st.success(f"{G['goal_already']} {gpa_val:.2f}")
            else:
                st.markdown(f"""
<div style="background:#EEF2FF;border:1.5px solid #C7D2FE;border-radius:14px;padding:18px 20px;margin-top:8px">
  <div style="font-size:13px;color:#4F46E5;font-weight:700;margin-bottom:6px">{G['goal_label']}</div>
  <div style="font-size:28px;font-weight:800;color:#1A1915;letter-spacing:-1px">{needed_credits:.1f} {G['col_credits'].lower()}</div>
  <div style="font-size:12px;color:#6366F1;margin-top:6px">
    {G['goal_sub_a']}{future_grade}{G['goal_sub_b']}{target_gpa:.2f}{G['goal_sub_c']} ({gpa_val:.2f} →)
  </div>
</div>
""", unsafe_allow_html=True)
    else:
        st.info(G["goal_no_course"])
