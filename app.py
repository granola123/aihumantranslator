import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
import io
import os
from docx import Document
from docx.shared import Pt
from fpdf import FPDF

st.set_page_config(
    page_title="Humanize AI",
    layout="wide",
    page_icon="✦",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# LANGUAGE STATE
# ══════════════════════════════════════════════════════════════════════════════
if "lang" not in st.session_state:
    st.session_state.lang = "ko"

L_ALL = {
    "ko": {
        # sidebar
        "brand_sub":       "Student Tools Suite",
        "nav_tools":       "도구",
        "nav_humanizer":   "AI 인간화",
        "nav_gpa":         "GPA 계산기",
        "nav_active":      "현재",
        "job_category":    "직무 카테고리",
        "jd_label":        "채용 공고 JD",
        "jd_hint":         "자소서 모드에서 사용됩니다",
        "jd_placeholder":  "공고 전문을 붙여넣으세요...\n한국어·영어 모두 지원합니다.",
        "pipeline_title":  "파이프라인",
        "pipe1":           "경험·키워드 추출",
        "pipe2":           "페르소나 초안 생성",
        "pipe3":           "불규칙성 주입",
        "pipe4":           "팩트 검수",
        # header
        "page_title":      "AI 문서 인간화",
        "card_humanizer":  "AI 인간화",
        "card_h_desc":     "자소서·CV·에세이를 AI 흔적 없이 자연스럽게 재작성",
        "card_h_cta":      "현재 페이지 ↑",
        "card_trans":      "번역",
        "card_t_desc":     "한국어 ↔ 영어 직무 문서 번역 (각 탭 내 토글)",
        "card_t_cta":      "각 탭에서 사용 가능",
        "card_gpa":        "GPA 계산기",
        "card_g_desc":     "4.0 / 4.3 스케일 · 무제한 과목 · 목표 GPA 계산기",
        "card_g_cta":      "바로 가기 →",
        # tabs
        "tab_cover":       "📝  자소서",
        "tab_cv":          "📄  CV / 이력서",
        "tab_essay":       "✍️  에세이",
        "banner_cover":    "자소서 모드",
        "banner_cover_d":  "우측에 채용 공고를 입력하면 <strong>ATS 키워드 최적화</strong>까지 함께 진행됩니다. AI가 쓴 것 같은 흔적을 제거하고 인간적인 문체로 바꿉니다.",
        "banner_cv":       "CV / 이력서 모드",
        "banner_cv_d":     "불릿 포인트 형식을 유지하면서 <strong>수동적 표현을 능동 동사로</strong> 바꿉니다. 문장 길이를 의도적으로 불균일하게 만들어 사람이 직접 쓴 것처럼 보이게 합니다.",
        "banner_essay":    "에세이 모드",
        "banner_essay_d":  "개인 에세이의 <strong>서사적 흐름과 감정 결</strong>을 살립니다. 사건 → 감정 → 깨달음의 구조를 유지하면서 AI 특유의 매끄러운 패턴을 제거합니다.",
        # mode switch
        "sw_human":        "✦ 인간화",
        "sw_trans":        "🌐 번역",
        # placeholders
        "ph_cover":        "한국어 또는 영어로 작성된 자소서 초안을 붙여넣으세요.\n자동으로 언어를 감지하여 최적화합니다.",
        "ph_cv":           "CV 또는 이력서 전문을 붙여넣으세요.\n불릿 포인트 형식 그대로 붙여넣으면 됩니다.",
        "ph_essay":        "개인 에세이 또는 Personal Statement를 붙여넣으세요.\n서사 흐름을 살리면서 AI 흔적을 제거합니다.",
        # run buttons
        "run_cover":       "자소서 첨삭 시작 →",
        "run_cv":          "CV 인간화 시작 →",
        "run_essay":       "에세이 인간화 시작 →",
        # errors
        "err_no_api":      "서버에 API 키가 설정되지 않았습니다. 관리자에게 문의하세요.",
        "err_no_jd":       "채용 공고(JD)를 우측 입력란에 붙여넣어 주세요.",
        "err_no_text":     "텍스트를 입력해주세요.",
        # progress
        "step1":           "경험·키워드 추출 중",
        "step2":           "인간화 초안 생성 중",
        "step3":           "불규칙성 주입 중",
        "step4":           "팩트 검수 중",
        "step_ai":         "AI 감지 확률 측정 중",
        "step_ai_pass":    "기준치 통과!",
        "step_ai_retry":   "30% 목표 미달, 재수정 중...",
        # result
        "done_ok":         "인간화 완료",
        "done_max":        "인간화 완료 (최대 재시도 도달)",
        "score_hist":      "AI 점수 변화",
        "metric_ai":       "AI 감지 확률",
        "metric_ats":      "ATS 점수",
        "metric_orig":     "원본 길이",
        "metric_diff":     "길이 변화",
        "before_label":    "원본",
        "before_sub":      "AI 감지 위험",
        "after_label":     "인간화 결과",
        "after_sub":       "AI 탐지 우회",
        "dl_label":        "다운로드",
        "copy_btn":        "클립보드 복사",
        "copy_done":       "복사 완료 ✓",
        "facts_expander":  "추출된 사실 데이터 보기",
        "char_ko":         "한국어",
        "char_unit":       "자",
        # translation UI
        "trans_dir_title": "번역 방향 선택",
        "trans_ko_en":     "🇰🇷 한국어  →  🇺🇸 영어",
        "trans_en_ko":     "🇺🇸 영어  →  🇰🇷 한국어",
        "trans_src_label": "원문 붙여넣기",
        "trans_btn":       "번역 시작",
        "trans_style":     "문체",
        "trans_done":      "번역 완료",
        "trans_before":    "원문",
        "trans_after":     "번역문",
        "trans_err":       "번역 오류",
        "trans_running":   "번역 중...",
        "trans_style_opt": "문체 최적화 적용 중",
        "trans_warn":      "방향 확인",
        "trans_detect":    "감지",
        "doc_cover":       "자기소개서",
        "doc_cv":          "이력서/CV",
        "doc_essay":       "개인 에세이",
        "jd_section_title": "채용 공고 입력 (JD)",
        "jd_optional":     "자소서 모드에서 ATS 키워드 최적화에 사용됩니다 (선택)",
        "cat_label":       "직무 분야",
        "beta_title":      "베타 서비스 안내",
        "beta_body":       "현재 베타 버전으로 운영 중입니다. AI 인간화 결과가 100% 완벽하지 않을 수 있으며, 일부 문장은 여전히 AI가 작성한 것처럼 감지될 수 있습니다. 결과물을 반드시 직접 검토 후 사용해 주세요.",
    },
    "en": {
        # sidebar
        "brand_sub":       "Student Tools Suite",
        "nav_tools":       "Tools",
        "nav_humanizer":   "AI Humanizer",
        "nav_gpa":         "GPA Calculator",
        "nav_active":      "Active",
        "job_category":    "Job Category",
        "jd_label":        "Job Description (JD)",
        "jd_hint":         "Used in Cover Letter mode",
        "jd_placeholder":  "Paste the full job posting here...\nSupports English & Korean.",
        "pipeline_title":  "Pipeline",
        "pipe1":           "Extract Facts & Keywords",
        "pipe2":           "Generate Persona Draft",
        "pipe3":           "Inject Irregularity",
        "pipe4":           "Fact Guardrail Check",
        # header
        "page_title":      "AI Document Humanizer",
        "card_humanizer":  "AI Humanizer",
        "card_h_desc":     "Rewrite cover letters, CVs & essays to bypass AI detection",
        "card_h_cta":      "Current page ↑",
        "card_trans":      "Translate",
        "card_t_desc":     "Korean ↔ English translation for job documents (toggle in each tab)",
        "card_t_cta":      "Available in each tab",
        "card_gpa":        "GPA Calculator",
        "card_g_desc":     "4.0 / 4.3 Scale · Unlimited Courses · Goal GPA Estimator",
        "card_g_cta":      "Go →",
        # tabs
        "tab_cover":       "📝  Cover Letter",
        "tab_cv":          "📄  CV / Resume",
        "tab_essay":       "✍️  Essay",
        "banner_cover":    "Cover Letter Mode",
        "banner_cover_d":  "Add a job description in the right panel to enable <strong>ATS keyword optimization</strong>. Removes AI-written patterns and rewrites in a genuine human voice.",
        "banner_cv":       "CV / Resume Mode",
        "banner_cv_d":     "Preserves bullet-point formatting while converting <strong>passive expressions to strong action verbs</strong>. Intentionally varies sentence length to read like a real person wrote it.",
        "banner_essay":    "Essay Mode",
        "banner_essay_d":  "Preserves the <strong>narrative arc and emotional tone</strong> of personal essays. Maintains the event → emotion → insight structure while removing AI-typical smooth patterns.",
        # mode switch
        "sw_human":        "✦ Humanize",
        "sw_trans":        "🌐 Translate",
        # placeholders
        "ph_cover":        "Paste your cover letter draft here (English or Korean).\nLanguage is detected automatically.",
        "ph_cv":           "Paste your full CV or resume here.\nBullet-point formatting is preserved.",
        "ph_essay":        "Paste your personal essay or Personal Statement here.\nNarrative flow is preserved while removing AI traces.",
        # run buttons
        "run_cover":       "Humanize Cover Letter →",
        "run_cv":          "Humanize CV →",
        "run_essay":       "Humanize Essay →",
        # errors
        "err_no_api":      "API key not configured on this server. Please contact the administrator.",
        "err_no_jd":       "Please add a Job Description in the sidebar.",
        "err_no_text":     "Please enter some text.",
        # progress
        "step1":           "Extracting facts & keywords",
        "step2":           "Generating humanized draft",
        "step3":           "Injecting irregularity",
        "step4":           "Running fact guardrail",
        "step_ai":         "Measuring AI detection probability",
        "step_ai_pass":    "Passed threshold!",
        "step_ai_retry":   "Target 30% not met — refining...",
        # result
        "done_ok":         "Humanization Complete",
        "done_max":        "Complete (max retries reached)",
        "score_hist":      "AI Score History",
        "metric_ai":       "AI Detection",
        "metric_ats":      "ATS Score",
        "metric_orig":     "Original Length",
        "metric_diff":     "Length Change",
        "before_label":    "Original",
        "before_sub":      "AI detection risk",
        "after_label":     "Humanized Result",
        "after_sub":       "AI detection bypassed",
        "dl_label":        "Download",
        "copy_btn":        "Copy to Clipboard",
        "copy_done":       "Copied ✓",
        "facts_expander":  "View extracted fact data",
        "char_ko":         "Korean",
        "char_unit":       "chars",
        # translation UI
        "trans_dir_title": "Select Translation Direction",
        "trans_ko_en":     "🇰🇷 Korean  →  🇺🇸 English",
        "trans_en_ko":     "🇺🇸 English  →  🇰🇷 Korean",
        "trans_src_label": "Paste source text",
        "trans_btn":       "Translate",
        "trans_style":     "Style",
        "trans_done":      "Translation Complete",
        "trans_before":    "Source",
        "trans_after":     "Translation",
        "trans_err":       "Translation error",
        "trans_running":   "Translating...",
        "trans_style_opt": "Applying document style optimization",
        "trans_warn":      "check direction",
        "trans_detect":    "Detected",
        "doc_cover":       "Cover Letter",
        "doc_cv":          "Resume/CV",
        "doc_essay":       "Personal Essay",
        "jd_section_title": "Job Description (JD)",
        "jd_optional":     "Used for ATS keyword optimization in Cover Letter mode (optional)",
        "cat_label":       "Job Field",
        "beta_title":      "Beta Notice",
        "beta_body":       "This tool is currently in beta. Humanization results may not be 100% perfect — some sentences may still be flagged as AI-written by detection tools. Please review all output carefully before use.",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
# GLOBAL CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Inter', 'Noto Sans KR', sans-serif !important; }

.stApp { background: #F7F8FC; }
.block-container { padding: 2rem 2.5rem 5rem !important; max-width: 1280px !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #E2E8F0 !important;
}
section[data-testid="stSidebar"] .block-container { padding: 2rem 1.25rem !important; }

/* ── Tabs ── */
div[data-testid="stTabs"] > div:first-child {
    border-bottom: 2px solid #E2E8F0 !important;
    gap: 0 !important; margin-bottom: 24px !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-size: 14px !important; font-weight: 600 !important; color: #94A3B8 !important;
    padding: 11px 20px !important; border-radius: 0 !important; background: transparent !important;
    border: none !important; border-bottom: 3px solid transparent !important;
    margin-bottom: -2px !important; transition: all 0.18s !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:hover {
    color: #1B3A6B !important; background: #EEF2FA !important;
}
div[data-testid="stTabs"] button[aria-selected="true"][data-baseweb="tab"] {
    color: #1B3A6B !important; border-bottom-color: #1B3A6B !important;
    font-weight: 700 !important;
}
div[data-testid="stTabs"] div[role="tabpanel"] { padding-top: 0 !important; }

/* ── Textarea ── */
div[data-testid="stTextArea"] textarea {
    background: #FFFFFF !important; color: #1E293B !important;
    border: 1.5px solid #CBD5E1 !important; border-radius: 10px !important;
    font-family: 'Inter', 'Noto Sans KR', sans-serif !important;
    font-size: 14.5px !important; line-height: 1.75 !important;
    padding: 14px 16px !important; resize: vertical !important;
    transition: border-color 0.18s, box-shadow 0.18s !important;
}
div[data-testid="stTextArea"] textarea:focus {
    border-color: #1B3A6B !important;
    box-shadow: 0 0 0 3px rgba(27,58,107,0.10) !important; outline: none !important;
}
div[data-testid="stTextArea"] label {
    font-size: 13px !important; font-weight: 700 !important;
    color: #374151 !important; margin-bottom: 6px !important;
}

/* ── Selectbox ── */
div[data-testid="stSelectbox"] > div > div {
    background: #FFFFFF !important; border: 1.5px solid #CBD5E1 !important;
    border-radius: 10px !important; color: #1E293B !important; font-size: 14px !important;
    transition: border-color 0.18s !important;
}
div[data-testid="stSelectbox"] label {
    font-size: 13px !important; font-weight: 700 !important; color: #374151 !important;
    letter-spacing: 0.2px !important;
}

/* ── Primary button ── */
.stButton > button[kind="primary"] {
    background: #1B3A6B !important; color: #FFFFFF !important;
    border: none !important; border-radius: 10px !important;
    padding: 14px 28px !important; font-size: 15px !important;
    font-weight: 700 !important; letter-spacing: -0.2px !important;
    transition: all 0.18s !important; height: auto !important;
    box-shadow: 0 2px 8px rgba(27,58,107,0.18) !important;
}
.stButton > button[kind="primary"]:hover {
    background: #163060 !important; transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(27,58,107,0.25) !important;
}
.stButton > button[kind="primary"]:active { transform: translateY(0) !important; }

/* ── Secondary button ── */
.stButton > button[kind="secondary"] {
    background: #FFFFFF !important; color: #475569 !important;
    border: 1.5px solid #CBD5E1 !important; border-radius: 10px !important;
    font-size: 14px !important; font-weight: 600 !important;
    padding: 12px 20px !important; transition: all 0.18s !important; height: auto !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #1B3A6B !important; color: #1B3A6B !important; background: #EEF2FA !important;
}

/* ── Download button ── */
.stDownloadButton > button {
    background: #FFFFFF !important; color: #475569 !important;
    border: 1.5px solid #CBD5E1 !important; border-radius: 9px !important;
    font-size: 13px !important; font-weight: 600 !important;
    padding: 8px 12px !important; transition: all 0.18s !important;
}
.stDownloadButton > button:hover {
    border-color: #1B3A6B !important; color: #1B3A6B !important; background: #EEF2FA !important;
}

/* ── Expander ── */
div[data-testid="stExpander"] {
    border: 1.5px solid #E2E8F0 !important; border-radius: 10px !important;
    background: #FFFFFF !important; overflow: hidden !important; box-shadow: none !important;
}
div[data-testid="stExpander"] summary {
    color: #475569 !important; font-weight: 600 !important;
    font-size: 13px !important; padding: 12px 16px !important;
}
div[data-testid="stExpander"] summary:hover { background: #F1F5F9 !important; }

/* ── Progress ── */
.stProgress > div > div { background: #E2E8F0 !important; border-radius: 99px !important; height: 6px !important; }
.stProgress > div > div > div { border-radius: 99px !important; height: 6px !important; background: #1B3A6B !important; }

/* ── Alert ── */
div[data-testid="stAlert"] { border-radius: 10px !important; font-size: 14px !important; font-weight: 500 !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

/* ── Metric ── */
div[data-testid="stMetric"] {
    background: #FFFFFF !important; border: 1px solid #E2E8F0 !important;
    border-radius: 10px !important; padding: 14px 16px !important;
}
div[data-testid="stMetricLabel"] { font-size: 12px !important; color: #64748B !important; font-weight: 600 !important; }
div[data-testid="stMetricValue"] { font-size: 24px !important; color: #1B3A6B !important; font-weight: 800 !important; }

/* ── Radio ── */
div[data-testid="stRadio"] label { font-size: 14px !important; color: #374151 !important; font-weight: 500 !important; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
ACTION_VERBS = {
    "Strategy / Finance": ["Formulated","Forecasted","Evaluated","Structured","Negotiated","Synthesized","Streamlined"],
    "Marketing":          ["Launched","Optimized","Analyzed","Cultivated","Amplified","Positioned","Executed"],
    "Engineering":        ["Developed","Architected","Automated","Deployed","Refactored","Integrated","Scaled"],
    "Product Management": ["Prioritized","Collaborated","Shipped","Defined","Validated","Iterated","Championed"],
    "Data / AI":          ["Modeled","Trained","Predicted","Visualized","Extracted","Transformed","Benchmarked"],
    "Operations / HR":    ["Coordinated","Implemented","Facilitated","Reduced","Improved","Managed","Standardized"],
}

CATEGORY_KEYS = [
    "Strategy / Finance", "Marketing", "Engineering",
    "Product Management", "Data / AI", "Operations / HR",
]
CATEGORY_LABELS = {
    "ko": ["전략 / 재무", "마케팅", "개발 / 엔지니어링", "제품 관리", "데이터 / AI", "운영 / 인사"],
    "en": ["Strategy / Finance", "Marketing", "Engineering", "Product Management", "Data / AI", "Operations / HR"],
}

TAB_META = {
    "cover": {"color": "#4F46E5", "light": "#EEF2FF", "info_bg": "#EEF2FF", "info_text": "#3730A3", "info_sub": "#6366F1"},
    "cv":    {"color": "#0891B2", "light": "#ECFEFF", "info_bg": "#ECFEFF", "info_text": "#164E63", "info_sub": "#0E7490"},
    "essay": {"color": "#7C3AED", "light": "#F5F3FF", "info_bg": "#F5F3FF", "info_text": "#4C1D95", "info_sub": "#7C3AED"},
}


# ══════════════════════════════════════════════════════════════════════════════
# UTILS
# ══════════════════════════════════════════════════════════════════════════════
GROQ_BASE_URL = "https://api.groq.com/openai/v1"
GROQ_MODEL = "llama-3.3-70b-versatile"

def get_client():
    """API key from Streamlit secrets or environment variable (Groq)."""
    key = ""
    try:
        key = st.secrets.get("groq_key", "")
    except Exception:
        pass
    if not key:
        key = os.environ.get("GROQ_API_KEY", "")
    return OpenAI(api_key=key, base_url=GROQ_BASE_URL) if key else None

def is_korean(text: str) -> bool:
    return sum(1 for c in text if '가' <= c <= '힣') / max(len(text), 1) > 0.1


# ══════════════════════════════════════════════════════════════════════════════
# AI PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
def step1_extract(client, jd, draft, category, mode):
    if mode == "cv":
        body = f"Extract structured data from this CV/Resume.\n\nDRAFT:\n{draft}\n\nOutput:\nEXPERIENCES: [role/company/duration/metrics]\nSKILLS: [tech+soft]\nEDUCATION: [degrees]\nACHIEVEMENTS: [quantified]\nTONE: [Korean/English]"
    elif mode == "essay":
        body = f"Extract from personal essay.\n\nDRAFT:\n{draft}\n\nOutput:\nCORE_NARRATIVE: [1-2 sentences]\nKEY_EVENTS: [specific events]\nEMOTIONS: [turning points]\nINSIGHTS: [lessons]\nTONE: [Korean/English]"
    else:
        body = f"Extract from cover letter.\n\nDRAFT:\n{draft}\n\nJD:\n{jd}\n\nOutput:\nEXPERIENCES: [achievements+numbers]\nSKILLS_DEMONSTRATED: [skills]\nMISSING_KEYWORDS: [JD keywords not in draft]\nMUST_INCLUDE: [top 5 JD keywords]\nTONE: [Korean/English]\nCATEGORY: {category}"
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"system","content":"Data extractor. Output ONLY structured data."},{"role":"user","content":body}],
        temperature=0.1,
    ).choices[0].message.content


def step2_persona(client, facts, draft, category, mode):
    verbs = ", ".join(ACTION_VERBS.get(category, []))
    kor = is_korean(draft)
    if mode == "cv":
        sys = "너는 10년 경력 헤드헌터야. AI 냄새 나는 표현을 전부 걷어내고 실제 사람이 쓴 것처럼 바꿔." if kor else "You're a senior headhunter rewriting a resume to sound genuinely human."
        task = (f"아래 데이터로 이력서를 인간화해. 불릿 유지, 숫자 유지, 수동→능동, 길이 불균일하게.\n금지어: 혁신적인, 탁월한, 최선, 열정\n\n[데이터]\n{facts}\n\n[원본]\n{draft}\n\n본문만 출력." if kor
                else f"Humanize this CV. Keep bullets+numbers. Active verbs. Vary lengths. No: passionate/innovative/leveraged.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\n[VERBS]\n{verbs}\n\nOutput CV text only.")
    elif mode == "essay":
        sys = "너는 대학원 입학사정관 출신 글쓰기 코치야. 서사 흐름을 살리고 AI 흔적을 제거해." if kor else "You're a writing coach who reads college essays. Human > polished."
        task = (f"에세이 인간화. 서사 유지(사건→감정→깨달음), 문장 길이 불규칙, 구어체 연결사 삽입, 열린 결말 허용.\n금지: 혁신,열정,도전,성장\n\n[데이터]\n{facts}\n\n[원본]\n{draft}\n\n본문만 출력." if kor
                else f"Humanize this essay. Keep arc (event→emotion→insight). Vary length. Use contractions. No: passionate/transformative/journey.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\nOutput only.")
    else:
        sys = ("너는 대기업 인사팀 출신 40대 컨설턴트야. 밤 11시, 피곤하지만 경험 풍부.\n문체: 짧은문장+긴문장 혼합, 구어체 연결사, 단락 길이 불균일.\n금지어: 혁신,열정,도전,시너지,기여,임하겠습니다" if kor
               else "You're a 35-yo ex-recruiter editing at 11 PM. Slightly imperfect. Mix short+long sentences. No: passionate/innovative/leverage/synergy/dynamic")
        task = (f"자소서 다시 써줘. 원본 없는 내용 추가 금지. '저는' 3연속 금지. 병렬구조 반복 금지. 숫자 유지.\n\n[사실]\n{facts}\n\n[원본]\n{draft}\n\n[동사]\n{verbs}\n\n본문만 출력." if kor
                else f"Rewrite ONLY using facts below. No 3+ 'I'-starts. Break repeated parallel. Preserve numbers.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\n[VERBS]\n{verbs}\n\nCover letter body only.")
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"system","content":sys},{"role":"user","content":task}],
        temperature=1.0, presence_penalty=0.7, frequency_penalty=0.5,
    ).choices[0].message.content


def step3_irregularity(client, text, draft, mode):
    kor = is_korean(draft)
    if mode == "cv":
        p = ("CV에서 AI 패턴 제거.\n1.같은 구조 불릿→2개 다르게 2.숫자 표기 다양화 3.마지막 불릿 미완성 느낌\n팩트 변경 금지. 수정본만." if kor
             else "Remove AI patterns from CV.\n1.Vary 2 bullets if all same structure 2.Mix number formats 3.Last bullet slightly open\nNo fact changes. Output only.")
    elif mode == "essay":
        p = ("에세이 AI 패턴 제거.\n1.긴 문장 2개→대시로 끊기 2.짧은 문장 2개 유지 3.모든 단락이 결론이면→1개 질문/여운으로 4.비슷한 단락 3개↑→1개 1문장으로\n팩트 변경 금지." if kor
             else "Remove AI patterns from essay.\n1.Split 2 longest sentences with em-dash 2.Keep 2 shortest verbatim 3.If every para ends with conclusion→1 open thought 4.3+ similar-length paras→collapse 1\nNo changes to facts. Output only.")
    else:
        p = ("자소서 AI 흔적 제거.\n1.긴 문장 3개→대시로 끊기 2.짧은 문장 2개 유지 3.~했습니다 3연속→1개 명사형 4.3항 나열→1개 비틀기 5.마지막 문장→열린 결말 6.같은 길이 단락 3개↑→1개 1문장\n팩트 변경 금지. 수정본만." if kor
             else "Remove AI patterns from cover letter.\n1.Split 3 longest w/ em-dash 2.Keep 2 shortest verbatim 3.3+ 'I' starts→change one 4.Repeated parallel→break one 5.Final line→trailing thought 6.3+ same-length paras→collapse one\nNo facts changed. Output only.")
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"user","content":f"{p}\n\n[글]\n{text}"}],
        temperature=0.5,
    ).choices[0].message.content


def step4_guardrail(client, original, final):
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role":"system","content":"Fact-checker. Remove invented facts not in original. Keep all style changes."},
            {"role":"user","content":f"ORIGINAL:\n{original}\n\nDRAFT:\n{final}\n\nOutput corrected text only."},
        ],
        temperature=0.1,
    ).choices[0].message.content


def ats_score(text, jd):
    jw = set(jd.lower().split())
    tw = set(text.lower().split())
    return min(100, int(len(jw & tw) / max(len(jw), 1) * 300)) if jw else 0


def step5_ai_score(client, text) -> int:
    prompt = f"""You are an AI detection expert. Analyze the following text and estimate the probability (0-100) that it was written by AI.

Criteria for HIGH AI probability (score closer to 100):
- Overly uniform sentence lengths
- Perfect parallel structures repeated multiple times
- Buzzwords: innovative, passionate, leverage, synergy, 혁신, 열정, 시너지
- Every paragraph ends with a tidy conclusion
- No hesitation, no imperfection, no personal voice
- Formulaic transitions: Furthermore, Additionally, In conclusion, 또한, 따라서, 결론적으로

Criteria for LOW AI probability (score closer to 0):
- Irregular sentence lengths (very short + very long mixed)
- Natural interruptions (em-dashes, incomplete thoughts)
- Colloquial connectors: Honestly, Here's the thing, 솔직히, 그때, 사실은
- Some paragraphs end without conclusion
- Feels like a real person wrote it at midnight

TEXT TO ANALYZE:
{text}

Respond with ONLY a single integer (0-100). Nothing else."""
    try:
        resp = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0, max_tokens=5,
        ).choices[0].message.content.strip()
        return max(0, min(100, int("".join(filter(str.isdigit, resp)) or "50")))
    except Exception:
        return 50


def step5_refine(client, text, original) -> str:
    kor = is_korean(original)
    if kor:
        prompt = f"""아래 글이 AI가 쓴 것 같다는 평가를 받았어. 더 인간적으로 바꿔줘.

[글]
{text}

적용할 수정 (모두 적용):
1. 가장 AI스러운 문장 3개 → 완전히 다른 구조로 재작성
2. 완벽하게 마무리된 단락 1개 → 열린 결말 또는 여운으로
3. 너무 매끄러운 전환어(또한, 따라서, 결론적으로) → 구어체로 교체
4. 가장 긴 문장 2개 → 대시(—)로 중간에 끊기
5. 문장 3개 이상 같은 길이 패턴 → 의도적으로 깨기
6. 자연스러운 망설임 또는 방향 전환 1곳 추가

팩트(숫자, 이름, 경험) 변경 절대 금지. 수정본만 출력."""
    else:
        prompt = f"""The text below was flagged as likely AI-written. Make it sound more human.

[TEXT]
{text}

Apply ALL of these:
1. Rewrite the 3 most AI-sounding sentences with completely different structures
2. Take 1 perfectly-concluded paragraph → rewrite ending as open thought or question
3. Replace smooth transitions (Furthermore, Additionally, In conclusion) → conversational alternatives
4. Split the 2 longest sentences with an em-dash at a natural break
5. Break any 3+ sentence pattern of similar length
6. Add 1 moment of natural hesitation or shift in direction

Never change facts (numbers, names, experiences). Output revised text only."""
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.9, presence_penalty=0.6, frequency_penalty=0.4,
    ).choices[0].message.content


# ══════════════════════════════════════════════════════════════════════════════
# FILE GENERATION
# ══════════════════════════════════════════════════════════════════════════════
def make_docx(text):
    doc = Document()
    s = doc.styles['Normal']
    s.font.name = '맑은 고딕' if is_korean(text) else 'Calibri'
    s.font.size = Pt(11)
    for p in text.split('\n'): doc.add_paragraph(p)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def make_pdf(text):
    pdf = FPDF(); pdf.add_page(); pdf.set_margins(20,20,20)
    loaded = False
    if is_korean(text):
        font_paths = [
            r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\NanumGothic.ttf",
            r"C:\Windows\Fonts\gulim.ttc",
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]
        for path in font_paths:
            if os.path.exists(path):
                try: pdf.add_font("K","",path,uni=True); pdf.set_font("K",size=11); loaded=True; break
                except: pass
    if not loaded: pdf.set_font("Helvetica",size=11)
    pdf.set_auto_page_break(auto=True,margin=20)
    for line in text.split('\n'): pdf.multi_cell(0,7,line or " ")
    return bytes(pdf.output())


# ══════════════════════════════════════════════════════════════════════════════
# COPY BUTTON
# ══════════════════════════════════════════════════════════════════════════════
def copy_button(text, key, L):
    safe = text.replace("\\","\\\\").replace("`","\\`").replace("$","\\$").replace('"','\\"')
    components.html(f"""
<button id="cb{key}" onclick="
  navigator.clipboard.writeText(`{safe}`).then(()=>{{
    const b=document.getElementById('cb{key}');
    b.textContent='{L['copy_done']}';
    b.style.background='#ECFDF5';b.style.borderColor='#10B981';b.style.color='#065F46';
    setTimeout(()=>{{b.textContent='{L['copy_btn']}';b.style.background='';b.style.borderColor='';b.style.color='';}},2000);
  }})" style="
  width:100%;padding:10px 16px;background:#FFFFFF;border:1.5px solid #DDDCDA;
  border-radius:9px;color:#5C5A55;font-size:13px;font-weight:600;cursor:pointer;
  font-family:Inter,'Noto Sans KR',sans-serif;transition:all 0.18s;
">{L['copy_btn']}</button>""", height=44)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    L = L_ALL[st.session_state.lang]

    # Logo
    st.markdown(f"""
<div style="display:flex;align-items:center;gap:10px;padding-bottom:16px;
  border-bottom:1px solid #EBEBEA;margin-bottom:16px">
  <div style="width:36px;height:36px;background:#1A1915;border-radius:10px;
    display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:17px;font-weight:800;flex-shrink:0">✦</div>
  <div>
    <div style="font-size:15px;font-weight:800;color:#1A1915;letter-spacing:-0.3px">Humanize AI</div>
    <div style="font-size:11px;color:#9B9790;margin-top:1px">{L['brand_sub']}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Language toggle ────────────────────────────────────────────────────────
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
    letter-spacing:0.6px;margin-bottom:10px">{L['nav_tools']}</p>
  <div style="display:flex;flex-direction:column;gap:4px">
    <div style="padding:10px 12px;border-radius:9px;background:#EEF2FA;
      border-left:3px solid #1B3A6B">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
        <span style="font-size:14px">✦</span>
        <span style="font-size:13px;font-weight:700;color:#1B3A6B">{L['nav_humanizer']}</span>
        <span style="margin-left:auto;font-size:10px;background:#1B3A6B;color:#fff;
          padding:2px 6px;border-radius:4px;font-weight:700">{L['nav_active'].upper()}</span>
      </div>
      <div style="font-size:11px;color:#6B7A99;line-height:1.5;padding-left:22px">
        {L['card_h_desc']}
      </div>
    </div>
    <a href="/GPA_Calculator" target="_self" style="text-decoration:none">
      <div style="padding:10px 12px;border-radius:9px;border:1px solid #E2E8F0;
        background:#FFFFFF;transition:all 0.15s;border-left:3px solid #047857"
        onmouseenter="this.style.background='#F0FDF4'"
        onmouseleave="this.style.background='#FFFFFF'">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
          <span style="font-size:14px">🎓</span>
          <span style="font-size:13px;font-weight:600;color:#1E293B">{L['nav_gpa']}</span>
          <span style="margin-left:auto;font-size:12px;color:#94A3B8">→</span>
        </div>
        <div style="font-size:11px;color:#6B7A99;line-height:1.5;padding-left:22px">
          {L['card_g_desc']}
        </div>
      </div>
    </a>
  </div>
</div>
""", unsafe_allow_html=True)




# ══════════════════════════════════════════════════════════════════════════════
# REFRESH L after sidebar (sidebar may have changed lang)
# ══════════════════════════════════════════════════════════════════════════════
L = L_ALL[st.session_state.lang]

# ══════════════════════════════════════════════════════════════════════════════
# MAIN HEADER + TOOL CARDS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="background:#FFFBEB;border:1.5px solid #FDE68A;border-radius:14px;
  padding:14px 18px;display:flex;align-items:flex-start;gap:12px;margin-bottom:20px">
  <div style="width:32px;height:32px;background:#F59E0B;border-radius:9px;
    display:flex;align-items:center;justify-content:center;font-size:15px;
    flex-shrink:0;color:#fff;font-weight:800">β</div>
  <div>
    <div style="font-size:13px;font-weight:800;color:#92400E;margin-bottom:3px">
      {L['beta_title']}
    </div>
    <div style="font-size:13px;color:#B45309;line-height:1.6">
      {L['beta_body']}
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div style="margin-bottom:24px;padding-bottom:20px;border-bottom:1.5px solid #E2E8F0">
  <div style="display:flex;align-items:center;gap:14px">
    <div style="width:42px;height:42px;background:#1B3A6B;
      border-radius:10px;display:flex;align-items:center;justify-content:center;
      font-size:18px;flex-shrink:0;color:#fff">✦</div>
    <h1 style="font-size:26px;font-weight:800;color:#1E293B;letter-spacing:-0.5px;margin:0;line-height:1.2">
      {L['page_title']}
    </h1>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MAIN LAYOUT: tabs (left) + JD panel (right)
# ══════════════════════════════════════════════════════════════════════════════
_lang = st.session_state.lang
_cat_labels = CATEGORY_LABELS[_lang]

# JD panel widgets rendered first so variables are ready before tabs
_main_col, _jd_col = st.columns([3, 2], gap="large")

with _jd_col:
    st.markdown(f"""
<div style="background:#F4F6FB;border:1.5px solid #D1D9E8;border-radius:14px;padding:22px 20px">
  <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px">
    <span style="font-size:20px">📋</span>
    <span style="font-size:15px;font-weight:800;color:#1B3A6B;letter-spacing:-0.3px">{L['jd_section_title']}</span>
  </div>
  <p style="font-size:12px;color:#6B7A99;margin:0 0 18px">{L['jd_optional']}</p>
</div>""", unsafe_allow_html=True)
    st.markdown("<div style='height:-14px'></div>", unsafe_allow_html=True)
    _sel_label = st.selectbox(
        L["cat_label"], _cat_labels, key="cat_sel",
    )
    job_category = CATEGORY_KEYS[_cat_labels.index(_sel_label)]
    job_description = st.text_area(
        L["jd_label"], height=340,
        placeholder=L["jd_placeholder"], key="jd_text",
    )

with _main_col:
    tab_cover, tab_cv, tab_essay = st.tabs([
        L["tab_cover"], L["tab_cv"], L["tab_essay"],
    ])


# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATION UI
# ══════════════════════════════════════════════════════════════════════════════
def render_translation(mode, tab_color, L):
    doc_type = L[f"doc_{mode}"]
    dir_key = f"trans_dir_{mode}"
    if dir_key not in st.session_state:
        st.session_state[dir_key] = "ko_en"

    st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
  <span style="font-size:18px">🌐</span>
  <span style="font-size:16px;font-weight:700;color:#1A1915">{L['trans_dir_title']}</span>
</div>""", unsafe_allow_html=True)

    direction_label = st.radio(
        "dir", options=[L["trans_ko_en"], L["trans_en_ko"]],
        index=0 if st.session_state[dir_key] == "ko_en" else 1,
        horizontal=True, key=f"dir_radio_{mode}", label_visibility="collapsed",
    )
    direction = "ko_en" if direction_label == L["trans_ko_en"] else "en_ko"
    st.session_state[dir_key] = direction

    src_lang = ("한국어" if direction == "ko_en" else "영어") if L == L_ALL["ko"] else ("Korean" if direction == "ko_en" else "English")
    tgt_lang = ("영어" if direction == "ko_en" else "한국어") if L == L_ALL["ko"] else ("English" if direction == "ko_en" else "Korean")
    src_flag = "🇰🇷" if direction == "ko_en" else "🇺🇸"
    tgt_flag = "🇺🇸" if direction == "ko_en" else "🇰🇷"

    st.markdown(f"""
<div style="display:flex;align-items:center;justify-content:center;gap:16px;
  margin:12px 0 18px;padding:14px;background:#FFFFFF;
  border:2px solid {tab_color}22;border-radius:14px">
  <div style="text-align:center">
    <div style="font-size:26px">{src_flag}</div>
    <div style="font-size:14px;font-weight:700;color:#1A1915;margin-top:2px">{src_lang}</div>
  </div>
  <div style="font-size:28px;color:{tab_color};font-weight:300">→</div>
  <div style="text-align:center">
    <div style="font-size:26px">{tgt_flag}</div>
    <div style="font-size:14px;font-weight:700;color:{tab_color};margin-top:2px">{tgt_lang}</div>
  </div>
  <div style="margin-left:16px;padding-left:16px;border-left:1px solid #EBEBEA">
    <div style="font-size:11px;color:#B8B6B2">{L['trans_style']}</div>
    <div style="font-size:12px;font-weight:600;color:#5C5A55">{doc_type}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown(f'<p style="font-size:13px;font-weight:600;color:#5C5A55;margin-bottom:6px">{src_flag} {src_lang} {L["trans_src_label"]}</p>', unsafe_allow_html=True)
    trans_input = st.text_area(
        "trans_src", height=240,
        placeholder=f"{src_flag} {src_lang}...",
        key=f"trans_input_{mode}", label_visibility="collapsed",
    )
    if len(trans_input) > 0:
        detected = ("한국어" if L == L_ALL["ko"] else "Korean") if is_korean(trans_input) else "English"
        wrong_dir = (direction == "ko_en" and not is_korean(trans_input)) or (direction == "en_ko" and is_korean(trans_input))
        warn = f" ⚠ {L['trans_warn']}" if wrong_dir else ""
        unit = L["char_unit"]
        st.markdown(f'<p style="font-size:12px;color:#B8B6B2;text-align:right;margin-top:4px">{len(trans_input):,}{unit} · {L["trans_detect"]}: {detected}{warn}</p>', unsafe_allow_html=True)

    if not st.button(f"{L['trans_btn']} {src_flag} → {tgt_flag}", type="primary", use_container_width=True, key=f"run_trans_{mode}"):
        return

    client = get_client()
    if not client:
        st.error(L["err_no_api"]); return
    if not trans_input.strip():
        st.error(L["err_no_text"]); return

    style_guide = {
        "cover": {
            "ko_en": "Korean cover letter → English. Sound like it was originally written in English. Use strong action verbs. Professional tone for English-speaking job markets.",
            "en_ko": "영어 자기소개서 → 한국어. 한국 취업 문화에 맞는 자연스러운 경어체(~했습니다). 직역 금지.",
        },
        "cv": {
            "ko_en": "Korean resume/CV → English. Keep bullet-point format. Use action verbs (Developed, Managed, etc.). Preserve all numbers exactly.",
            "en_ko": "영어 이력서 → 한국어. 불릿 포인트 유지. 숫자·수치 그대로. 간결한 능동 표현.",
        },
        "essay": {
            "ko_en": "Korean personal essay → English. Preserve narrative voice and emotional tone. Feel personal and authentic.",
            "en_ko": "영어 에세이 → 한국어. 서사 흐름과 감정 톤 유지. 자연스러운 한국어 표현 우선.",
        },
    }
    system_prompt = f"""Professional translator for job application documents ({doc_type}).
Direction: {src_lang} → {tgt_lang}
Style: {style_guide[mode][direction]}
Rules: Output ONLY translated text. Preserve formatting, line breaks, bullet points, all numbers."""

    bar = st.progress(0)
    ph = st.empty()
    ph.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:10px;
  padding:12px 16px;display:flex;align-items:center;gap:12px;margin-top:8px">
  <div style="width:28px;height:28px;border-radius:8px;background:{tab_color};
    color:#fff;font-size:13px;display:flex;align-items:center;justify-content:center">🌐</div>
  <div>
    <div style="font-size:13px;font-weight:600;color:#1A1915">{src_lang} → {tgt_lang} {L['trans_running']}</div>
    <div style="font-size:11px;color:#B8B6B2">{L['trans_style_opt']}</div>
  </div>
</div>""", unsafe_allow_html=True)
    bar.progress(40)

    try:
        result = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role":"system","content":system_prompt},{"role":"user","content":trans_input}],
            temperature=0.3,
        ).choices[0].message.content
        bar.progress(100)
    except Exception as e:
        bar.empty(); ph.empty()
        st.error(f"{L['trans_err']}: {e}"); return

    bar.empty(); ph.empty()

    st.markdown(f"""
<div style="background:#F0FDF4;border:1.5px solid #A7F3D0;border-radius:12px;
  padding:12px 16px;display:flex;align-items:center;gap:10px;margin:14px 0">
  <div style="width:28px;height:28px;background:#10B981;border-radius:8px;
    color:#fff;font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center">✓</div>
  <div>
    <div style="font-size:13px;font-weight:700;color:#065F46">{L['trans_done']}</div>
    <div style="font-size:11px;color:#6EE7B7">{src_lang} → {tgt_lang} · {doc_type}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    tl, tr = st.columns(2, gap="medium")
    with tl:
        st.markdown(f'<div style="font-size:12px;font-weight:700;color:#9B9790;margin-bottom:8px">{L["trans_before"]} ({src_lang})</div>', unsafe_allow_html=True)
        st.text_area("t_before", trans_input, height=360, disabled=True, label_visibility="collapsed", key=f"tb_{mode}")
    with tr:
        st.markdown(f'<div style="font-size:12px;font-weight:700;color:#9B9790;margin-bottom:8px">{L["trans_after"]} ({tgt_lang})</div>', unsafe_allow_html=True)
        st.text_area("t_after", result, height=320, label_visibility="collapsed", key=f"ta_{mode}")
        copy_button(result, f"trans_{mode}", L)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        td1, td2, td3 = st.columns(3)
        with td1: st.download_button("TXT", result.encode("utf-8"), "translation.txt", "text/plain", use_container_width=True, key=f"tdl_txt_{mode}")
        with td2:
            try: st.download_button("DOCX", make_docx(result), "translation.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key=f"tdl_docx_{mode}")
            except: pass
        with td3:
            try: st.download_button("PDF", make_pdf(result), "translation.pdf", "application/pdf", use_container_width=True, key=f"tdl_pdf_{mode}")
            except: pass


# ══════════════════════════════════════════════════════════════════════════════
# RENDER TAB
# ══════════════════════════════════════════════════════════════════════════════
def render_tab(mode, tab_color, L, job_description, job_category):
    # ── Humanize mode ──────────────────────────────────────────────────────────
    PLACEHOLDERS = {"cover": L["ph_cover"], "cv": L["ph_cv"], "essay": L["ph_essay"]}
    input_text = st.text_area(
        "input", height=240, placeholder=PLACEHOLDERS[mode],
        key=f"input_{mode}", label_visibility="collapsed",
    )

    char_count = len(input_text)
    if char_count > 0:
        lang_detected = L["char_ko"] if is_korean(input_text) else "English"
        st.markdown(f'<p style="font-size:12px;color:#B8B6B2;text-align:right;margin-top:4px">{char_count:,}{L["char_unit"]} · {lang_detected}</p>', unsafe_allow_html=True)

    run_labels = {"cover": L["run_cover"], "cv": L["run_cv"], "essay": L["run_essay"]}
    clicked = st.button(run_labels[mode], type="primary", use_container_width=True, key=f"run_{mode}")

    if not clicked:
        return

    client = get_client()
    if not client:
        st.error(L["err_no_api"]); return
    if mode == "cover" and not job_description.strip():
        st.error(L["err_no_jd"]); return
    if not input_text.strip():
        st.error(L["err_no_text"]); return

    lang = L["char_ko"] if is_korean(input_text) else "English"

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    bar = st.progress(0)
    step_ph = st.empty()

    SPINNER = "<style>@keyframes spin{to{transform:rotate(360deg)}}</style>"
    SPIN_DIV = f'<div style="width:16px;height:16px;border:2px solid {tab_color};border-top-color:transparent;border-radius:50%;animation:spin 0.8s linear infinite"></div>'

    def show_step(label, step_str, sub=None):
        step_ph.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:10px;
  padding:12px 16px;display:flex;align-items:center;gap:12px;margin-top:8px">
  <div style="width:28px;height:28px;border-radius:8px;background:{tab_color};
    color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;
    justify-content:center;flex-shrink:0;white-space:nowrap">{step_str}</div>
  <div>
    <div style="font-size:13px;font-weight:600;color:#1A1915">{label}</div>
    <div style="font-size:11px;color:#B8B6B2;margin-top:1px">{sub or lang}</div>
  </div>
  <div style="margin-left:auto">{SPIN_DIV}</div>
</div>{SPINNER}""", unsafe_allow_html=True)

    def show_score_step(label, step_str, score=None, attempt=None):
        if score is not None:
            score_col = "#EF4444" if score > 50 else "#F59E0B" if score > 30 else "#10B981"
            score_badge = f'<div style="margin-left:auto;font-size:13px;font-weight:800;color:{score_col}">{score}%</div>'
            sub_txt = f"AI {score}% — {L['step_ai_retry'] if score > 30 else L['step_ai_pass']}"
        else:
            score_badge = f'<div style="margin-left:auto">{SPIN_DIV}</div>'
            sub_txt = f"#{attempt}" if attempt else ""
        step_ph.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:10px;
  padding:12px 16px;display:flex;align-items:center;gap:12px;margin-top:8px">
  <div style="width:28px;height:28px;border-radius:8px;background:#F59E0B;
    color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;
    justify-content:center;flex-shrink:0">{step_str}</div>
  <div>
    <div style="font-size:13px;font-weight:600;color:#1A1915">{label}</div>
    <div style="font-size:11px;color:#B8B6B2;margin-top:1px">{sub_txt}</div>
  </div>
  {score_badge}
</div>{SPINNER}""", unsafe_allow_html=True)

    try:
        show_step(L["step1"], "1/4"); facts = step1_extract(client, job_description, input_text, job_category, mode); bar.progress(15)
        show_step(L["step2"], "2/4"); draft2 = step2_persona(client, facts, input_text, job_category, mode); bar.progress(30)
        show_step(L["step3"], "3/4"); irreg = step3_irregularity(client, draft2, input_text, mode); bar.progress(45)
        show_step(L["step4"], "4/4"); final = step4_guardrail(client, input_text, irreg); bar.progress(55)

        MAX_TRIES = 4
        ai_history = []
        for attempt in range(1, MAX_TRIES + 1):
            show_score_step(f"{L['step_ai']} ({attempt}/{MAX_TRIES})", "AI", attempt=attempt)
            score = step5_ai_score(client, final)
            ai_history.append(score)
            bar.progress(min(55 + int((attempt / MAX_TRIES) * 40), 94))
            if score <= 30:
                show_score_step(f"AI {score}% — {L['step_ai_pass']}", "AI", score=score); break
            show_score_step(f"AI {score}% — {L['step_ai_retry']}", "AI", score=score)
            final = step5_refine(client, final, input_text)

        bar.progress(100)
        final_ai_score = ai_history[-1]

    except Exception as e:
        bar.empty(); step_ph.empty()
        st.error(f"Error: {e}"); raise

    bar.empty(); step_ph.empty()

    # ── Result banner ──
    diff = len(final) - len(input_text)
    sign = "+" if diff >= 0 else ""
    diff_col = "#10B981" if diff >= 0 else "#EF4444"
    passed = final_ai_score <= 30
    banner_bg   = "#F0FDF4" if passed else "#FFFBEB"
    banner_bord = "#A7F3D0" if passed else "#FDE68A"
    banner_icon_bg = "#10B981" if passed else "#F59E0B"
    banner_title_col = "#065F46" if passed else "#78350F"
    banner_sub_col   = "#6EE7B7" if passed else "#FCD34D"
    banner_title = L["done_ok"] if passed else L["done_max"]
    banner_sub = f"AI {final_ai_score}% · {lang} · ×{len(ai_history)}"

    history_dots = ""
    for i, s in enumerate(ai_history):
        dot_col = "#10B981" if s <= 30 else "#F59E0B" if s <= 60 else "#EF4444"
        arrow = " → " if i < len(ai_history) - 1 else ""
        history_dots += f'<span style="font-weight:700;color:{dot_col}">{s}%</span><span style="color:#B8B6B2;font-size:11px">{arrow}</span>'

    st.markdown(f"""
<div style="background:{banner_bg};border:1.5px solid {banner_bord};border-radius:12px;
  padding:14px 18px;display:flex;align-items:center;gap:12px;margin:16px 0">
  <div style="width:32px;height:32px;background:{banner_icon_bg};border-radius:8px;
    display:flex;align-items:center;justify-content:center;font-size:16px;
    color:#fff;font-weight:700;flex-shrink:0">{'✓' if passed else '!'}</div>
  <div style="flex:1">
    <div style="font-size:14px;font-weight:700;color:{banner_title_col}">{banner_title}</div>
    <div style="font-size:12px;color:{banner_sub_col};margin-top:2px">{banner_sub}</div>
  </div>
  <div style="font-size:12px;text-align:right">
    <div style="font-size:10px;font-weight:600;color:#9B9790;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:3px">{L['score_hist']}</div>
    <div>{history_dots}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Metric cards ──
    ai_score_col = "#10B981" if final_ai_score <= 30 else "#F59E0B" if final_ai_score <= 60 else "#EF4444"
    if mode == "cover" and job_description.strip():
        ats = ats_score(final, job_description)
        ats_col = "#10B981" if ats >= 60 else "#F59E0B" if ats >= 35 else "#EF4444"
        m1,m2,m3,m4 = st.columns(4)
        cards = [
            (f"{final_ai_score}", "%", L["metric_ai"], ai_score_col),
            (str(ats), "/100", L["metric_ats"], ats_col),
            (f"{len(input_text):,}", L["char_unit"], L["metric_orig"], "#1A1915"),
            (f"{sign}{diff:,}", L["char_unit"], L["metric_diff"], diff_col),
        ]
        cols_list = [m1,m2,m3,m4]
    else:
        m1,m2,m3 = st.columns(3)
        cards = [
            (f"{final_ai_score}", "%", L["metric_ai"], ai_score_col),
            (f"{len(input_text):,}", L["char_unit"], L["metric_orig"], "#1A1915"),
            (f"{sign}{diff:,}", L["char_unit"], L["metric_diff"], diff_col),
        ]
        cols_list = [m1,m2,m3]

    for col, (val, unit, label, col_color) in zip(cols_list, cards):
        with col:
            st.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:14px;
  padding:18px;text-align:center;margin-bottom:4px">
  <div style="font-size:10px;font-weight:700;color:#9B9790;text-transform:uppercase;
    letter-spacing:0.6px;margin-bottom:8px">{label}</div>
  <div style="display:flex;align-items:baseline;justify-content:center;gap:3px">
    <span style="font-size:30px;font-weight:800;color:{col_color};letter-spacing:-1.5px">{val}</span>
    <span style="font-size:12px;color:#B8B6B2;font-weight:500">{unit}</span>
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Before / After ──
    col_l, col_r = st.columns(2, gap="medium")
    with col_l:
        st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
  <div style="width:8px;height:8px;border-radius:50%;background:#EF4444;flex-shrink:0"></div>
  <span style="font-size:13px;font-weight:700;color:#1A1915">{L['before_label']}</span>
  <span style="font-size:11px;color:#B8B6B2;font-weight:400">{L['before_sub']}</span>
</div>""", unsafe_allow_html=True)
        st.text_area("before_disp", input_text, height=420, disabled=True, label_visibility="collapsed", key=f"ba_before_{mode}")

    with col_r:
        st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
  <div style="width:8px;height:8px;border-radius:50%;background:{tab_color};flex-shrink:0"></div>
  <span style="font-size:13px;font-weight:700;color:#1A1915">{L['after_label']}</span>
  <span style="font-size:11px;color:#B8B6B2;font-weight:400">{L['after_sub']}</span>
</div>""", unsafe_allow_html=True)
        st.text_area("after_disp", final, height=380, label_visibility="collapsed", key=f"ba_after_{mode}")
        copy_button(final, f"{mode}_main", L)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(f'<p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">{L["dl_label"]}</p>', unsafe_allow_html=True)
        fname = {"cover":"cover_letter","cv":"cv","essay":"essay"}[mode]
        d1,d2,d3,d4 = st.columns(4)
        with d1: st.download_button("TXT", final.encode("utf-8"), f"{fname}.txt","text/plain",use_container_width=True,key=f"dl_txt_{mode}")
        with d2:
            try: st.download_button("DOCX",make_docx(final),f"{fname}.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True,key=f"dl_docx_{mode}")
            except: st.caption("DOCX error")
        with d3:
            try: st.download_button("PDF",make_pdf(final),f"{fname}.pdf","application/pdf",use_container_width=True,key=f"dl_pdf_{mode}")
            except: st.caption("PDF error")
        with d4:
            try: st.download_button("DOC",make_docx(final),f"{fname}.doc","application/msword",use_container_width=True,key=f"dl_doc_{mode}")
            except: st.caption("DOC error")

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    with st.expander(L["facts_expander"]):
        st.code(facts, language="yaml")


# ══════════════════════════════════════════════════════════════════════════════
# TAB RENDER
# ══════════════════════════════════════════════════════════════════════════════
def tab_info_banner(icon, title, desc, bg, text_col, sub_col):
    st.markdown(f"""
<div style="background:{bg};border-radius:12px;padding:14px 18px;margin-bottom:20px;
  display:flex;align-items:flex-start;gap:12px">
  <div style="font-size:22px;flex-shrink:0">{icon}</div>
  <div>
    <div style="font-size:14px;font-weight:700;color:{text_col};margin-bottom:3px">{title}</div>
    <div style="font-size:12px;color:{sub_col};line-height:1.6">{desc}</div>
  </div>
</div>""", unsafe_allow_html=True)


with tab_cover:
    m = TAB_META["cover"]
    tab_info_banner("📝", L["banner_cover"], L["banner_cover_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("cover", m["color"], L, job_description, job_category)

with tab_cv:
    m = TAB_META["cv"]
    tab_info_banner("📄", L["banner_cv"], L["banner_cv_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("cv", m["color"], L, job_description, job_category)

with tab_essay:
    m = TAB_META["essay"]
    tab_info_banner("✍️", L["banner_essay"], L["banner_essay_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("essay", m["color"], L, job_description, job_category)
