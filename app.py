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
    page_icon="??,
    initial_sidebar_state="expanded",
)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# LANGUAGE STATE
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
if "lang" not in st.session_state:
    st.session_state.lang = "ko"

L_ALL = {
    "ko": {
        # sidebar
        "brand_sub":       "Student Tools Suite",
        "nav_tools":       "?꾧뎄",
        "nav_humanizer":   "AI ?멸컙??,
        "nav_gpa":         "GPA 怨꾩궛湲?,
        "nav_active":      "?꾩옱",
        "job_category":    "吏곷Т 移댄뀒怨좊━",
        "jd_label":        "梨꾩슜 怨듦퀬 JD",
        "jd_hint":         "?먯냼??紐⑤뱶?먯꽌 ?ъ슜?⑸땲??,
        "jd_placeholder":  "怨듦퀬 ?꾨Ц??遺숈뿬?ｌ쑝?몄슂...\n?쒓뎅?는룹쁺??紐⑤몢 吏?먰빀?덈떎.",
        "pipeline_title":  "?뚯씠?꾨씪??,
        "pipe1":           "寃쏀뿕쨌?ㅼ썙??異붿텧",
        "pipe2":           "?섎Ⅴ?뚮굹 珥덉븞 ?앹꽦",
        "pipe3":           "遺덇퇋移숈꽦 二쇱엯",
        "pipe4":           "?⑺듃 寃??,
        # header
        "page_title":      "AI 臾몄꽌 ?멸컙??,
        "card_humanizer":  "AI ?멸컙??,
        "card_h_desc":     "?먯냼?쑣텰V쨌?먯꽭?대? AI ?붿쟻 ?놁씠 ?먯뿰?ㅻ읇寃??ъ옉??,
        "card_h_cta":      "?꾩옱 ?섏씠吏 ??,
        "card_trans":      "踰덉뿭",
        "card_t_desc":     "?쒓뎅?????곸뼱 吏곷Т 臾몄꽌 踰덉뿭 (媛??????좉?)",
        "card_t_cta":      "媛???뿉???ъ슜 媛??,
        "card_gpa":        "GPA 怨꾩궛湲?,
        "card_g_desc":     "4.0 / 4.3 ?ㅼ???쨌 臾댁젣??怨쇰ぉ 쨌 紐⑺몴 GPA 怨꾩궛湲?,
        "card_g_cta":      "諛붾줈 媛湲???,
        # tabs
        "tab_cover":       "?뱷  ?먯냼??,
        "tab_cv":          "?뱞  CV / ?대젰??,
        "tab_essay":       "?랃툘  ?먯꽭??,
        "banner_cover":    "?먯냼??紐⑤뱶",
        "banner_cover_d":  "?ъ씠?쒕컮?먯꽌 梨꾩슜 怨듦퀬瑜??낅젰?섎㈃ <strong>ATS ?ㅼ썙??理쒖쟻??/strong>源뚯? ?④퍡 吏꾪뻾?⑸땲?? AI媛 ??寃?媛숈? ?붿쟻???쒓굅?섍퀬 ?멸컙?곸씤 臾몄껜濡?諛붽퓠?덈떎.",
        "banner_cv":       "CV / ?대젰??紐⑤뱶",
        "banner_cv_d":     "遺덈┸ ?ъ씤???뺤떇???좎??섎㈃??<strong>?섎룞???쒗쁽???λ룞 ?숈궗濡?/strong> 諛붽퓠?덈떎. 臾몄옣 湲몄씠瑜??섎룄?곸쑝濡?遺덇퇏?쇳븯寃?留뚮뱾???щ엺??吏곸젒 ??寃껋쿂??蹂댁씠寃??⑸땲??",
        "banner_essay":    "?먯꽭??紐⑤뱶",
        "banner_essay_d":  "媛쒖씤 ?먯꽭?댁쓽 <strong>?쒖궗???먮쫫怨?媛먯젙 寃?/strong>???대┰?덈떎. ?ш굔 ??媛먯젙 ??源⑤떖?뚯쓽 援ъ“瑜??좎??섎㈃??AI ?뱀쑀??留ㅻ걚?ъ슫 ?⑦꽩???쒓굅?⑸땲??",
        # mode switch
        "sw_human":        "???멸컙??,
        "sw_trans":        "?뙋 踰덉뿭",
        # placeholders
        "ph_cover":        "?쒓뎅???먮뒗 ?곸뼱濡??묒꽦???먯냼??珥덉븞??遺숈뿬?ｌ쑝?몄슂.\n?먮룞?쇰줈 ?몄뼱瑜?媛먯??섏뿬 理쒖쟻?뷀빀?덈떎.",
        "ph_cv":           "CV ?먮뒗 ?대젰???꾨Ц??遺숈뿬?ｌ쑝?몄슂.\n遺덈┸ ?ъ씤???뺤떇 洹몃?濡?遺숈뿬?ｌ쑝硫??⑸땲??",
        "ph_essay":        "媛쒖씤 ?먯꽭???먮뒗 Personal Statement瑜?遺숈뿬?ｌ쑝?몄슂.\n?쒖궗 ?먮쫫???대━硫댁꽌 AI ?붿쟻???쒓굅?⑸땲??",
        # run buttons
        "run_cover":       "?먯냼??泥⑥궘 ?쒖옉 ??,
        "run_cv":          "CV ?멸컙???쒖옉 ??,
        "run_essay":       "?먯꽭???멸컙???쒖옉 ??,
        # errors
        "err_no_api":      "?쒕쾭??API ?ㅺ? ?ㅼ젙?섏? ?딆븯?듬땲?? 愿由ъ옄?먭쾶 臾몄쓽?섏꽭??",
        "err_no_jd":       "?ъ씠?쒕컮?먯꽌 梨꾩슜 怨듦퀬(JD)瑜??낅젰?댁＜?몄슂.",
        "err_no_text":     "?띿뒪?몃? ?낅젰?댁＜?몄슂.",
        # progress
        "step1":           "寃쏀뿕쨌?ㅼ썙??異붿텧 以?,
        "step2":           "?멸컙??珥덉븞 ?앹꽦 以?,
        "step3":           "遺덇퇋移숈꽦 二쇱엯 以?,
        "step4":           "?⑺듃 寃??以?,
        "step_ai":         "AI 媛먯? ?뺣쪧 痢≪젙 以?,
        "step_ai_pass":    "湲곗?移??듦낵!",
        "step_ai_retry":   "30% 紐⑺몴 誘몃떖, ?ъ닔??以?..",
        # result
        "done_ok":         "?멸컙???꾨즺",
        "done_max":        "?멸컙???꾨즺 (理쒕? ?ъ떆???꾨떖)",
        "score_hist":      "AI ?먯닔 蹂??,
        "metric_ai":       "AI 媛먯? ?뺣쪧",
        "metric_ats":      "ATS ?먯닔",
        "metric_orig":     "?먮낯 湲몄씠",
        "metric_diff":     "湲몄씠 蹂??,
        "before_label":    "?먮낯",
        "before_sub":      "AI 媛먯? ?꾪뿕",
        "after_label":     "?멸컙??寃곌낵",
        "after_sub":       "AI ?먯? ?고쉶",
        "dl_label":        "?ㅼ슫濡쒕뱶",
        "copy_btn":        "?대┰蹂대뱶 蹂듭궗",
        "copy_done":       "蹂듭궗 ?꾨즺 ??,
        "facts_expander":  "異붿텧???ъ떎 ?곗씠??蹂닿린",
        "char_ko":         "?쒓뎅??,
        "char_unit":       "??,
        # translation UI
        "trans_dir_title": "踰덉뿭 諛⑺뼢 ?좏깮",
        "trans_ko_en":     "?눖?눟 ?쒓뎅?? ?? ?눣?눡 ?곸뼱",
        "trans_en_ko":     "?눣?눡 ?곸뼱  ?? ?눖?눟 ?쒓뎅??,
        "trans_src_label": "?먮Ц 遺숈뿬?ｊ린",
        "trans_btn":       "踰덉뿭 ?쒖옉",
        "trans_style":     "臾몄껜",
        "trans_done":      "踰덉뿭 ?꾨즺",
        "trans_before":    "?먮Ц",
        "trans_after":     "踰덉뿭臾?,
        "trans_err":       "踰덉뿭 ?ㅻ쪟",
        "trans_running":   "踰덉뿭 以?..",
        "trans_style_opt": "臾몄껜 理쒖쟻???곸슜 以?,
        "trans_warn":      "諛⑺뼢 ?뺤씤",
        "trans_detect":    "媛먯?",
        "doc_cover":       "?먭린?뚭컻??,
        "doc_cv":          "?대젰??CV",
        "doc_essay":       "媛쒖씤 ?먯꽭??,
        "beta_title":      "踰좏? ?쒕퉬???덈궡",
        "beta_body":       "?꾩옱 踰좏? 踰꾩쟾?쇰줈 ?댁쁺 以묒엯?덈떎. AI ?멸컙??寃곌낵媛 100% ?꾨꼍?섏? ?딆쓣 ???덉쑝硫? ?쇰? 臾몄옣? ?ъ쟾??AI媛 ?묒꽦??寃껋쿂??媛먯??????덉뒿?덈떎. 寃곌낵臾쇱쓣 諛섎뱶??吏곸젒 寃?????ъ슜??二쇱꽭??",
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
        "card_h_cta":      "Current page ??,
        "card_trans":      "Translate",
        "card_t_desc":     "Korean ??English translation for job documents (toggle in each tab)",
        "card_t_cta":      "Available in each tab",
        "card_gpa":        "GPA Calculator",
        "card_g_desc":     "4.0 / 4.3 Scale 쨌 Unlimited Courses 쨌 Goal GPA Estimator",
        "card_g_cta":      "Go ??,
        # tabs
        "tab_cover":       "?뱷  Cover Letter",
        "tab_cv":          "?뱞  CV / Resume",
        "tab_essay":       "?랃툘  Essay",
        "banner_cover":    "Cover Letter Mode",
        "banner_cover_d":  "Add a job description in the sidebar to enable <strong>ATS keyword optimization</strong>. Removes AI-written patterns and rewrites in a genuine human voice.",
        "banner_cv":       "CV / Resume Mode",
        "banner_cv_d":     "Preserves bullet-point formatting while converting <strong>passive expressions to strong action verbs</strong>. Intentionally varies sentence length to read like a real person wrote it.",
        "banner_essay":    "Essay Mode",
        "banner_essay_d":  "Preserves the <strong>narrative arc and emotional tone</strong> of personal essays. Maintains the event ??emotion ??insight structure while removing AI-typical smooth patterns.",
        # mode switch
        "sw_human":        "??Humanize",
        "sw_trans":        "?뙋 Translate",
        # placeholders
        "ph_cover":        "Paste your cover letter draft here (English or Korean).\nLanguage is detected automatically.",
        "ph_cv":           "Paste your full CV or resume here.\nBullet-point formatting is preserved.",
        "ph_essay":        "Paste your personal essay or Personal Statement here.\nNarrative flow is preserved while removing AI traces.",
        # run buttons
        "run_cover":       "Humanize Cover Letter ??,
        "run_cv":          "Humanize CV ??,
        "run_essay":       "Humanize Essay ??,
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
        "step_ai_retry":   "Target 30% not met ??refining...",
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
        "copy_done":       "Copied ??,
        "facts_expander":  "View extracted fact data",
        "char_ko":         "Korean",
        "char_unit":       "chars",
        # translation UI
        "trans_dir_title": "Select Translation Direction",
        "trans_ko_en":     "?눖?눟 Korean  ?? ?눣?눡 English",
        "trans_en_ko":     "?눣?눡 English  ?? ?눖?눟 Korean",
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
        "beta_title":      "Beta Notice",
        "beta_body":       "This tool is currently in beta. Humanization results may not be 100% perfect ??some sentences may still be flagged as AI-written by detection tools. Please review all output carefully before use.",
    },
}

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# GLOBAL CSS
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&family=Inter:wght@400;500;600;700;800&display=swap');

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Inter', 'Noto Sans KR', sans-serif !important; }

.stApp { background: #FAFAF8; }
.block-container { padding: 2.5rem 2.5rem 5rem !important; max-width: 1200px !important; }

section[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #EBEBEA !important;
}
section[data-testid="stSidebar"] .block-container { padding: 2rem 1.25rem !important; }

/* tabs */
div[data-testid="stTabs"] > div:first-child {
    border-bottom: 2px solid #EBEBEA !important; gap: 0 !important; margin-bottom: 28px !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-size: 15px !important; font-weight: 600 !important; color: #9B9790 !important;
    padding: 12px 22px !important; border-radius: 0 !important; background: transparent !important;
    border: none !important; border-bottom: 3px solid transparent !important;
    margin-bottom: -2px !important; transition: all 0.18s !important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:hover {
    color: #1A1915 !important; background: #F5F4F1 !important;
}
div[data-testid="stTabs"] button[aria-selected="true"][data-baseweb="tab"] {
    color: #1A1915 !important; border-bottom-color: #4F46E5 !important;
}
div[data-testid="stTabs"] div[role="tabpanel"] { padding-top: 0 !important; }

/* textarea */
div[data-testid="stTextArea"] textarea {
    background: #FFFFFF !important; color: #1A1915 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 12px !important;
    font-family: 'Inter', 'Noto Sans KR', sans-serif !important;
    font-size: 14.5px !important; line-height: 1.75 !important;
    padding: 14px 16px !important; resize: vertical !important;
    transition: border-color 0.18s, box-shadow 0.18s !important;
}
div[data-testid="stTextArea"] textarea:focus {
    border-color: #4F46E5 !important;
    box-shadow: 0 0 0 3px rgba(79,70,229,0.10) !important; outline: none !important;
}
div[data-testid="stTextArea"] label {
    font-size: 13px !important; font-weight: 600 !important;
    color: #5C5A55 !important; margin-bottom: 6px !important;
}

/* text input */
div[data-testid="stTextInput"] input {
    background: #F7F6F3 !important; color: #1A1915 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 10px !important;
    font-size: 14px !important; padding: 10px 12px !important; transition: all 0.18s !important;
}
div[data-testid="stTextInput"] input:focus {
    background: #fff !important; border-color: #4F46E5 !important;
    box-shadow: 0 0 0 3px rgba(79,70,229,0.10) !important;
}
div[data-testid="stTextInput"] label {
    font-size: 12px !important; font-weight: 700 !important; color: #5C5A55 !important;
    text-transform: uppercase !important; letter-spacing: 0.5px !important;
}

/* selectbox */
div[data-testid="stSelectbox"] > div > div {
    background: #F7F6F3 !important; border: 1.5px solid #DDDCDA !important;
    border-radius: 10px !important; color: #1A1915 !important; font-size: 14px !important;
    transition: border-color 0.18s !important;
}
div[data-testid="stSelectbox"] label {
    font-size: 12px !important; font-weight: 700 !important; color: #5C5A55 !important;
    text-transform: uppercase !important; letter-spacing: 0.5px !important;
}

/* primary button */
.stButton > button[kind="primary"] {
    background: #1A1915 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 12px !important;
    padding: 15px 28px !important; font-size: 15px !important;
    font-weight: 700 !important; letter-spacing: -0.2px !important;
    transition: all 0.18s !important; height: auto !important;
}
.stButton > button[kind="primary"]:hover {
    background: #2D2C28 !important; transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"]:active { transform: translateY(0) !important; }

/* secondary button */
.stButton > button[kind="secondary"] {
    background: #FFFFFF !important; color: #5C5A55 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 10px !important;
    font-size: 14px !important; font-weight: 600 !important;
    padding: 12px 20px !important; transition: all 0.18s !important; height: auto !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #1A1915 !important; color: #1A1915 !important; background: #F5F4F1 !important;
}

/* download button */
.stDownloadButton > button {
    background: #FFFFFF !important; color: #5C5A55 !important;
    border: 1.5px solid #DDDCDA !important; border-radius: 9px !important;
    font-size: 13px !important; font-weight: 600 !important;
    padding: 8px 12px !important; transition: all 0.18s !important;
}
.stDownloadButton > button:hover {
    border-color: #1A1915 !important; color: #1A1915 !important; background: #F5F4F1 !important;
}

/* expander */
div[data-testid="stExpander"] {
    border: 1.5px solid #EBEBEA !important; border-radius: 12px !important;
    background: #FFFFFF !important; overflow: hidden !important; box-shadow: none !important;
}
div[data-testid="stExpander"] summary {
    color: #5C5A55 !important; font-weight: 600 !important;
    font-size: 13px !important; padding: 12px 16px !important;
}
div[data-testid="stExpander"] summary:hover { background: #F7F6F3 !important; }

/* progress */
.stProgress > div > div { background: #EBEBEA !important; border-radius: 99px !important; height: 6px !important; }
.stProgress > div > div > div { border-radius: 99px !important; height: 6px !important; }

/* alert */
div[data-testid="stAlert"] { border-radius: 12px !important; font-size: 14px !important; font-weight: 500 !important; }

/* scrollbar */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #DDDCDA; border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: #B8B6B2; }

/* language toggle pill */
.lang-toggle {
    display: flex; background: #F0EFEC; border-radius: 10px;
    padding: 3px; gap: 3px; margin-bottom: 20px;
}
.lang-btn {
    flex: 1; border: none; cursor: pointer; border-radius: 7px;
    padding: 8px 0; font-size: 13px; font-weight: 700;
    transition: all 0.18s; font-family: 'Inter', sans-serif;
}
</style>
""", unsafe_allow_html=True)


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# CONSTANTS
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
ACTION_VERBS = {
    "Strategy / Finance": ["Formulated","Forecasted","Evaluated","Structured","Negotiated","Synthesized","Streamlined"],
    "Marketing":          ["Launched","Optimized","Analyzed","Cultivated","Amplified","Positioned","Executed"],
    "Engineering":        ["Developed","Architected","Automated","Deployed","Refactored","Integrated","Scaled"],
    "Product Management": ["Prioritized","Collaborated","Shipped","Defined","Validated","Iterated","Championed"],
    "Data / AI":          ["Modeled","Trained","Predicted","Visualized","Extracted","Transformed","Benchmarked"],
    "Operations / HR":    ["Coordinated","Implemented","Facilitated","Reduced","Improved","Managed","Standardized"],
}

TAB_META = {
    "cover": {"color": "#4F46E5", "light": "#EEF2FF", "info_bg": "#EEF2FF", "info_text": "#3730A3", "info_sub": "#6366F1"},
    "cv":    {"color": "#0891B2", "light": "#ECFEFF", "info_bg": "#ECFEFF", "info_text": "#164E63", "info_sub": "#0E7490"},
    "essay": {"color": "#7C3AED", "light": "#F5F3FF", "info_bg": "#F5F3FF", "info_text": "#4C1D95", "info_sub": "#7C3AED"},
}


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# UTILS
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
    return sum(1 for c in text if '媛' <= c <= '??) / max(len(text), 1) > 0.1


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# AI PIPELINE
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
        sys = "?덈뒗 10??寃쎈젰 ?ㅻ뱶?뚰꽣?? AI ?꾩깉 ?섎뒗 ?쒗쁽???꾨? 嫄룹뼱?닿퀬 ?ㅼ젣 ?щ엺????寃껋쿂??諛붽퓭." if kor else "You're a senior headhunter rewriting a resume to sound genuinely human."
        task = (f"?꾨옒 ?곗씠?곕줈 ?대젰?쒕? ?멸컙?뷀빐. 遺덈┸ ?좎?, ?レ옄 ?좎?, ?섎룞?믩뒫?? 湲몄씠 遺덇퇏?쇳븯寃?\n湲덉??? ?곸떊?곸씤, ?곸썡?? 理쒖꽑, ?댁젙\n\n[?곗씠??\n{facts}\n\n[?먮낯]\n{draft}\n\n蹂몃Ц留?異쒕젰." if kor
                else f"Humanize this CV. Keep bullets+numbers. Active verbs. Vary lengths. No: passionate/innovative/leveraged.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\n[VERBS]\n{verbs}\n\nOutput CV text only.")
    elif mode == "essay":
        sys = "?덈뒗 ??숈썝 ?낇븰?ъ젙愿 異쒖떊 湲?곌린 肄붿튂?? ?쒖궗 ?먮쫫???대━怨?AI ?붿쟻???쒓굅??" if kor else "You're a writing coach who reads college essays. Human > polished."
        task = (f"?먯꽭???멸컙?? ?쒖궗 ?좎?(?ш굔?믨컧?뺚넂源⑤떖??, 臾몄옣 湲몄씠 遺덇퇋移? 援ъ뼱泥??곌껐???쎌엯, ?대┛ 寃곕쭚 ?덉슜.\n湲덉?: ?곸떊,?댁젙,?꾩쟾,?깆옣\n\n[?곗씠??\n{facts}\n\n[?먮낯]\n{draft}\n\n蹂몃Ц留?異쒕젰." if kor
                else f"Humanize this essay. Keep arc (event?뭙motion?뭝nsight). Vary length. Use contractions. No: passionate/transformative/journey.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\nOutput only.")
    else:
        sys = ("?덈뒗 ?湲곗뾽 ?몄궗? 異쒖떊 40? 而⑥꽕?댄듃?? 諛?11?? ?쇨낀?섏?留?寃쏀뿕 ?띾?.\n臾몄껜: 吏㏃?臾몄옣+湲대Ц???쇳빀, 援ъ뼱泥??곌껐?? ?⑤씫 湲몄씠 遺덇퇏??\n湲덉??? ?곸떊,?댁젙,?꾩쟾,?쒕꼫吏,湲곗뿬,?꾪븯寃좎뒿?덈떎" if kor
               else "You're a 35-yo ex-recruiter editing at 11 PM. Slightly imperfect. Mix short+long sentences. No: passionate/innovative/leverage/synergy/dynamic")
        task = (f"?먯냼???ㅼ떆 ?⑥쨾. ?먮낯 ?녿뒗 ?댁슜 異붽? 湲덉?. '??? 3?곗냽 湲덉?. 蹂묐젹援ъ“ 諛섎났 湲덉?. ?レ옄 ?좎?.\n\n[?ъ떎]\n{facts}\n\n[?먮낯]\n{draft}\n\n[?숈궗]\n{verbs}\n\n蹂몃Ц留?異쒕젰." if kor
                else f"Rewrite ONLY using facts below. No 3+ 'I'-starts. Break repeated parallel. Preserve numbers.\n\n[FACTS]\n{facts}\n\n[ORIGINAL]\n{draft}\n\n[VERBS]\n{verbs}\n\nCover letter body only.")
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"system","content":sys},{"role":"user","content":task}],
        temperature=1.0, presence_penalty=0.7, frequency_penalty=0.5,
    ).choices[0].message.content


def step3_irregularity(client, text, draft, mode):
    kor = is_korean(draft)
    if mode == "cv":
        p = ("CV?먯꽌 AI ?⑦꽩 ?쒓굅.\n1.媛숈? 援ъ“ 遺덈┸??媛??ㅻⅤ寃?2.?レ옄 ?쒓린 ?ㅼ뼇??3.留덉?留?遺덈┸ 誘몄셿???먮굦\n?⑺듃 蹂寃?湲덉?. ?섏젙蹂몃쭔." if kor
             else "Remove AI patterns from CV.\n1.Vary 2 bullets if all same structure 2.Mix number formats 3.Last bullet slightly open\nNo fact changes. Output only.")
    elif mode == "essay":
        p = ("?먯꽭??AI ?⑦꽩 ?쒓굅.\n1.湲?臾몄옣 2媛쒋넂??쒕줈 ?딄린 2.吏㏃? 臾몄옣 2媛??좎? 3.紐⑤뱺 ?⑤씫??寃곕줎?대㈃??媛?吏덈Ц/?ъ슫?쇰줈 4.鍮꾩듂???⑤씫 3媛쒋넁??媛?1臾몄옣?쇰줈\n?⑺듃 蹂寃?湲덉?." if kor
             else "Remove AI patterns from essay.\n1.Split 2 longest sentences with em-dash 2.Keep 2 shortest verbatim 3.If every para ends with conclusion?? open thought 4.3+ similar-length paras?뭖ollapse 1\nNo changes to facts. Output only.")
    else:
        p = ("?먯냼??AI ?붿쟻 ?쒓굅.\n1.湲?臾몄옣 3媛쒋넂??쒕줈 ?딄린 2.吏㏃? 臾몄옣 2媛??좎? 3.~?덉뒿?덈떎 3?곗냽??媛?紐낆궗??4.3???섏뿴??媛?鍮꾪?湲?5.留덉?留?臾몄옣?믪뿴由?寃곕쭚 6.媛숈? 湲몄씠 ?⑤씫 3媛쒋넁??媛?1臾몄옣\n?⑺듃 蹂寃?湲덉?. ?섏젙蹂몃쭔." if kor
             else "Remove AI patterns from cover letter.\n1.Split 3 longest w/ em-dash 2.Keep 2 shortest verbatim 3.3+ 'I' starts?뭖hange one 4.Repeated parallel?뭕reak one 5.Final line?뭪railing thought 6.3+ same-length paras?뭖ollapse one\nNo facts changed. Output only.")
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"user","content":f"{p}\n\n[湲]\n{text}"}],
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
- Buzzwords: innovative, passionate, leverage, synergy, ?곸떊, ?댁젙, ?쒕꼫吏
- Every paragraph ends with a tidy conclusion
- No hesitation, no imperfection, no personal voice
- Formulaic transitions: Furthermore, Additionally, In conclusion, ?먰븳, ?곕씪?? 寃곕줎?곸쑝濡?
Criteria for LOW AI probability (score closer to 0):
- Irregular sentence lengths (very short + very long mixed)
- Natural interruptions (em-dashes, incomplete thoughts)
- Colloquial connectors: Honestly, Here's the thing, ?붿쭅?? 洹몃븣, ?ъ떎?
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
        prompt = f"""?꾨옒 湲??AI媛 ??寃?媛숇떎???됯?瑜?諛쏆븯?? ???멸컙?곸쑝濡?諛붽퓭以?

[湲]
{text}

?곸슜???섏젙 (紐⑤몢 ?곸슜):
1. 媛??AI?ㅻ윭??臾몄옣 3媛????꾩쟾???ㅻⅨ 援ъ“濡??ъ옉??2. ?꾨꼍?섍쾶 留덈Т由щ맂 ?⑤씫 1媛????대┛ 寃곕쭚 ?먮뒗 ?ъ슫?쇰줈
3. ?덈Т 留ㅻ걚?ъ슫 ?꾪솚???먰븳, ?곕씪?? 寃곕줎?곸쑝濡? ??援ъ뼱泥대줈 援먯껜
4. 媛??湲?臾몄옣 2媛????????濡?以묎컙???딄린
5. 臾몄옣 3媛??댁긽 媛숈? 湲몄씠 ?⑦꽩 ???섎룄?곸쑝濡?源④린
6. ?먯뿰?ㅻ윭??留앹꽕???먮뒗 諛⑺뼢 ?꾪솚 1怨?異붽?

?⑺듃(?レ옄, ?대쫫, 寃쏀뿕) 蹂寃??덈? 湲덉?. ?섏젙蹂몃쭔 異쒕젰."""
    else:
        prompt = f"""The text below was flagged as likely AI-written. Make it sound more human.

[TEXT]
{text}

Apply ALL of these:
1. Rewrite the 3 most AI-sounding sentences with completely different structures
2. Take 1 perfectly-concluded paragraph ??rewrite ending as open thought or question
3. Replace smooth transitions (Furthermore, Additionally, In conclusion) ??conversational alternatives
4. Split the 2 longest sentences with an em-dash at a natural break
5. Break any 3+ sentence pattern of similar length
6. Add 1 moment of natural hesitation or shift in direction

Never change facts (numbers, names, experiences). Output revised text only."""
    return client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.9, presence_penalty=0.6, frequency_penalty=0.4,
    ).choices[0].message.content


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# FILE GENERATION
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def make_docx(text):
    doc = Document()
    s = doc.styles['Normal']
    s.font.name = '留묒? 怨좊뵓' if is_korean(text) else 'Calibri'
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


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# COPY BUTTON
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# SIDEBAR
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
with st.sidebar:
    L = L_ALL[st.session_state.lang]

    # Logo
    st.markdown(f"""
<div style="display:flex;align-items:center;gap:10px;padding-bottom:16px;
  border-bottom:1px solid #EBEBEA;margin-bottom:16px">
  <div style="width:36px;height:36px;background:#1A1915;border-radius:10px;
    display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:17px;font-weight:800;flex-shrink:0">??/div>
  <div>
    <div style="font-size:15px;font-weight:800;color:#1A1915;letter-spacing:-0.3px">Humanize AI</div>
    <div style="font-size:11px;color:#9B9790;margin-top:1px">{L['brand_sub']}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ?? Language toggle ????????????????????????????????????????????????????????
    lc1, lc2 = st.columns(2, gap="small")
    with lc1:
        if st.button("?눖?눟  ?쒓뎅??, key="lang_ko", use_container_width=True,
                     type="primary" if st.session_state.lang == "ko" else "secondary"):
            st.session_state.lang = "ko"; st.rerun()
    with lc2:
        if st.button("?눣?눡  English", key="lang_en", use_container_width=True,
                     type="primary" if st.session_state.lang == "en" else "secondary"):
            st.session_state.lang = "en"; st.rerun()

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # Navigation
    st.markdown(f"""
<div style="margin-bottom:20px;padding-bottom:20px;border-bottom:1px solid #EBEBEA">
  <p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;
    letter-spacing:0.6px;margin-bottom:10px">{L['nav_tools']}</p>
  <div style="display:flex;flex-direction:column;gap:2px">
    <div style="display:flex;align-items:center;gap:10px;padding:9px 12px;
      border-radius:9px;background:#F0EFEC;">
      <span style="font-size:15px">??/span>
      <span style="font-size:13px;font-weight:700;color:#1A1915">{L['nav_humanizer']}</span>
      <span style="margin-left:auto;font-size:10px;background:#1A1915;color:#fff;
        padding:2px 6px;border-radius:4px;font-weight:700">{L['nav_active'].upper()}</span>
    </div>
    <a href="/GPA_Calculator" target="_self" style="display:flex;align-items:center;gap:10px;
      padding:9px 12px;border-radius:9px;text-decoration:none;transition:all 0.15s;color:#5C5A55;">
      <span style="font-size:15px">?럳</span>
      <span style="font-size:13px;font-weight:600;color:#5C5A55">{L['nav_gpa']}</span>
    </a>
  </div>
</div>
""", unsafe_allow_html=True)

    # Job category
    st.markdown(f'<p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px">{L["job_category"]}</p>', unsafe_allow_html=True)
    job_category = st.selectbox("cat", list(ACTION_VERBS.keys()), label_visibility="collapsed")

    # JD
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown(f'<p style="font-size:11px;font-weight:700;color:#9B9790;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px">{L["jd_label"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p style="font-size:11px;color:#B8B6B2;margin-bottom:6px">{L["jd_hint"]}</p>', unsafe_allow_html=True)
    job_description = st.text_area("jd", height=200,
        placeholder=L["jd_placeholder"], label_visibility="collapsed")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# REFRESH L after sidebar (sidebar may have changed lang)
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
L = L_ALL[st.session_state.lang]

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# MAIN HEADER + TOOL CARDS
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
st.markdown(f"""
<div style="background:#FFFBEB;border:1.5px solid #FDE68A;border-radius:14px;
  padding:14px 18px;display:flex;align-items:flex-start;gap:12px;margin-bottom:20px">
  <div style="width:32px;height:32px;background:#F59E0B;border-radius:9px;
    display:flex;align-items:center;justify-content:center;font-size:15px;
    flex-shrink:0;color:#fff;font-weight:800">棺</div>
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
<div style="margin-bottom:28px;padding-bottom:20px;border-bottom:1px solid #EBEBEA">
  <div style="display:flex;align-items:center;gap:12px">
    <div style="width:40px;height:40px;background:#1A1915;
      border-radius:10px;display:flex;align-items:center;justify-content:center;
      font-size:18px;flex-shrink:0;color:#fff">??/div>
    <h1 style="font-size:28px;font-weight:800;color:#1A1915;letter-spacing:-0.6px;margin:0">
      {L['page_title']}
    </h1>
  </div>
</div>

<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:28px">
  <div style="background:#FFFFFF;border:1.5px solid #DDDCDA;border-radius:10px;padding:16px 18px">
    <div style="font-size:18px;margin-bottom:8px">??/div>
    <div style="font-size:13px;font-weight:700;color:#1A1915;margin-bottom:4px">{L['card_humanizer']}</div>
    <div style="font-size:11px;color:#9B9790;line-height:1.5">{L['card_h_desc']}</div>
    <div style="margin-top:10px;font-size:11px;font-weight:600;color:#5C5A55">{L['card_h_cta']} ??/div>
  </div>
  <div style="background:#FFFFFF;border:1.5px solid #DDDCDA;border-radius:10px;padding:16px 18px">
    <div style="font-size:18px;margin-bottom:8px">?뙋</div>
    <div style="font-size:13px;font-weight:700;color:#1A1915;margin-bottom:4px">{L['card_trans']}</div>
    <div style="font-size:11px;color:#9B9790;line-height:1.5">{L['card_t_desc']}</div>
    <div style="margin-top:10px;font-size:11px;font-weight:600;color:#5C5A55">{L['card_t_cta']} ??/div>
  </div>
  <a href="/GPA_Calculator" target="_self" style="text-decoration:none">
    <div style="background:#FFFFFF;border:1.5px solid #DDDCDA;border-radius:10px;padding:16px 18px;
      cursor:pointer;transition:border-color 0.15s;height:100%;box-sizing:border-box"
      onmouseenter="this.style.borderColor='#1A1915'" onmouseleave="this.style.borderColor='#DDDCDA'">
      <div style="font-size:18px;margin-bottom:8px">?럳</div>
      <div style="font-size:13px;font-weight:700;color:#1A1915;margin-bottom:4px">{L['card_gpa']}</div>
      <div style="font-size:11px;color:#9B9790;line-height:1.5">{L['card_g_desc']}</div>
      <div style="margin-top:10px;font-size:11px;font-weight:600;color:#5C5A55">{L['card_g_cta']} ??/div>
    </div>
  </a>
</div>
""", unsafe_allow_html=True)

# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# TABS
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
tab_cover, tab_cv, tab_essay = st.tabs([
    L["tab_cover"], L["tab_cv"], L["tab_essay"],
])


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# TRANSLATION UI
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def render_translation(mode, tab_color, L):
    doc_type = L[f"doc_{mode}"]
    dir_key = f"trans_dir_{mode}"
    if dir_key not in st.session_state:
        st.session_state[dir_key] = "ko_en"

    st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
  <span style="font-size:18px">?뙋</span>
  <span style="font-size:16px;font-weight:700;color:#1A1915">{L['trans_dir_title']}</span>
</div>""", unsafe_allow_html=True)

    direction_label = st.radio(
        "dir", options=[L["trans_ko_en"], L["trans_en_ko"]],
        index=0 if st.session_state[dir_key] == "ko_en" else 1,
        horizontal=True, key=f"dir_radio_{mode}", label_visibility="collapsed",
    )
    direction = "ko_en" if direction_label == L["trans_ko_en"] else "en_ko"
    st.session_state[dir_key] = direction

    src_lang = ("?쒓뎅?? if direction == "ko_en" else "?곸뼱") if L == L_ALL["ko"] else ("Korean" if direction == "ko_en" else "English")
    tgt_lang = ("?곸뼱" if direction == "ko_en" else "?쒓뎅??) if L == L_ALL["ko"] else ("English" if direction == "ko_en" else "Korean")
    src_flag = "?눖?눟" if direction == "ko_en" else "?눣?눡"
    tgt_flag = "?눣?눡" if direction == "ko_en" else "?눖?눟"

    st.markdown(f"""
<div style="display:flex;align-items:center;justify-content:center;gap:16px;
  margin:12px 0 18px;padding:14px;background:#FFFFFF;
  border:2px solid {tab_color}22;border-radius:14px">
  <div style="text-align:center">
    <div style="font-size:26px">{src_flag}</div>
    <div style="font-size:14px;font-weight:700;color:#1A1915;margin-top:2px">{src_lang}</div>
  </div>
  <div style="font-size:28px;color:{tab_color};font-weight:300">??/div>
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
        detected = ("?쒓뎅?? if L == L_ALL["ko"] else "Korean") if is_korean(trans_input) else "English"
        wrong_dir = (direction == "ko_en" and not is_korean(trans_input)) or (direction == "en_ko" and is_korean(trans_input))
        warn = f" ??{L['trans_warn']}" if wrong_dir else ""
        unit = L["char_unit"]
        st.markdown(f'<p style="font-size:12px;color:#B8B6B2;text-align:right;margin-top:4px">{len(trans_input):,}{unit} 쨌 {L["trans_detect"]}: {detected}{warn}</p>', unsafe_allow_html=True)

    if not st.button(f"{L['trans_btn']} {src_flag} ??{tgt_flag}", type="primary", use_container_width=True, key=f"run_trans_{mode}"):
        return

    client = get_client()
    if not client:
        st.error(L["err_no_api"]); return
    if not trans_input.strip():
        st.error(L["err_no_text"]); return

    style_guide = {
        "cover": {
            "ko_en": "Korean cover letter ??English. Sound like it was originally written in English. Use strong action verbs. Professional tone for English-speaking job markets.",
            "en_ko": "?곸뼱 ?먭린?뚭컻?????쒓뎅?? ?쒓뎅 痍⑥뾽 臾명솕??留욌뒗 ?먯뿰?ㅻ윭??寃쎌뼱泥?~?덉뒿?덈떎). 吏곸뿭 湲덉?.",
        },
        "cv": {
            "ko_en": "Korean resume/CV ??English. Keep bullet-point format. Use action verbs (Developed, Managed, etc.). Preserve all numbers exactly.",
            "en_ko": "?곸뼱 ?대젰?????쒓뎅?? 遺덈┸ ?ъ씤???좎?. ?レ옄쨌?섏튂 洹몃?濡? 媛꾧껐???λ룞 ?쒗쁽.",
        },
        "essay": {
            "ko_en": "Korean personal essay ??English. Preserve narrative voice and emotional tone. Feel personal and authentic.",
            "en_ko": "?곸뼱 ?먯꽭?????쒓뎅?? ?쒖궗 ?먮쫫怨?媛먯젙 ???좎?. ?먯뿰?ㅻ윭???쒓뎅???쒗쁽 ?곗꽑.",
        },
    }
    system_prompt = f"""Professional translator for job application documents ({doc_type}).
Direction: {src_lang} ??{tgt_lang}
Style: {style_guide[mode][direction]}
Rules: Output ONLY translated text. Preserve formatting, line breaks, bullet points, all numbers."""

    bar = st.progress(0)
    ph = st.empty()
    ph.markdown(f"""
<div style="background:#FFFFFF;border:1.5px solid #EBEBEA;border-radius:10px;
  padding:12px 16px;display:flex;align-items:center;gap:12px;margin-top:8px">
  <div style="width:28px;height:28px;border-radius:8px;background:{tab_color};
    color:#fff;font-size:13px;display:flex;align-items:center;justify-content:center">?뙋</div>
  <div>
    <div style="font-size:13px;font-weight:600;color:#1A1915">{src_lang} ??{tgt_lang} {L['trans_running']}</div>
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
    color:#fff;font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center">??/div>
  <div>
    <div style="font-size:13px;font-weight:700;color:#065F46">{L['trans_done']}</div>
    <div style="font-size:11px;color:#6EE7B7">{src_lang} ??{tgt_lang} 쨌 {doc_type}</div>
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


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# RENDER TAB
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
def render_tab(mode, tab_color, L, job_description, job_category):
    tog_key = f"show_trans_{mode}"
    if tog_key not in st.session_state:
        st.session_state[tog_key] = False

    is_trans = st.session_state[tog_key]

    sw1, sw2 = st.columns(2, gap="small")
    with sw1:
        if st.button(L["sw_human"], key=f"sw_human_{mode}", use_container_width=True,
                     type="primary" if not is_trans else "secondary"):
            st.session_state[tog_key] = False; st.rerun()
    with sw2:
        if st.button(L["sw_trans"], key=f"sw_trans_{mode}", use_container_width=True,
                     type="primary" if is_trans else "secondary"):
            st.session_state[tog_key] = True; st.rerun()

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    if st.session_state[tog_key]:
        render_translation(mode, tab_color, L)
        return

    # ?? Humanize mode ??????????????????????????????????????????????????????????
    PLACEHOLDERS = {"cover": L["ph_cover"], "cv": L["ph_cv"], "essay": L["ph_essay"]}
    input_text = st.text_area(
        "input", height=240, placeholder=PLACEHOLDERS[mode],
        key=f"input_{mode}", label_visibility="collapsed",
    )

    char_count = len(input_text)
    if char_count > 0:
        lang_detected = L["char_ko"] if is_korean(input_text) else "English"
        st.markdown(f'<p style="font-size:12px;color:#B8B6B2;text-align:right;margin-top:4px">{char_count:,}{L["char_unit"]} 쨌 {lang_detected}</p>', unsafe_allow_html=True)

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
            sub_txt = f"AI {score}% ??{L['step_ai_retry'] if score > 30 else L['step_ai_pass']}"
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
                show_score_step(f"AI {score}% ??{L['step_ai_pass']}", "AI", score=score); break
            show_score_step(f"AI {score}% ??{L['step_ai_retry']}", "AI", score=score)
            final = step5_refine(client, final, input_text)

        bar.progress(100)
        final_ai_score = ai_history[-1]

    except Exception as e:
        bar.empty(); step_ph.empty()
        st.error(f"Error: {e}"); raise

    bar.empty(); step_ph.empty()

    # ?? Result banner ??
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
    banner_sub = f"AI {final_ai_score}% 쨌 {lang} 쨌 횞{len(ai_history)}"

    history_dots = ""
    for i, s in enumerate(ai_history):
        dot_col = "#10B981" if s <= 30 else "#F59E0B" if s <= 60 else "#EF4444"
        arrow = " ??" if i < len(ai_history) - 1 else ""
        history_dots += f'<span style="font-weight:700;color:{dot_col}">{s}%</span><span style="color:#B8B6B2;font-size:11px">{arrow}</span>'

    st.markdown(f"""
<div style="background:{banner_bg};border:1.5px solid {banner_bord};border-radius:12px;
  padding:14px 18px;display:flex;align-items:center;gap:12px;margin:16px 0">
  <div style="width:32px;height:32px;background:{banner_icon_bg};border-radius:8px;
    display:flex;align-items:center;justify-content:center;font-size:16px;
    color:#fff;font-weight:700;flex-shrink:0">{'?? if passed else '!'}</div>
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

    # ?? Metric cards ??
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

    # ?? Before / After ??
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


# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
# TAB RENDER
# ?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧?먥븧
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
    tab_info_banner("?뱷", L["banner_cover"], L["banner_cover_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("cover", m["color"], L, job_description, job_category)

with tab_cv:
    m = TAB_META["cv"]
    tab_info_banner("?뱞", L["banner_cv"], L["banner_cv_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("cv", m["color"], L, job_description, job_category)

with tab_essay:
    m = TAB_META["essay"]
    tab_info_banner("?랃툘", L["banner_essay"], L["banner_essay_d"], m["info_bg"], m["info_text"], m["info_sub"])
    render_tab("essay", m["color"], L, job_description, job_category)
