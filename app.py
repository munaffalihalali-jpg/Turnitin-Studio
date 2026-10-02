"""
========================================================================================
NassGuard Engine — v7.5.0 (Bilingual Open Source - Adaptive Theme & Global Repositories)
المحرك الإحصائي المتكامل لفحص الاستلال وكشف النصوص المولدة بالذكاء الاصطناعي
========================================================================================
التعديلات الجوهرية:
- دمج مستوعبات عالمية (CORE, DOAJ, arXiv, PubMed, SSRN, IASJ).
- دعم كامل للغات والاتجاهات (RTL للعربية / LTR للإنجليزية).
- الوضع المظلم والمضيء التلقائي المتوافق مع الرسوم البيانية.
"""

from __future__ import annotations
import streamlit as st
import hashlib, html, io, json, math, os, re, time
from collections import Counter
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from typing import Optional, List, Tuple, Dict, Any
from urllib.parse import quote, urlparse, parse_qs
import urllib.request
import xml.etree.ElementTree as ET
import requests

try:
    import pypdf
    PYPDF_READY = True
except ImportError:
    PYPDF_READY = False

try:
    import docx
    DOCX_READY = True
except ImportError:
    DOCX_READY = False

try:
    import plotly.graph_objects as go
    PLOTLY_READY = True
except ImportError:
    PLOTLY_READY = False

try:
    from bs4 import BeautifulSoup
    BS4_READY = True
except ImportError:
    BS4_READY = False

# ==========================================
# إعدادات التطبيق الأساسية
# ==========================================
APP_NAME = "NassGuard"
APP_VERSION = "7.5.0"
HEADERS = {"User-Agent": f"{APP_NAME}-IntegrityEngine/7.5 (contact@opensource.org)"}

st.set_page_config(
    page_title=f"{APP_NAME} | Academic Integrity",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# القاموس اللغوي (Translation Dictionary)
# ==========================================
TRANSLATIONS = {
    "ar": {
        "app_title": "منظومة التدقيق الأكاديمي المسبق ونواة الكشف الإحصائي",
        "engine_info": "⚙ المحرك: تحليل إحصائي ثنائي اللغة NLP وربط عالمي",
        "status_info": "🟢 الحالة: متصل بالمستوعبات الدولية (Core, DOAJ, arXiv...)",
        "sidebar_draft": "📄 بيانات المسودة",
        "author_name": "اسم الباحث:",
        "paper_title": "عنوان الورقة العلمية:",
        "upload_section": "📁 إيداع ملف البحث",
        "upload_label": "رفع المستند (DOCX, PDF, TXT):",
        "upload_success": "تم إيداع الملف بنجاح ({words} كلمة).",
        "algo_settings": "⚙ معايير الخوارزمية",
        "word_threshold": "عتبة السلسلة المتطابقة (الكلمات):",
        "mitigation_toggle": "تفعيل موازنة الأبحاث العلمية الصارمة",
        "exclude_bib": "استبعاد قائمة المراجع",
        "tab_canvas": "📑 مستعرض البحث والمطابقات",
        "tab_ai": "🟣 كاشف الذكاء الاصطناعي (NLP)",
        "tab_analytics": "📊 مصفوفة التحليل",
        "tab_humanizer": "✍️ المحرر الأكاديمي",
        "text_input_label": "متن البحث المفحوص:",
        "text_input_placeholder": "الصق نص البحث الأكاديمي (عربي أو إنجليزي) هنا...",
        "total_words": "إجمالي الكلمات:",
        "digital_footprint": "البصمة الرقمية:",
        "engine_caption": "يعتمد الكاشف على مصفوفة قياس التباين الإحصائي وتلازم الألفاظ ثنائية اللغة.",
        "run_btn": "🚀 تشغيل الفحص المزدوج الشامل (استلال + ذكاء اصطناعي)",
        "short_text_warn": "النص قصير جداً؛ يُشترط وجود 15 كلمة على الأقل لبدء الفحص.",
        "status_running": "جاري تنفيذ التحليل الإحصائي واسترجاع البيانات من المستوعبات العالمية والمحلية...",
        "step_1": "1. استخراج المتن وفصل المراجع...",
        "step_2": "2. تشغيل كاشف الذكاء الاصطناعي NLP...",
        "step_3": "3. استعلام المستودعات الدولية (Crossref, CORE, DOAJ, arXiv) والمحلية (IASJ)...",
        "status_complete": "اكتمل الفحص الشامل بنجاح!",
        "excluded_bib_label": "المراجع المستبعدة:",
        "audit_results": "نتائج التدقيق الرقمي",
        "sim_index": "مؤشر الاستلال اللفظي",
        "sim_desc": "{matched} كلمة مطابقة من أصل {total}",
        "ai_index": "مؤشر الاشتباه التوليدي (AI)",
        "ai_desc": "{flagged} مقطع مشبوه من أصل {total}",
        "retrieved_sources": "المصادر المسترجعة:",
        "ai_tab_title": "التحليل الإحصائي المعمق لمؤشرات التوليد الآلي",
        "ai_tab_caption": "تحليل احتمالي يستند إلى معامل الاختلاف التدفقي (Burstiness) وكثافة القوالب التوليدية.",
        "ai_metric_1": "مؤشر الذكاء الاصطناعي الكلي",
        "ai_metric_2": "مؤشر التوقع والتلازم",
        "ai_metric_3": "معامل التباين التدفقي (CV)",
        "delta_ai_normal": "طبيعي",
        "delta_ai_high": "مرتفع",
        "delta_pred_normal": "توقع طبيعي",
        "delta_pred_high": "توقع نمطي",
        "delta_cv_human": "تدفق بشري",
        "delta_cv_ai": "رتابة آلية",
        "no_ai_chunks": "لم يرصد المحرك أي مقاطع ذات نمطية آلية متطابقة.",
        "suspicious_chunks_title": "المقاطع المرصودة ذات الاشتباه الإحصائي المرتفع:",
        "chunk_id": "مقطع مشتبه به #{id}",
        "ai_prob": "احتمال توليدي:",
        "burstiness": "التباين التدفقي:",
        "cliche_hits": "القوالب الآلية المرصودة:",
        "perplexity": "درجة التوقع:",
        "req_run_analytics": "يرجى تشغيل الفحص أولاً لتوليد الرسوم البيانية.",
        "pie_authentic": "محتوى أصيل",
        "pie_plagiarized": "استلال منسوخ",
        "pie_ai": "نص مشبوه آلياً",
        "doc_stats": "إحصائيات الوثيقة الموزونة:",
        "total_sentences": "إجمالي الجمل:",
        "suspicious_segments": "المقاطع المشبوهة:",
        "mean_cv": "معامل التباين العام (CV):",
        "hum_title": "إعادة البناء الأسلوبي وأنسنة النصوص محلياً",
        "hum_caption": "أداة فورية تكسر الرتابة، ترفع التباين التدفقي، وتستبدل القوالب الآلية.",
        "hum_input": "النص المراد مراجعته وتعديله:",
        "hum_btn": "✨ تشغيل الأنسنة ورفع التباين اللغوي الآن",
        "hum_warn": "الرجاء إدخال نص أولاً.",
        "hum_success": "تمت إعادة الهيكلة بنجاح مع رفع تباين الجمل وكسر الرتابة!",
        "hum_output": "النص بعد الأنسنة الأكاديمية:",
        "hum_apply": "📥 اعتماد النص المحسن وفحصه فوراً",
        "dir": "rtl",
        "align": "right"
    },
    "en": {
        "app_title": "Advanced Academic Integrity & Statistical NLP Engine",
        "engine_info": "⚙ Engine: Bilingual Quad-Metric NLP & Global Connect",
        "status_info": "🟢 Status: Connected to Global Repositories (Core, DOAJ...)",
        "sidebar_draft": "📄 Draft Information",
        "author_name": "Author Name:",
        "paper_title": "Paper Title:",
        "upload_section": "📁 Upload Document",
        "upload_label": "Upload File (DOCX, PDF, TXT):",
        "upload_success": "File uploaded successfully ({words} words).",
        "algo_settings": "⚙ Algorithm Settings",
        "word_threshold": "Match String Threshold (Words):",
        "mitigation_toggle": "Enable Strict Academic Mitigation",
        "exclude_bib": "Exclude Bibliography",
        "tab_canvas": "📑 Document & Matches",
        "tab_ai": "🟣 AI Detector (NLP)",
        "tab_analytics": "📊 Analytics & Charts",
        "tab_humanizer": "✍️ Academic Humanizer",
        "text_input_label": "Document Text:",
        "text_input_placeholder": "Paste your academic text here (English or Arabic)...",
        "total_words": "Total Words:",
        "digital_footprint": "Digital Footprint:",
        "engine_caption": "Engine uses statistical burstiness and bilingual collocation matrices.",
        "run_btn": "🚀 Run Comprehensive Scan (Plagiarism + AI)",
        "short_text_warn": "Text is too short; at least 15 words are required.",
        "status_running": "Running statistical analysis and retrieving data from global repositories...",
        "step_1": "1. Extracting text and parsing bibliography...",
        "step_2": "2. Running AI detection engine...",
        "step_3": "3. Querying global (CORE, DOAJ, arXiv) & regional repositories...",
        "status_complete": "Comprehensive scan completed successfully!",
        "excluded_bib_label": "Excluded Bibliography:",
        "audit_results": "Digital Audit Results",
        "sim_index": "Similarity Index",
        "sim_desc": "{matched} matched words out of {total}",
        "ai_index": "AI Generation Index",
        "ai_desc": "{flagged} suspicious chunks out of {total}",
        "retrieved_sources": "Retrieved Sources:",
        "ai_tab_title": "In-Depth Statistical Analysis of AI Signatures",
        "ai_tab_caption": "Probabilistic analysis based on burstiness CV and generative cliche density.",
        "ai_metric_1": "Global AI Index",
        "ai_metric_2": "Predictability Index",
        "ai_metric_3": "Burstiness Coefficient (CV)",
        "delta_ai_normal": "Normal",
        "delta_ai_high": "High",
        "delta_pred_normal": "Natural",
        "delta_pred_high": "Stereotypical",
        "delta_cv_human": "Human Flow",
        "delta_cv_ai": "Machine Monotony",
        "no_ai_chunks": "No identical AI patterns detected in the text.",
        "suspicious_chunks_title": "Flagged Chunks with High Statistical Suspicion:",
        "chunk_id": "Suspicious Chunk #{id}",
        "ai_prob": "AI Probability:",
        "burstiness": "Burstiness:",
        "cliche_hits": "Detected Cliches:",
        "perplexity": "Predictability Score:",
        "req_run_analytics": "Please run the scan first to generate analytics.",
        "pie_authentic": "Authentic Content",
        "pie_plagiarized": "Plagiarized Text",
        "pie_ai": "AI-Generated Text",
        "doc_stats": "Weighted Document Statistics:",
        "total_sentences": "Total Sentences:",
        "suspicious_segments": "Suspicious Segments:",
        "mean_cv": "Mean Burstiness (CV):",
        "hum_title": "Stylistic Reconstruction & Local Humanization",
        "hum_caption": "An instant tool to break monotony, increase burstiness, and replace AI cliches.",
        "hum_input": "Text to review and modify:",
        "hum_btn": "✨ Run Humanization & Increase Variance",
        "hum_warn": "Please enter some text first.",
        "hum_success": "Restructuring completed successfully! Variance increased.",
        "hum_output": "Text after academic humanization:",
        "hum_apply": "📥 Apply modified text and scan",
        "dir": "ltr",
        "align": "left"
    }
}

def t(key, **kwargs):
    text = TRANSLATIONS[st.session_state.lang][key]
    if kwargs:
        return text.format(**kwargs)
    return text

# ==========================================
# تهيئة لغة الواجهة والتهيئة المتغيرة
# ==========================================
if "lang" not in st.session_state:
    st.session_state.lang = "ar"
if "doc_text" not in st.session_state:
    st.session_state.doc_text = ""
if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False
if "sim_data" not in st.session_state:
    st.session_state.sim_data = None
if "ai_report" not in st.session_state:
    st.session_state.ai_report = None

lang_selection = st.sidebar.radio("🌐 Language / اللغة", ["العربية", "English"], index=0 if st.session_state.lang == "ar" else 1)
new_lang = "ar" if lang_selection == "العربية" else "en"
if new_lang != st.session_state.lang:
    st.session_state.lang = new_lang
    st.rerun()

DIR = t("dir")
ALIGN = t("align")

HIGHLIGHT_PALETTE = [
    {"bg": "#FFE4E6", "border": "#F43F5E", "text": "#881337", "badge": "#E11D48"},
    {"bg": "#DBEAFE", "border": "#3B82F6", "text": "#1E3A8A", "badge": "#2563EB"},
    {"bg": "#FEF3C7", "border": "#F59E0B", "text": "#78350F", "badge": "#D97706"},
    {"bg": "#D1FAE5", "border": "#10B981", "text": "#064E3B", "badge": "#059669"},
    {"bg": "#EDE9FE", "border": "#8B5CF6", "text": "#4C1D95", "badge": "#7C3AED"},
    {"bg": "#FFEDD5", "border": "#F97316", "text": "#7C2D12", "badge": "#EA580C"},
]

STOP_WORDS = {
    "من", "في", "على", "الى", "إلى", "عن", "مع", "هذا", "هذه", "ذلك", "تلك", "هو", "هي",
    "هم", "هن", "كما", "وقد", "وهو", "وهي", "ان", "إن", "أن", "ما", "لا", "لم", "لن", "ثم",
    "او", "أو", "أي", "كل", "بين", "بعد", "قبل", "ضمن", "عند", "كان", "كانت", "يكون",
    "تكون", "تم", "قد", "بها", "به", "له", "لها", "التي", "الذي", "الذين", "حيث", "و", "ف", "ب", "ك", "ل",
    "the", "of", "in", "on", "and", "or", "to", "for", "is", "are", "was", "were"
}

# ==========================================
# CSS الديناميكي المتكيف
# ==========================================
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Tajawal:wght@400;500;700;800&display=swap');
    
    :root {{
        --app-bg: #F4F7FA;
        --card-bg: #FFFFFF;
        --text-main: #1E293B;
        --text-muted: #64748B;
        --border-color: #E2E8F0;
        --card-hover: #F8FAFC;
        --panel-shadow: rgba(0,0,0,0.04);
        --mark-text-ai: #5B21B6;
        --mark-bg-ai: #F5F3FF;
    }}

    @media (prefers-color-scheme: dark) {{
        :root {{
            --app-bg: #0E1117;
            --card-bg: #1E1E2E;
            --text-main: #E2E8F0;
            --text-muted: #94A3B8;
            --border-color: #334155;
            --card-hover: #27273A;
            --panel-shadow: rgba(0,0,0,0.25);
            --mark-text-ai: #E9D5FF;
            --mark-bg-ai: #4C1D95;
        }}
    }}

    html, body, [class*="css"] {{
        font-family: 'Tajawal', 'Cairo', -apple-system, BlinkMacSystemFont, sans-serif !important;
        direction: {DIR} !important;
        text-align: {ALIGN} !important;
    }}

    .stApp {{
        background-color: var(--app-bg) !important;
    }}

    section[data-testid="stSidebar"] {{
        direction: {DIR} !important;
        text-align: {ALIGN} !important;
        background: var(--card-bg) !important;
        border-left: { '1px solid var(--border-color)' if DIR == 'rtl' else 'none' } !important;
        border-right: { '1px solid var(--border-color)' if DIR == 'ltr' else 'none' } !important;
        box-shadow: { '-5px 0 25px' if DIR == 'rtl' else '5px 0 25px' } var(--panel-shadow);
    }}

    .app-bar {{
        background: linear-gradient(135deg, #312E81 0%, #4F46E5 100%);
        color: #FFFFFF;
        padding: 24px 35px;
        border-radius: 16px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 25px;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.1);
        flex-direction: { 'row' if DIR == 'ltr' else 'row-reverse' };
    }}

    .paper-canvas {{
        background-color: var(--card-bg) !important;
        color: var(--text-main) !important;
        border: 1px solid var(--border-color);
        box-shadow: 0 12px 35px -5px var(--panel-shadow);
        border-radius: 16px;
        padding: 55px 65px;
        line-height: 2.1;
        font-size: 17px;
        font-family: 'Cairo', serif;
        direction: {DIR} !important;
        text-align: justify !important;
        min-height: 750px;
    }}

    .mark-plag {{
        display: inline;
        padding: 3px 6px;
        border-radius: 6px;
        font-weight: 700;
        transition: all 0.2s ease;
        color: #111827 !important; 
    }}

    .mark-ai {{
        display: inline;
        background-color: var(--mark-bg-ai) !important;
        color: var(--mark-text-ai) !important;
        border-bottom: 3px solid #8B5CF6 !important;
        padding: 3px 7px;
        border-radius: 6px;
        font-weight: 700;
    }}

    .pill-badge {{
        display: inline-block;
        color: #FFFFFF !important;
        font-size: 11px;
        font-weight: 800;
        min-width: 20px;
        height: 20px;
        line-height: 20px;
        text-align: center;
        border-radius: 5px;
        margin: 0 4px;
        vertical-align: middle;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}

    .panel-card {{
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 8px 25px var(--panel-shadow);
        color: var(--text-main);
    }}

    .source-card {{
        background: var(--card-hover);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 14px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        direction: {DIR};
    }}
    
    .source-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 15px var(--panel-shadow);
        border-color: #8B5CF6;
    }}

    .metric-num {{
        font-size: 52px;
        font-weight: 900;
        line-height: 1;
        font-family: 'Cairo', sans-serif;
        text-shadow: 0 4px 10px var(--panel-shadow);
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
        direction: {DIR};
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        background-color: transparent;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: var(--card-bg);
        border-bottom: 3px solid #4F46E5 !important;
        font-weight: 700;
        color: var(--text-main) !important;
    }}

    .suspicious-box {{
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        border-right: { '5px solid #8B5CF6' if DIR == 'rtl' else 'none' };
        border-left: { '5px solid #8B5CF6' if DIR == 'ltr' else 'none' };
        padding: 18px;
        border-radius: 10px;
        margin-bottom: 15px;
        box-shadow: 0 4px 10px var(--panel-shadow);
        text-align: {ALIGN};
    }}
</style>
""", unsafe_allow_html=True)

@dataclass
class Source:
    source_id: str
    title: str
    url: str
    source_type: str
    text: str = ""
    scope: str = "metadata"
    words: int = 0
    fetch_error: str = ""
    status: int = 0
    doi: Optional[str] = None

@dataclass
class ChunkReport:
    chunk_id: int
    start_char: int
    end_char: int
    text: str
    sentence_count: int
    word_count: int
    perplexity_proxy: float
    burstiness: float
    ai_probability: float
    is_suspicious: bool
    diagnostics: Dict[str, Any]

@dataclass
class DocumentAnalysisReport:
    document_length_words: int
    document_length_sentences: int
    total_chunks: int
    flagged_chunks_count: int
    global_ai_percentage: float
    mean_perplexity: float
    mean_burstiness: float
    suspicious_segments: List[ChunkReport]

# =======================================================================
# المحرك الإحصائي
# =======================================================================
class AdvancedAIDetector:
    def __init__(self, academic_mode: bool = True):
        self.academic_mode = academic_mode
        self.ai_markers = [
            r"علاوة على ذلك", r"في هذا الصدد", r"تجدر الإشارة إلى", r"جدير بالذكر",
            r"مما لا شك فيه", r"يلعب[\s\w]*دوراً[\s\w]*محورياً", r"تلعب[\s\w]*دوراً[\s\w]*محورياً",
            r"وفي الختام", r"بشكل متسارع", r"مما يساهم في", r"على صعيد آخر",
            r"في إطار", r"يعد من أهم", r"تعد من أهم", r"من الأهمية بمكان",
            r"بدقة متناهية", r"بسرعة فائقة", r"استراتيجيات إدارة المخاطر",
            
            r"in today's rapidly", r"rapidly evolving", r"furthermore", r"moreover",
            r"in conclusion", r"it is important to note", r"plays a crucial role",
            r"transformative force", r"ultimately", r"delve into", r"a testament to",
            r"comprehensive understanding", r"navigating the complexities", r"shed light on"
        ]

    def segment_sentences_with_spans(self, text: str) -> List[Tuple[str, int, int]]:
        pattern = re.compile(r"([^.!?؟;\n]+[.!?؟;\n]+|[^.!?؟;\n]+$)")
        results = []
        for match in pattern.finditer(text):
            sent = match.group().strip()
            if len(sent.split()) >= 3:
                results.append((sent, match.start(), match.end()))
        return results

    def compute_burstiness(self, sentences_list: List[str]) -> Tuple[float, float, float]:
        lengths = [len(s.split()) for s in sentences_list if s.strip()]
        if len(lengths) < 2:
            return 0.20, float(lengths[0]) if lengths else 0.0, 0.0
        mean = sum(lengths) / len(lengths)
        variance = sum((l - mean) ** 2 for l in lengths) / len(lengths)
        std_dev = math.sqrt(variance)
        cv = std_dev / (mean + 1e-5)
        return cv, mean, std_dev

    def compute_cliche_density(self, text: str, sentences_list: List[str]) -> Tuple[float, int]:
        if not sentences_list:
            return 0.0, 0
        hits = 0
        for sent in sentences_list:
            if any(re.search(marker, sent, re.IGNORECASE) for marker in self.ai_markers):
                hits += 1
        density = hits / len(sentences_list)
        return density, hits

    def compute_predictability_proxy(self, text: str) -> float:
        tokens = [w for w in text.split() if w not in STOP_WORDS]
        if not tokens:
            return 50.0
        ttr = len(set(tokens)) / len(tokens)
        return round((1.0 - ttr) * 100.0, 2)

    def evaluate_text_prob(self, text: str, sents: List[str], global_cv: float) -> Tuple[float, Dict[str, Any]]:
        burst_risk = 1.0 / (1.0 + math.exp((global_cv - 0.32) / 0.05))
        cliche_density, hit_count = self.compute_cliche_density(text, sents)
        cliche_risk = min(1.0, cliche_density * 1.5)
        pred = self.compute_predictability_proxy(text)
        pred_risk = min(1.0, max(0.1, pred / 50.0))

        combined_prob = (0.45 * burst_risk) + (0.40 * cliche_risk) + (0.15 * pred_risk)

        discount = 0.0
        if self.academic_mode:
            academic_markers = [
                r"استناداً إلى", r"يتضح من خلال", r"في ضوء ما سبق", r"وفقاً للمنهجية",
                r"based on the", r"results indicate", r"empirical evidence", r"this study", r"in contrast to"
            ]
            for m in academic_markers:
                if re.search(m, text, re.IGNORECASE):
                    discount += 0.04
            if global_cv > 0.45:
                discount += 0.10

        final_prob = max(0.0, min(1.0, combined_prob - discount))

        return final_prob, {
            "burst_risk": round(burst_risk, 3),
            "cliche_risk": round(cliche_risk, 3),
            "cliche_hits": hit_count,
            "pred_proxy": pred
        }

    def analyze_document(self, text: str) -> DocumentAnalysisReport:
        all_sents_spans = self.segment_sentences_with_spans(text)
        if not all_sents_spans:
            return DocumentAnalysisReport(0, 0, 0, 0, 0.0, 0.0, 0.0, [])

        raw_all_sents = [s[0] for s in all_sents_spans]
        global_cv, global_mean, global_std = self.compute_burstiness(raw_all_sents)

        chunks = []
        n = len(all_sents_spans)
        i = 0
        while i < n:
            chunk_sents = []
            wc = 0
            j = i
            while j < n and wc < 80:
                chunk_sents.append(all_sents_spans[j])
                wc += len(all_sents_spans[j][0].split())
                j += 1
            if chunk_sents:
                c_text = " ".join(s[0] for s in chunk_sents)
                chunks.append((c_text, chunk_sents[0][1], chunk_sents[-1][2], [s[0] for s in chunk_sents]))
            step = max(1, (j - i) - 1)
            i += step

        reports = []
        for idx, (c_text, start, end, sents) in enumerate(chunks, 1):
            prob, diag = self.evaluate_text_prob(c_text, sents, global_cv)
            is_susp = prob >= 0.50
            reports.append(ChunkReport(
                chunk_id=idx, start_char=start, end_char=end, text=c_text,
                sentence_count=len(sents), word_count=len(c_text.split()),
                perplexity_proxy=diag["pred_proxy"], burstiness=round(global_cv, 4),
                ai_probability=round(prob, 4), is_suspicious=is_susp, diagnostics=diag
            ))

        suspicious = [c for c in reports if c.is_suspicious]

        if suspicious:
            spans = [(c.start_char, c.end_char) for c in suspicious]
            merged = self._merge_intervals(spans)
            flagged_chars = sum(b - a for a, b in merged)
            total_chars = max(len(text.strip()), 1)
            char_pct = (flagged_chars / total_chars) * 100.0
            avg_susp_prob = sum(c.ai_probability for c in suspicious) / len(suspicious)
            global_ai_pct = round(char_pct * avg_susp_prob, 1)
        else:
            global_ai_pct = 0.0

        mean_ppl = sum(c.perplexity_proxy for c in reports) / len(reports)

        return DocumentAnalysisReport(
            document_length_words=len(text.split()), document_length_sentences=len(all_sents_spans),
            total_chunks=len(reports), flagged_chunks_count=len(suspicious),
            global_ai_percentage=min(100.0, global_ai_pct), mean_perplexity=round(mean_ppl, 2),
            mean_burstiness=round(global_cv, 4), suspicious_segments=suspicious
        )

    @staticmethod
    def _merge_intervals(intervals: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
        if not intervals: return []
        sorted_intervals = sorted(intervals, key=lambda x: x[0])
        merged = [sorted_intervals[0]]
        for cur in sorted_intervals[1:]:
            prev_s, prev_e = merged[-1]
            if cur[0] <= prev_e:
                merged[-1] = (prev_s, max(prev_e, cur[1]))
            else:
                merged.append(cur)
        return merged

def sha256_text(x: str | bytes) -> str:
    b = x.encode("utf-8", "ignore") if isinstance(x, str) else x
    return hashlib.sha256(b).hexdigest()

def norm(t_str: str) -> str:
    t_str = t_str or ""
    t_str = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", t_str)
    t_str = re.sub(r"[إأآٱ]", "ا", t_str).replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    t_str = re.sub(r"[.،,:;؛!?؟()\[\]{}<>«»“”\"']+", " ", t_str)
    return re.sub(r"\s+", " ", t_str).strip().lower()

def get_tokens(t_str: str) -> list[str]:
    return re.findall(r"[\w\u0600-\u06FF]+", norm(t_str), re.UNICODE)

def wc(t_str: str) -> int:
    return len(get_tokens(t_str))

def sentences(t_str: str) -> list[str]:
    return [x.strip() for x in re.split(r"(?<=[.!?؟؛])\s+|\n+", t_str or "") if wc(x) >= 4]

def extract_file(name: str, data: bytes) -> tuple[str, list[str]]:
    n = name.lower()
    if n.endswith(".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
            try: return data.decode(enc), []
            except UnicodeDecodeError: pass
        return data.decode("utf-8", "replace"), []
    if n.endswith(".pdf") and PYPDF_READY:
        try:
            r = pypdf.PdfReader(io.BytesIO(data))
            out = [p.extract_text() for p in r.pages if p.extract_text()]
            return "\n\n".join(out), []
        except Exception as e: return "", [f"PDF Error: {e}"]
    if n.endswith(".docx") and DOCX_READY:
        try:
            d = docx.Document(io.BytesIO(data))
            out = [p.text for p in d.paragraphs if p.text.strip()]
            return "\n".join(out), []
        except Exception as e: return "", [f"DOCX Error: {e}"]
    return "", ["Unsupported format"]

def split_bib(t_str: str) -> tuple[str, str]:
    m = re.search(r"(?im)(?:^|\n)[ \t]*(?:references|bibliography|المراجع|المصادر|المصادر والمراجع)[ \t]*[:\-]?[ \t]*(?:\n|$)", t_str or "")
    return (t_str[:m.start()].rstrip(), t_str[m.start():].strip()) if m else (t_str, "")

# =======================================================================
# دوال استرجاع البيانات المضافة حديثاً (Global & Regional Repositories)
# =======================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def query_academic_sources(q: str) -> list[dict]:
    results = []
    try:
        r = requests.get(f"https://api.crossref.org/works?query.bibliographic={quote(q)}&rows=2", headers=HEADERS, timeout=6)
        if r.status_code == 200:
            for item in r.json().get("message", {}).get("items", []):
                t_title = (item.get("title") or [""])[0]
                ab = item.get("abstract") or ""
                clean_ab = re.sub(r"<[^>]+>", " ", ab)
                if t_title:
                    results.append({"title": f"Crossref: {t_title[:60]}", "url": item.get("URL") or "", "text": f"{t_title}. {clean_ab}", "type": "Publications"})
    except: pass

    try:
        r_w = requests.get(f"https://ar.wikipedia.org/w/api.php?action=query&list=search&srsearch={quote(q)}&format=json&utf8=1&srlimit=2", headers=HEADERS, timeout=6)
        if r_w.status_code == 200:
            for item in r_w.json().get("query", {}).get("search", []):
                t_title = item.get("title") or ""
                if t_title:
                    r_txt = requests.get(f"https://ar.wikipedia.org/w/api.php?action=parse&page={quote(t_title)}&prop=text&format=json&formatversion=2", headers=HEADERS, timeout=6)
                    body_txt = ""
                    if r_txt.status_code == 200:
                        raw_h = (r_txt.json().get("parse") or {}).get("text") or ""
                        body_txt = re.sub(r"(?is)<[^>]+>", " ", raw_h)
                    results.append({"title": f"{t_title} — Wikipedia", "url": f"https://ar.wikipedia.org/wiki/{quote(t_title.replace(' ', '_'))}", "text": body_txt, "type": "Internet Sources"})
    except: pass
    return results

def query_global_repositories(q: str) -> list[dict]:
    results = []
    encoded_q = quote(q)

    # 1. arXiv API
    try:
        url = f"http://export.arxiv.org/api/query?search_query=all:{encoded_q}&max_results=2"
        response = urllib.request.urlopen(url, timeout=6).read()
        root = ET.fromstring(response)
        ns = {'atom': 'http://www.w3.org/2005/Atom'}
        for entry in root.findall('atom:entry', ns):
            title = entry.find('atom:title', ns).text.replace('\n', ' ')
            summary = entry.find('atom:summary', ns).text.replace('\n', ' ')
            link = entry.find('atom:id', ns).text
            results.append({"title": f"arXiv: {title[:60]}", "url": link, "text": f"{title}. {summary}", "type": "Preprints"})
    except Exception: pass

    # 2. DOAJ API
    try:
        r_doaj = requests.get(f"https://doaj.org/api/v1/search/articles/{encoded_q}?pageSize=2", headers=HEADERS, timeout=6)
        if r_doaj.status_code == 200:
            for item in r_doaj.json().get("results", []):
                bibjson = item.get("bibjson", {})
                title = bibjson.get("title", "")
                abstract = bibjson.get("abstract", "")
                url = bibjson.get("link", [{"url": ""}])[0].get("url", "")
                if title:
                    results.append({"title": f"DOAJ: {title[:60]}", "url": url, "text": f"{title}. {abstract}", "type": "Open Access Journals"})
    except Exception: pass

    # 3. PubMed Central (PMC)
    try:
        pmc_search = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pmc&term={encoded_q}&retmax=2&retmode=json"
        r_pmc = requests.get(pmc_search, headers=HEADERS, timeout=6)
        if r_pmc.status_code == 200:
            ids = r_pmc.json().get("esearchresult", {}).get("idlist", [])
            if ids:
                summary_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pmc&id={','.join(ids)}&retmode=json"
                r_sum = requests.get(summary_url, headers=HEADERS, timeout=6)
                if r_sum.status_code == 200:
                    data = r_sum.json().get("result", {})
                    for uid in ids:
                        article = data.get(uid, {})
                        title = article.get("title", "")
                        if title:
                            results.append({"title": f"PubMed: {title[:60]}", "url": f"https://www.ncbi.nlm.nih.gov/pmc/articles/{uid}/", "text": title, "type": "Medical/Bio"})
    except Exception: pass

    # 4. CORE API
    try:
        r_core = requests.get(f"https://api.core.ac.uk/v3/search/works/?q={encoded_q}&limit=2", headers=HEADERS, timeout=6)
        if r_core.status_code == 200:
            for item in r_core.json().get("results", []):
                title = item.get("title", "")
                abstract = item.get("abstract", "") or ""
                url = item.get("downloadUrl") or ""
                if title:
                    results.append({"title": f"CORE: {title[:60]}", "url": url, "text": f"{title}. {abstract}", "type": "Global Research"})
    except Exception: pass

    return results

def query_regional_arabic_repositories(q: str) -> list[dict]:
    results = []
    if not BS4_READY:
        return results # إذا لم تكن المكتبة مثبتة يتم التجاوز بأمان
    
    # 1. IASJ (المجلات الأكاديمية العلمية العراقية)
    try:
        iasj_url = f"https://www.iasj.net/iasj/search?query={quote(q)}"
        r_iasj = requests.get(iasj_url, headers=HEADERS, timeout=7)
        if r_iasj.status_code == 200:
            soup = BeautifulSoup(r_iasj.text, 'lxml')
            for item in soup.select('.result-item')[:2]:
                title_tag = item.select_one('h4 a')
                abstract_tag = item.select_one('.abstract')
                if title_tag:
                    title = title_tag.text.strip()
                    url = "https://www.iasj.net" + title_tag['href']
                    abstract = abstract_tag.text.strip() if abstract_tag else ""
                    results.append({"title": f"IASJ: {title[:60]}", "url": url, "text": f"{title}. {abstract}", "type": "Arabic Journals"})
    except Exception: pass

    # 2. SSRN
    try:
        ssrn_url = f"https://papers.ssrn.com/sol3/results.cfm?txtKey_Words={quote(q)}"
        r_ssrn = requests.get(ssrn_url, headers=HEADERS, timeout=7)
        if r_ssrn.status_code == 200:
            soup = BeautifulSoup(r_ssrn.text, 'lxml')
            for item in soup.select('.title.optClickTitle')[:2]:
                title = item.text.strip()
                url = item.get('href', '')
                if title:
                    results.append({"title": f"SSRN: {title[:60]}", "url": url, "text": title, "type": "Preprints/Social"})
    except Exception: pass

    return results

def find_plagiarism_matches(body: str, sources: list[Source], min_words: int, exclude_quotes: bool) -> tuple[list[dict], set[int]]:
    tokens = get_tokens(body)
    total_tokens = len(tokens)
    matched_indices = set()
    matches = []
    usable = [s for s in sources if s.words >= 20]

    src_ngrams = {}
    for s in usable:
        s_toks = get_tokens(s.text)
        ng = {" ".join(s_toks[i:i + min_words]) for i in range(len(s_toks) - min_words + 1)}
        src_ngrams[s.source_id] = (ng, s_toks)

    i = 0
    while i <= total_tokens - min_words:
        gram = " ".join(tokens[i:i + min_words])
        found_source = None
        for sid, (ng_set, _) in src_ngrams.items():
            if gram in ng_set:
                found_source = sid
                break

        if found_source:
            j = i + min_words
            while j < total_tokens:
                ext = " ".join(tokens[i:j + 1])
                full_s_text = " ".join(src_ngrams[found_source][1])
                if ext in full_s_text: j += 1
                else: break
            matched_indices.update(range(i, j))
            matches.append({
                "source_id": found_source, "start": i, "end": j,
                "length": j - i, "phrase": " ".join(tokens[i:j])
            })
            i = j
        else:
            i += 1
    return matches, matched_indices

def run_local_humanizer(text: str) -> str:
    replacements = {
        "علاوة على ذلك": "وفي سياق متصل", "في هذا الصدد": "وعليه",
        "تجدر الإشارة إلى": "ومن الجدير بالملاحظة أن", "جدير بالذكر": "ومما يستوجب الوقوف عنده أن",
        "مما لا شك فيه": "ومن الثابت تحليلياً", "يلعب دورا محوريا": "يمثل ركيزة جوهرية في",
        "تعد من أهم": "تُصنف ضمن أبرز",
        "in conclusion": "to summarize", "furthermore": "additionally",
        "it is important to note": "notably", "plays a crucial role": "is essential to"
    }
    cleaned = text
    for k, v in replacements.items():
        cleaned = cleaned.replace(k, v)

    sents = sentences(cleaned)
    reworked = []
    for s in sents:
        words = s.split()
        if len(words) > 22:
            mid = len(words) // 2
            if st.session_state.lang == "ar":
                reworked.append(" ".join(words[:mid]) + "؛")
                reworked.append("الأمر الذي يؤكد أن " + " ".join(words[mid:]))
            else:
                reworked.append(" ".join(words[:mid]) + ";")
                reworked.append("which highlights that " + " ".join(words[mid:]))
        else:
            reworked.append(s)
    return " ".join(reworked)

# =======================================================================
# واجهة المستخدم الرئيسية (Main UI)
# =======================================================================
def main():
    st.markdown(f"""
    <div class="app-bar">
        <div style="text-align: {ALIGN};">
            <div style="font-size: 28px; font-weight: 800; letter-spacing: 0.5px; font-family: sans-serif;">{APP_NAME}</div>
            <div style="font-size: 14px; opacity: 0.9; margin-top: 5px; font-weight: 500;">{t('app_title')} — {APP_VERSION}</div>
        </div>
        <div style="text-align: left; font-size: 12px; direction: ltr; font-family: monospace; background: rgba(0,0,0,0.15); padding: 8px 12px; border-radius: 8px;">
            <div style="margin-bottom: 3px;">{t('engine_info')}</div>
            <div>{t('status_info')}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(f"<h3 style='color:var(--text-main); font-weight: 800;'>{t('sidebar_draft')}</h3>", unsafe_allow_html=True)
        st.text_input(t('author_name'), "")
        st.text_input(t('paper_title'), "")

        st.divider()
        st.markdown(f"<h4 style='color:var(--text-main); font-weight: 700;'>{t('upload_section')}</h4>", unsafe_allow_html=True)
        up_file = st.file_uploader(t('upload_label'), type=["docx", "pdf", "txt"])
        if up_file:
            data = up_file.read()
            extracted, warns = extract_file(up_file.name, data)
            if extracted and extracted != st.session_state.doc_text:
                st.session_state.doc_text = extracted
                st.session_state.analysis_done = False
                st.sidebar.success(t('upload_success', words=wc(extracted)))
                st.rerun()

        st.divider()
        st.markdown(f"<h4 style='color:var(--text-main); font-weight: 700;'>{t('algo_settings')}</h4>", unsafe_allow_html=True)
        min_words = st.slider(t('word_threshold'), 4, 10, 6)
        academic_mitigation_toggle = st.checkbox(t('mitigation_toggle'), value=True)
        exclude_bib = st.checkbox(t('exclude_bib'), value=True)

    tab_studio, tab_ai, tab_analytics, tab_humanizer = st.tabs([
        t('tab_canvas'), t('tab_ai'), t('tab_analytics'), t('tab_humanizer')
    ])

    with tab_studio:
        c1, c2 = st.columns([3, 1])
        with c1:
            st.session_state.doc_text = st.text_area(
                t('text_input_label'),
                value=st.session_state.doc_text,
                height=160,
                placeholder=t('text_input_placeholder')
            )
        with c2:
            t_words = wc(st.session_state.doc_text)
            st.markdown(f"**{t('total_words')}** `{t_words}`")
            st.markdown(f"**{t('digital_footprint')}** `{sha256_text(st.session_state.doc_text)[:12]}...`")
            st.caption(t('engine_caption'))

        if st.button(t('run_btn'), use_container_width=True):
            if wc(st.session_state.doc_text) < 15:
                st.warning(t('short_text_warn'))
            else:
                with st.status(t('status_running'), expanded=True) as status:
                    st.write(t('step_1'))
                    body, bib = split_bib(st.session_state.doc_text) if exclude_bib else (st.session_state.doc_text, "")

                    st.write(t('step_2'))
                    ai_engine = AdvancedAIDetector(academic_mode=academic_mitigation_toggle)
                    ai_res = ai_engine.analyze_document(body)

                    st.write(t('step_3'))
                    sents = [s for s in sentences(body) if wc(s) >= 6]
                    queries = [" ".join([w for w in get_tokens(s) if w not in STOP_WORDS][:6]) for s in sents[:5]]
                    records = []
                    for q in queries:
                        # استدعاء جميع الدوال والمستوعبات المضافة وتجميع نتائجها
                        records.extend(query_academic_sources(q))
                        records.extend(query_global_repositories(q))
                        records.extend(query_regional_arabic_repositories(q))

                    sources_dict = {}
                    for r in records:
                        k = norm(r["title"])[:150]
                        if k not in sources_dict:
                            sources_dict[k] = Source(
                                source_id=sha256_text(k)[:10], title=r["title"], url=r["url"],
                                source_type=r["type"], text=r["text"], words=wc(r["text"])
                            )
                    sources = list(sources_dict.values())
                    matches, matched_indices = find_plagiarism_matches(body, sources, min_words, exclude_quotes=True)

                    total_doc_words = max(wc(body), 1)
                    sim_pct = round((len(matched_indices) / total_doc_words) * 100, 1)

                    st.session_state.sim_data = {
                        "body": body, "bib": bib, "sources": sources, "matches": matches,
                        "similarity_percent": sim_pct, "matched_words": len(matched_indices), "total_words": total_doc_words
                    }
                    st.session_state.ai_report = ai_res
                    st.session_state.analysis_done = True
                    status.update(label=t('status_complete'), state="complete")

        if st.session_state.analysis_done:
            sim = st.session_state.sim_data
            air = st.session_state.ai_report
            col_paper, col_sidebar = st.columns([7, 3])

            with col_paper:
                raw_sentences = sentences(sim["body"])
                rendered = []
                sm = {s.source_id: (idx + 1, HIGHLIGHT_PALETTE[idx % len(HIGHLIGHT_PALETTE)]) for idx, s in enumerate(sim["sources"])}

                for s_idx, sent in enumerate(raw_sentences, 1):
                    s_norm = norm(sent)
                    plag_match = [m for m in sim["matches"] if m["phrase"] in s_norm or s_norm in m["phrase"]]
                    is_ai_chunk = any(sent in seg.text or seg.text in sent for seg in air.suspicious_segments)

                    safe_sent = html.escape(sent)
                    if plag_match:
                        m_obj = plag_match[0]
                        lbl, color = sm.get(m_obj["source_id"], (1, HIGHLIGHT_PALETTE[0]))
                        rendered.append(
                            f'<mark class="mark-plag" style="background-color:{color["bg"]}; border-bottom:2px solid {color["border"]};">'
                            f'{safe_sent}<span class="pill-badge" style="background-color:{color["badge"]};">{lbl}</span></mark>'
                        )
                    elif is_ai_chunk:
                        rendered.append(
                            f'<mark class="mark-ai">'
                            f'{safe_sent}<span class="pill-badge" style="background-color:#8B5CF6;">AI</span></mark>'
                        )
                    else:
                        rendered.append(f'<span style="color:var(--text-muted); font-size:12px; font-family:monospace; margin:0 4px;">[{s_idx}]</span> {safe_sent}')

                canvas_html = "<br><br>".join(rendered)
                if sim.get("bib"):
                    canvas_html += f'<div style="color:var(--text-muted); font-size:15px; margin-top:40px; border-top:2px dashed var(--border-color); padding-top:20px;"><b>{t("excluded_bib_label")}</b><br>{html.escape(sim["bib"])}</div>'
                st.markdown(f'<div class="paper-canvas">{canvas_html}</div>', unsafe_allow_html=True)

            with col_sidebar:
                st.markdown('<div class="panel-card">', unsafe_allow_html=True)
                st.markdown(f"<h4 style='margin:0 0 20px 0; font-weight:800; text-align:center;'>{t('audit_results')}</h4>", unsafe_allow_html=True)

                sim_color = "#E11D48" if sim["similarity_percent"] > 20 else "#059669"
                ai_color = "#8B5CF6" if air.global_ai_percentage > 25 else "#059669"

                st.markdown(f"""
                <div style="text-align: center; border-bottom: 1px solid var(--border-color); padding-bottom: 18px; margin-bottom:18px;">
                    <div class="metric-num" style="color: {sim_color};">{sim["similarity_percent"]}%</div>
                    <div style="font-size: 14px; font-weight: 700; color: var(--text-main); margin-top: 5px;">{t('sim_index')}</div>
                    <div style="font-size: 12px; color: var(--text-muted);">{t('sim_desc', matched=sim['matched_words'], total=sim['total_words'])}</div>
                </div>
                <div style="text-align: center; border-bottom: 1px solid var(--border-color); padding-bottom: 18px; margin-bottom:18px;">
                    <div class="metric-num" style="color: {ai_color};">{air.global_ai_percentage}%</div>
                    <div style="font-size: 14px; font-weight: 700; color: var(--text-main); margin-top: 5px;">{t('ai_index')}</div>
                    <div style="font-size: 12px; color: var(--text-muted);">{t('ai_desc', flagged=air.flagged_chunks_count, total=air.total_chunks)}</div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown(f"<div style='font-size: 13px; font-weight:800; color:var(--text-main); margin: 15px 0 10px;'>{t('retrieved_sources')}</div>", unsafe_allow_html=True)
                for idx, s in enumerate(sim["sources"][:5], 1):
                    c_pal = HIGHLIGHT_PALETTE[(idx - 1) % len(HIGHLIGHT_PALETTE)]
                    st.markdown(f"""
                    <div class="source-card">
                        <div class="pill-badge" style="background:{c_pal['badge']}; font-size:12px; height:24px; min-width:24px; line-height:24px;">{idx}</div>
                        <div style="flex:1; overflow:hidden;">
                            <div style="font-weight:800; font-size:13px; color:var(--text-main); line-height:1.4; margin-bottom:3px;" title="{html.escape(s.title)}">{html.escape(s.title[:38])}...</div>
                            <div style="font-size:11px; color:var(--text-muted); font-weight:600;">{s.source_type}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

    with tab_ai:
        st.markdown(f"<h4 style='color:var(--text-main); font-weight:800;'>{t('ai_tab_title')}</h4>", unsafe_allow_html=True)
        st.caption(t('ai_tab_caption'))

        if not st.session_state.analysis_done:
            st.info(t('req_run_analytics'))
        else:
            air = st.session_state.ai_report
            ca1, ca2, ca3 = st.columns(3)
            ca1.metric(t('ai_metric_1'), f"{air.global_ai_percentage}%", delta=t('delta_ai_normal') if air.global_ai_percentage < 20 else t('delta_ai_high'), delta_color="inverse")
            ca2.metric(t('ai_metric_2'), f"{air.mean_perplexity}%", delta=t('delta_pred_normal') if air.mean_perplexity < 40 else t('delta_pred_high'), delta_color="inverse")
            ca3.metric(t('ai_metric_3'), f"{air.mean_burstiness}", delta=t('delta_cv_human') if air.mean_burstiness > 0.35 else t('delta_cv_ai'))

            st.divider()
            st.markdown(f"##### {t('suspicious_chunks_title')}")
            if not air.suspicious_segments:
                st.success(t('no_ai_chunks'))
            else:
                for seg in air.suspicious_segments:
                    st.markdown(f"""
                    <div class="suspicious-box">
                        <div style="font-weight:800; color:var(--mark-text-ai); font-size:15px; margin-bottom:8px;">{t('chunk_id', id=seg.chunk_id)} | {t('ai_prob')} {seg.ai_probability * 100:.1f}%</div>
                        <div style="margin:10px 0; font-size:15px; line-height:1.8;"><mark style="background:var(--mark-bg-ai); padding:4px 8px; border-radius:4px; color:var(--mark-text-ai);">{html.escape(seg.text)}</mark></div>
                        <div style="font-size:12.5px; color:var(--text-muted); background:var(--card-hover); padding:8px 12px; border-radius:6px; display:inline-block; margin-top:5px; border: 1px solid var(--border-color);">
                            {t('burstiness')} <b style="color:var(--text-main);">{seg.burstiness}</b> &nbsp;|&nbsp; {t('cliche_hits')} <b style="color:var(--text-main);">{seg.diagnostics.get('cliche_hits', 0)}</b> &nbsp;|&nbsp; {t('perplexity')} <b style="color:var(--text-main);">{seg.perplexity_proxy}%</b>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

    with tab_analytics:
        if not st.session_state.analysis_done:
            st.info(t('req_run_analytics'))
        else:
            sim = st.session_state.sim_data
            air = st.session_state.ai_report
            c_left, c_right = st.columns(2)

            with c_left:
                if PLOTLY_READY:
                    fig = go.Figure(data=[go.Pie(
                        labels=[t('pie_authentic'), t('pie_plagiarized'), t('pie_ai')],
                        values=[max(0.0, 100.0 - sim["similarity_percent"] - air.global_ai_percentage), sim["similarity_percent"], air.global_ai_percentage],
                        hole=0.65,
                        marker=dict(colors=["#10B981", "#E11D48", "#8B5CF6"], line=dict(color="rgba(0,0,0,0)", width=0))
                    )])
                    fig.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20), showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    fig.update_traces(textfont=dict(color='#FFFFFF')) 
                    st.plotly_chart(fig, use_container_width=True)

            with c_right:
                st.markdown(f"""
                <div class="panel-card">
                    <h5 style="margin-top:0; font-weight:800; font-size:18px;">{t('doc_stats')}</h5>
                    <div style="margin-top: 15px; font-size: 15px; color:var(--text-muted); line-height: 2;">
                        <div><span style="display:inline-block; width:15px; color:#4F46E5;">■</span> <b style="color:var(--text-main);">{t('total_words')}</b> {sim['total_words']}</div>
                        <div><span style="display:inline-block; width:15px; color:#4F46E5;">■</span> <b style="color:var(--text-main);">{t('total_sentences')}</b> {air.document_length_sentences}</div>
                        <div><span style="display:inline-block; width:15px; color:#4F46E5;">■</span> <b style="color:var(--text-main);">{t('suspicious_segments')}</b> {air.flagged_chunks_count} / {air.total_chunks}</div>
                        <div><span style="display:inline-block; width:15px; color:#4F46E5;">■</span> <b style="color:var(--text-main);">{t('mean_cv')}</b> {air.mean_burstiness}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

    with tab_humanizer:
        st.markdown(f"<h4 style='color:var(--text-main); font-weight:800;'>{t('hum_title')}</h4>", unsafe_allow_html=True)
        st.caption(t('hum_caption'))

        h_input = st.text_area(t('hum_input'), value=st.session_state.doc_text, height=140)
        if st.button(t('hum_btn')):
            if not h_input.strip():
                st.warning(t('hum_warn'))
            else:
                reworked = run_local_humanizer(h_input)
                st.success(t('hum_success'))
                st.text_area(t('hum_output'), value=reworked, height=180)
                if st.button(t('hum_apply')):
                    st.session_state.doc_text = reworked
                    st.session_state.analysis_done = False
                    st.rerun()

if __name__ == "__main__":
    main()
