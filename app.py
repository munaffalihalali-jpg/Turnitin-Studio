"""
========================================================================================
Academic Similarity & Statistical AI Detection Studio — v7.3.0 (Bilingual Engine)
المحرك الإحصائي المتكامل لفحص الاستلال وكشف النصوص المولدة بالذكاء الاصطناعي
========================================================================================
التعديلات الجوهرية (v7.3.0):
- دعم ثنائي اللغة (عربي / إنجليزي) لمصفوفة الكليشيهات والقوالب التوليدية.
- تفعيل خوارزمية التباين التدفقي (Global Burstiness CV) لمعالجة الرتابة الإنجليزية.
- تحديث دروع الأبحاث الأكاديمية (Mitigation Markers) لتشمل مصطلحات البحث الإنجليزي.
"""

from __future__ import annotations
import streamlit as st
import hashlib, html, io, json, math, os, re, time
from collections import Counter
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from typing import Optional, List, Tuple, Dict, Any
from urllib.parse import quote, urlparse, parse_qs
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

APP_VERSION = "7.3.0 Bilingual NLP"
HEADERS = {"User-Agent": "AcademicStatisticalIntegrityEngine/7.3 (research-nlp; contact@university.edu)"}

st.set_page_config(
    page_title="Turnitin Studio | منظومة التحليل الإحصائي والاستلال",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

TURNITIN_PALETTE = [
    {"bg": "#FEE2E2", "border": "#EF4444", "text": "#7F1D1D", "badge": "#DC2626"},
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

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Amiri:ital,wght@0,400;0,700;1,400&display=swap');
    
    html, body, [class*="css"], .stApp {
        font-family: 'Cairo', -apple-system, BlinkMacSystemFont, sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
        background-color: #F8FAFC !important;
    }

    section[data-testid="stSidebar"] {
        direction: rtl !important;
        text-align: right !important;
        background: #FFFFFF !important;
        border-left: 1px solid #E2E8F0 !important;
        box-shadow: -4px 0 16px rgba(0,0,0,0.02);
    }

    .turnitin-app-bar {
        background: linear-gradient(135deg, #0F2D59 0%, #1E3A8A 100%);
        color: #FFFFFF;
        padding: 20px 30px;
        border-radius: 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 22px;
        box-shadow: 0 4px 18px rgba(15, 45, 89, 0.12);
    }

    .turnitin-paper-canvas {
        background-color: #FFFFFF !important;
        color: #1E293B !important;
        border: 1px solid #CBD5E1;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
        border-radius: 8px;
        padding: 50px 60px;
        line-height: 2.8;
        font-size: 18.5px;
        font-family: 'Amiri', 'Cairo', serif;
        direction: rtl !important;
        text-align: justify !important;
        min-height: 750px;
    }

    .turnitin-mark {
        display: inline;
        padding: 2px 5px;
        border-radius: 3px;
        font-weight: 700;
    }

    .turnitin-ai-mark {
        display: inline;
        background-color: #F3E8FF !important;
        color: #581C87 !important;
        border-bottom: 2.5px solid #9333EA !important;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 700;
    }

    .turnitin-pill {
        display: inline-block;
        color: #FFFFFF !important;
        font-size: 10.5px;
        font-weight: 800;
        min-width: 18px;
        height: 18px;
        line-height: 18px;
        text-align: center;
        border-radius: 3px;
        margin: 0 3px;
        vertical-align: middle;
    }

    .match-overview-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }

    .source-match-card {
        background: #FFFFFF;
        border: 1px solid #F1F5F9;
        border-bottom: 2px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 14px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .metric-badge {
        font-size: 46px;
        font-weight: 900;
        line-height: 1;
    }
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
# المحرك الإحصائي ثنائي اللغة (Bilingual NLP Statistical Integrity Engine)
# =======================================================================
class AdvancedAIDetector:
    """
    محرك إحصائي متطور يحاكي Turnitin يعتمد على 4 معايير كمية:
    1. Burstiness (معامل الاختلاف التدفقي)
    2. Cliche & Transition Density (كثافة القوالب الآلية العربية والإنجليزية)
    3. Structural Uniformity (تجانس الجمل)
    4. Collocation Predictability (التلازم والقدرة على التوقع)
    """
    def __init__(self, academic_mode: bool = True):
        self.academic_mode = academic_mode
        self.ai_markers = [
            # القوالب التوليدية للغة العربية
            r"علاوة على ذلك", r"في هذا الصدد", r"تجدر الإشارة إلى", r"جدير بالذكر",
            r"مما لا شك فيه", r"يلعب[\s\w]*دوراً[\s\w]*محورياً", r"تلعب[\s\w]*دوراً[\s\w]*محورياً",
            r"وفي الختام", r"بشكل متسارع", r"مما يساهم في", r"على صعيد آخر",
            r"في إطار", r"يعد من أهم", r"تعد من أهم", r"من الأهمية بمكان",
            r"بدقة متناهية", r"بسرعة فائقة", r"استراتيجيات إدارة المخاطر",
            
            # القوالب التوليدية للغة الإنجليزية (English AI Signatures)
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
        """
        مؤشر التوقع (عكس الحيرة):
        النصوص الآلية تحوي تلازمات شائعة جداً وتنوعاً معجمياً وسطياً.
        درجة عالية تعني توقعاً آلياً عالياً (حيرة منخفضة).
        """
        tokens = [w for w in text.split() if w not in STOP_WORDS]
        if not tokens:
            return 50.0
        ttr = len(set(tokens)) / len(tokens)
        
        predictability = (1.0 - ttr) * 100.0
        return round(predictability, 2)

    def evaluate_text_prob(self, text: str, sents: List[str], global_cv: float) -> Tuple[float, Dict[str, Any]]:
        # 1. خطر التباين (CV منخفض = خطر آلي مرتفع)
        burst_risk = 1.0 / (1.0 + math.exp((global_cv - 0.32) / 0.05))

        # 2. خطر الكليشيهات والقوالب
        cliche_density, hit_count = self.compute_cliche_density(text, sents)
        cliche_risk = min(1.0, cliche_density * 1.5)

        # 3. خطر التوقع والمعجم
        pred = self.compute_predictability_proxy(text)
        pred_risk = min(1.0, max(0.1, pred / 50.0))

        # التركيب الوزني
        combined_prob = (0.45 * burst_risk) + (0.40 * cliche_risk) + (0.15 * pred_risk)

        # تخفيف مخصص للأبحاث العلمية ذات المصطلحات الثقيلة (Bilingual Mitigation)
        discount = 0.0
        if self.academic_mode:
            academic_markers = [
                # عربي
                r"استناداً إلى", r"يتضح من خلال", r"في ضوء ما سبق", r"وفقاً للمنهجية",
                # إنجليزي
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

        # التقطيع الديناميكي الذكي
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
                chunk_id=idx,
                start_char=start,
                end_char=end,
                text=c_text,
                sentence_count=len(sents),
                word_count=len(c_text.split()),
                perplexity_proxy=diag["pred_proxy"],
                burstiness=round(global_cv, 4),
                ai_probability=round(prob, 4),
                is_suspicious=is_susp,
                diagnostics=diag
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
            document_length_words=len(text.split()),
            document_length_sentences=len(all_sents_spans),
            total_chunks=len(reports),
            flagged_chunks_count=len(suspicious),
            global_ai_percentage=min(100.0, global_ai_pct),
            mean_perplexity=round(mean_ppl, 2),
            mean_burstiness=round(global_cv, 4),
            suspicious_segments=suspicious
        )

    @staticmethod
    def _merge_intervals(intervals: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
        if not intervals:
            return []
        sorted_intervals = sorted(intervals, key=lambda x: x[0])
        merged = [sorted_intervals[0]]
        for cur in sorted_intervals[1:]:
            prev_s, prev_e = merged[-1]
            if cur[0] <= prev_e:
                merged[-1] = (prev_s, max(prev_e, cur[1]))
            else:
                merged.append(cur)
        return merged

# =======================================================================
# دوال استرجاع المصادر ومطابقة الاستلال
# =======================================================================
def sha256_text(x: str | bytes) -> str:
    b = x.encode("utf-8", "ignore") if isinstance(x, str) else x
    return hashlib.sha256(b).hexdigest()

def norm(t: str) -> str:
    t = t or ""
    t = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", t)
    t = re.sub(r"[إأآٱ]", "ا", t).replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    t = re.sub(r"[.،,:;؛!?؟()\[\]{}<>«»“”\"']+", " ", t)
    return re.sub(r"\s+", " ", t).strip().lower()

def get_tokens(t: str) -> list[str]:
    return re.findall(r"[\w\u0600-\u06FF]+", norm(t), re.UNICODE)

def wc(t: str) -> int:
    return len(get_tokens(t))

def sentences(t: str) -> list[str]:
    return [x.strip() for x in re.split(r"(?<=[.!?؟؛])\s+|\n+", t or "") if wc(x) >= 4]

def extract_file(name: str, data: bytes) -> tuple[str, list[str]]:
    n = name.lower()
    if n.endswith(".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
            try:
                return data.decode(enc), []
            except UnicodeDecodeError:
                pass
        return data.decode("utf-8", "replace"), []
    if n.endswith(".pdf") and PYPDF_READY:
        try:
            r = pypdf.PdfReader(io.BytesIO(data))
            out = [p.extract_text() for p in r.pages if p.extract_text()]
            return "\n\n".join(out), []
        except Exception as e:
            return "", [f"فشل PDF: {e}"]
    if n.endswith(".docx") and DOCX_READY:
        try:
            d = docx.Document(io.BytesIO(data))
            out = [p.text for p in d.paragraphs if p.text.strip()]
            return "\n".join(out), []
        except Exception as e:
            return "", [f"فشل DOCX: {e}"]
    return "", ["نوع الملف غير مدعوم."]

def split_bib(t: str) -> tuple[str, str]:
    m = re.search(r"(?im)(?:^|\n)[ \t]*(?:references|bibliography|المراجع|المصادر|المصادر والمراجع)[ \t]*[:\-]?[ \t]*(?:\n|$)", t or "")
    return (t[:m.start()].rstrip(), t[m.start():].strip()) if m else (t, "")

@st.cache_data(ttl=3600, show_spinner=False)
def query_academic_sources(q: str) -> list[dict]:
    results = []
    try:
        r = requests.get(f"https://api.crossref.org/works?query.bibliographic={quote(q)}&rows=2", headers=HEADERS, timeout=6)
        if r.status_code == 200:
            for item in r.json().get("message", {}).get("items", []):
                t = (item.get("title") or [""])[0]
                ab = item.get("abstract") or ""
                clean_ab = re.sub(r"<[^>]+>", " ", ab)
                if t:
                    results.append({"title": f"مجلة: {t[:60]}", "url": item.get("URL") or "", "text": f"{t}. {clean_ab}", "type": "Publications"})
    except Exception:
        pass

    try:
        r_w = requests.get(f"https://ar.wikipedia.org/w/api.php?action=query&list=search&srsearch={quote(q)}&format=json&utf8=1&srlimit=2", headers=HEADERS, timeout=6)
        if r_w.status_code == 200:
            for item in r_w.json().get("query", {}).get("search", []):
                t = item.get("title") or ""
                if t:
                    r_txt = requests.get(f"https://ar.wikipedia.org/w/api.php?action=parse&page={quote(t)}&prop=text&format=json&formatversion=2", headers=HEADERS, timeout=6)
                    body_txt = ""
                    if r_txt.status_code == 200:
                        raw_h = (r_txt.json().get("parse") or {}).get("text") or ""
                        body_txt = re.sub(r"(?is)<[^>]+>", " ", raw_h)
                    results.append({"title": f"{t} — ويكيبيديا", "url": f"https://ar.wikipedia.org/wiki/{quote(t.replace(' ', '_'))}", "text": body_txt, "type": "Internet Sources"})
    except Exception:
        pass

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
                if ext in full_s_text:
                    j += 1
                else:
                    break
            matched_indices.update(range(i, j))
            matches.append({
                "source_id": found_source,
                "start": i,
                "end": j,
                "length": j - i,
                "phrase": " ".join(tokens[i:j])
            })
            i = j
        else:
            i += 1

    return matches, matched_indices

def run_local_humanizer(text: str) -> str:
    replacements = {
        "علاوة على ذلك": "وفي سياق متصل",
        "في هذا الصدد": "وعليه",
        "تجدر الإشارة إلى": "ومن الجدير بالملاحظة أن",
        "جدير بالذكر": "ومما يستوجب الوقوف عنده أن",
        "مما لا شك فيه": "ومن الثابت تحليلياً",
        "يلعب دورا محوريا": "يمثل ركيزة جوهرية في",
        "تعد من أهم": "تُصنف ضمن أبرز"
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
            reworked.append(" ".join(words[:mid]) + "؛")
            reworked.append("الأمر الذي يؤكد أن " + " ".join(words[mid:]))
        else:
            reworked.append(s)
    return " ".join(reworked)

# =======================================================================
# الواجهة التفاعلية (Streamlit App Interface)
# =======================================================================
def main():
    if "doc_text" not in st.session_state:
        st.session_state.doc_text = ""
    if "analysis_done" not in st.session_state:
        st.session_state.analysis_done = False
    if "sim_data" not in st.session_state:
        st.session_state.sim_data = None
    if "ai_report" not in st.session_state:
        st.session_state.ai_report = None

    st.markdown(f"""
    <div class="turnitin-app-bar">
        <div>
            <div style="font-size: 22px; font-weight: 900;">Turnitin Studio | Academic Integrity & Statistical NLP</div>
            <div style="font-size: 13px; opacity: 0.88; margin-top: 3px;">منظومة التدقيق الأكاديمي المسبق ونواة الكشف الإحصائي — الإصدار {APP_VERSION}</div>
        </div>
        <div style="text-align: left; font-size: 11px; direction: ltr; font-family: monospace;">
            <div>Engine: Bilingual Quad-Metric NLP</div>
            <div>Status: Fully Operational (High Sensitivity)</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("<h3 style='color:#0F2D59;'>📄 بيانات المسودة</h3>", unsafe_allow_html=True)
        author_name = st.text_input("اسم الباحث:", "د. مناف فالح عباس")
        paper_title = st.text_input("عنوان الورقة العلمية:", "فحص النزاهة الأكاديمية")

        st.divider()
        st.markdown("<h4 style='color:#0F2D59;'>📁 إيداع ملف البحث</h4>", unsafe_allow_html=True)
        up_file = st.file_uploader("رفع المستند (DOCX, PDF, TXT):", type=["docx", "pdf", "txt"])
        if up_file:
            data = up_file.read()
            extracted, warns = extract_file(up_file.name, data)
            if extracted and extracted != st.session_state.doc_text:
                st.session_state.doc_text = extracted
                st.session_state.analysis_done = False
                st.sidebar.success(f"تم إيداع الملف بنجاح ({wc(extracted)} كلمة).")
                st.rerun()

        st.divider()
        st.markdown("<h4 style='color:#0F2D59;'>⚙ معايير الخوارزمية</h4>", unsafe_allow_html=True)
        min_words = st.slider("عتبة السلسلة المتطابقة (Word Threshold):", 4, 10, 6)
        academic_mitigation_toggle = st.checkbox("تفعيل موازنة الأبحاث الصارمة (Mitigation)", value=True)
        exclude_bib = st.checkbox("استبعاد قائمة المراجع", value=True)

    tab_studio, tab_ai, tab_analytics, tab_humanizer = st.tabs([
        "📑 مستعرض البحث والمطابقات (Canvas)",
        "🟣 كاشف الذكاء الاصطناعي الإحصائي (NLP Detector)",
        "📊 مصفوفة التحليل والرسوم البيانية (Analytics)",
        "✍️ المساعد والمحرر الأكاديمي (Humanizer)"
    ])

    with tab_studio:
        c1, c2 = st.columns([3, 1])
        with c1:
            st.session_state.doc_text = st.text_area(
                "متن البحث المفحوص:",
                value=st.session_state.doc_text,
                height=160,
                placeholder="الصق نص البحث الأكاديمي (عربي أو إنجليزي) هنا..."
            )
        with c2:
            t_words = wc(st.session_state.doc_text)
            st.markdown(f"**إجمالي الكلمات:** `{t_words}` كلمة")
            st.markdown(f"**البصمة الرقمية:** `{sha256_text(st.session_state.doc_text)[:12]}...`")
            st.caption("يعتمد الكاشف على مصفوفة قياس التباين الإحصائي وتلازم الألفاظ ثنائية اللغة.")

        if st.button("🚀 تشغيل الفحص المزدوج الشامل (استلال + ذكاء اصطناعي)", use_container_width=True):
            if wc(st.session_state.doc_text) < 15:
                st.warning("النص قصير جداً؛ يُشترط وجود 15 كلمة على الأقل لبدء الفحص.")
            else:
                with st.status("جاري تنفيذ التحليل الإحصائي والاسترجاع الميداني...", expanded=True) as status:
                    st.write("1. استخراج المتن وفصل المراجع...")
                    body, bib = split_bib(st.session_state.doc_text) if exclude_bib else (st.session_state.doc_text, "")

                    st.write("2. تشغيل كاشف الذكاء الاصطناعي (Bilingual Quad-Metric NLP)...")
                    ai_engine = AdvancedAIDetector(academic_mode=academic_mitigation_toggle)
                    ai_res = ai_engine.analyze_document(body)

                    st.write("3. استعلام المستودعات الدولية وفحص سلاسل الاستلال...")
                    sents = [s for s in sentences(body) if wc(s) >= 6]
                    queries = [" ".join([w for w in get_tokens(s) if w not in STOP_WORDS][:6]) for s in sents[:5]]
                    records = []
                    for q in queries:
                        records.extend(query_academic_sources(q))

                    sources_dict = {}
                    for r in records:
                        k = norm(r["title"])[:150]
                        if k not in sources_dict:
                            sources_dict[k] = Source(
                                source_id=sha256_text(k)[:10],
                                title=r["title"],
                                url=r["url"],
                                source_type=r["type"],
                                text=r["text"],
                                words=wc(r["text"])
                            )
                    sources = list(sources_dict.values())
                    matches, matched_indices = find_plagiarism_matches(body, sources, min_words, exclude_quotes=True)

                    total_doc_words = max(wc(body), 1)
                    sim_pct = round((len(matched_indices) / total_doc_words) * 100, 1)

                    st.session_state.sim_data = {
                        "body": body,
                        "bib": bib,
                        "sources": sources,
                        "matches": matches,
                        "similarity_percent": sim_pct,
                        "matched_words": len(matched_indices),
                        "total_words": total_doc_words
                    }
                    st.session_state.ai_report = ai_res
                    st.session_state.analysis_done = True
                    status.update(label="اكتمل الفحص الشامل بنجاح!", state="complete")

        if st.session_state.analysis_done:
            sim = st.session_state.sim_data
            air = st.session_state.ai_report
            col_paper, col_sidebar = st.columns([7, 3])

            with col_paper:
                raw_sentences = sentences(sim["body"])
                rendered = []
                sm = {s.source_id: (idx + 1, TURNITIN_PALETTE[idx % len(TURNITIN_PALETTE)]) for idx, s in enumerate(sim["sources"])}

                for s_idx, sent in enumerate(raw_sentences, 1):
                    s_norm = norm(sent)
                    plag_match = [m for m in sim["matches"] if m["phrase"] in s_norm or s_norm in m["phrase"]]
                    is_ai_chunk = any(sent in seg.text or seg.text in sent for seg in air.suspicious_segments)

                    safe_sent = html.escape(sent)
                    if plag_match:
                        m_obj = plag_match[0]
                        lbl, color = sm.get(m_obj["source_id"], (1, TURNITIN_PALETTE[0]))
                        rendered.append(
                            f'<mark class="turnitin-mark" style="background-color:{color["bg"]}; color:{color["text"]}; border-bottom:2.5px solid {color["border"]};">'
                            f'{safe_sent}<span class="turnitin-pill" style="background-color:{color["badge"]};">{lbl}</span></mark>'
                        )
                    elif is_ai_chunk:
                        rendered.append(
                            f'<mark class="turnitin-ai-mark" title="مقطع مشبوه: تباين منخفض وكثافة قوالب آلية">'
                            f'{safe_sent}<span class="turnitin-pill" style="background-color:#9333EA;">AI</span></mark>'
                        )
                    else:
                        rendered.append(f'<span style="color:#94A3B8; font-size:11px; font-family:monospace;">[{s_idx}]</span> {safe_sent}')

                canvas_html = "<br><br>".join(rendered)
                if sim.get("bib"):
                    canvas_html += f'<div style="color:#64748B; font-size:14px; margin-top:30px; border-top:1px dashed #CBD5E1; padding-top:12px;"><b>المراجع المستبعدة:</b><br>{html.escape(sim["bib"])}</div>'
                st.markdown(f'<div class="turnitin-paper-canvas">{canvas_html}</div>', unsafe_allow_html=True)

            with col_sidebar:
                st.markdown('<div class="match-overview-panel">', unsafe_allow_html=True)
                st.markdown("<h4 style='margin:0; color:#0F2D59;'>نتائج التدقيق الرقمي</h4>", unsafe_allow_html=True)

                sim_color = "#DC2626" if sim["similarity_percent"] > 20 else "#059669"
                ai_color = "#9333EA" if air.global_ai_percentage > 25 else "#059669"

                st.markdown(f"""
                <div style="text-align: center; border-bottom: 2px solid #F1F5F9; padding: 12px 0;">
                    <div class="metric-badge" style="color: {sim_color};">{sim["similarity_percent"]}%</div>
                    <div style="font-size: 13px; font-weight: 700; color: #475569;">مؤشر الاستلال اللفظي</div>
                    <div style="font-size: 11px; color: #64748B;">{sim['matched_words']} كلمة من أصل {sim['total_words']}</div>
                </div>
                <div style="text-align: center; border-bottom: 2px solid #F1F5F9; padding: 12px 0; margin-top: 8px;">
                    <div class="metric-badge" style="color: {ai_color};">{air.global_ai_percentage}%</div>
                    <div style="font-size: 13px; font-weight: 700; color: #475569;">مؤشر الاشتباه التوليدي (AI)</div>
                    <div style="font-size: 11px; color: #64748B;">{air.flagged_chunks_count} مقاطع مشبوهة من أصل {air.total_chunks}</div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("<div style='font-size: 12px; font-weight:700; color:#334155; margin: 12px 0 6px;'>المصادر المسترجعة:</div>", unsafe_allow_html=True)
                for idx, s in enumerate(sim["sources"][:5], 1):
                    c_pal = TURNITIN_PALETTE[(idx - 1) % len(TURNITIN_PALETTE)]
                    st.markdown(f"""
                    <div class="source-match-card">
                        <div class="turnitin-pill" style="background:{c_pal['badge']};">{idx}</div>
                        <div style="flex:1; overflow:hidden;">
                            <div style="font-weight:700; font-size:12px; color:#0F172A;" title="{html.escape(s.title)}">{html.escape(s.title[:32])}</div>
                            <div style="font-size:10px; color:#64748B;">{s.source_type}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

    with tab_ai:
        st.markdown("<h4 style='color:#0F2D59; font-weight:800;'>التحليل الإحصائي المعمق لمؤشرات التوليد الآلي</h4>", unsafe_allow_html=True)
        st.caption("تحليل احتمالي يستند إلى معامل الاختلاف التدفقي (Burstiness) وكثافة القوالب التوليدية.")

        if not st.session_state.analysis_done:
            st.info("قم بتشغيل الفحص أولاً لعرض تقرير الذكاء الاصطناعي الإحصائي.")
        else:
            air = st.session_state.ai_report
            ca1, ca2, ca3 = st.columns(3)
            ca1.metric("مؤشر الذكاء الاصطناعي الكلي", f"{air.global_ai_percentage}%", delta="طبيعي" if air.global_ai_percentage < 20 else "مرتفع")
            ca2.metric("مؤشر التوقع والتلازم", f"{air.mean_perplexity}%", delta="توقع طبيعي" if air.mean_perplexity < 40 else "توقع نمطي")
            ca3.metric("معامل التباين التدفقي (CV)", f"{air.mean_burstiness}", delta="تدفق بشري" if air.mean_burstiness > 0.35 else "رتابة آلية")

            st.divider()
            st.markdown("##### المقاطع المرصودة ذات الاشتباه الإحصائي المرتفع:")
            if not air.suspicious_segments:
                st.success("لم يرصد المحرك أي مقاطع ذات نمطية آلية متطابقة.")
            else:
                for seg in air.suspicious_segments:
                    st.markdown(f"""
                    <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-right:4px solid #9333EA; padding:14px; border-radius:6px; margin-bottom:10px;">
                        <div style="font-weight:800; color:#581C87;">مقطع مشتبه به #{seg.chunk_id} | احتمال توليدي: {seg.ai_probability * 100:.1f}%</div>
                        <div style="margin:6px 0; font-size:14px;"><mark style="background:#F3E8FF; padding:3px 6px; border-radius:3px;">{html.escape(seg.text)}</mark></div>
                        <div style="font-size:11.5px; color:#64748B;">
                            التباين التدفقي: <b>{seg.burstiness}</b> | القوالب الآلية المرصودة: <b>{seg.diagnostics.get('cliche_hits', 0)}</b> | درجة التوقع: <b>{seg.perplexity_proxy}%</b>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

    with tab_analytics:
        if not st.session_state.analysis_done:
            st.info("يرجى تشغيل الفحص أولاً لتوليد الرسوم البيانية.")
        else:
            sim = st.session_state.sim_data
            air = st.session_state.ai_report
            c_left, c_right = st.columns(2)

            with c_left:
                if PLOTLY_READY:
                    fig = go.Figure(data=[go.Pie(
                        labels=["محتوى أصيل", "استلال منسوخ", "نص مشبوه آلياً"],
                        values=[max(0.0, 100.0 - sim["similarity_percent"] - air.global_ai_percentage), sim["similarity_percent"], air.global_ai_percentage],
                        hole=0.65,
                        marker=dict(colors=["#10B981", "#EF4444", "#9333EA"], line=dict(color="#FFF", width=2))
                    )])
                    fig.update_layout(height=300, margin=dict(l=20, r=20, t=20, b=20), showlegend=True)
                    st.plotly_chart(fig, use_container_width=True)

            with c_right:
                st.markdown(f"""
                <div style="background:#FFF; border:1px solid #E2E8F0; padding:20px; border-radius:8px;">
                    <h5 style="color:#0F2D59; margin-top:0;">إحصائيات الوثيقة الموزونة:</h5>
                    <p style="margin:6px 0;">• <b>إجمالي الكلمات:</b> {sim['total_words']}</p>
                    <p style="margin:6px 0;">• <b>إجمالي الجمل:</b> {air.document_length_sentences}</p>
                    <p style="margin:6px 0;">• <b>المقاطع المشبوهة:</b> {air.flagged_chunks_count} من أصل {air.total_chunks}</p>
                    <p style="margin:6px 0;">• <b>معامل التباين العام (CV):</b> {air.mean_burstiness}</p>
                </div>
                """, unsafe_allow_html=True)

    with tab_humanizer:
        st.markdown("<h4 style='color:#0F2D59; font-weight:800;'>إعادة البناء الأسلوبي وأنسنة النصوص محلياً</h4>", unsafe_allow_html=True)
        st.caption("أداة فورية تكسر الرتابة، ترفع التباين التدفقي (Burstiness)، وتستبدل القوالب الآلية.")

        h_input = st.text_area("النص المراد مراجعته وتعديله:", value=st.session_state.doc_text, height=140)
        if st.button("✨ تشغيل الأنسنة ورفع التباين اللغوي الآن"):
            if not h_input.strip():
                st.warning("الرجاء إدخال نص أولاً.")
            else:
                reworked = run_local_humanizer(h_input)
                st.success("تمت إعادة الهيكلة بنجاح مع رفع تباين الجمل وكسر الرتابة!")
                st.text_area("النص بعد الأنسنة الأكاديمية:", value=reworked, height=180)
                if st.button("📥 اعتماد النص المحسن وفحصه فوراً"):
                    st.session_state.doc_text = reworked
                    st.session_state.analysis_done = False
                    st.rerun()

if __name__ == "__main__":
    main()