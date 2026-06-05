import re
import os
import json
import time
import threading
import streamlit as st

# ─── مسارات الملفات ────────────────────────────────────────────────────────────
SHEETS_ID      = "1QirRmdv5LI9FsDjOeTAhQfIP5jH4vd2Y"
SHEETS_URL     = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=xlsx"
LOCAL_BACKUP   = os.path.join(os.path.dirname(__file__), "foods.xlsx")
DATA_CACHE     = os.path.join(os.path.dirname(__file__), "data_cache.json")
REFRESH_SEC    = 120

# ══════════════════════════════════════════════════════════════════════════════
# DataManager — singleton
# ══════════════════════════════════════════════════════════════════════════════
class DataManager:
    def __init__(self):
        self._lock    = threading.Lock()
        self._grams:  list = []
        self._alts:   list = []
        self._version: int = 0

        g, a = self._read_disk_cache()
        if g is not None:
            self._grams, self._alts = g, a
            self._version = 1

        if not self._grams and not self._alts:
            g, a = self._fetch_remote()
            if g is None:
                g, a = self._fetch_local()
            if g:
                self._grams, self._alts = g, a
                self._version = 1
                self._write_disk_cache(g, a)

        t = threading.Thread(target=self._bg_loop, daemon=True)
        t.start()

    def get(self):
        with self._lock:
            return list(self._grams), list(self._alts), self._version

    @staticmethod
    def _write_disk_cache(grams, alts):
        try:
            with open(DATA_CACHE, "w", encoding="utf-8") as f:
                json.dump({"grams": grams, "alts": alts}, f, ensure_ascii=False, separators=(',', ':'))
        except Exception:
            pass

    @staticmethod
    def _read_disk_cache():
        try:
            with open(DATA_CACHE, "r", encoding="utf-8") as f:
                d = json.load(f)
            return d.get("grams", []), d.get("alts", [])
        except Exception:
            return None, None

    @staticmethod
    def _parse(df):
        df.columns = df.columns.str.strip()
        known = {
            "النوع":  "name",
            "الفئة":  "section",
            "الكمية": "amount",
        }
        carb_col   = next((c for c in df.columns if c not in known and not str(c).startswith("Unnamed")), None)
        carb_label = carb_col.strip() if carb_col else "معامل الكارب"
        rename = {k: v for k, v in known.items() if k in df.columns}
        if carb_col:
            rename[carb_col] = "carbs_raw"
        df = df.rename(columns=rename)
        if "name" not in df.columns:
            return [], []
        df = df.dropna(subset=["name"])
        df["name"]      = df["name"].astype(str).str.strip()
        df["section"]   = df["section"].astype(str).str.strip()   if "section"   in df.columns else ""
        df["amount"]    = df["amount"].astype(str).str.strip()    if "amount"    in df.columns else ""
        df["carbs_raw"] = df["carbs_raw"].astype(str).str.strip() if "carbs_raw" in df.columns else ""
        df["carb_label"] = carb_label
        df["section_ar"] = df["section"].apply(
            lambda s: "البدائل" if "بدائل" in s else ("الجرام" if "جرام" in s else s.strip())
        )
        df["carbs_clean"] = df["carbs_raw"].apply(_clean_carb)
        df["_norm_name"]  = df["name"].apply(_normalize)
        df["_norm_sec"]   = df["section_ar"].apply(_normalize)
        grams = df[df["section"].str.contains("جرام",  na=False)].to_dict("records")
        alts  = df[df["section"].str.contains("بدائل", na=False)].to_dict("records")
        grams.sort(key=lambda x: str(x.get("name", "")))
        alts.sort(key=lambda x:  str(x.get("name", "")))
        return grams, alts

    @staticmethod
    def _fetch_remote():
        try:
            import io
            import requests
            import pandas as pd
            r = requests.get(SHEETS_URL, timeout=12)
            if r.status_code == 200:
                df = pd.read_excel(io.BytesIO(r.content), engine="openpyxl")
                return DataManager._parse(df)
        except Exception:
            pass
        return None, None

    @staticmethod
    def _fetch_local():
        try:
            import pandas as pd
            df = pd.read_excel(LOCAL_BACKUP, engine="openpyxl")
            return DataManager._parse(df)
        except Exception:
            return [], []

    def _bg_loop(self):
        while True:
            g, a = self._fetch_remote()
            if g is not None:
                with self._lock:
                    self._grams   = g
                    self._alts    = a
                    self._version += 1
                self._write_disk_cache(g, a)
            time.sleep(REFRESH_SEC)

@st.cache_resource
def get_manager() -> DataManager:
    return DataManager()

# ══════════════════════════════════════════════════════════════════════════════
# دوال مساعدة
# ══════════════════════════════════════════════════════════════════════════════
def _normalize(text: str) -> str:
    text = str(text).strip().lower()
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)
    text = re.sub(r'[أإآٱ]', 'ا', text)
    text = re.sub(r'ة', 'ه', text)
    text = re.sub(r'ؤ', 'و', text)
    text = re.sub(r'[ئى]', 'ي', text)
    return text

def _clean_carb(val) -> str:
    if val is None:
        return "—"
    t = str(val).strip()
    if t in ("", "nan", "None"):
        return "—"
    t = t.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
    return t.replace('٫', '.').replace('٬', ',')

def _matches(q_norm: str, food: dict) -> bool:
    return q_norm in food.get("_norm_name", "") or q_norm in food.get("_norm_sec", "")

# ══════════════════════════════════════════════════════════════════════════════
# HTML builders
# ══════════════════════════════════════════════════════════════════════════════
def _build_gram_card_html(food: dict) -> str:
    name  = food.get("name", "")
    sec   = food.get("section_ar", "الجرام")
    carbs = food.get("carbs_clean", "—")
    return (
        f"<div class='card-gram'>"
        f"<div class='cg-name'>{name}</div>"
        f"<span class='cg-cat'>📂 {sec}</span>"
        f"<div class='cg-carb'>"
        f"<span class='cg-carb-num'>{carbs}</span>"
        f"</div></div>"
    )

def _build_alt_card_html(food: dict) -> str:
    name   = food.get("name", "")
    sec    = food.get("section_ar", "البدائل")
    amount = food.get("amount", "")
    carbs  = food.get("carbs_clean", "—")
    amt_ok = amount and amount not in ("nan", "None", "")
    pill_amount = (
        f"<div class='ca-pill'><span class='ca-pill-label'>الكمية</span>"
        f"<span class='ca-pill-value'>{amount}</span></div>"
        if amt_ok else ""
    )
    return (
        f"<div class='card-alt'>"
        f"<div class='ca-top'><div class='ca-name'>{name}</div>"
        f"<span class='ca-cat'>📂 {sec}</span></div>"
        f"<div class='ca-body'>{pill_amount}"
        f"<div class='ca-pill'><span class='ca-pill-label'>الكربوهيدرات</span>"
        f"<span class='ca-pill-value'>{carbs}</span></div>"
        f"</div></div>"
    )

def _get_card_html_cache(section: str, foods: list, version: int, search: str) -> list[str]:
    cache_key = f"_html_{section}_{version}_{search}"
    if cache_key not in st.session_state:
        builder = _build_gram_card_html if section == "gram" else _build_alt_card_html
        st.session_state[cache_key] = [builder(f) for f in foods]
    return st.session_state[cache_key]

# ══════════════════════════════════════════════════════════════════════════════
# إعداد الصفحة + CSS
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="كارب الأطعمة",
    page_icon="🥗",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)

st.markdown("""
<style>
*, html, body, [class*="css"] {
    font-family: 'Cairo', 'Segoe UI', Tahoma, Arial, sans-serif !important;
}
.stApp { direction: rtl; background: #f0f2f8; }
h1,h2,h3,h4,h5,h6,p,label,span,div { direction: rtl; text-align: right; }

section[data-testid="stSidebar"],
[data-testid="collapsedControl"] { display: none !important; }
#MainMenu, header[data-testid="stHeader"], .stDeployButton { display: none !important; }

.block-container {
    padding-top: 2rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 860px !important;
}

.stTextInput > div > div > input {
    direction: rtl !important;
    text-align: right !important;
    font-size: 1rem !important;
    border-radius: 14px !important;
    border: 2px solid #dde3f0 !important;
    padding: 12px 18px !important;
    background: #ffffff !important;
    box-shadow: 0 2px 8px rgba(26,35,126,0.07) !important;
}
.stTextInput > div > div > input:focus {
    border-color: #1a237e !important;
    box-shadow: 0 2px 14px rgba(26,35,126,0.16) !important;
}

.search-summary {
    background: #e8f0fe;
    border-right: 4px solid #1a237e;
    border-radius: 10px;
    padding: 10px 16px;
    margin-bottom: 16px;
    font-size: 0.88rem;
    font-weight: 700;
    color: #1a237e;
    direction: rtl;
    display: flex;
    gap: 18px;
    flex-wrap: wrap;
    align-items: center;
}

.sec-hdr {
    border-radius: 14px;
    padding: 14px 20px;
    margin-bottom: 22px;
    display: flex;
    align-items: center;
    gap: 10px;
    color: #fff;
    flex-wrap: nowrap;
}
.sec-hdr-gram { background: linear-gradient(135deg,#1a237e,#3949ab); }
.sec-hdr-alt  { background: linear-gradient(135deg,#1b5e20,#388e3c); }
.sec-hdr-icon { font-size: 1.4rem; flex-shrink: 0; }
.sec-hdr-text { font-size: 1rem; font-weight: 900; flex: 1; white-space: nowrap; }

.card-gram {
    background: #fff;
    border-radius: 16px;
    padding: 18px 20px;
    border: 1px solid #e3e8f4;
    box-shadow: 0 2px 8px rgba(26,35,126,0.06);
    margin-bottom: 2px;
    transition: box-shadow 0.2s;
}
.card-gram:hover { box-shadow: 0 4px 18px rgba(26,35,126,0.13); }
.cg-name { font-size: 1.05rem; font-weight: 800; color: #1a237e; margin-bottom: 8px; line-height: 1.3; }
.cg-cat {
    display: inline-block;
    background: #e8f0fe;
    color: #3949ab;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.78rem;
    font-weight: 700;
    margin-bottom: 14px;
}
.cg-carb {
    background: #f4f6ff;
    border-radius: 10px;
    padding: 10px 14px;
    margin-top: 4px;
    display: flex;
    flex-direction: row-reverse;
    align-items: baseline;
    justify-content: flex-end;
    gap: 8px;
}
.cg-carb-num { font-size: 1.3rem; font-weight: 900; color: #1a237e; }

.card-alt {
    background: #fff;
    border-radius: 16px;
    padding: 0;
    border: 1px solid #dcf0e3;
    box-shadow: 0 2px 8px rgba(27,94,32,0.06);
    margin-bottom: 2px;
    overflow: hidden;
    transition: box-shadow 0.2s;
}
.card-alt:hover { box-shadow: 0 4px 18px rgba(27,94,32,0.13); }
.ca-top {
    background: linear-gradient(135deg,#f1faf3,#e8f5e9);
    padding: 14px 18px 10px;
    border-bottom: 1px solid #dcf0e3;
    border-radius: 16px 16px 0 0;
}
.ca-name { font-size: 1.05rem; font-weight: 800; color: #1b5e20; margin-bottom: 6px; }
.ca-cat {
    display: inline-block;
    background: #c8e6c9;
    color: #1b5e20;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.78rem;
    font-weight: 700;
}
.ca-body { padding: 12px 18px 14px; display: flex; gap: 10px; flex-wrap: wrap; }
.ca-pill {
    display: flex;
    flex-direction: column;
    align-items: center;
    background: #f8fffe;
    border: 1px solid #e0f2e8;
    border-radius: 12px;
    padding: 8px 16px;
    flex: 1;
    min-width: 90px;
}
.ca-pill-label { font-size: 0.72rem; color: #888; margin-bottom: 3px; }
.ca-pill-value { font-size: 1rem; font-weight: 800; color: #1b5e20; }

.empty-msg { text-align: center; color: #bbb; padding: 50px 0; font-size: 0.95rem; }

@media (max-width: 640px) {
    [data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        width: 100% !important; min-width: 100% !important; flex: 0 0 100% !important;
    }
    h1 { font-size: 1.35rem !important; }
    .block-container { padding: 1rem 0.75rem !important; }
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# جلب البيانات
# ══════════════════════════════════════════════════════════════════════════════
GRAMS_FOODS, ALT_FOODS, DATA_VERSION = get_manager().get()
total_all = len(GRAMS_FOODS) + len(ALT_FOODS)

# ─── العنوان ──────────────────────────────────────────────────────────────────
st.markdown(
    "<h1 style='color:#1a237e;margin-bottom:16px;text-align:center;'>🥗 كارب الأطعمة</h1>",
    unsafe_allow_html=True,
)

# ─── البحث ────────────────────────────────────────────────────────────────────
global_search = st.text_input(
    label="بحث", placeholder="ابحث...",
    label_visibility="collapsed", key="gs",
)

if global_search:
    search_key = f"_search_{DATA_VERSION}_{global_search}"
    if search_key not in st.session_state:
        q = _normalize(global_search)
        st.session_state[search_key] = (
            [f for f in GRAMS_FOODS if _matches(q, f)],
            [f for f in ALT_FOODS   if _matches(q, f)],
        )
    fg, fa = st.session_state[search_key]
    st.markdown(f"""
<div class='search-summary'>
    <span>🔎 «{global_search}»</span>
    <span>⚖️ الجرام: {len(fg)}</span>
    <span>🔄 البدائل: {len(fa)}</span>
    <span>الإجمالي: {len(fg)+len(fa)}</span>
</div>""", unsafe_allow_html=True)
else:
    fg = GRAMS_FOODS

fa = ALT_FOODS

st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

# ─── التبويبات ────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs([
    "⚖️ حساب الكربوهيدرات بالجرام",
    "🔄 بدائل الكربوهيدرات",
])

# ─── دوال رسم البطاقات ────────────────────────────────────────────────────────
def render_gram_cards(foods: list, version: int, search: str):
    html_list = _get_card_html_cache("gram", foods, version, search)
    for i in range(0, len(foods), 2):
        cols = st.columns(2, gap="small")
        for j, col in enumerate(cols):
            idx = i + j
            if idx >= len(foods):
                break
            with col:
                st.markdown(html_list[idx], unsafe_allow_html=True)

def render_alt_cards(foods: list, version: int, search: str):
    html_list = _get_card_html_cache("alt", foods, version, search)
    for i in range(0, len(foods), 2):
        cols = st.columns(2, gap="small")
        for j, col in enumerate(cols):
            idx = i + j
            if idx >= len(foods):
                break
            with col:
                st.markdown(html_list[idx], unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# القسم الأول — الجرام
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("""
<div class='sec-hdr sec-hdr-gram'>
  <span class='sec-hdr-icon'>⚖️</span>
  <span class='sec-hdr-text'>حساب الكربوهيدرات بالجرام</span>
</div>""", unsafe_allow_html=True)
    if total_all == 0:
        st.info("لا توجد بيانات بعد.")
    elif not fg:
        st.markdown("<div class='empty-msg'>لا توجد نتائج مطابقة</div>", unsafe_allow_html=True)
    else:
        render_gram_cards(fg, DATA_VERSION, global_search)

# ══════════════════════════════════════════════════════════════════════════════
# القسم الثاني — البدائل
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
<div class='sec-hdr sec-hdr-alt'>
  <span class='sec-hdr-icon'>🔄</span>
  <span class='sec-hdr-text'>بدائل الكربوهيدرات</span>
</div>""", unsafe_allow_html=True)
    if total_all == 0:
        st.info("لا توجد بيانات بعد.")
    elif not fa:
        st.markdown("<div class='empty-msg'>لا توجد نتائج مطابقة</div>", unsafe_allow_html=True)
    else:
        render_alt_cards(fa, DATA_VERSION, global_search)
