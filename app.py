import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
import base64
import hashlib
import json as _json
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from io import BytesIO

# ══════════════════════════════════════════════════════
#  ANTHROPIC API CONFIG
# ══════════════════════════════════════════════════════
import os
import anthropic

# Leer API key desde múltiples fuentes posibles
def get_anthropic_client():
    api_key = None
    
    # Intentar desde st.secrets
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except:
        pass
    
    # Intentar desde variable de entorno
    if not api_key:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
    
    if not api_key:
        st.error("❌ ANTHROPIC_API_KEY no encontrada en secrets ni variables de entorno")
        return None
    
    return anthropic.Anthropic(api_key=api_key)

# ══════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(
    page_title="NEXA — Plataforma Documental",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════
#  GLOBAL CSS — NEXA v2 Premium Dark SaaS
# ══════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

/* ── RESET & BASE ─────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body, .stApp {
    background: #F6F8FB !important;
    color: #0A0F1E !important;
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* ── OCULTAR CHROME STREAMLIT ─────────────────────── */
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"]  { display: none !important; }
footer                         { visibility: hidden !important; }
.stDeployButton                { display: none !important; }
#MainMenu                      { visibility: hidden !important; }

/* ── ÁREA PRINCIPAL ───────────────────────────────── */
[data-testid="stAppViewContainer"] > [data-testid="stMain"] {
    background: #F6F8FB !important;
}
.block-container {
    padding-top: 1.5rem !important;
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
    padding-bottom: 4rem !important;
    max-width: 100% !important;
    background: #F6F8FB !important;
}

/* ── SIDEBAR ──────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #0A0F1E !important;
    border-right: 1px solid rgba(255,255,255,0.05) !important;
    min-width: 260px !important;
    max-width: 260px !important;
}
[data-testid="stSidebar"] > div:first-child {
    padding: 0 !important;
}
[data-testid="stSidebarContent"] {
    padding: 0 !important;
}
[data-testid="stSidebarCollapseButton"] {
    display: none !important;
}

/* ── TABS (NAVEGACIÓN PRINCIPAL) ──────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.08) !important;
    border-radius: 12px !important;
    padding: 4px !important;
    gap: 4px !important;
    box-shadow: 0 1px 3px rgba(10,15,30,0.06) !important;
    margin-bottom: 28px !important;
    flex-wrap: wrap !important;
    justify-content: center !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #8494A8 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    border-radius: 8px !important;
    padding: 8px 18px !important;
    border: none !important;
    transition: all .2s ease !important;
    white-space: nowrap !important;
    letter-spacing: 0.1px !important;
}
.stTabs [data-baseweb="tab"]:hover {
    background: rgba(0,194,203,0.08) !important;
    color: #00C2CB !important;
    transform: translateY(-1px) !important;
}
.stTabs [aria-selected="true"][data-baseweb="tab"] {
    background: #0A0F1E !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
    box-shadow: 0 2px 12px rgba(0,194,203,0.2) !important;
}
.stTabs [data-baseweb="tab-highlight"] { display: none !important; }
.stTabs [data-baseweb="tab-border"]    { display: none !important; }
.stTabs [data-testid="stTabsContent"]  { padding-top: 0 !important; }

/* ── TIPOGRAFÍA ───────────────────────────────────── */
h1 { color: #0A0F1E !important; font-weight: 800 !important; font-size: 1.6rem !important; letter-spacing: -0.5px !important; }
h2 { color: #1A2234 !important; font-weight: 700 !important; }
h3 { color: #2D3A52 !important; font-weight: 600 !important; }
p  { color: #556070 !important; line-height: 1.6 !important; }
strong { color: #0A0F1E !important; }

/* ── BOTONES PRIMARIOS ────────────────────────────── */
[data-testid="baseButton-primary"],
button[kind="primary"] {
    background: linear-gradient(135deg, #00C2CB 0%, #0099FF 100%) !important;
    border: none !important;
    color: #0A0F1E !important;
    font-weight: 700 !important;
    border-radius: 10px !important;
    font-size: 14px !important;
    letter-spacing: 0.2px !important;
    box-shadow: 0 4px 16px rgba(0,194,203,0.25) !important;
    transition: all .2s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
}
[data-testid="baseButton-primary"]:hover,
button[kind="primary"]:hover {
    background: linear-gradient(135deg, #00ADB5 0%, #0088EE 100%) !important;
    box-shadow: 0 6px 28px rgba(0,194,203,0.46) !important;
    transform: translateY(-2px) !important;
}
[data-testid="baseButton-primary"]:active,
button[kind="primary"]:active { transform: translateY(0) !important; }

/* ── BOTONES SECUNDARIOS ──────────────────────────── */
[data-testid="baseButton-secondary"],
button[kind="secondary"] {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.12) !important;
    color: #2D3A52 !important;
    border-radius: 10px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    transition: all .18s ease !important;
}
[data-testid="baseButton-secondary"]:hover,
button[kind="secondary"]:hover {
    background: rgba(0,194,203,0.08) !important;
    border-color: rgba(0,194,203,0.45) !important;
    color: #00C2CB !important;
    transform: translateY(-1px) !important;
}

/* ── DOWNLOAD BUTTON ──────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #00C2CB 0%, #0099FF 100%) !important;
    border: none !important;
    color: #0A0F1E !important;
    font-weight: 700 !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 16px rgba(0,194,203,0.25) !important;
    transition: all .2s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    box-shadow: 0 6px 28px rgba(0,194,203,0.46) !important;
    transform: translateY(-2px) !important;
}

/* ── INPUTS ───────────────────────────────────────── */
[data-testid="stTextInput"] label {
    color: #556070 !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
}
[data-testid="stTextInput"] > div > div > input {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.12) !important;
    color: #0A0F1E !important;
    border-radius: 10px !important;
    padding: 10px 14px !important;
    font-size: 14px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    caret-color: #00C2CB !important;
    transition: border-color .15s, box-shadow .15s !important;
}
[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #00C2CB !important;
    box-shadow: 0 0 0 3px rgba(0,194,203,0.12) !important;
    outline: none !important;
}
[data-testid="stTextInput"] > div > div > input::placeholder { color: #8494A8 !important; }

/* ── NUMBER INPUT ─────────────────────────────────── */
[data-testid="stNumberInput"] input {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.12) !important;
    color: #0A0F1E !important;
    border-radius: 10px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

/* ── SELECT / RADIO ───────────────────────────────── */
[data-testid="stSelectbox"] > div > div {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.12) !important;
    border-radius: 10px !important;
    color: #0A0F1E !important;
}

/* ── FILE UPLOADER ────────────────────────────────── */
[data-testid="stFileUploader"] section {
    background: rgba(0,194,203,0.02) !important;
    border: 2px dashed rgba(0,194,203,0.25) !important;
    border-radius: 16px !important;
    transition: all .25s ease !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #00C2CB !important;
    background: rgba(0,194,203,0.05) !important;
    box-shadow: 0 0 30px rgba(0,194,203,0.06) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #556070 !important;
    font-size: 13px !important;
}
[data-testid="stFileUploaderDropzone"] svg { color: #8494A8 !important; }

/* ── CONTAINERS CON BORDE ─────────────────────────── */
[data-testid="stVerticalBlockBorderWrapper"] > div {
    background: #FFFFFF !important;
    border: 1px solid rgba(10,15,30,0.08) !important;
    border-radius: 16px !important;
    transition: all .2s ease !important;
    box-shadow: 0 1px 3px rgba(10,15,30,0.06) !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div:hover {
    border-color: rgba(0,194,203,0.28) !important;
    box-shadow: 0 4px 16px rgba(10,15,30,0.1) !important;
}

/* ── DATA EDITOR ──────────────────────────────────── */
[data-testid="stDataEditor"] {
    border: 1px solid rgba(0,194,203,0.14) !important;
    border-radius: 12px !important;
    overflow: hidden !important;
}

/* ── ALERTAS ──────────────────────────────────────── */
[data-testid="stAlert"] {
    background: rgba(0,194,203,0.05) !important;
    border: 1px solid rgba(0,194,203,0.16) !important;
    border-radius: 12px !important;
    color: #556070 !important;
    font-size: 13px !important;
}
[data-testid="stAlert"] p      { color: #556070 !important; }
[data-testid="stAlert"] strong { color: #00C2CB !important; }

/* ── PROGRESS / SPINNER ───────────────────────────── */
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #00C2CB, #0099FF) !important;
    border-radius: 4px !important;
}
[data-testid="stSpinner"] > div > div { border-top-color: #00C2CB !important; }

/* ── EXPANDER ─────────────────────────────────────── */
[data-testid="stExpander"] {
    background: #FFFFFF !important;
    border: 1px solid rgba(0,194,203,0.1) !important;
    border-radius: 12px !important;
    overflow: hidden !important;
}
[data-testid="stExpanderDetails"] { background: #FFFFFF !important; }

/* ── HR / CAPTIONS ────────────────────────────────── */
hr { border-color: rgba(0,194,203,0.08) !important; }
.stCaption, [data-testid="stCaptionContainer"] { color: #8494A8 !important; font-size: 11px !important; }
[data-testid="stMarkdownContainer"] p { color: #556070 !important; }

/* ══════════════════════════════════════════════════════
   COMPONENTES REUTILIZABLES NEXA v2
══════════════════════════════════════════════════════ */

/* Cabecera de página */
.nx-page-header {
    padding: 6px 0 24px 0;
    border-bottom: 1px solid rgba(0,194,203,0.1);
    margin-bottom: 28px;
    position: relative;
}
.nx-page-header::after {
    content: '';
    position: absolute;
    bottom: -1px;
    left: 0;
    width: 80px;
    height: 2px;
    background: linear-gradient(90deg, #00C2CB 0%, transparent 100%);
    border-radius: 2px;
}
.nx-page-title {
    font-size: 24px;
    font-weight: 800;
    color: #0A0F1E;
    line-height: 1.2;
    letter-spacing: -0.5px;
}
.nx-page-sub {
    font-size: 14px;
    color: #556070;
    margin-top: 6px;
    line-height: 1.6;
    max-width: 700px;
}
.nx-page-sub strong { color: #2D3A52 !important; }

/* Barra de pasos */
.nx-steps { display: flex; align-items: center; justify-content: center; padding: 8px 0 24px 0; }
.nx-step  { display: flex; flex-direction: column; align-items: center; gap: 5px; min-width: 90px; }
.nx-circle {
    width: 42px; height: 42px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 15px; font-weight: 700;
}
.nx-circle.done   { background: linear-gradient(135deg,#00C2CB,#0099FF); color:#0A0F1E; box-shadow: 0 0 18px rgba(0,194,203,.5); }
.nx-circle.active { background: linear-gradient(135deg,#00C2CB,#0088EE); color:#0A0F1E; box-shadow: 0 0 24px rgba(0,194,203,.7); animation: pulse-glow 2s infinite; }
.nx-circle.idle   { background: #F6F8FB; color: #8494A8; border: 2px solid rgba(10,15,30,0.1); }
.nx-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .7px; color: #8494A8; }
.nx-label.active, .nx-label.done { color: #00C2CB; }
.nx-line { flex:1; height:2px; max-width:68px; border-radius:2px; margin-bottom:18px; }
.nx-line.done { background: linear-gradient(90deg,#00C2CB,#0099FF); }
.nx-line.idle { background: rgba(10,15,30,0.08); }

@keyframes pulse-glow {
    0%,100% { box-shadow: 0 0 24px rgba(0,194,203,.7); }
    50%      { box-shadow: 0 0 36px rgba(0,194,203,1), 0 0 60px rgba(0,194,203,.3); }
}

/* Encabezados de sección */
.nx-section {
    font-size: 11px; font-weight: 700; color: #00C2CB;
    text-transform: uppercase; letter-spacing: 1.5px;
    margin: 24px 0 12px 0;
    display: flex; align-items: center; gap: 10px;
}
.nx-section::after {
    content: ''; flex:1; height:1px;
    background: linear-gradient(90deg, rgba(0,194,203,.3) 0%, transparent 100%);
}

/* Estado vacío */
.nx-empty {
    text-align: center; padding: 56px 24px;
    background: #F6F8FB;
    border-radius: 20px;
    border: 2px dashed rgba(0,194,203,0.16); margin-top: 12px;
}
.nx-empty-icon { font-size: 52px; margin-bottom: 14px; filter: grayscale(0.3); }
.nx-empty-text { font-size: 16px; color: #556070; font-weight: 600; }
.nx-empty-sub  { font-size: 13px; color: #8494A8; margin-top: 8px; }

/* Tarjeta de éxito */
.nx-success-card {
    background: linear-gradient(135deg, #F6F8FB 0%, #FFFFFF 50%, #F6F8FB 100%);
    border: 1px solid rgba(0,194,203,0.28);
    border-radius: 20px; padding: 36px; text-align: center; margin: 14px 0;
    box-shadow: 0 8px 40px rgba(0,194,203,0.08);
    animation: fadeInUp 0.4s ease;
}
.nx-success-icon  { font-size: 52px; margin-bottom: 12px; }
.nx-success-title { font-size: 22px; font-weight: 800; color: #00C2CB; margin-bottom: 8px; letter-spacing: -0.3px; }
.nx-success-sub   { font-size: 14px; color: #556070; }
.nx-success-sub strong { color: #2D3A52 !important; }

@keyframes fadeInUp {
    from { opacity:0; transform:translateY(20px); }
    to   { opacity:1; transform:translateY(0); }
}

/* Sección exportar */
.nx-export-bar {
    border-top: 1px solid rgba(0,194,203,0.1);
    background: linear-gradient(0deg, rgba(246,248,251,0.95) 0%, transparent 100%);
    padding: 20px 0 4px 0;
    margin-top: 20px;
}
.nx-export-label {
    font-size: 11px; font-weight: 700; color: #556070;
    text-transform: uppercase; letter-spacing: 1.2px;
    margin-bottom: 12px;
}

/* Próximamente */
.nx-coming-soon {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; padding: 100px 24px; text-align: center;
}
.nx-cs-icon  { font-size: 60px; margin-bottom: 20px; }
.nx-cs-title { font-size: 22px; font-weight: 800; color: #2D3A52; margin-bottom: 10px; letter-spacing: -0.3px; }
.nx-cs-sub   { font-size: 14px; color: #8494A8; line-height: 1.6; }
.nx-cs-badge {
    display: inline-flex; align-items: center; gap: 6px; margin-top: 20px;
    font-size: 11px; font-weight: 700; color: #00C2CB;
    background: rgba(0,194,203,0.08); border: 1px solid rgba(0,194,203,0.22);
    padding: 6px 18px; border-radius: 20px; letter-spacing: 1px;
    text-transform: uppercase;
}
.nx-cs-badge::before { content: '⚡'; }

/* ── SIDEBAR INTERNO ──────────────────────────────── */
.nx-sidebar {
    padding: 20px 16px;
    height: 100%;
    display: flex;
    flex-direction: column;
}
.nx-sidebar-logo {
    display: flex; align-items: center; justify-content: center;
    padding: 20px 16px 24px;
    border-bottom: 1px solid rgba(255,255,255,0.06);
    margin-bottom: 20px;
}
.nx-sidebar-version {
    font-size: 10px; font-weight: 600; color: rgba(255,255,255,0.3);
    text-transform: uppercase; letter-spacing: 1.2px;
    text-align: center; margin-top: 6px;
}
.nx-sidebar-badge {
    display: inline-flex; align-items: center; gap: 5px;
    background: rgba(0,194,203,0.12);
    border: 1px solid rgba(0,194,203,0.2);
    border-radius: 20px; padding: 3px 12px;
    font-size: 10px; font-weight: 700; color: #00C2CB;
    letter-spacing: 0.8px; text-transform: uppercase;
}
.nx-stat-box {
    background: rgba(0,194,203,0.05);
    border: 1px solid rgba(0,194,203,0.1);
    border-radius: 12px; padding: 12px 14px; margin-bottom: 8px;
    display: flex; align-items: center; gap: 10px;
}
.nx-stat-icon { font-size: 20px; }
.nx-stat-label { font-size: 11px; color: rgba(255,255,255,0.4); font-weight: 500; }
.nx-stat-value { font-size: 16px; font-weight: 700; color: #FFFFFF; }

/* ── TOOL CARDS (MODO SELECCIÓN) ──────────────────── */
.nx-tool-card {
    background: #FFFFFF;
    border: 1px solid rgba(10,15,30,0.08);
    border-radius: 16px; padding: 20px 16px;
    cursor: pointer; transition: all .2s ease;
    text-align: center; min-height: 120px;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    box-shadow: 0 1px 3px rgba(10,15,30,0.06);
}
.nx-tool-card:hover {
    border-color: rgba(0,194,203,0.35);
    transform: translateY(-3px);
    box-shadow: 0 12px 40px rgba(10,15,30,0.12);
}
.nx-tool-card.active {
    border-color: #00C2CB;
    background: linear-gradient(145deg, rgba(0,194,203,0.08), rgba(0,194,203,0.04));
    box-shadow: 0 8px 32px rgba(0,194,203,0.15);
}

/* ── PANEL LATERAL DERECHO ────────────────────────── */
.nx-right-panel {
    background: #FFFFFF;
    border: 1px solid rgba(10,15,30,0.08);
    border-radius: 16px; padding: 20px 16px;
    position: sticky; top: 80px;
    box-shadow: 0 1px 3px rgba(10,15,30,0.06);
}
.nx-panel-title {
    font-size: 13px; font-weight: 700; color: #0A0F1E;
    margin-bottom: 16px; display: flex; align-items: center; gap: 8px;
}

/* ── MINIATURAS PDF ───────────────────────────────── */
.nx-thumb {
    border-radius: 10px; overflow: hidden;
    transition: transform .2s ease;
}
.nx-thumb:hover { transform: scale(1.02); }

/* ── LOGIN ────────────────────────────────────────── */
.nx-login-card {
    background: #FFFFFF;
    border: 1px solid rgba(0,194,203,0.18);
    border-radius: 24px; padding: 48px 40px; text-align: center;
    box-shadow: 0 20px 80px rgba(10,15,30,0.12), 0 0 0 1px rgba(0,194,203,0.05);
}

</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  IMPORTS OPCIONALES
# ══════════════════════════════════════════════════════
try:
    import fitz
    FITZ_OK = True
except ImportError:
    FITZ_OK = False

# ══════════════════════════════════════════════════════
#  HELPERS DE MINIATURAS
# ══════════════════════════════════════════════════════
def _gen_thumbs(pdf_bytes: bytes, scale: float = 1.2):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    mat = fitz.Matrix(scale, scale)
    result = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        result.append(base64.b64encode(pix.tobytes("png")).decode())
    n = len(doc)
    doc.close()
    return result, n


def _ensure_thumbs(prefix: str, pdf_bytes: bytes, sel_key: str = None):
    sig = hashlib.md5(
        pdf_bytes[:2048] + len(pdf_bytes).to_bytes(8, "big") + pdf_bytes[-2048:]
    ).hexdigest()[:12]
    if st.session_state.get(f"{prefix}_sig") != sig:
        with st.spinner("Generando miniaturas…"):
            thumbs, n = _gen_thumbs(pdf_bytes)
        st.session_state[f"{prefix}_sig"]    = sig
        st.session_state[f"{prefix}_thumbs"] = thumbs
        st.session_state[f"{prefix}_n"]      = n
        if sel_key:
            st.session_state[sel_key] = set()
    return st.session_state[f"{prefix}_thumbs"], st.session_state[f"{prefix}_n"]


def _thumb_card(b64: str, page_num: int,
                selected: bool = False,
                mode: str = "delete",
                badge_label: str = None,
                badge_bg: str = "rgba(27,159,216,0.85)",
                badge_fg: str = "#fff") -> str:
    if selected and mode == "delete":
        border  = "2px solid rgba(231,76,60,0.9)"
        overlay = ('<div style="position:absolute;inset:0;background:rgba(231,76,60,0.18);'
                   'border-radius:8px;pointer-events:none;"></div>')
        pin = ('<div style="position:absolute;top:5px;right:5px;background:rgba(231,76,60,0.95);'
               'color:#fff;width:22px;height:22px;border-radius:50%;display:flex;'
               'align-items:center;justify-content:center;font-size:13px;'
               'font-weight:700;pointer-events:none;">✕</div>')
    elif selected and mode == "split":
        border  = "2px solid rgba(255,210,0,0.9)"
        overlay = ('<div style="position:absolute;inset:0;background:rgba(255,210,0,0.1);'
                   'border-radius:8px;pointer-events:none;"></div>')
        pin = ('<div style="position:absolute;top:5px;right:5px;background:rgba(255,210,0,0.9);'
               'color:#000;width:22px;height:22px;border-radius:50%;display:flex;'
               'align-items:center;justify-content:center;font-size:12px;'
               'font-weight:700;pointer-events:none;">✂</div>')
    else:
        border, overlay, pin = "2px solid rgba(27,159,216,0.2)", "", ""

    badge = ""
    if badge_label:
        badge = (f'<div style="position:absolute;top:5px;left:5px;background:{badge_bg};'
                 f'color:{badge_fg};padding:1px 7px;border-radius:8px;font-size:9px;'
                 f'font-weight:700;pointer-events:none;">{badge_label}</div>')

    return (f'<div style="border:{border};border-radius:12px;padding:6px;'
            f'background:linear-gradient(145deg,#09172A,#0A1626);'
            f'position:relative;margin-bottom:4px;transition:transform .2s;">'
            f'{overlay}{pin}{badge}'
            f'<img src="data:image/png;base64,{b64}" '
            f'style="width:100%;border-radius:7px;display:block;" draggable="false"/>'
            f'<div style="text-align:center;font-size:10px;color:#2E5878;'
            f'margin-top:5px;font-weight:600;letter-spacing:0.3px;">Pág. {page_num}</div>'
            f'</div>')


def _parse_fabric_color(color_str):
    if not color_str or color_str in ("", "transparent"):
        return None
    cs = str(color_str).strip()
    try:
        if cs.startswith("#") and len(cs) >= 7:
            return (int(cs[1:3],16)/255, int(cs[3:5],16)/255, int(cs[5:7],16)/255)
        if "rgba" in cs:
            p = cs[cs.index("(")+1:cs.rindex(")")].split(",")
            if float(p[3].strip()) < 0.03:
                return None
            return (float(p[0])/255, float(p[1])/255, float(p[2].strip())/255)
        if "rgb" in cs:
            p = cs[cs.index("(")+1:cs.rindex(")")].split(",")
            return (float(p[0])/255, float(p[1])/255, float(p[2].strip())/255)
    except Exception:
        pass
    return (0, 0, 0)


def _extract_path_pts(obj, sx, sy):
    pts = []
    path_data = obj.get("path", [])
    off_x = obj.get("left", 0)
    off_y = obj.get("top", 0)
    po = obj.get("pathOffset")
    if isinstance(po, dict):
        off_x += po.get("x", 0)
        off_y += po.get("y", 0)
    for cmd in path_data:
        if not cmd:
            continue
        t = cmd[0]
        if t in ("M", "L") and len(cmd) >= 3:
            pts.append(fitz.Point((off_x + cmd[1]) * sx, (off_y + cmd[2]) * sy))
        elif t == "Q" and len(cmd) >= 5:
            pts.append(fitz.Point((off_x + cmd[3]) * sx, (off_y + cmd[4]) * sy))
        elif t == "C" and len(cmd) >= 7:
            pts.append(fitz.Point((off_x + cmd[5]) * sx, (off_y + cmd[6]) * sy))
    return pts

# ══════════════════════════════════════════════════════
#  CANVAS HTML — EDITOR PDF AVANZADO (estilo Adobe)
# ══════════════════════════════════════════════════════
_CANVAS_TMPL = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    background: #060F1D;
    font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", sans-serif;
    overflow: hidden; user-select: none;
}

/* ── TOOLBAR PRINCIPAL ── */
#toolbar {
    display: flex; align-items: center; gap: 4px;
    padding: 6px 10px;
    background: linear-gradient(180deg, #0E1E30 0%, #0A1828 100%);
    border-bottom: 1px solid #152535;
    flex-wrap: wrap; min-height: 44px;
}

/* ── GRUPOS DE HERRAMIENTAS ── */
.tb-group {
    display: flex; align-items: center; gap: 3px;
    padding: 2px 6px;
    background: rgba(27,159,216,0.04);
    border: 1px solid rgba(27,159,216,0.1);
    border-radius: 8px;
}
.tb-group-label {
    font-size: 9px; color: #2A4A6A; text-transform: uppercase;
    letter-spacing: 0.8px; padding: 0 4px 2px;
    display: block; width: 100%; text-align: center;
}

.tb-btn {
    background: transparent; border: 1px solid transparent;
    color: #8AAEC8; border-radius: 7px; padding: 5px 9px;
    font-size: 12px; cursor: pointer; white-space: nowrap;
    transition: all .15s ease; display: flex; align-items: center; gap: 4px;
    font-family: inherit;
}
.tb-btn:hover { background: rgba(27,159,216,0.1); color: #C8E4F0; border-color: rgba(27,159,216,0.2); }
.tb-btn.on {
    background: linear-gradient(135deg, #1B9FD8, #1580B8);
    border-color: #1B9FD8; color: #fff;
    box-shadow: 0 2px 10px rgba(27,159,216,0.3);
}
.tb-btn.em { background: #183650; border-color: #1B9FD8; color: #C8E4F0; font-weight: 700; }
.tb-btn.danger:hover { background: rgba(231,76,60,0.15); border-color: rgba(231,76,60,0.3); color: #E74C3C; }

.sep { width: 1px; background: rgba(27,159,216,0.15); height: 24px; margin: 0 4px; flex-shrink: 0; }

/* ── TOOLBAR TEXTO ── */
#tb-text {
    display: none; align-items: center; gap: 5px; padding: 5px 10px;
    background: #09192B; border-bottom: 1px solid #152535; flex-wrap: wrap;
}

/* ── PALETA ── */
.pal { display: flex; gap: 3px; align-items: center; }
.dot {
    width: 16px; height: 16px; border-radius: 50%; cursor: pointer; flex-shrink: 0;
    border: 2px solid transparent; transition: transform .12s, border-color .12s;
}
.dot:hover { transform: scale(1.35); }
.dot.sel { border-color: #fff; box-shadow: 0 0 0 2px #1B9FD8; }
.dot.wh { border-color: #555; }

.inp {
    width: 44px; background: #07111C; color: #C8E4F0;
    border: 1px solid rgba(27,159,216,0.25); border-radius: 6px;
    font-size: 12px; padding: 3px 5px; font-family: inherit;
}
.sel-f {
    background: #07111C; color: #C8E4F0;
    border: 1px solid rgba(27,159,216,0.25); border-radius: 6px;
    font-size: 12px; padding: 3px 6px; cursor: pointer; font-family: inherit;
}

/* ── INPUT TEXTO ── */
#txt-input-area {
    display: none; padding: 6px 10px;
    background: #09192B; border-bottom: 1px solid #1B4060;
    align-items: center; gap: 6px;
}

/* ── ZOOM CONTROL ── */
#zoom-lbl {
    font-size: 11px; color: #C8E4F0; white-space: nowrap;
    min-width: 40px; text-align: center; font-weight: 600;
}

/* ── CANVAS WRAPPER ── */
#cvwrap { position: relative; overflow: hidden; width: 100%; }
canvas  { display: block; cursor: default; }

/* ── STATUS BAR ── */
#status-bar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 4px 12px; background: #071018;
    border-top: 1px solid #0D1E2E; min-height: 24px;
}
#status     { font-size: 10px; color: #2A4A6A; }
#coord-info { font-size: 10px; color: #1A3A55; font-family: monospace; }

/* ── HIGHLIGHT ANNOTATION ── */
.hl-color-btn {
    width: 20px; height: 20px; border-radius: 4px; cursor: pointer;
    border: 2px solid transparent; transition: all .1s;
    flex-shrink: 0;
}
.hl-color-btn:hover { transform: scale(1.2); }
.hl-color-btn.sel   { border-color: #fff; box-shadow: 0 0 0 2px #1B9FD8; }

</style>
</head>
<body>

<!-- ── TOOLBAR PRINCIPAL ── -->
<div id="toolbar">

    <!-- Selección & Texto -->
    <div class="tb-group">
        <button class="tb-btn on" id="b-tf"   onclick="T('tf')"  title="Mover/Seleccionar (V)">&#9654; Selec.</button>
        <button class="tb-btn"   id="b-text"  onclick="T('text')"title="Añadir texto (T)">&#9998; Texto</button>
    </div>

    <div class="sep"></div>

    <!-- Formas -->
    <div class="tb-group">
        <button class="tb-btn" id="b-rect"  onclick="T('rect')" title="Rectángulo (R)">&#9633; Rect</button>
        <button class="tb-btn" id="b-ell"   onclick="T('ell')"  title="Elipse (E)">&#9711; Elipse</button>
        <button class="tb-btn" id="b-line"  onclick="T('line')" title="Línea (L)">&#8599; Línea</button>
        <button class="tb-btn" id="b-arrow" onclick="T('arrow')"title="Flecha">&#10152; Flecha</button>
    </div>

    <div class="sep"></div>

    <!-- Anotaciones -->
    <div class="tb-group">
        <button class="tb-btn" id="b-hl"     onclick="T('hl')"    title="Resaltar (H)">&#9678; Resaltar</button>
        <button class="tb-btn" id="b-redact" onclick="T('redact')"title="Redactar/Tachar">&#9646; Redactar</button>
        <button class="tb-btn" id="b-stamp"  onclick="T('stamp')" title="Sello">&#10003; Sello</button>
    </div>

    <div class="sep"></div>

    <!-- Color & Trazo -->
    <span class="lbl" style="font-size:10px;color:#2A4A6A;white-space:nowrap;">Color:</span>
    <div class="pal" id="pal1"></div>
    <div class="sep"></div>
    <span class="lbl" style="font-size:10px;color:#2A4A6A;">Trazo:</span>
    <input type="number" class="inp" id="sw" min="0.5" max="30" value="2" step="0.5">
    <label style="display:flex;align-items:center;gap:4px;cursor:pointer;font-size:11px;color:#8AAEC8;">
        <input type="checkbox" id="useFill" style="accent-color:#1B9FD8;cursor:pointer"> Relleno
    </label>
    <label style="display:flex;align-items:center;gap:4px;cursor:pointer;font-size:11px;color:#8AAEC8;">
        <input type="number" class="inp" id="opacity-inp" min="10" max="100" value="100" style="width:44px;"> %
    </label>

    <div class="sep"></div>

    <!-- Zoom -->
    <button class="tb-btn" onclick="zoomOut()" title="Alejar">&#128269;-</button>
    <span id="zoom-lbl">100%</span>
    <button class="tb-btn" onclick="zoomIn()"  title="Acercar">&#128269;+</button>
    <button class="tb-btn" onclick="setZoom(1)"title="Reset zoom">1:1</button>

    <div class="sep"></div>

    <!-- Acciones -->
    <button id="undo-btn" class="tb-btn" onclick="doUndo()" title="Deshacer (Ctrl+Z)">&#8634; Deshacer</button>
    <button id="del-btn"  class="tb-btn danger" onclick="delSel()" style="display:none;" title="Eliminar (Del)">&#10005; Borrar</button>

    <div class="sep"></div>

    <!-- Stamps coloreados -->
    <button class="tb-btn" onclick="addStamp('APROBADO','#27AE60')">&#10003; Aprobado</button>
    <button class="tb-btn" onclick="addStamp('CONFIDENCIAL','#E74C3C')">&#9679; Confidencial</button>
    <button class="tb-btn" onclick="addStamp('BORRADOR','#F39C12')">&#9999; Borrador</button>

    <div class="sep"></div>
    <button id="save-btn" class="tb-btn on" onclick="saveDat()" title="Guardar en Streamlit">&#128190; Guardar página</button>
</div>

<!-- ── TOOLBAR TEXTO ── -->
<div id="tb-text">
    <select class="sel-f" id="ff">
        <option value="Arial">Inter / Arial</option>
        <option value="Times New Roman">Times New Roman</option>
        <option value="Helvetica">Helvetica</option>
        <option value="Courier New">Courier New</option>
        <option value="Georgia">Georgia</option>
    </select>
    <input type="number" class="inp" id="fs" min="6" max="288" value="16" style="width:50px">
    <button class="tb-btn" id="b-bold" onclick="togBold()" title="Negrita"><b>N</b></button>
    <button class="tb-btn" id="b-ital" onclick="togItal()" title="Cursiva"><i>K</i></button>
    <button class="tb-btn" id="b-under" onclick="togUnder()" title="Subrayado"><u>S</u></button>
    <div class="sep"></div>
    <span style="font-size:10px;color:#4A7A9C;">Color texto:</span>
    <div class="pal" id="pal2"></div>
</div>

<!-- ── INPUT TEXTO ── -->
<div id="txt-input-area">
    <span style="font-size:11px;color:#4A7A9C;white-space:nowrap;">&#128204; Texto:</span>
    <input type="text" id="txt-inp" placeholder="Escribe el texto aquí..."
        onkeydown="if(event.key==='Enter'){commitTxt();}else if(event.key==='Escape'){cancelTxt();}"
        style="flex:1;background:#07111C;color:#C8E4F0;border:1px solid #1B9FD8;
               border-radius:7px;font-size:13px;padding:5px 10px;outline:none;min-width:200px;
               font-family:inherit;">
    <button onclick="commitTxt()"
        style="background:linear-gradient(135deg,#1B9FD8,#1580B8);border:none;color:#fff;
               border-radius:7px;padding:5px 16px;font-size:12px;font-weight:700;cursor:pointer;">
        Agregar
    </button>
    <button onclick="cancelTxt()"
        style="background:#07111C;border:1px solid #1B4060;color:#C8E4F0;
               border-radius:7px;padding:4px 10px;font-size:12px;cursor:pointer;">
        &#10005;
    </button>
</div>

<!-- ── CANVAS ── -->
<div id="cvwrap" style="width:100%;overflow:hidden;">
    <canvas id="cv" width="___CW___" height="___CH___"
        style="display:block;transform-origin:top left;"></canvas>
</div>

<!-- ── STATUS BAR ── -->
<div id="status-bar">
    <span id="status">Herramienta: Mover/Seleccionar</span>
    <span id="coord-info">X: 0  Y: 0</span>
</div>

<script>
var CW=___CW___,CH=___CH___,HR=4;
var zoom=1, ZOOMS=[0.25,0.5,0.75,1,1.25,1.5,1.75,2,2.5,3];
var cv=document.getElementById('cv'), ctx=cv.getContext('2d');
var els=___INIT___, history_=[];
var tool='tf', sel=-1, drag=null, drawing=null;
var aCol='#1B9FD8', tCol='#000000', bold=false, ital=false, underline=false;
var pendTx=null, pendTy=null, pendIdx=null;
var hlColors=['rgba(255,235,59,0.4)','rgba(76,175,80,0.4)','rgba(33,150,243,0.4)','rgba(255,87,34,0.4)'];
var curHlColor=hlColors[0];

// ── Background ─────────────────────────────────────────────────────────────
var bg=new Image(); bg.onload=draw;
bg.src='data:image/png;base64,___BG___';

// ── Paleta de colores ───────────────────────────────────────────────────────
var COLS=[
    '#000000','#FFFFFF','#E74C3C','#1B9FD8','#27AE60','#F1C40F',
    '#E67E22','#9B59B6','#95A5A6','#00BCD4','#E91E63','#795548',
    '#FF5722','#3F51B5','#009688','#CDDC39'
];

function mkPal(id, cb) {
    var p=document.getElementById(id);
    COLS.forEach(function(c){
        var d=document.createElement('div');
        d.className='dot'+(c==='#FFFFFF'?' wh':'');
        d.style.background=c; d.dataset.c=c;
        d.onclick=function(){ cb(c); selPal(id,c); };
        p.appendChild(d);
    });
}
function selPal(id,col){
    document.getElementById(id).querySelectorAll('.dot').forEach(function(d){
        d.classList.toggle('sel', d.dataset.c===col);
    });
}
mkPal('pal1',function(c){
    aCol=c;
    if(sel>=0&&sel<els.length){
        var e=els[sel];
        if(e.type!=='text'){ e.sc=c; if(document.getElementById('useFill').checked) e.fc=opq(c); draw(); }
    }
});
mkPal('pal2',function(c){
    tCol=c;
    if(sel>=0&&sel<els.length&&els[sel].type==='text'){ els[sel].tc=c; draw(); }
});
selPal('pal1','#1B9FD8'); selPal('pal2','#000000');

function opq(c){
    if(!c||c.length<7) return '';
    var r=parseInt(c.slice(1,3),16),g=parseInt(c.slice(3,5),16),b=parseInt(c.slice(5,7),16);
    return 'rgba('+r+','+g+','+b+',1)';
}
function getFill(){ return document.getElementById('useFill').checked ? opq(aCol) : ''; }
function getSW(){   return parseFloat(document.getElementById('sw').value)||2; }
function getFF(){   return document.getElementById('ff').value; }
function getFS(){   return parseInt(document.getElementById('fs').value)||16; }
function getOp(){   return (parseInt(document.getElementById('opacity-inp').value)||100)/100; }

// ── History (Undo) ────────────────────────────────────────────────────────
function pushHistory(){ history_.push(JSON.stringify(els)); if(history_.length>50) history_.shift(); }
function doUndo(){
    if(history_.length>0){ els=JSON.parse(history_.pop()); sel=-1; draw(); }
}

// ── Texto bold/ital/under ──────────────────────────────────────────────────
function togBold(){
    bold=!bold; document.getElementById('b-bold').classList.toggle('em',bold);
    if(sel>=0&&els[sel]&&els[sel].type==='text'){ els[sel].bold=bold; remeas(els[sel]); draw(); }
}
function togItal(){
    ital=!ital; document.getElementById('b-ital').classList.toggle('em',ital);
    if(sel>=0&&els[sel]&&els[sel].type==='text'){ els[sel].ital=ital; remeas(els[sel]); draw(); }
}
function togUnder(){
    underline=!underline; document.getElementById('b-under').classList.toggle('em',underline);
    if(sel>=0&&els[sel]&&els[sel].type==='text'){ els[sel].under=underline; draw(); }
}
function txF(e){ return (e.bold?'bold ':'')+( e.ital?'italic ':'')+e.fs+'px '+e.ff; }
function remeas(e){
    ctx.save(); ctx.font=txF(e);
    var m=ctx.measureText(e.txt||' ');
    e.w=Math.max(m.width,10); e.h=e.fs*1.4;
    ctx.restore();
}

// ── Stamps ─────────────────────────────────────────────────────────────────
function addStamp(txt, color){
    pushHistory();
    var x=CW/2-80, y=CH/2-20;
    var st_={type:'stamp',x:x,y:y,w:160,h:44,txt:txt,color:color,opacity:0.85};
    els.push(st_); sel=els.length-1;
    document.getElementById('del-btn').style.display='';
    draw();
}

// ── Tool setter ────────────────────────────────────────────────────────────
var TIPS={
    tf:'Clic para seleccionar, arrastra para mover/redimensionar',
    rect:'Clic y arrastra: rectángulo',
    ell:'Clic y arrastra: elipse',
    line:'Clic y arrastra: línea',
    arrow:'Clic y arrastra: flecha',
    text:'Clic en el canvas para agregar texto; doble clic en texto existente para editar',
    hl:'Clic y arrastra: resaltado',
    redact:'Clic y arrastra: redacción/tachado',
    stamp:'Clic para insertar sello personalizado'
};
function T(t){
    tool=t;
    document.querySelectorAll('.tb-btn').forEach(function(b){ b.classList.remove('on'); });
    var b=document.getElementById('b-'+t); if(b) b.classList.add('on');
    sel=-1; document.getElementById('del-btn').style.display='none';
    document.getElementById('tb-text').style.display=(t==='text')?'flex':'none';
    document.getElementById('status').textContent='Herramienta: '+(TIPS[t]||t);
    cv.style.cursor=(t==='tf')?'default':(t==='text'?'text':'crosshair');
    draw();
}

// ── Handles ────────────────────────────────────────────────────────────────
function gH(e){
    if(e.type==='line'||e.type==='arrow') return [{id:'p1',x:e.x1,y:e.y1,cur:'move'},{id:'p2',x:e.x2,y:e.y2,cur:'move'}];
    var x=e.x,y=e.y,w=e.w||20,h=e.h||(e.fs?e.fs*1.4:20)||20;
    return [
        {id:'tl',x:x,    y:y,    cur:'nw-resize'},
        {id:'tm',x:x+w/2,y:y,    cur:'n-resize'},
        {id:'tr',x:x+w,  y:y,    cur:'ne-resize'},
        {id:'ml',x:x,    y:y+h/2,cur:'w-resize'},
        {id:'mr',x:x+w,  y:y+h/2,cur:'e-resize'},
        {id:'bl',x:x,    y:y+h,  cur:'sw-resize'},
        {id:'bm',x:x+w/2,y:y+h,  cur:'s-resize'},
        {id:'br',x:x+w,  y:y+h,  cur:'se-resize'}
    ];
}
function hitH(e,mx,my){
    var hs=gH(e);
    for(var i=0;i<hs.length;i++){ var h=hs[i]; if(Math.abs(mx-h.x)<=6&&Math.abs(my-h.y)<=6) return h; }
    return null;
}
function hitE(e,mx,my){
    if(e._hid) return false;
    var w=e.w||20, h=e.h||(e.fs?e.fs*1.4:20)||20;
    if(e.type==='rect'||e.type==='hl'||e.type==='redact'||e.type==='stamp')
        return mx>=e.x&&mx<=e.x+w&&my>=e.y&&my<=e.y+h;
    if(e.type==='ell'){
        var cx=e.x+w/2,cy=e.y+h/2,rx=Math.abs(w/2)+6,ry=Math.abs(h/2)+6;
        return (mx-cx)*(mx-cx)/(rx*rx)+(my-cy)*(my-cy)/(ry*ry)<=1;
    }
    if(e.type==='line'||e.type==='arrow'){
        var dx=e.x2-e.x1,dy=e.y2-e.y1,len=Math.sqrt(dx*dx+dy*dy);
        if(len<1) return Math.hypot(mx-e.x1,my-e.y1)<10;
        var t=Math.max(0,Math.min(1,((mx-e.x1)*dx+(my-e.y1)*dy)/(len*len)));
        return Math.hypot(mx-(e.x1+t*dx),my-(e.y1+t*dy))<8;
    }
    if(e.type==='text') return mx>=e.x&&mx<=e.x+w&&my>=e.y&&my<=e.y+h;
    return false;
}

// ── Draw ───────────────────────────────────────────────────────────────────
function dE(e,isSel,isDraft){
    if(e._hid) return;
    ctx.save();
    var op=e.opacity!==undefined?e.opacity:1;
    ctx.globalAlpha=op;
    if(isDraft) ctx.setLineDash([5,4]);

    if(e.type==='rect'){
        if(e.fc){ ctx.fillStyle=e.fc; ctx.fillRect(e.x,e.y,e.w,e.h); }
        ctx.strokeStyle=e.sc||'#1B9FD8'; ctx.lineWidth=e.sw||2; ctx.strokeRect(e.x,e.y,e.w,e.h);
    } else if(e.type==='ell'){
        ctx.beginPath();
        ctx.ellipse(e.x+(e.w||20)/2,e.y+(e.h||20)/2,
            Math.max(1,Math.abs((e.w||20)/2)),Math.max(1,Math.abs((e.h||20)/2)),0,0,Math.PI*2);
        if(e.fc){ ctx.fillStyle=e.fc; ctx.fill(); }
        ctx.strokeStyle=e.sc||'#1B9FD8'; ctx.lineWidth=e.sw||2; ctx.stroke();
    } else if(e.type==='line'){
        ctx.beginPath(); ctx.moveTo(e.x1,e.y1); ctx.lineTo(e.x2,e.y2);
        ctx.strokeStyle=e.sc||'#1B9FD8'; ctx.lineWidth=e.sw||2; ctx.stroke();
    } else if(e.type==='arrow'){
        // Línea
        ctx.beginPath(); ctx.moveTo(e.x1,e.y1); ctx.lineTo(e.x2,e.y2);
        ctx.strokeStyle=e.sc||'#1B9FD8'; ctx.lineWidth=e.sw||2; ctx.stroke();
        // Punta
        var angle=Math.atan2(e.y2-e.y1,e.x2-e.x1);
        var al=14+((e.sw||2)*2);
        ctx.beginPath();
        ctx.moveTo(e.x2,e.y2);
        ctx.lineTo(e.x2-al*Math.cos(angle-0.45),e.y2-al*Math.sin(angle-0.45));
        ctx.lineTo(e.x2-al*Math.cos(angle+0.45),e.y2-al*Math.sin(angle+0.45));
        ctx.closePath();
        ctx.fillStyle=e.sc||'#1B9FD8'; ctx.fill();
    } else if(e.type==='text'){
        ctx.font=txF(e); ctx.fillStyle=e.tc||'#000';
        ctx.textBaseline='top'; ctx.fillText(e.txt||'',e.x,e.y);
        if(e.under){
            var th=e.h||e.fs*1.4;
            ctx.beginPath(); ctx.moveTo(e.x,e.y+e.fs+2); ctx.lineTo(e.x+e.w,e.y+e.fs+2);
            ctx.strokeStyle=e.tc||'#000'; ctx.lineWidth=1.5; ctx.stroke();
        }
    } else if(e.type==='hl'){
        ctx.fillStyle=e.color||'rgba(255,235,59,0.4)';
        ctx.fillRect(e.x,e.y,e.w,e.h);
    } else if(e.type==='redact'){
        // Fondo negro opaco
        ctx.fillStyle=e.color||'#000';
        ctx.fillRect(e.x,e.y,e.w,e.h);
    } else if(e.type==='stamp'){
        // Borde exterior
        ctx.strokeStyle=e.color||'#E74C3C';
        ctx.lineWidth=3; ctx.setLineDash([]);
        ctx.strokeRect(e.x,e.y,e.w,e.h);
        // Fondo semitransparente
        ctx.fillStyle=(e.color||'#E74C3C').replace(')',',0.1)').replace('rgb','rgba')||'rgba(231,76,60,0.1)';
        try{ ctx.fillRect(e.x,e.y,e.w,e.h); } catch(x){}
        // Texto
        ctx.font='bold '+Math.min(22,e.h*0.55)+'px Arial';
        ctx.fillStyle=e.color||'#E74C3C';
        ctx.textAlign='center'; ctx.textBaseline='middle';
        ctx.fillText(e.txt||'SELLO',e.x+e.w/2,e.y+e.h/2);
        ctx.textAlign='left'; ctx.textBaseline='top';
    }

    ctx.restore();

    if(isSel){
        ctx.save();
        ctx.strokeStyle='#1B9FD8'; ctx.lineWidth=1.5; ctx.setLineDash([5,4]);
        if(e.type==='line'||e.type==='arrow'){
            ctx.beginPath(); ctx.moveTo(e.x1,e.y1); ctx.lineTo(e.x2,e.y2); ctx.stroke();
        } else {
            var bw=e.w||20, bh=e.h||(e.fs?e.fs*1.4:20)||20;
            ctx.strokeRect(e.x-3,e.y-3,bw+6,bh+6);
        }
        ctx.setLineDash([]);
        var hs=gH(e);
        for(var i=0;i<hs.length;i++){
            var hx=hs[i].x,hy=hs[i].y,hs2=4;
            ctx.fillStyle='#1B9FD8'; ctx.fillRect(hx-hs2,hy-hs2,hs2*2,hs2*2);
            ctx.strokeStyle='#fff'; ctx.lineWidth=1; ctx.strokeRect(hx-hs2,hy-hs2,hs2*2,hs2*2);
        }
        ctx.restore();
    }
}

function draw(){
    ctx.clearRect(0,0,CW,CH);
    if(bg.complete&&bg.naturalWidth>0) ctx.drawImage(bg,0,0,CW,CH);
    else{ ctx.fillStyle='#f5f5f5'; ctx.fillRect(0,0,CW,CH); }
    for(var i=0;i<els.length;i++) dE(els[i],i===sel,false);
    if(drawing) dE(drawing,false,true);
    if(pendTx!==null){
        ctx.save(); ctx.strokeStyle='#1B9FD8'; ctx.lineWidth=2;
        ctx.beginPath(); ctx.moveTo(pendTx-12,pendTy); ctx.lineTo(pendTx+12,pendTy);
        ctx.moveTo(pendTx,pendTy-12); ctx.lineTo(pendTx,pendTy+12);
        ctx.stroke();
        ctx.beginPath(); ctx.arc(pendTx,pendTy,4,0,Math.PI*2);
        ctx.fillStyle='#1B9FD8'; ctx.fill();
        ctx.restore();
    }
}

// Init text
for(var _i=0;_i<els.length;_i++){ if(els[_i].type==='text') remeas(els[_i]); }
applyZoom();

// ── XY helper ──────────────────────────────────────────────────────────────
function xy(e){
    var r=cv.getBoundingClientRect();
    return{x:(e.clientX-r.left)/zoom*(CW/r.width),y:(e.clientY-r.top)/zoom*(CH/r.height)};
}
function applyZoom(){
    cv.style.transform='scale('+zoom+')';
    cv.style.transformOrigin='top left';
    var wrap=document.getElementById('cvwrap');
    wrap.style.width=(CW*zoom)+'px';
    wrap.style.height=(CH*zoom)+'px';
    wrap.style.overflow=zoom>1?'auto':'hidden';
    document.getElementById('zoom-lbl').textContent=Math.round(zoom*100)+'%';
}
function setZoom(z){ zoom=z; applyZoom(); }
function zoomIn(){
    var zi=ZOOMS.indexOf(zoom);
    if(zi<0){ zoom=1; } else if(zi<ZOOMS.length-1){ zoom=ZOOMS[zi+1]; }
    applyZoom();
}
function zoomOut(){
    var zi=ZOOMS.indexOf(zoom);
    if(zi<0){ zoom=1; } else if(zi>0){ zoom=ZOOMS[zi-1]; }
    applyZoom();
}

// ── Mouse events ───────────────────────────────────────────────────────────
cv.addEventListener('mousedown',function(e){
    var p=xy(e);
    if(tool==='tf'){
        if(sel>=0&&sel<els.length){
            var hh=hitH(els[sel],p.x,p.y);
            if(hh){ drag={t:'r',idx:sel,hid:hh.id,ox:p.x,oy:p.y,s0:JSON.parse(JSON.stringify(els[sel]))}; return; }
        }
        for(var i=els.length-1;i>=0;i--){
            if(hitE(els[i],p.x,p.y)){
                sel=i; drag={t:'m',idx:i,ox:p.x,oy:p.y,s0:JSON.parse(JSON.stringify(els[i]))};
                document.getElementById('del-btn').style.display='';
                syncTB(els[i]); draw(); return;
            }
        }
        sel=-1; document.getElementById('del-btn').style.display='none';
        document.getElementById('tb-text').style.display='none';
        draw(); return;
    }
    if(tool==='text'){ startTxt(p.x,p.y,null); return; }
    if(tool==='stamp'){
        pushHistory();
        var st_={type:'stamp',x:p.x-80,y:p.y-22,w:160,h:44,txt:'SELLO',color:aCol,opacity:getOp()};
        els.push(st_); sel=els.length-1;
        document.getElementById('del-btn').style.display='';
        draw(); return;
    }
    var fc=getFill();
    if(tool==='rect')   drawing={type:'rect',  x:p.x,y:p.y,w:0,h:0,sc:aCol,sw:getSW(),fc:fc,opacity:getOp()};
    if(tool==='ell')    drawing={type:'ell',   x:p.x,y:p.y,w:0,h:0,sc:aCol,sw:getSW(),fc:fc,opacity:getOp()};
    if(tool==='line')   drawing={type:'line',  x1:p.x,y1:p.y,x2:p.x,y2:p.y,sc:aCol,sw:getSW(),opacity:getOp()};
    if(tool==='arrow')  drawing={type:'arrow', x1:p.x,y1:p.y,x2:p.x,y2:p.y,sc:aCol,sw:getSW(),opacity:getOp()};
    if(tool==='hl')     drawing={type:'hl',    x:p.x,y:p.y,w:0,h:Math.max(getSW()*4,20),color:curHlColor,opacity:0.6};
    if(tool==='redact') drawing={type:'redact',x:p.x,y:p.y,w:0,h:0,color:'#000000',opacity:1};
});

cv.addEventListener('mousemove',function(e){
    var p=xy(e);
    document.getElementById('coord-info').textContent='X: '+Math.round(p.x)+'  Y: '+Math.round(p.y);
    if(drag){ doDrag(p); return; }
    if(drawing){ doDrawing(p); return; }
    if(tool!=='tf'){ cv.style.cursor='crosshair'; return; }
    if(sel>=0&&sel<els.length){
        var hh=hitH(els[sel],p.x,p.y);
        if(hh){ cv.style.cursor=hh.cur; return; }
        if(hitE(els[sel],p.x,p.y)){ cv.style.cursor='move'; return; }
    }
    for(var i=els.length-1;i>=0;i--){ if(hitE(els[i],p.x,p.y)){ cv.style.cursor='move'; return; } }
    cv.style.cursor='default';
});

cv.addEventListener('mouseup',function(){
    if(drag){ drag=null; return; }
    if(drawing){
        var d=drawing; drawing=null; var MIN=3;
        if(d.type==='rect'||d.type==='ell'||d.type==='redact'){
            if(Math.abs(d.w)>MIN||Math.abs(d.h)>MIN){
                var rx=Math.min(d.x,d.x+d.w),ry=Math.min(d.y,d.y+d.h);
                pushHistory();
                els.push({type:d.type,x:rx,y:ry,w:Math.abs(d.w),h:Math.abs(d.h),
                           sc:d.sc,sw:d.sw,fc:d.fc,color:d.color,opacity:d.opacity});
            }
        }
        if(d.type==='hl'){
            if(Math.abs(d.w)>MIN){
                pushHistory();
                var hx=Math.min(d.x,d.x+d.w);
                els.push({type:'hl',x:hx,y:d.y,w:Math.abs(d.w),h:d.h,color:d.color,opacity:d.opacity});
            }
        }
        if((d.type==='line'||d.type==='arrow')&&Math.hypot(d.x2-d.x1,d.y2-d.y1)>MIN){
            pushHistory();
            els.push(d);
        }
        draw();
    }
});

cv.addEventListener('dblclick',function(e){
    var p=xy(e);
    for(var i=els.length-1;i>=0;i--){
        if(els[i].type==='text'&&hitE(els[i],p.x,p.y)){ editTxt(i); return; }
    }
});

cv.addEventListener('wheel',function(e){
    if(e.ctrlKey||e.metaKey){
        e.preventDefault();
        if(e.deltaY<0) zoomIn(); else zoomOut();
    }
},{passive:false});

// ── Drag logic ─────────────────────────────────────────────────────────────
function doDrawing(p){
    if(drawing.type==='rect'||drawing.type==='ell'||drawing.type==='redact'){
        drawing.w=p.x-drawing.x; drawing.h=p.y-drawing.y;
    }
    if(drawing.type==='hl'){ drawing.w=p.x-drawing.x; }
    if(drawing.type==='line'||drawing.type==='arrow'){ drawing.x2=p.x; drawing.y2=p.y; }
    draw();
}
function doDrag(p){
    var e=els[drag.idx],s=drag.s0,dx=p.x-drag.ox,dy=p.y-drag.oy,h=drag.hid;
    if(drag.t==='m'){
        if(e.type==='line'||e.type==='arrow'){ e.x1=s.x1+dx; e.y1=s.y1+dy; e.x2=s.x2+dx; e.y2=s.y2+dy; }
        else{ e.x=s.x+dx; e.y=s.y+dy; }
    } else {
        if(e.type==='line'||e.type==='arrow'){
            if(h==='p1'){ e.x1=s.x1+dx; e.y1=s.y1+dy; } else{ e.x2=s.x2+dx; e.y2=s.y2+dy; }
        } else if(e.type==='text'){
            var sw=s.w||20,sh=s.h||(s.fs*1.4)||20;
            var nw=sw,nh=sh;
            if(h==='tl'||h==='ml'||h==='bl') nw=sw-dx;
            if(h==='tr'||h==='mr'||h==='br') nw=sw+dx;
            if(h==='tl'||h==='tm'||h==='tr') nh=sh-dy;
            if(h==='bl'||h==='bm'||h==='br') nh=sh+dy;
            if(nw>8&&nh>4){
                var scl=Math.max(nw/sw,nh/sh);
                e.fs=Math.max(6,Math.round(s.fs*scl));
                remeas(e);
                if(h==='tl'||h==='tm'||h==='tr') e.y=s.y+(sh-e.h);
                if(h==='tl'||h==='ml'||h==='bl') e.x=s.x+(sw-e.w);
            }
        } else {
            var nx=s.x,ny=s.y,nw2=s.w,nh2=s.h;
            if(h==='tl'){ nx=s.x+dx; ny=s.y+dy; nw2=s.w-dx; nh2=s.h-dy; }
            else if(h==='tm'){ ny=s.y+dy; nh2=s.h-dy; }
            else if(h==='tr'){ ny=s.y+dy; nw2=s.w+dx; nh2=s.h-dy; }
            else if(h==='ml'){ nx=s.x+dx; nw2=s.w-dx; }
            else if(h==='mr'){ nw2=s.w+dx; }
            else if(h==='bl'){ nx=s.x+dx; nw2=s.w-dx; nh2=s.h+dy; }
            else if(h==='bm'){ nh2=s.h+dy; }
            else if(h==='br'){ nw2=s.w+dx; nh2=s.h+dy; }
            if(nw2>4){ e.x=nx; e.w=nw2; } if(nh2>4){ e.y=ny; e.h=nh2; }
        }
    }
    draw();
}

// ── Text editing ────────────────────────────────────────────────────────────
function syncTB(e){
    if(!e||e.type!=='text') return;
    document.getElementById('ff').value=e.ff||'Arial';
    document.getElementById('fs').value=e.fs||16;
    bold=!!e.bold; ital=!!e.ital; underline=!!e.under; tCol=e.tc||'#000000';
    document.getElementById('b-bold').classList.toggle('em',bold);
    document.getElementById('b-ital').classList.toggle('em',ital);
    document.getElementById('b-under').classList.toggle('em',underline);
    selPal('pal2',tCol);
    document.getElementById('tb-text').style.display='flex';
}
function editTxt(idx){
    sel=idx; syncTB(els[idx]);
    var e=els[idx]; e._hid=true;
    startTxt(e.x,e.y,idx);
}
function startTxt(cx,cy,eidx){
    pendTx=cx; pendTy=cy; pendIdx=eidx;
    var inp=document.getElementById('txt-inp');
    inp.value=(eidx!==null&&eidx>=0&&els[eidx])?els[eidx].txt:'';
    document.getElementById('txt-input-area').style.display='flex';
    inp.focus();
    if(inp.value) inp.setSelectionRange(0,inp.value.length);
    draw();
}
function commitTxt(){
    var v=document.getElementById('txt-inp').value.trim();
    if(pendIdx!==null&&pendIdx>=0&&pendIdx<els.length){
        var ex=els[pendIdx]; ex._hid=false;
        if(v){ ex.txt=v; ex.ff=getFF(); ex.fs=getFS(); ex.tc=tCol; ex.bold=bold; ex.ital=ital; ex.under=underline; remeas(ex); }
        else{ pushHistory(); els.splice(pendIdx,1); sel=-1; }
    } else if(v&&pendTx!==null){
        var ne={type:'text',x:pendTx,y:pendTy,txt:v,ff:getFF(),fs:getFS(),tc:tCol,bold:bold,ital:ital,under:underline,w:0,h:0,opacity:1};
        pushHistory(); remeas(ne); els.push(ne);
    }
    pendTx=null; pendTy=null; pendIdx=null;
    document.getElementById('txt-input-area').style.display='none';
    document.getElementById('txt-inp').value='';
    draw();
}
function cancelTxt(){
    if(pendIdx!==null&&pendIdx>=0&&pendIdx<els.length) els[pendIdx]._hid=false;
    pendTx=null; pendTy=null; pendIdx=null;
    document.getElementById('txt-input-area').style.display='none';
    document.getElementById('txt-inp').value='';
    draw();
}

// ── Keys ────────────────────────────────────────────────────────────────────
document.addEventListener('keydown',function(ev){
    if(ev.target.tagName==='INPUT'||ev.target.tagName==='TEXTAREA'||ev.target.tagName==='SELECT') return;
    if((ev.key==='Delete'||ev.key==='Backspace')&&sel>=0){ ev.preventDefault(); delSel(); }
    if(ev.key==='Escape'){ sel=-1; document.getElementById('del-btn').style.display='none'; draw(); }
    if((ev.ctrlKey||ev.metaKey)&&ev.key==='z'){ ev.preventDefault(); doUndo(); }
    if(ev.key==='v'||ev.key==='V') T('tf');
    if(ev.key==='r'||ev.key==='R') T('rect');
    if(ev.key==='e'||ev.key==='E') T('ell');
    if(ev.key==='l'||ev.key==='L') T('line');
    if(ev.key==='t'||ev.key==='T') T('text');
    if(ev.key==='h'||ev.key==='H') T('hl');
});
function delSel(){
    if(sel>=0&&sel<els.length){ pushHistory(); els.splice(sel,1); sel=-1; document.getElementById('del-btn').style.display='none'; draw(); }
}

// ── Save to Streamlit ───────────────────────────────────────────────────────
function saveDat(){
    var payload=JSON.stringify({canvas_width:CW,canvas_height:CH,elements:els});
    try{
        var pd=window.parent.document,ta=null;
        var ws=pd.querySelectorAll('[data-testid="stTextArea"]');
        for(var i=0;i<ws.length;i++){
            var l=ws[i].querySelector('label');
            if(l&&l.textContent.trim()==='NEXA_CANVAS_DATA'){ ta=ws[i].querySelector('textarea'); break; }
        }
        if(!ta) ta=pd.querySelector('textarea[aria-label="NEXA_CANVAS_DATA"]');
        if(ta){
            var sv=Object.getOwnPropertyDescriptor(window.parent.HTMLTextAreaElement.prototype,'value').set;
            sv.call(ta,payload);
            ta.dispatchEvent(new window.parent.Event('input',{bubbles:true}));
            ta.dispatchEvent(new window.parent.Event('change',{bubbles:true}));
            setTimeout(function(){ ta.dispatchEvent(new window.parent.Event('blur',{bubbles:true})); },120);
            document.getElementById('status').textContent='✓ '+els.length+' elemento(s) guardado(s) correctamente';
        } else {
            document.getElementById('status').textContent='⚠ Campo no encontrado. Recarga la página si persiste.';
        }
    } catch(err){
        document.getElementById('status').textContent='Error: '+err.message;
    }
}
</script>
</body>
</html>"""

# ══════════════════════════════════════════════════════
#  SIDEBAR — Logo + Navegación + Estadísticas
# ══════════════════════════════════════════════════════
with st.sidebar:
    # Logo
    ruta_logo = None
    for ext in ["logo.png","logo.jpg","logo.jpeg"]:
        if os.path.exists(ext):
            ruta_logo = ext
            break

    st.markdown('<div style="height:24px;"></div>', unsafe_allow_html=True)

    if ruta_logo:
        try:
            col_l, col_m, col_r = st.columns([0.5, 2, 0.5])
            with col_m:
                st.image(ruta_logo, use_container_width=True)
        except Exception:
            pass
    else:
        st.markdown("""
        <div style="text-align:center;padding:8px 0 16px;">
            <div style="font-size:32px;font-weight:900;
                background:linear-gradient(135deg,#1B9FD8,#1DE0C0);
                -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                background-clip:text;letter-spacing:-1px;">NEXA</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align:center;margin-bottom:24px;">
        <span style="display:inline-flex;align-items:center;gap:5px;
            background:rgba(27,159,216,0.08);border:1px solid rgba(27,159,216,0.15);
            border-radius:20px;padding:4px 14px;font-size:10px;font-weight:700;
            color:#1B9FD8;letter-spacing:0.8px;text-transform:uppercase;">
            ⚡ v2.0 Premium
        </span>
    </div>
    <div style="height:1px;background:linear-gradient(90deg,transparent,rgba(27,159,216,0.2),transparent);
        margin-bottom:24px;"></div>
    """, unsafe_allow_html=True)

    # Estadísticas de sesión
    st.markdown("""
    <div style="font-size:10px;font-weight:700;color:#1A3A56;text-transform:uppercase;
        letter-spacing:1.2px;margin-bottom:12px;">📊 Sesión actual</div>
    """, unsafe_allow_html=True)

    docs_procesados = st.session_state.get("_nx_docs_count", 0)
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.markdown(f"""
        <div style="background:rgba(27,159,216,0.06);border:1px solid rgba(27,159,216,0.1);
            border-radius:12px;padding:12px 10px;text-align:center;margin-bottom:8px;">
            <div style="font-size:22px;font-weight:800;color:#1B9FD8;">{docs_procesados}</div>
            <div style="font-size:10px;color:#2A4A6A;margin-top:2px;">Docs</div>
        </div>""", unsafe_allow_html=True)
    with col_s2:
        total_els = sum(len(v) for v in st.session_state.get("ed_canvas_data", {}).values())
        st.markdown(f"""
        <div style="background:rgba(27,159,216,0.06);border:1px solid rgba(27,159,216,0.1);
            border-radius:12px;padding:12px 10px;text-align:center;margin-bottom:8px;">
            <div style="font-size:22px;font-weight:800;color:#1DE0C0;">{total_els}</div>
            <div style="font-size:10px;color:#2A4A6A;margin-top:2px;">Ediciones</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style="height:1px;background:linear-gradient(90deg,transparent,rgba(27,159,216,0.15),transparent);
        margin:16px 0 20px;"></div>
    <div style="font-size:10px;font-weight:700;color:#1A3A56;text-transform:uppercase;
        letter-spacing:1.2px;margin-bottom:12px;">🛠 Herramientas</div>
    """, unsafe_allow_html=True)

    # Menú de herramientas en sidebar
    tools_menu = [
        ("🗂️", "Nexíficar Masivo",  "Ensambla PDFs desde Excel + ZIP"),
        ("📄", "Nexíficar PDFs",    "Une múltiples PDFs uno a uno"),
        ("✂️", "Dividir PDF",       "Divide por páginas o rangos"),
        ("🗜️", "Comprimir PDF",     "Próximamente"),
        ("🔗", "Merge PDF",         "Próximamente"),
        ("✏️", "Editar PDF",        "Editor visual tipo Adobe"),
        ("🗑️", "Eliminar Páginas",  "Elimina páginas visualmente"),
    ]
    for icon, name, desc in tools_menu:
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:10px;padding:8px 10px;
            border-radius:10px;margin-bottom:3px;cursor:pointer;
            transition:background .15s;opacity:0.85;"
            onmouseover="this.style.background='rgba(27,159,216,0.1)'"
            onmouseout="this.style.background='transparent'">
            <span style="font-size:16px;">{icon}</span>
            <div>
                <div style="font-size:12px;font-weight:600;color:#8AAEC8;">{name}</div>
                <div style="font-size:10px;color:#1A3456;margin-top:1px;">{desc}</div>
            </div>
        </div>""", unsafe_allow_html=True)

    # Footer sidebar
    st.markdown("""
    <div style="position:absolute;bottom:20px;left:0;right:0;padding:0 16px;text-align:center;">
        <div style="height:1px;background:linear-gradient(90deg,transparent,rgba(27,159,216,0.15),transparent);
            margin-bottom:14px;"></div>
        <div style="font-size:10px;color:#0F2035;">
            NEXA © 2025 · Plataforma Documental<br>
            <span style="color:#1A3A56;">Transformación de Procesos</span>
        </div>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  AUTENTICACIÓN
# ══════════════════════════════════════════════════════
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown("""
        <div style="background:linear-gradient(145deg,#07111C,#0A1626);
                    border:1px solid rgba(27,159,216,0.2);
                    border-radius:24px;padding:48px 40px;text-align:center;
                    box-shadow:0 20px 80px rgba(0,0,0,0.5),0 0 0 1px rgba(27,159,216,0.05);">
            <div style="font-size:52px;margin-bottom:16px;">🔒</div>
            <div style="font-size:22px;font-weight:800;color:#E8F4FF;
                        margin-bottom:8px;letter-spacing:-0.3px;">
                Acceso Restringido
            </div>
            <div style="font-size:14px;color:#2A4A6A;margin-bottom:28px;line-height:1.6;">
                Ingresa tu código de acceso para acceder<br>a la plataforma NEXA.
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        password = st.text_input(
            "Código de acceso:",
            type="password",
            placeholder="••••••••••••",
            label_visibility="collapsed"
        )
        if st.button("⚡ Entrar a NEXA", type="primary", use_container_width=True):
            try:
                claves_validas = list(st.secrets["accesos"].values())
                if password in claves_validas:
                    st.session_state.autenticado = True
                    st.session_state["_nx_docs_count"] = 0
                    st.rerun()
                else:
                    st.error("❌ Código incorrecto o inactivo.")
            except Exception:
                st.warning("⚠️ La bóveda de contraseñas no está configurada en Streamlit.")
    st.stop()

# ══════════════════════════════════════════════════════
#  HEADER PRINCIPAL
# ══════════════════════════════════════════════════════
st.markdown("""
<div style="display:flex;align-items:center;justify-content:space-between;
    padding:0 0 24px 0;border-bottom:1px solid rgba(27,159,216,0.08);
    margin-bottom:8px;">
    <div>
        <div style="font-size:11px;font-weight:700;color:#1A3A56;text-transform:uppercase;
            letter-spacing:1.5px;margin-bottom:6px;">⚡ NEXA Platform</div>
        <div style="font-size:26px;font-weight:900;color:#E8F4FF;letter-spacing:-0.8px;">
            Gestión Documental <span style="background:linear-gradient(135deg,#1B9FD8,#1DE0C0);
            -webkit-background-clip:text;-webkit-text-fill-color:transparent;
            background-clip:text;">Inteligente</span>
        </div>
    </div>
    <div style="display:flex;gap:10px;align-items:center;">
        <div style="font-size:12px;color:#2A4A6A;">Plataforma PDF Empresarial</div>
        <div style="width:8px;height:8px;border-radius:50%;background:#1DE0C0;
            box-shadow:0 0 10px rgba(29,224,192,0.6);animation:pulse-glow 2s infinite;"></div>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  FUNCIONES VALIDADOR IA
# ══════════════════════════════════════════════════════
def extraer_texto_pdf(archivo_bytes):
    try:
        import fitz
        doc = fitz.open(stream=archivo_bytes, filetype="pdf")
        texto_completo = ""
        paginas = []
        for i, page in enumerate(doc):
            texto_pagina = page.get_text()
            texto_completo += f"\n--- PÁGINA {i+1} ---\n{texto_pagina}"
            paginas.append({"pagina": i+1, "texto": texto_pagina})
        n = len(doc)
        doc.close()
        return texto_completo, paginas, n
    except Exception as e:
        return "", [], 0

def validar_con_ia(texto_pdf, tipo_tramite, criterios_extra=""):
    import anthropic
    import json
    
    prompts = {
        "Desembolso Subsidio Colsubsidio": f"""Eres experto validador de documentos de desembolso de subsidios de vivienda en Colombia para Colsubsidio.
Analiza el texto del expediente y verifica CADA criterio:

1. CERTIFICADO DE TRADICIÓN Y LIBERTAD (CTL):
   - ¿Fecha de expedición menor a 30 días?
   - ¿Vigencia del subsidio POSTERIOR a fecha de firma de escritura?
   - ¿Especifica que es vivienda VIS?
   - ¿Número de matrícula inmobiliaria presente?

2. ACTA DE ENTREGA:
   - ¿Fecha del acta POSTERIOR o igual a fecha de suscripción de escritura?
   - ¿Firmada por todas las partes?

3. AUTORIZACIÓN DE DESEMBOLSO:
   - ¿Valor del subsidio coincide con la escritura pública?
   - ¿Firmada por todos los beneficiarios mayores de edad?

4. ESCRITURA PÚBLICA:
   - ¿Valor de compraventa consistente con autorización de desembolso?
   - ¿Especifica que es VIS?

5. CERTIFICADO DE EXISTENCIA DE LA VIVIENDA:
   - ¿Está presente?

CRITERIOS ADICIONALES: {criterios_extra}

Responde SOLO con JSON válido, sin texto adicional, sin backticks:
{{{{
  "score": <0-100>,
  "resultado_general": "<APROBADO|REQUIERE_REVISION|RECHAZADO>",
  "resumen": "<frase explicando resultado>",
  "hallazgos": [
    {{{{
      "criterio": "<nombre>",
      "documento": "<documento revisado>",
      "estado": "<CUMPLE|NO_CUMPLE|NO_ENCONTRADO|ADVERTENCIA>",
      "severidad": "<CRITICO|MEDIO|BAJO|OK>",
      "observacion": "<hallazgo específico>",
      "pagina_estimada": <número o null>
    }}}}
  ],
  "documentos_faltantes": ["<lista>"],
  "recomendaciones": ["<lista de acciones>"]
}}}}""",
        "Desembolso Caja Compensación": f"""Eres experto validador de documentos de desembolso de subsidios de vivienda en Colombia.
Verifica: CTL vigencia max 30 días, acta entrega posterior a escritura, valor subsidio consistente, firmas beneficiarios, certificado existencia vivienda.
CRITERIOS ADICIONALES: {criterios_extra}
Responde SOLO con JSON válido sin backticks con la misma estructura definida.""",
        "Personalizado": f"""Eres experto validador de documentos de radicación en Colombia.
Valida el expediente según estos criterios: {criterios_extra}
Responde SOLO con JSON válido sin backticks:
{{"score":0,"resultado_general":"REQUIERE_REVISION","resumen":"","hallazgos":[],"documentos_faltantes":[],"recomendaciones":[]}}"""
    }
    
    prompt = prompts.get(tipo_tramite, prompts["Personalizado"])
    client = get_anthropic_client()
    if not client:
        raise Exception("No se pudo inicializar el cliente de Anthropic")
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": f"{prompt}\n\nTEXTO DEL EXPEDIENTE:\n{texto_pdf[:15000]}"}]
    )
    respuesta = message.content[0].text.strip()
    if "```" in respuesta:
        partes = respuesta.split("```")
        for parte in partes:
            parte = parte.strip()
            if parte.startswith("json"):
                parte = parte[4:].strip()
            if parte.startswith("{"):
                respuesta = parte
                break
    return json.loads(respuesta)

# ══════════════════════════════════════════════════════
#  TABS PRINCIPALES
# ══════════════════════════════════════════════════════
tabs = st.tabs([
    "🗂️ Nexíficar Masivo",
    "📄 Nexíficar PDFs",
    "✂️ Dividir PDF",
    "🗜️ Comprimir PDF",
    "🔗 Merge PDF",
    "✏️ Editar PDF",
    "🗑️ Eliminar Páginas",
    "🤖 Validador IA",
])

# ══════════════════════════════════════════════════════
#  TAB 0 — NEXÍFICAR MASIVO
# ══════════════════════════════════════════════════════
with tabs[0]:

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">🗂️ Nexíficar PDFs Masivamente</div>
        <div class="nx-page-sub">Ensambla cientos de expedientes al mismo tiempo usando tu
        <strong>Plantilla de Excel</strong> y archivos <strong>ZIP</strong>,
        o renombra documentos de forma automática.</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="nx-section">📊 Archivos de entrada</div>', unsafe_allow_html=True)
    archivo_excel = st.file_uploader("Plantilla de Excel de Mapeo (.xlsx)", type=["xlsx"])
    archivos_zip  = st.file_uploader(
        "Archivos ZIP con los PDFs (puedes seleccionar varios)",
        type=["zip"], accept_multiple_files=True
    )
    st.markdown("---")

    def buscar_archivo_en_dir(nombre, dir_raiz):
        for root, _, archs in os.walk(dir_raiz):
            if nombre in archs: return os.path.join(root, nombre)
        return None

    def parse_paginas(inst_str):
        parsed = []
        if pd.isna(inst_str) or str(inst_str).strip() == '': return parsed
        for inst in str(inst_str).split(';'):
            if ':' not in inst: continue
            p, pos = inst.split(':')
            try: pos_final = int(pos)
            except: continue
            if p.lower() == 'completo': parsed.append(('completo', pos_final))
            elif '-' in p:
                try: parsed.append((list(range(int(p.split('-')[0])-1, int(p.split('-')[1]))), pos_final))
                except: pass
            elif ',' in p:
                try: parsed.append(([int(x.strip())-1 for x in p.split(',')], pos_final))
                except: pass
            else:
                try: parsed.append(([int(p)-1], pos_final))
                except: pass
        return parsed

    if st.button("🚀 Nexíficar Documentos Masivamente", type="primary", use_container_width=True):
        if not archivo_excel or not archivos_zip:
            st.warning("⚠️ Por favor, sube el Excel y al menos un archivo ZIP para comenzar.")
        else:
            with st.spinner("Nexíficando documentos mágicamente… Esto puede tomar unos segundos."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    ruta_origen = os.path.join(temp_dir, 'origen')
                    ruta_salida = os.path.join(temp_dir, 'salida')
                    os.makedirs(ruta_origen)
                    os.makedirs(ruta_salida)

                    for zip_file in archivos_zip:
                        with zipfile.ZipFile(zip_file, 'r') as z:
                            z.extractall(ruta_origen)

                    try:
                        df = pd.read_excel(archivo_excel)
                        columnas_archivo = [col for col in df.columns if str(col).startswith('Archivo_')]

                        if 'Nombre_Salida' not in df.columns:
                            st.error("❌ El Excel debe tener una columna llamada 'Nombre_Salida'.")
                        else:
                            barra   = st.progress(0)
                            exitos  = 0
                            errores = []

                            for idx, row in df.iterrows():
                                nombre_salida = str(row['Nombre_Salida']).strip()
                                if pd.isna(nombre_salida) or nombre_salida == 'nan' or not nombre_salida: continue
                                if not nombre_salida.lower().endswith('.pdf'): nombre_salida += '.pdf'

                                ruta_final      = os.path.join(ruta_salida, nombre_salida)
                                max_pos         = 0
                                docs_a_procesar = []

                                for col_arch in columnas_archivo:
                                    num_index = col_arch.split('_')[1]
                                    col_inst  = f'Instrucciones_{num_index}'
                                    if col_inst in df.columns:
                                        nombre_doc    = str(row.get(col_arch, '')).strip()
                                        instrucciones = str(row.get(col_inst, '')).strip()
                                        if nombre_doc and nombre_doc != 'nan':
                                            parsed = parse_paginas(instrucciones)
                                            for _, pos in parsed: max_pos = max(max_pos, pos)
                                            docs_a_procesar.append((nombre_doc, parsed))

                                if max_pos > 0:
                                    paginas_pos = [[] for _ in range(max_pos + 1)]
                                    error_fila  = False

                                    for n_doc, p_inst in docs_a_procesar:
                                        r_doc = buscar_archivo_en_dir(n_doc, ruta_origen)
                                        if not r_doc:
                                            errores.append(f"No se encontró '{n_doc}' para crear '{nombre_salida}'.")
                                            error_fila = True
                                            break
                                        try:
                                            reader = PdfReader(r_doc)
                                            for p_spec, p_final in p_inst:
                                                if p_spec == 'completo':
                                                    for i in range(len(reader.pages)):
                                                        paginas_pos[p_final].append(reader.pages[i])
                                                else:
                                                    for i in p_spec:
                                                        if i < len(reader.pages):
                                                            paginas_pos[p_final].append(reader.pages[i])
                                        except Exception as e:
                                            errores.append(f"Error leyendo '{n_doc}': {e}")
                                            error_fila = True
                                            break

                                    if not error_fila:
                                        writer = PdfWriter()
                                        for pos_idx in range(1, max_pos + 1):
                                            for p_obj in paginas_pos[pos_idx]: writer.add_page(p_obj)
                                        if len(writer.pages) > 0:
                                            with open(ruta_final, "wb") as f: writer.write(f)
                                            exitos += 1

                                barra.progress((idx + 1) / len(df))

                            if exitos > 0:
                                st.session_state["_nx_docs_count"] = st.session_state.get("_nx_docs_count",0) + exitos
                                st.success(f"🎉 ¡Proceso finalizado! Se nexíficaron {exitos} documentos con éxito.")
                                zip_final = os.path.join(temp_dir, 'NEXA_Resultados.zip')
                                with zipfile.ZipFile(zip_final, 'w') as z:
                                    for r, _, archs in os.walk(ruta_salida):
                                        for a in archs: z.write(os.path.join(r, a), a)
                                with open(zip_final, "rb") as fp:
                                    st.download_button(
                                        label="⬇️ Descargar Resultados (ZIP)",
                                        data=fp,
                                        file_name="NEXA_Resultados.zip",
                                        mime="application/zip",
                                        type="primary"
                                    )
                            else:
                                st.error("No se pudo generar ningún documento. Revisa tu Excel y que los PDFs existan.")

                            if errores:
                                with st.expander("⚠️ Ver detalles de advertencias"):
                                    for err in set(errores): st.write(err)

                    except Exception as e:
                        st.error(f"❌ Error leyendo el Excel: {e}")

# ══════════════════════════════════════════════════════
#  TAB 1 — NEXÍFICAR PDFs
# ══════════════════════════════════════════════════════
with tabs[1]:

    try:
        import fitz
        FITZ_OK = True
    except ImportError:
        FITZ_OK = False

    for _k, _v in [("nx_done", False), ("nx_buffer", None),
                   ("nx_nombre", "Documento_Unificado.pdf"),
                   ("nx_order", []), ("nx_files_sig", ""),
                   ("nx_editor_ver", 0)]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    def _render_steps(step):
        cfg = [("1", "Subir"), ("2", "Ordenar"), ("3", "Unificar")]
        html = '<div class="nx-steps">'
        for i, (num, lbl) in enumerate(cfg):
            if   i + 1 < step:  cs, ls, ic = "done",   "done",   "✓"
            elif i + 1 == step: cs, ls, ic = "active", "active", num
            else:                cs, ls, ic = "idle",   "",       num
            html += (f'<div class="nx-step">'
                     f'<div class="nx-circle {cs}">{ic}</div>'
                     f'<span class="nx-label {ls}">{lbl}</span>'
                     f'</div>')
            if i < len(cfg) - 1:
                html += f'<div class="nx-line {"done" if i+2<=step else "idle"}"></div>'
        return html + '</div>'

    def _get_thumb(pdf_bytes):
        if not FITZ_OK: return ""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(0.9, 0.9), alpha=False)
            return base64.b64encode(pix.tobytes("png")).decode()
        except Exception:
            return ""

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">📄🔗📄 Nexíficar PDFs</div>
        <div class="nx-page-sub">Sube varios PDFs sueltos y únelos en <strong>un solo archivo</strong>,
        con previsualización de miniatura y reordenamiento por número de posición.</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="nx-section">📂 Paso 1 — Subir PDFs</div>', unsafe_allow_html=True)
    archivos_subidos = st.file_uploader(
        "Selecciona o arrastra tus archivos PDF aquí",
        type=["pdf"],
        accept_multiple_files=True
    )

    if not archivos_subidos:
        st.session_state.nx_done      = False
        st.session_state.nx_buffer    = None
        st.session_state.nx_order     = []
        st.session_state.nx_files_sig = ""

    step = 1 if not archivos_subidos else (3 if st.session_state.nx_done else 2)
    st.markdown(_render_steps(step), unsafe_allow_html=True)

    if not archivos_subidos:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">📂</div>
            <div class="nx-empty-text">Usa el selector de arriba para cargar tus PDFs</div>
            <div class="nx-empty-sub">Puedes seleccionar múltiples archivos a la vez</div>
        </div>""", unsafe_allow_html=True)
    else:
        file_info_map = {}
        for arch in archivos_subidos:
            arch.seek(0)
            raw = arch.read()
            arch.seek(0)
            try:
                pages = len(PdfReader(BytesIO(raw)).pages)
            except Exception:
                pages = "?"
            kb  = len(raw) / 1024
            sz  = f"{kb:.1f} KB" if kb < 1024 else f"{kb/1024:.1f} MB"
            file_info_map[arch.name] = {
                "arch": arch, "name": arch.name, "raw": raw,
                "pages": pages, "size": sz, "thumb": _get_thumb(raw),
            }

        files_sig = hashlib.md5("".join(sorted(file_info_map.keys())).encode()).hexdigest()[:8]

        if st.session_state.nx_files_sig != files_sig:
            st.session_state.nx_files_sig = files_sig
            st.session_state.nx_order     = list(file_info_map.keys())
            st.session_state.nx_done      = False
            st.session_state.nx_buffer    = None

        st.session_state.nx_order = [n for n in st.session_state.nx_order if n in file_info_map]
        orden = st.session_state.nx_order

        _nx_col, _nx_panel = st.columns([3, 1])

        with _nx_panel:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:15px;font-weight:700;color:#C8E4F0;margin-bottom:12px;">📄🔗📄 Nexíficar PDFs</div>',
                    unsafe_allow_html=True
                )
                _n_arch = len(orden)
                st.markdown(
                    f'<div style="font-size:12px;color:#4A7A9C;margin-bottom:8px;">'
                    f'{_n_arch} archivo{"s" if _n_arch!=1 else ""} cargado{"s" if _n_arch!=1 else ""}</div>',
                    unsafe_allow_html=True
                )
                if st.session_state.nx_done and st.session_state.nx_buffer:
                    st.download_button(
                        label="⬇️ Descargar PDF",
                        data=st.session_state.nx_buffer,
                        file_name=st.session_state.nx_nombre,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                        key="nx_dl_panel"
                    )

        with _nx_col:
            if not st.session_state.nx_done:
                st.markdown('<div class="nx-section">📋 Paso 2 — Ajusta el orden y selecciona los PDFs</div>',
                            unsafe_allow_html=True)
                st.info(
                    "Edita los números de **Orden** para cambiar la posición. "
                    "Desmarca **✓** para excluir un PDF del resultado. "
                    "Pulsa **Aplicar orden** para confirmar.",
                    icon="ℹ️"
                )

                order_df = pd.DataFrame([
                    {
                        "Orden":   i + 1,
                        "✓":       True,
                        "Archivo": name,
                        "Páginas": str(file_info_map[name]["pages"]),
                        "Tamaño":  file_info_map[name]["size"],
                    }
                    for i, name in enumerate(orden)
                ])

                edited_df = st.data_editor(
                    order_df,
                    column_config={
                        "Orden":   st.column_config.NumberColumn(
                                       "Orden", min_value=1, max_value=len(orden), step=1, width="small"),
                        "✓":       st.column_config.CheckboxColumn("✓", width="small"),
                        "Archivo": st.column_config.TextColumn("Archivo", disabled=True),
                        "Páginas": st.column_config.TextColumn("Páginas", disabled=True, width="small"),
                        "Tamaño":  st.column_config.TextColumn("Tamaño",  disabled=True, width="small"),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"order_editor_{st.session_state.nx_editor_ver}",
                )

                if st.button("↺ Aplicar orden", use_container_width=True):
                    nuevos = (
                        edited_df[edited_df["✓"]]
                        .sort_values("Orden")["Archivo"]
                        .tolist()
                    )
                    if not nuevos:
                        st.warning("⚠️ Debes incluir al menos un PDF.")
                    else:
                        st.session_state.nx_order      = nuevos
                        st.session_state.nx_editor_ver += 1
                        st.rerun()

                if orden:
                    st.markdown('<div class="nx-section">🖼️ Vista previa del orden actual</div>',
                                unsafe_allow_html=True)
                    CARDS_PER_ROW = 4
                    for row_start in range(0, len(orden), CARDS_PER_ROW):
                        row_slice = orden[row_start : row_start + CARDS_PER_ROW]
                        cols = st.columns(CARDS_PER_ROW)
                        for col_offset, name in enumerate(row_slice):
                            fi  = file_info_map[name]
                            idx = row_start + col_offset
                            with cols[col_offset]:
                                with st.container(border=True):
                                    if fi["thumb"]:
                                        st.image(base64.b64decode(fi["thumb"]), use_container_width=True)
                                    else:
                                        st.markdown(
                                            '<div style="background:#07111C;border-radius:10px;'
                                            'height:100px;display:flex;align-items:center;'
                                            'justify-content:center;font-size:32px;">📄</div>',
                                            unsafe_allow_html=True
                                        )
                                    short = (fi["name"][:22] + "…") if len(fi["name"]) > 22 else fi["name"]
                                    st.markdown(
                                        f'<div style="font-size:12px;font-weight:600;color:#C8E4F0;'
                                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'
                                        f'margin:6px 0 2px 0;">'
                                        f'<span style="background:linear-gradient(135deg,#1B9FD8,#1580B8);'
                                        f'color:#fff;border-radius:50%;padding:1px 7px;margin-right:5px;'
                                        f'font-size:11px;font-weight:700;">{idx + 1}</span>'
                                        f'{short}</div>',
                                        unsafe_allow_html=True
                                    )
                                    st.caption(f"📄 {fi['pages']} pág. · 💾 {fi['size']}")

                st.markdown('<div class="nx-export-bar">', unsafe_allow_html=True)
                st.markdown('<div class="nx-export-label">💾 Nombre del PDF final y exportación</div>',
                            unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

                col_name, col_btn = st.columns([3, 1])
                with col_name:
                    nombre_final = st.text_input(
                        "Nombre:",
                        "Documento_Unificado.pdf",
                        label_visibility="collapsed",
                        placeholder="Nombre_del_archivo_final.pdf"
                    )
                    if not nombre_final.lower().endswith(".pdf"):
                        nombre_final += ".pdf"
                with col_btn:
                    st.markdown("<br>", unsafe_allow_html=True)
                    nexificar = st.button(
                        f"🔗 Nexíficar {len(orden)} PDFs",
                        type="primary",
                        use_container_width=True,
                        disabled=(len(orden) == 0)
                    )

                if nexificar:
                    if not orden:
                        st.warning("⚠️ No hay documentos para unir.")
                    else:
                        with st.spinner("Nexíficando documentos conservando la calidad original…"):
                            try:
                                merger = PdfMerger()
                                for n in orden:
                                    a = file_info_map[n]["arch"]
                                    a.seek(0)
                                    merger.append(a)
                                buf = BytesIO()
                                merger.write(buf)
                                merger.close()
                                buf.seek(0)
                                st.session_state.nx_buffer = buf.getvalue()
                                st.session_state.nx_nombre = nombre_final
                                st.session_state.nx_done   = True
                                st.session_state["_nx_docs_count"] = st.session_state.get("_nx_docs_count",0) + 1
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Error al unir los archivos: {e}")
            else:
                total = len(orden)
                st.markdown(f"""
                <div class="nx-success-card">
                    <div class="nx-success-icon">🎉</div>
                    <div class="nx-success-title">¡Nexíficación completada!</div>
                    <div class="nx-success-sub">
                        {total} PDF{'s' if total != 1 else ''} unidos en
                        <strong>{st.session_state.nx_nombre}</strong>
                    </div>
                </div>""", unsafe_allow_html=True)

                st.download_button(
                    label="⬇️ Descargar PDF Unificado",
                    data=st.session_state.nx_buffer,
                    file_name=st.session_state.nx_nombre,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🔄 Nexíficar otros PDFs", use_container_width=True):
                    st.session_state.nx_done       = False
                    st.session_state.nx_buffer     = None
                    st.session_state.nx_order      = []
                    st.session_state.nx_files_sig  = ""
                    st.session_state.nx_editor_ver = 0
                    st.rerun()

# ══════════════════════════════════════════════════════
#  TAB 2 — DIVIDIR PDF
# ══════════════════════════════════════════════════════
with tabs[2]:

    for _k, _v in [
        ("sp_mode", "A"), ("sp_sel_b", set()),
        ("sp_pages_per_part", 2), ("sp_result", None),
        ("sp_sig", ""), ("sp_thumbs", []), ("sp_n", 0),
    ]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">✂️ Dividir PDF</div>
        <div class="nx-page-sub">Elige cómo quieres dividir tu PDF, selecciona las páginas
        visualmente y descarga todas las partes como ZIP.</div>
    </div>
    """, unsafe_allow_html=True)

    sp_file = st.file_uploader("Sube el PDF a dividir", type=["pdf"], key="sp_uploader")

    if not sp_file:
        st.session_state.sp_result = None
        st.session_state.sp_sel_b  = set()
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">✂️</div>
            <div class="nx-empty-text">Sube un PDF para empezar a dividirlo</div>
            <div class="nx-empty-sub">Extrae páginas, selecciona rangos o divide en partes iguales</div>
        </div>""", unsafe_allow_html=True)
    else:
        if not FITZ_OK:
            st.error("⚠️ PyMuPDF no está instalado. Agrégalo a requirements.txt.")
        else:
            sp_bytes = sp_file.read()
            thumbs, n_pages = _ensure_thumbs("sp", sp_bytes)

            st.markdown('<div class="nx-section">⚙️ Elige el modo de división</div>', unsafe_allow_html=True)

            _SP_MODES = [
                ("A", "📄", "Extraer todas las páginas",
                 "Genera un ZIP con cada página del PDF como documento individual."),
                ("B", "☑️", "Seleccionar páginas",
                 "Elige páginas o rangos (ej: 1,3,5-8) y extráelas en un ZIP."),
                ("C", "🔢", "Dividir en partes iguales",
                 "Define cuántas páginas por parte y divide el PDF uniformemente."),
            ]
            mode_cols = st.columns(3)
            for (mid, icon, title, desc), col in zip(_SP_MODES, mode_cols):
                active = st.session_state.sp_mode == mid
                bdr = "2px solid #1B9FD8" if active else "1px solid rgba(27,159,216,0.1)"
                bg_c = "rgba(27,159,216,0.1)" if active else "linear-gradient(145deg,#09172A,#0A1626)"
                col.markdown(
                    f'<div style="background:{bg_c};border:{bdr};border-radius:16px;'
                    f'padding:20px 16px;text-align:center;min-height:130px;">'
                    f'<div style="font-size:32px;margin-bottom:10px;">{icon}</div>'
                    f'<div style="font-size:13px;font-weight:700;color:#C8E4F0;margin-bottom:6px;">{title}</div>'
                    f'<div style="font-size:11px;color:#2A4A6A;line-height:1.5;">{desc}</div>'
                    f'</div>', unsafe_allow_html=True
                )
                if col.button(
                    "✓ Seleccionado" if active else "Seleccionar",
                    key=f"sp_mode_{mid}", use_container_width=True,
                    type="primary" if active else "secondary"
                ):
                    st.session_state.sp_mode = mid
                    st.session_state.sp_result = None
                    st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)

            grid_col, panel_col = st.columns([3, 1])
            mode = st.session_state.sp_mode
            _SP_SEC = ["#1B9FD8","#27AE60","#E74C3C","#F39C12","#9B59B6","#1ABC9C","#E91E63","#FF9800"]

            with panel_col:
                st.markdown(
                    '<div style="background:linear-gradient(145deg,#07111C,#09172A);'
                    'border:1px solid rgba(27,159,216,0.15);border-radius:16px;padding:20px 16px;">',
                    unsafe_allow_html=True
                )
                _mode_names = {"A": ("📄","Extraer todo"), "B": ("☑️","Por páginas"), "C": ("🔢","Partes iguales")}
                _ic, _lb = _mode_names[mode]
                st.markdown(
                    f'<div style="font-size:11px;font-weight:700;color:#1B9FD8;text-transform:uppercase;'
                    f'letter-spacing:1px;margin-bottom:10px;">Modo activo</div>'
                    f'<div style="font-size:24px;margin-bottom:4px;">{_ic}</div>'
                    f'<div style="font-size:13px;font-weight:700;color:#C8E4F0;margin-bottom:16px;">{_lb}</div>',
                    unsafe_allow_html=True
                )

                if mode == "B":
                    n_sel_b = len(st.session_state.sp_sel_b)
                    st.markdown(
                        f'<div style="font-size:32px;font-weight:800;color:#1B9FD8;text-align:center;'
                        f'margin-bottom:4px;">{n_sel_b}</div>'
                        f'<div style="font-size:11px;color:#2A4A6A;text-align:center;'
                        f'margin-bottom:16px;">páginas seleccionadas</div>',
                        unsafe_allow_html=True
                    )
                    sp_custom = st.text_input(
                        "Rangos (ej: 1,3,5-8)", value="",
                        placeholder="1,3,5-8", key="sp_custom_text",
                        label_visibility="collapsed"
                    )
                    if st.button("Aplicar rangos", use_container_width=True, key="sp_apply_range"):
                        try:
                            new_sel = set()
                            for part in sp_custom.split(","):
                                part = part.strip()
                                if not part: continue
                                if "-" in part:
                                    a, b = part.split("-", 1)
                                    new_sel.update(range(int(a)-1, int(b)))
                                else:
                                    new_sel.add(int(part)-1)
                            st.session_state.sp_sel_b = {p for p in new_sel if 0 <= p < n_pages}
                            st.rerun()
                        except Exception:
                            st.error("Formato inválido.")
                    if st.button("Limpiar", use_container_width=True, key="sp_clear_b"):
                        st.session_state.sp_sel_b = set()
                        st.rerun()

                elif mode == "C":
                    _max_ppp = max(1, n_pages - 1)
                    _cur_ppp = min(st.session_state.sp_pages_per_part, _max_ppp)
                    _ppp = st.number_input(
                        "Páginas por parte", min_value=1, max_value=_max_ppp,
                        value=_cur_ppp, key="sp_ppp"
                    )
                    st.session_state.sp_pages_per_part = _ppp
                    _n_parts_c = (n_pages + _ppp - 1) // _ppp
                    st.markdown(
                        f'<div style="font-size:12px;color:#2A4A6A;margin-top:8px;">'
                        f'→ {_n_parts_c} parte{"s" if _n_parts_c>1 else ""}</div>',
                        unsafe_allow_html=True
                    )

                st.markdown('</div>', unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                _can_split = (
                    mode == "A" or
                    (mode == "B" and len(st.session_state.sp_sel_b) > 0) or
                    mode == "C"
                )
                if st.button("✂️ Dividir PDF", type="primary",
                             use_container_width=True, disabled=not _can_split,
                             key="sp_do_split"):
                    with st.spinner("Dividiendo…"):
                        try:
                            reader  = PdfReader(BytesIO(sp_bytes))
                            zip_buf = BytesIO()
                            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                if mode == "A":
                                    for pi in range(n_pages):
                                        w = PdfWriter()
                                        w.add_page(reader.pages[pi])
                                        pb = BytesIO(); w.write(pb)
                                        zf.writestr(f"pagina_{pi+1:03d}.pdf", pb.getvalue())
                                elif mode == "B":
                                    for pi in sorted(st.session_state.sp_sel_b):
                                        w = PdfWriter()
                                        w.add_page(reader.pages[pi])
                                        pb = BytesIO(); w.write(pb)
                                        zf.writestr(f"pagina_{pi+1:03d}.pdf", pb.getvalue())
                                elif mode == "C":
                                    _ppp2 = st.session_state.sp_pages_per_part
                                    _part = 1
                                    for start in range(0, n_pages, _ppp2):
                                        w = PdfWriter()
                                        for pi in range(start, min(start + _ppp2, n_pages)):
                                            w.add_page(reader.pages[pi])
                                        pb = BytesIO(); w.write(pb)
                                        zf.writestr(f"parte_{_part:02d}.pdf", pb.getvalue())
                                        _part += 1
                            zip_buf.seek(0)
                            st.session_state.sp_result = zip_buf.getvalue()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

                if st.session_state.sp_result:
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.download_button(
                        label="⬇️ Descargar ZIP",
                        data=st.session_state.sp_result,
                        file_name=f"{sp_file.name.replace('.pdf','')}_dividido.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🔄 Dividir otro PDF", use_container_width=True, key="sp_reset"):
                        st.session_state.sp_result         = None
                        st.session_state.sp_sel_b          = set()
                        st.session_state.sp_mode           = "A"
                        st.session_state.sp_pages_per_part = 2
                        st.session_state.sp_sig            = ""
                        st.session_state.sp_thumbs         = []
                        st.session_state.sp_n              = 0
                        st.rerun()

            with grid_col:
                st.markdown('<div class="nx-section">🖼️ Páginas del PDF</div>', unsafe_allow_html=True)
                COLS = 6
                for row_start in range(0, n_pages, COLS):
                    row_cols = st.columns(COLS)
                    for ci, pi in enumerate(range(row_start, min(row_start + COLS, n_pages))):
                        with row_cols[ci]:
                            if mode == "B":
                                is_sel = pi in st.session_state.sp_sel_b
                                bdr_c  = "2px solid #27AE60" if is_sel else "2px solid rgba(27,159,216,0.2)"
                                ov     = ('<div style="position:absolute;inset:0;'
                                          'background:rgba(39,174,96,0.15);border-radius:8px;'
                                          'pointer-events:none;"></div>') if is_sel else ""
                                chk    = ('<div style="position:absolute;top:5px;left:5px;'
                                          'background:#27AE60;color:#fff;width:20px;height:20px;'
                                          'border-radius:5px;display:flex;align-items:center;'
                                          'justify-content:center;font-size:12px;font-weight:700;'
                                          'pointer-events:none;">✓</div>') if is_sel else (
                                          '<div style="position:absolute;top:5px;left:5px;'
                                          'background:rgba(255,255,255,0.06);'
                                          'border:2px solid rgba(27,159,216,0.25);'
                                          'width:20px;height:20px;border-radius:5px;'
                                          'pointer-events:none;"></div>'
                                )
                                st.markdown(
                                    f'<div style="border:{bdr_c};border-radius:12px;padding:5px;'
                                    f'background:linear-gradient(145deg,#09172A,#0A1626);'
                                    f'position:relative;margin-bottom:4px;">'
                                    f'{ov}{chk}'
                                    f'<img src="data:image/png;base64,{thumbs[pi]}" '
                                    f'style="width:100%;border-radius:7px;display:block;" draggable="false"/>'
                                    f'<div style="text-align:center;font-size:10px;color:#2E5878;'
                                    f'margin-top:4px;font-weight:600;">{pi+1}</div></div>',
                                    unsafe_allow_html=True
                                )
                                if is_sel:
                                    if st.button("✓", key=f"sp_b_{pi}", use_container_width=True):
                                        st.session_state.sp_sel_b.discard(pi)
                                        st.rerun()
                                else:
                                    if st.button("+", key=f"sp_b_{pi}", use_container_width=True):
                                        st.session_state.sp_sel_b.add(pi)
                                        st.rerun()
                            elif mode == "C":
                                _ppp3 = st.session_state.sp_pages_per_part
                                _part_num = pi // _ppp3
                                _bg_c = _SP_SEC[_part_num % len(_SP_SEC)]
                                st.markdown(
                                    _thumb_card(thumbs[pi], pi + 1, selected=False,
                                               badge_label=f"P{_part_num+1}",
                                               badge_bg=_bg_c, badge_fg="#fff"),
                                    unsafe_allow_html=True
                                )
                            else:
                                st.markdown(_thumb_card(thumbs[pi], pi + 1), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  TAB 3 — COMPRIMIR PDF
# ══════════════════════════════════════════════════════
with tabs[3]:
    _cs3_col, _cs3_panel = st.columns([3, 1])
    with _cs3_col:
        st.markdown("""
    <div class="nx-coming-soon">
        <div class="nx-cs-icon">🗜️</div>
        <div class="nx-cs-title">Comprimir PDF</div>
        <div class="nx-cs-sub">Reduce el tamaño de tus PDFs sin perder calidad visible.</div>
        <div class="nx-cs-badge">Próximamente</div>
    </div>""", unsafe_allow_html=True)
    with _cs3_panel:
        with st.container(border=True):
            st.markdown('<div style="font-size:14px;font-weight:700;color:#C8E4F0;margin-bottom:8px;">🗜️ Comprimir PDF</div>', unsafe_allow_html=True)
            st.markdown('<div style="font-size:12px;color:#2A4A6A;">Disponible próximamente.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  TAB 4 — MERGE PDF
# ══════════════════════════════════════════════════════
with tabs[4]:
    _cs4_col, _cs4_panel = st.columns([3, 1])
    with _cs4_col:
        st.markdown("""
    <div class="nx-coming-soon">
        <div class="nx-cs-icon">🔗</div>
        <div class="nx-cs-title">Merge PDF</div>
        <div class="nx-cs-sub">Combina PDFs con opciones avanzadas de intercalado y portada.</div>
        <div class="nx-cs-badge">Próximamente</div>
    </div>""", unsafe_allow_html=True)
    with _cs4_panel:
        with st.container(border=True):
            st.markdown('<div style="font-size:14px;font-weight:700;color:#C8E4F0;margin-bottom:8px;">🔗 Merge PDF</div>', unsafe_allow_html=True)
            st.markdown('<div style="font-size:12px;color:#2A4A6A;">Disponible próximamente.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
#  TAB 5 — EDITAR PDF (AVANZADO, ESTILO ADOBE)
# ══════════════════════════════════════════════════════
with tabs[5]:

    for _k, _v in [
        ("ed_sig",         ""),
        ("ed_thumbs",      []),
        ("ed_n",           0),
        ("ed_cur_page",    0),
        ("ed_canvas_data", {}),
        ("ed_result",      None),
    ]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">&#9999;&#65039; Editar PDF</div>
        <div class="nx-page-sub">
            Editor visual avanzado estilo Adobe Acrobat: añade
            <strong>texto</strong>, <strong>formas</strong>, <strong>flechas</strong>,
            <strong>resaltados</strong>, <strong>sellos</strong> y <strong>redacciones</strong>
            directamente sobre el PDF. Guarda página a página y descarga el resultado final.
        </div>
    </div>
    """, unsafe_allow_html=True)

    ed_file = st.file_uploader("Sube el PDF a editar", type=["pdf"], key="ed_uploader")

    if not ed_file:
        st.session_state.ed_canvas_data = {}
        st.session_state.ed_result      = None
        st.session_state.ed_cur_page    = 0
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">&#9999;&#65039;</div>
            <div class="nx-empty-text">Sube un PDF para empezar a editarlo</div>
            <div class="nx-empty-sub">Editor visual con herramientas de anotación avanzadas — estilo Adobe Acrobat</div>
        </div>""", unsafe_allow_html=True)
    else:
        if not FITZ_OK:
            st.error("PyMuPDF no está instalado.")
        else:
            ed_bytes = ed_file.read()
            thumbs_ed, n_pages_ed = _ensure_thumbs("ed", ed_bytes)

            cur_page = max(0, min(st.session_state.ed_cur_page, n_pages_ed - 1))
            st.session_state.ed_cur_page = cur_page

            CANVAS_W = 820
            _doc_ed  = fitz.open(stream=ed_bytes, filetype="pdf")
            _pg_ed   = _doc_ed[cur_page]
            pw_r     = _pg_ed.rect.width
            ph_r     = _pg_ed.rect.height
            canvas_h = int(ph_r * CANVAS_W / pw_r)
            _pix_bg  = _pg_ed.get_pixmap(
                matrix=fitz.Matrix(CANVAS_W / pw_r, canvas_h / ph_r), alpha=False
            )
            _doc_ed.close()
            _bg_b64 = base64.b64encode(_pix_bg.tobytes("png")).decode()

            _cur_els   = st.session_state.ed_canvas_data.get(cur_page, [])
            _init_json = _json.dumps(_cur_els)

            _canvas_html = (
                _CANVAS_TMPL
                .replace("___BG___",   _bg_b64)
                .replace("___CW___",   str(CANVAS_W))
                .replace("___CH___",   str(canvas_h))
                .replace("___INIT___", _init_json)
            )

            # Layout: nav izq | canvas | panel der
            nav_col, canvas_col, props_col = st.columns([1, 4, 1.5])

            with nav_col:
                st.markdown(
                    '<div style="font-size:11px;font-weight:700;color:#1B9FD8;'
                    'text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;">P&aacute;ginas</div>',
                    unsafe_allow_html=True
                )
                for pi in range(n_pages_ed):
                    _is_cur = pi == cur_page
                    _n_el   = len(st.session_state.ed_canvas_data.get(pi, []))
                    _bdr_pn = "2px solid #1B9FD8" if _is_cur else "1px solid rgba(27,159,216,0.12)"
                    _bg_pn  = "rgba(27,159,216,0.12)" if _is_cur else "linear-gradient(145deg,#09172A,#0A1626)"
                    _bdg_pn = (f'<div style="position:absolute;bottom:4px;right:4px;'
                                f'background:linear-gradient(135deg,#1B9FD8,#1580B8);'
                                f'color:#fff;font-size:9px;font-weight:700;padding:1px 6px;'
                                f'border-radius:6px;">{_n_el}</div>') if _n_el > 0 else ""
                    st.markdown(
                        f'<div style="border:{_bdr_pn};border-radius:10px;padding:5px;'
                        f'background:{_bg_pn};position:relative;margin-bottom:5px;">'
                        f'<img src="data:image/png;base64,{thumbs_ed[pi]}" '
                        f'style="width:100%;border-radius:6px;display:block;"/>'
                        f'{_bdg_pn}'
                        f'<div style="text-align:center;font-size:10px;color:#2E5878;'
                        f'margin-top:3px;font-weight:600;">{pi+1}</div></div>',
                        unsafe_allow_html=True
                    )
                    if not _is_cur:
                        if st.button("Ver", key=f"ed_nav_{pi}", use_container_width=True):
                            st.session_state.ed_cur_page = pi
                            st.rerun()

            with props_col:
                with st.container(border=True):
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:700;color:#C8E4F0;margin-bottom:4px;">'
                        f'&#9999; Editor PDF</div>'
                        f'<div style="font-size:11px;color:#4A7A9C;margin-bottom:12px;">'
                        f'P&aacute;gina {cur_page+1} de {n_pages_ed}</div>',
                        unsafe_allow_html=True
                    )

                    _total_els = sum(len(v) for v in st.session_state.ed_canvas_data.values())
                    st.markdown(
                        f'<div style="background:rgba(27,159,216,0.06);border:1px solid rgba(27,159,216,0.12);'
                        f'border-radius:10px;padding:10px;text-align:center;margin-bottom:12px;">'
                        f'<div style="font-size:24px;font-weight:800;color:#1B9FD8;">{_total_els}</div>'
                        f'<div style="font-size:10px;color:#2A4A6A;margin-top:2px;">elementos totales</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        '<div style="font-size:11px;color:#4A7A9C;margin-bottom:10px;line-height:1.5;">'
                        'Usa <strong style="color:#1B9FD8">Guardar página</strong> '
                        'en la barra del canvas para confirmar cada página. '
                        'Luego aplica todos los cambios al PDF:</div>',
                        unsafe_allow_html=True
                    )

                    # Atajos de teclado
                    with st.expander("⌨️ Atajos de teclado"):
                        st.markdown("""
                        | Tecla | Acción |
                        |-------|--------|
                        | `V` | Seleccionar |
                        | `R` | Rectángulo |
                        | `E` | Elipse |
                        | `L` | Línea |
                        | `T` | Texto |
                        | `H` | Resaltar |
                        | `Del` | Eliminar |
                        | `Ctrl+Z` | Deshacer |
                        | `Ctrl+Wheel` | Zoom |
                        """)

                    if st.button("&#128196; Aplicar al PDF", type="primary",
                                  use_container_width=True, key="ed_save_pdf",
                                  disabled=_total_els == 0):
                        with st.spinner("Generando PDF con anotaciones…"):
                            try:
                                _doc_sv = fitz.open(stream=ed_bytes, filetype="pdf")
                                for _pi, _page_els in st.session_state.ed_canvas_data.items():
                                    if not _page_els:
                                        continue
                                    _pg_sv = _doc_sv[_pi]
                                    _pw_sv = _pg_sv.rect.width
                                    _ph_sv = _pg_sv.rect.height
                                    _ch_sv = int(_ph_sv * CANVAS_W / _pw_sv)
                                    _sxp   = _pw_sv / CANVAS_W
                                    _syp   = _ph_sv / _ch_sv
                                    for _el in _page_els:
                                        _et  = _el.get("type", "")
                                        _sc  = _parse_fabric_color(_el.get("sc", "#000000"))
                                        _lw  = max(0.5, (_el.get("sw", 2)) * _sxp)
                                        _fcs = _el.get("fc", "")
                                        _fc  = _parse_fabric_color(_fcs) if _fcs else None
                                        if _et == "rect":
                                            _rx = _el["x"]*_sxp; _ry = _el["y"]*_syp
                                            _rw = _el["w"]*_sxp; _rh = _el["h"]*_syp
                                            _pg_sv.draw_rect(
                                                fitz.Rect(_rx,_ry,_rx+_rw,_ry+_rh),
                                                color=_sc, fill=_fc, width=_lw
                                            )
                                        elif _et == "ell":
                                            _rx = _el["x"]*_sxp; _ry = _el["y"]*_syp
                                            _rw = _el["w"]*_sxp; _rh = _el["h"]*_syp
                                            _pg_sv.draw_oval(
                                                fitz.Rect(_rx,_ry,_rx+_rw,_ry+_rh),
                                                color=_sc, fill=_fc, width=_lw
                                            )
                                        elif _et in ("line", "arrow"):
                                            _pg_sv.draw_line(
                                                fitz.Point(_el["x1"]*_sxp, _el["y1"]*_syp),
                                                fitz.Point(_el["x2"]*_sxp, _el["y2"]*_syp),
                                                color=_sc, width=_lw
                                            )
                                        elif _et == "text":
                                            _fmap = {
                                                "Arial":"helv","Helvetica":"helv",
                                                "Inter / Arial":"helv",
                                                "Times New Roman":"tiro","Times":"tiro","Georgia":"tiro",
                                                "Courier New":"cour","Courier":"cour"
                                            }
                                            _tc    = _parse_fabric_color(_el.get("tc","#000000"))
                                            _fs_pt = max(4, _el.get("fs",14) * _sxp)
                                            _py    = _el["y"] * _syp + _fs_pt * 0.8
                                            _pg_sv.insert_text(
                                                fitz.Point(_el["x"]*_sxp, _py),
                                                _el.get("txt",""),
                                                fontname=_fmap.get(_el.get("ff","Arial"),"helv"),
                                                fontsize=_fs_pt,
                                                color=_tc or (0,0,0)
                                            )
                                        elif _et == "hl":
                                            _hx = _el["x"]*_sxp; _hy = _el["y"]*_syp
                                            _hw = _el["w"]*_sxp; _hh = _el["h"]*_syp
                                            _pg_sv.draw_rect(
                                                fitz.Rect(_hx,_hy,_hx+_hw,_hy+_hh),
                                                color=(1,0.9,0.1), fill=(1,0.9,0.1), width=0, fill_opacity=0.35
                                            )
                                        elif _et == "redact":
                                            _rx2 = _el["x"]*_sxp; _ry2 = _el["y"]*_syp
                                            _rw2 = _el["w"]*_sxp; _rh2 = _el["h"]*_syp
                                            _pg_sv.draw_rect(
                                                fitz.Rect(_rx2,_ry2,_rx2+_rw2,_ry2+_rh2),
                                                color=(0,0,0), fill=(0,0,0), width=0
                                            )
                                        elif _et == "stamp":
                                            _sx2 = _el["x"]*_sxp; _sy2 = _el["y"]*_syp
                                            _sw2 = _el["w"]*_sxp; _sh2 = _el["h"]*_syp
                                            _stc = _parse_fabric_color(_el.get("color","#E74C3C"))
                                            _pg_sv.draw_rect(
                                                fitz.Rect(_sx2,_sy2,_sx2+_sw2,_sy2+_sh2),
                                                color=_stc, width=2.5
                                            )
                                            _sfs = min(16, _sh2*0.5)
                                            _pg_sv.insert_text(
                                                fitz.Point(_sx2+10, _sy2+_sh2/2-_sfs*0.4),
                                                _el.get("txt","SELLO"),
                                                fontname="helv", fontsize=_sfs,
                                                color=_stc
                                            )

                                _buf_sv = BytesIO()
                                _doc_sv.save(_buf_sv)
                                _doc_sv.close()
                                _buf_sv.seek(0)
                                st.session_state.ed_result = _buf_sv.getvalue()
                                st.session_state["_nx_docs_count"] = st.session_state.get("_nx_docs_count",0) + 1
                                st.rerun()
                            except Exception as _e:
                                st.error(f"Error: {_e}")

                    if st.session_state.ed_result:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.download_button(
                            label="⬇️ Descargar PDF Editado",
                            data=st.session_state.ed_result,
                            file_name=ed_file.name.replace(".pdf","_editado.pdf"),
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True
                        )
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("&#128260; Reiniciar edición", use_container_width=True, key="ed_rst"):
                            st.session_state.ed_canvas_data = {}
                            st.session_state.ed_result      = None
                            st.session_state.ed_cur_page    = 0
                            st.rerun()

                # Receptor oculto de datos del canvas
                st.markdown("---")
                st.markdown('<style>[data-testid="stTextArea"]{display:none!important;}hr{display:none!important;}</style>', unsafe_allow_html=True)
                _raw_json = st.text_area(
                    "NEXA_CANVAS_DATA",
                    key="ed_raw_json",
                    height=50,
                    label_visibility="collapsed",
                    placeholder="(datos del canvas)"
                )
                if _raw_json and _raw_json.strip().startswith("{"):
                    try:
                        _cdata    = _json.loads(_raw_json.strip())
                        _page_els = _cdata.get("elements", [])
                        st.session_state.ed_canvas_data[cur_page] = _page_els
                    except Exception:
                        pass
                    st.session_state["ed_raw_json"] = ""
                    st.rerun()

            with canvas_col:
                st.components.v1.html(_canvas_html, height=canvas_h + 260, scrolling=False)

# ══════════════════════════════════════════════════════
#  TAB 6 — ELIMINAR PÁGINAS
# ══════════════════════════════════════════════════════
with tabs[6]:

    for _k, _v in [("ep_sel", set()), ("ep_result", None),
                   ("ep_sig", ""), ("ep_thumbs", []), ("ep_n", 0)]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">🗑️ Eliminar Páginas</div>
        <div class="nx-page-sub">Selecciona visualmente las páginas que quieres
        <strong>eliminar</strong> y descarga el PDF resultante.</div>
    </div>
    """, unsafe_allow_html=True)

    ep_file = st.file_uploader("Sube el PDF", type=["pdf"], key="ep_uploader")

    if not ep_file:
        st.session_state.ep_sel    = set()
        st.session_state.ep_result = None
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">🗑️</div>
            <div class="nx-empty-text">Sube un PDF para eliminar páginas</div>
            <div class="nx-empty-sub">Verás todas las páginas y podrás seleccionar las que quieres borrar</div>
        </div>""", unsafe_allow_html=True)
    else:
        if not FITZ_OK:
            st.error("⚠️ PyMuPDF no está instalado.")
        else:
            _ep_col, _ep_panel = st.columns([3, 1])

            with _ep_panel:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:15px;font-weight:700;color:#C8E4F0;margin-bottom:12px;">🗑️ Eliminar páginas</div>',
                        unsafe_allow_html=True
                    )
                    _ep_n_pag = st.session_state.get("ep_n", 0)
                    if _ep_n_pag > 0:
                        st.markdown(
                            f'<div style="font-size:12px;color:#4A7A9C;margin-bottom:4px;">'
                            f'Páginas totales: <strong style="color:#C8E4F0;">{_ep_n_pag}</strong></div>',
                            unsafe_allow_html=True
                        )
                    st.markdown('<div style="font-size:11px;color:#4A7A9C;margin:10px 0 4px 0;">Páginas para quitar:</div>', unsafe_allow_html=True)
                    _ep_rng = st.text_input(
                        "Páginas para quitar",
                        key="ep_rng_txt",
                        placeholder="ejemplo: 1,5-8",
                        label_visibility="collapsed"
                    )
                    _ep_n_pag2 = st.session_state.get("ep_n", 0)
                    if _ep_rng and _ep_n_pag2 > 0:
                        try:
                            _new_sel = set()
                            for _part in _ep_rng.replace(" ", "").split(","):
                                if "-" in _part:
                                    _a, _b = _part.split("-", 1)
                                    _new_sel.update(range(int(_a)-1, int(_b)))
                                elif _part:
                                    _new_sel.add(int(_part)-1)
                            _new_sel = {p for p in _new_sel if 0 <= p < _ep_n_pag2}
                            if _new_sel != st.session_state.ep_sel:
                                st.session_state.ep_sel = _new_sel
                                st.rerun()
                        except Exception:
                            st.warning("Formato inválido", icon="⚠️")

                    _ep_n_sel2  = len(st.session_state.ep_sel)
                    _ep_n_pag3  = st.session_state.get("ep_n", 0)
                    st.markdown(
                        f'<div style="font-size:12px;color:{"#E74C3C" if _ep_n_sel2>0 else "#4A7A9C"};margin:10px 0;">'
                        f'{"🗑️ " if _ep_n_sel2>0 else ""}{_ep_n_sel2} página{"s" if _ep_n_sel2!=1 else ""} seleccionada{"s" if _ep_n_sel2!=1 else ""}</div>',
                        unsafe_allow_html=True
                    )
                    _ep_keep2 = _ep_n_pag3 - _ep_n_sel2
                    if _ep_n_sel2 > 0 and _ep_keep2 > 0 and not st.session_state.get("ep_result"):
                        if st.button(
                            f"🗑️ Eliminar {_ep_n_sel2} página{'s' if _ep_n_sel2>1 else ''}",
                            type="primary", use_container_width=True, key="ep_do_delete_panel",
                        ):
                            with st.spinner("Generando PDF…"):
                                try:
                                    _ep_raw = st.session_state.get("_ep_bytes_cache", b"")
                                    if _ep_raw:
                                        _rdr2 = PdfReader(BytesIO(_ep_raw))
                                        _wrt2 = PdfWriter()
                                        _ep_tot = _ep_n_pag3
                                        for _pi2 in range(_ep_tot):
                                            if _pi2 not in st.session_state.ep_sel:
                                                _wrt2.add_page(_rdr2.pages[_pi2])
                                        _buf2 = BytesIO()
                                        _wrt2.write(_buf2)
                                        _buf2.seek(0)
                                        st.session_state.ep_result = _buf2.getvalue()
                                        st.rerun()
                                except Exception as _e2:
                                    st.error(f"Error: {_e2}")
                    elif _ep_n_sel2 == 0:
                        st.info("Selecciona páginas.", icon="ℹ️")
                    elif _ep_n_pag3 > 0 and _ep_keep2 == 0:
                        st.warning("Conserva al menos una.", icon="⚠️")

            with _ep_col:
                ep_bytes = ep_file.read()
                st.session_state["_ep_bytes_cache"] = ep_bytes
                thumbs_ep, n_pages_ep = _ensure_thumbs("ep", ep_bytes, "ep_sel")
                n_sel = len(st.session_state.ep_sel)

                if st.session_state.ep_result:
                    ep_pages_kept = n_pages_ep - len(st.session_state.ep_sel)
                    ep_name = ep_file.name.replace(".pdf", "_sin_paginas.pdf")
                    st.markdown("""
                    <div class="nx-success-card">
                        <div class="nx-success-icon">✅</div>
                        <div class="nx-success-title">¡Páginas eliminadas!</div>
                        <div class="nx-success-sub">El PDF resultante está listo para descargar.</div>
                    </div>""", unsafe_allow_html=True)
                    st.download_button(
                        label=f"⬇️ Descargar PDF ({ep_pages_kept} páginas)",
                        data=st.session_state.ep_result,
                        file_name=ep_name,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🔄 Eliminar páginas de otro PDF", use_container_width=True, key="ep_reset"):
                        st.session_state.ep_result = None
                        st.session_state.ep_sel    = set()
                        st.session_state.ep_sig    = ""
                        st.session_state.ep_thumbs = []
                        st.session_state.ep_n      = 0
                        st.rerun()
                else:
                    st.markdown('<div class="nx-section">🎯 Selección rápida por rango</div>', unsafe_allow_html=True)
                    _rc1, _rc2, _rc3, _rc4 = st.columns([1, 1, 1, 1])
                    with _rc1:
                        ep_from = st.number_input("Desde página", 1, n_pages_ep, 1, key="ep_from")
                    with _rc2:
                        ep_to   = st.number_input("Hasta página", 1, n_pages_ep, n_pages_ep, key="ep_to")
                    with _rc3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("Seleccionar rango", use_container_width=True, key="ep_sel_range"):
                            for i in range(int(ep_from)-1, int(ep_to)):
                                st.session_state.ep_sel.add(i)
                            st.rerun()
                    with _rc4:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("Deseleccionar rango", use_container_width=True, key="ep_desel_range"):
                            for i in range(int(ep_from)-1, int(ep_to)):
                                st.session_state.ep_sel.discard(i)
                            st.rerun()

                    st.markdown('<div class="nx-section">🗑️ Selecciona páginas para eliminar</div>', unsafe_allow_html=True)
                    col_info, col_all, col_none = st.columns([3, 1, 1])
                    with col_info:
                        if n_sel == 0:
                            sel_txt = f"Ninguna página seleccionada de {n_pages_ep}"
                        else:
                            sel_txt = (f"**{n_sel}** página{'s' if n_sel!=1 else ''} "
                                       f"seleccionada{'s' if n_sel!=1 else ''} para eliminar")
                        st.info(sel_txt, icon="🗑️")
                    with col_all:
                        if st.button("Seleccionar todas", use_container_width=True, key="ep_all"):
                            st.session_state.ep_sel = set(range(n_pages_ep))
                            st.rerun()
                    with col_none:
                        if st.button("Deseleccionar todas", use_container_width=True, key="ep_none"):
                            st.session_state.ep_sel = set()
                            st.rerun()

                    COLS = 6
                    for row_start in range(0, n_pages_ep, COLS):
                        cols_ep = st.columns(COLS)
                        for ci, pi in enumerate(range(row_start, min(row_start + COLS, n_pages_ep))):
                            is_sel = pi in st.session_state.ep_sel
                            with cols_ep[ci]:
                                st.markdown(
                                    _thumb_card(thumbs_ep[pi], pi + 1, selected=is_sel, mode="delete"),
                                    unsafe_allow_html=True
                                )
                                if is_sel:
                                    if st.button("✕", key=f"ep_{pi}", use_container_width=True):
                                        st.session_state.ep_sel.discard(pi)
                                        st.rerun()
                                else:
                                    if st.button("☐", key=f"ep_{pi}", use_container_width=True):
                                        st.session_state.ep_sel.add(pi)
                                        st.rerun()

                    st.markdown("---")
                    pages_to_keep = n_pages_ep - n_sel
                    if n_sel == 0:
                        st.info("Selecciona al menos una página para eliminar.", icon="ℹ️")
                    elif pages_to_keep == 0:
                        st.warning("⚠️ No puedes eliminar todas las páginas. Debes conservar al menos una.")
                    else:
                        if st.button(
                            f"🗑️ Eliminar {n_sel} página{'s' if n_sel>1 else ''} seleccionada{'s' if n_sel>1 else ''} "
                            f"(quedan {pages_to_keep})",
                            type="primary", use_container_width=True, key="ep_do_delete"
                        ):
                            with st.spinner("Generando PDF…"):
                                try:
                                    reader = PdfReader(BytesIO(ep_bytes))
                                    writer = PdfWriter()
                                    for pi in range(n_pages_ep):
                                        if pi not in st.session_state.ep_sel:
                                            writer.add_page(reader.pages[pi])
                                    ep_buf = BytesIO()
                                    writer.write(ep_buf)
                                    ep_buf.seek(0)
                                    st.session_state.ep_result = ep_buf.getvalue()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error: {e}")

# ══════════════════════════════════════════════════════
#  TAB 7 — VALIDADOR IA
# ══════════════════════════════════════════════════════
with tabs[7]:

    st.markdown("""
    <div class="nx-page-header">
      <div class="nx-page-title">🤖 Validador <span style="color:#00C2CB">Inteligente</span> de Documentos</div>
      <div class="nx-page-sub">Valida automáticamente si tu expediente cumple los requisitos antes de radicar. <strong>Reduce devoluciones a cero.</strong></div>
    </div>
    """, unsafe_allow_html=True)

    col_conf, col_upload = st.columns([1, 1], gap="large")

    with col_conf:
        st.markdown('<div class="nx-section">⚙️ Configuración</div>', unsafe_allow_html=True)
        tipo_tramite = st.selectbox("Tipo de trámite", ["Desembolso Subsidio Colsubsidio", "Desembolso Caja Compensación", "Personalizado"])
        entidad = st.text_input("Entidad receptora", placeholder="Ej: Colsubsidio, Comfama...")
        criterios_extra = st.text_area("Criterios adicionales (opcional)", placeholder="Agrega criterios específicos de validación...", height=140)

    with col_upload:
        st.markdown('<div class="nx-section">📄 Expediente</div>', unsafe_allow_html=True)
        archivo_pdf = st.file_uploader("Sube el expediente PDF completo", type=["pdf"], key="validador_ia_uploader")
        if archivo_pdf:
            st.success(f"✅ {archivo_pdf.name} cargado")
            pdf_bytes = archivo_pdf.read()
            texto_pdf, paginas_pdf, num_paginas = extraer_texto_pdf(pdf_bytes)
            st.info(f"📄 {num_paginas} páginas listas para análisis")

    st.markdown("---")
    col_b1, col_b2, col_b3 = st.columns([1,2,1])
    with col_b2:
        btn_validar = st.button("🤖 Ejecutar Validación IA", type="primary", use_container_width=True, disabled=not archivo_pdf if 'archivo_pdf' in dir() else True)

    if 'btn_validar' in dir() and btn_validar and archivo_pdf:
        if not texto_pdf:
            st.error("❌ No se pudo extraer texto del PDF.")
        else:
            with st.spinner("🔍 Analizando expediente con IA... 20-30 segundos"):
                try:
                    resultado = validar_con_ia(texto_pdf, tipo_tramite, criterios_extra)
                    st.session_state["ultimo_resultado"] = resultado
                    st.session_state["ultimo_archivo"] = archivo_pdf.name
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    resultado = None

            if resultado:
                score = resultado.get("score", 0)
                res_general = resultado.get("resultado_general", "REQUIERE_REVISION")
                resumen = resultado.get("resumen", "")
                colores = {"APROBADO":"#00C48C","REQUIERE_REVISION":"#F5A623","RECHAZADO":"#F04747"}
                emojis = {"APROBADO":"✅","REQUIERE_REVISION":"⚠️","RECHAZADO":"❌"}
                color = colores.get(res_general, "#F5A623")
                emoji = emojis.get(res_general, "⚠️")

                st.markdown(f"""
                <div style="background:linear-gradient(135deg,#0A0F1E,#1A2234);border-radius:16px;padding:28px;margin:16px 0;text-align:center;border:1px solid rgba(0,194,203,0.2);">
                    <div style="font-size:48px;margin-bottom:8px;">{emoji}</div>
                    <div style="font-size:26px;font-weight:800;color:{color};margin-bottom:6px;">{res_general.replace('_',' ')}</div>
                    <div style="font-size:44px;font-weight:800;color:#fff;line-height:1;">{score}<span style="font-size:20px;color:rgba(255,255,255,0.4)">/100</span></div>
                    <div style="font-size:13px;color:rgba(255,255,255,0.5);margin-top:8px;">{resumen}</div>
                </div>
                """, unsafe_allow_html=True)

                hallazgos = resultado.get("hallazgos", [])
                if hallazgos:
                    st.markdown('<div class="nx-section">🔍 Hallazgos</div>', unsafe_allow_html=True)
                    c1,c2,c3,c4 = st.columns(4)
                    c1.metric("🔴 Críticos", len([h for h in hallazgos if h.get("severidad")=="CRITICO"]))
                    c2.metric("🟡 Medios", len([h for h in hallazgos if h.get("severidad")=="MEDIO"]))
                    c3.metric("🔵 Bajos", len([h for h in hallazgos if h.get("severidad")=="BAJO"]))
                    c4.metric("✅ OK", len([h for h in hallazgos if h.get("severidad")=="OK"]))

                    for h in hallazgos:
                        sev = h.get("severidad","BAJO")
                        col_sev = {"CRITICO":"#F04747","MEDIO":"#F5A623","BAJO":"#1B6FE8","OK":"#00C48C"}.get(sev,"#8494A8")
                        em = {"CUMPLE":"✅","NO_CUMPLE":"❌","NO_ENCONTRADO":"❓","ADVERTENCIA":"⚠️"}.get(h.get("estado",""),"❓")
                        pag = f" · Pág. {h.get('pagina_estimada')}" if h.get("pagina_estimada") else ""
                        st.markdown(f"""
                        <div style="background:#fff;border:1px solid rgba(10,15,30,0.08);border-left:4px solid {col_sev};border-radius:10px;padding:14px 16px;margin-bottom:8px;">
                            <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;">
                                <div style="flex:1;">
                                    <div style="font-size:13px;font-weight:700;color:#0A0F1E;margin-bottom:2px;">{em} {h.get('criterio','')}</div>
                                    <div style="font-size:11px;color:#8494A8;margin-bottom:5px;">📄 {h.get('documento','')}{pag}</div>
                                    <div style="font-size:12px;color:#556070;line-height:1.5;">{h.get('observacion','')}</div>
                                </div>
                                <div style="background:{col_sev};color:#fff;font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap;">{sev}</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                faltantes = resultado.get("documentos_faltantes", [])
                if faltantes:
                    st.markdown('<div class="nx-section">📋 Documentos faltantes</div>', unsafe_allow_html=True)
                    for d in faltantes:
                        st.markdown(f'<div style="background:rgba(240,71,71,0.05);border:1px solid rgba(240,71,71,0.2);border-radius:8px;padding:10px 14px;margin-bottom:6px;font-size:13px;color:#F04747;">❌ {d}</div>', unsafe_allow_html=True)

                recomendaciones = resultado.get("recomendaciones", [])
                if recomendaciones:
                    st.markdown('<div class="nx-section">💡 Acciones antes de radicar</div>', unsafe_allow_html=True)
                    for i, r in enumerate(recomendaciones, 1):
                        st.markdown(f'<div style="background:rgba(0,194,203,0.04);border:1px solid rgba(0,194,203,0.15);border-radius:8px;padding:10px 14px;margin-bottom:6px;font-size:13px;color:#0A0F1E;"><strong style="color:#00A8B0">{i}.</strong> {r}</div>', unsafe_allow_html=True)

                st.markdown("---")
                import json as json_lib
                reporte = f"""REPORTE VALIDACIÓN NEXA
{'='*50}
Archivo: {archivo_pdf.name}
Trámite: {tipo_tramite}
Entidad: {entidad or 'No especificada'}
Resultado: {res_general}
Score: {score}/100
Resumen: {resumen}

HALLAZGOS:
{chr(10).join([f"- [{h.get('severidad')}] {h.get('criterio')}: {h.get('observacion')}" for h in hallazgos])}

FALTANTES:
{chr(10).join([f"- {d}" for d in faltantes]) if faltantes else "Ninguno"}

RECOMENDACIONES:
{chr(10).join([f"{i+1}. {r}" for i, r in enumerate(recomendaciones)])}"""

                cd1, cd2 = st.columns(2)
                with cd1:
                    st.download_button("📄 Descargar reporte TXT", data=reporte, file_name=f"validacion_{archivo_pdf.name.replace('.pdf','')}.txt", mime="text/plain", use_container_width=True)
                with cd2:
                    st.download_button("📊 Descargar reporte JSON", data=json_lib.dumps(resultado, ensure_ascii=False, indent=2), file_name=f"validacion_{archivo_pdf.name.replace('.pdf','')}.json", mime="application/json", use_container_width=True)

    elif not archivo_pdf if 'archivo_pdf' in dir() else True:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">🤖</div>
            <div class="nx-empty-text">Sube un expediente PDF para comenzar</div>
            <div class="nx-empty-sub">El Validador IA analizará cada documento y te dirá exactamente qué corregir antes de radicar</div>
        </div>
        """, unsafe_allow_html=True)

