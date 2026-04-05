import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
import base64
import hashlib
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="NEXA - Transformación de Procesos", page_icon="⚙️", layout="wide")

# ==========================================
# NEXA DESIGN SYSTEM — CSS Global
# ==========================================
st.markdown("""
<style>
/* ══════════════════════════════════════════════════════
   BASE
══════════════════════════════════════════════════════ */
html, body, .stApp {
    background-color: #060F1D !important;
    color: #C8E4F0 !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
}

/* Ocultar chrome de Streamlit */
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"]  { display: none !important; }
#MainMenu                      { visibility: hidden !important; }
footer                         { visibility: hidden !important; }
.stDeployButton                { display: none !important; }

/* Ocultar sidebar completamente */
[data-testid="stSidebar"]      { display: none !important; }

/* Área principal */
[data-testid="stAppViewContainer"] > [data-testid="stMain"] {
    background: #060F1D !important;
}
.block-container {
    padding-top: 0 !important;
    padding-bottom: 3rem !important;
    max-width: 1200px !important;
}

/* ══════════════════════════════════════════════════════
   BARRA DE NAVEGACIÓN (TABS)
══════════════════════════════════════════════════════ */
.stTabs [data-baseweb="tab-list"] {
    background: #07111C !important;
    border-radius: 10px !important;
    padding: 5px 6px !important;
    gap: 3px !important;
    border: 1px solid rgba(27,159,216,0.1) !important;
    margin-bottom: 24px !important;
    flex-wrap: wrap !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #2A4060 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    border-radius: 7px !important;
    padding: 7px 15px !important;
    border: none !important;
    transition: background .15s, color .15s !important;
    white-space: nowrap !important;
}
.stTabs [data-baseweb="tab"]:hover {
    background: rgba(27,159,216,0.08) !important;
    color: #5A9FC4 !important;
}
.stTabs [aria-selected="true"][data-baseweb="tab"] {
    background: rgba(27,159,216,0.14) !important;
    color: #1B9FD8 !important;
    font-weight: 700 !important;
}
.stTabs [data-baseweb="tab-highlight"] {
    display: none !important;
}
.stTabs [data-baseweb="tab-border"] {
    display: none !important;
}
/* Contenido de los tabs */
.stTabs [data-testid="stTabsContent"] {
    padding-top: 0 !important;
}

/* ══════════════════════════════════════════════════════
   TIPOGRAFÍA
══════════════════════════════════════════════════════ */
h1 { color: #E0F0FF !important; font-weight: 700 !important; font-size: 1.55rem !important; letter-spacing: -0.2px !important; }
h2 { color: #B8D8EE !important; }
h3 { color: #9ABDD8 !important; }
p  { color: #8AAEC8 !important; }
strong { color: #C8E4F0 !important; }

/* ══════════════════════════════════════════════════════
   BOTONES
══════════════════════════════════════════════════════ */
[data-testid="baseButton-primary"],
button[kind="primary"] {
    background: #1B9FD8 !important;
    border: none !important;
    color: #fff !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    letter-spacing: 0.2px !important;
    box-shadow: 0 2px 14px rgba(27,159,216,0.28) !important;
    transition: all .18s ease !important;
}
[data-testid="baseButton-primary"]:hover,
button[kind="primary"]:hover {
    background: #1790C5 !important;
    box-shadow: 0 4px 24px rgba(27,159,216,0.46) !important;
    transform: translateY(-1px) !important;
}
[data-testid="baseButton-primary"]:active,
button[kind="primary"]:active { transform: translateY(0) !important; }

[data-testid="baseButton-secondary"],
button[kind="secondary"] {
    background: rgba(27,159,216,0.07) !important;
    border: 1px solid rgba(27,159,216,0.22) !important;
    color: #4A8FB0 !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    transition: all .18s ease !important;
}
[data-testid="baseButton-secondary"]:hover,
button[kind="secondary"]:hover {
    background: rgba(27,159,216,0.13) !important;
    border-color: rgba(27,159,216,0.45) !important;
    color: #1B9FD8 !important;
}

/* ══════════════════════════════════════════════════════
   INPUTS DE TEXTO
══════════════════════════════════════════════════════ */
[data-testid="stTextInput"] label { color: #2E5070 !important; font-size: 12px !important; font-weight: 600 !important; }
[data-testid="stTextInput"] > div > div > input {
    background: #07111C !important;
    border: 1px solid rgba(27,159,216,0.18) !important;
    color: #C8E4F0 !important;
    border-radius: 8px !important;
    padding: 10px 14px !important;
    font-size: 14px !important;
    caret-color: #1B9FD8 !important;
    transition: border-color .15s, box-shadow .15s !important;
}
[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #1B9FD8 !important;
    box-shadow: 0 0 0 3px rgba(27,159,216,0.1) !important;
    outline: none !important;
}
[data-testid="stTextInput"] > div > div > input::placeholder { color: #1A3450 !important; }

/* ══════════════════════════════════════════════════════
   FILE UPLOADER
══════════════════════════════════════════════════════ */
[data-testid="stFileUploader"] section {
    background: rgba(27,159,216,0.03) !important;
    border: 2px dashed rgba(27,159,216,0.25) !important;
    border-radius: 12px !important;
    transition: border-color .2s, background .2s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: rgba(27,159,216,0.5) !important;
    background: rgba(27,159,216,0.06) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #1E3A58 !important;
    font-size: 13px !important;
}
[data-testid="stFileUploaderDropzone"] svg { color: #1B4060 !important; }

/* ══════════════════════════════════════════════════════
   CONTAINERS CON BORDE
══════════════════════════════════════════════════════ */
[data-testid="stVerticalBlockBorderWrapper"] > div {
    background: #0A1626 !important;
    border: 1px solid rgba(27,159,216,0.12) !important;
    border-radius: 12px !important;
    transition: border-color .18s !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div:hover {
    border-color: rgba(27,159,216,0.28) !important;
}

/* ══════════════════════════════════════════════════════
   DATA EDITOR
══════════════════════════════════════════════════════ */
[data-testid="stDataEditor"] {
    border: 1px solid rgba(27,159,216,0.14) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}

/* ══════════════════════════════════════════════════════
   ALERTAS / INFO
══════════════════════════════════════════════════════ */
[data-testid="stAlert"] {
    background: rgba(27,159,216,0.05) !important;
    border: 1px solid rgba(27,159,216,0.16) !important;
    border-radius: 10px !important;
    color: #3A6888 !important;
    font-size: 13px !important;
}
[data-testid="stAlert"] p      { color: #3A6888 !important; }
[data-testid="stAlert"] strong { color: #1B9FD8 !important; }

/* ══════════════════════════════════════════════════════
   PROGRESS / SPINNER / EXPANDER
══════════════════════════════════════════════════════ */
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #1B9FD8 0%, #1DE0C0 100%) !important;
}
[data-testid="stSpinner"] > div > div { border-top-color: #1B9FD8 !important; }
[data-testid="stExpander"] {
    background: #0A1626 !important;
    border: 1px solid rgba(27,159,216,0.12) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}
[data-testid="stExpanderDetails"] { background: #0A1626 !important; }

/* ══════════════════════════════════════════════════════
   CAPTIONS / HR / DOWNLOAD
══════════════════════════════════════════════════════ */
.stCaption, [data-testid="stCaptionContainer"] { color: #1E3A58 !important; font-size: 11px !important; }
hr { border-color: rgba(27,159,216,0.1) !important; }
[data-testid="stMarkdownContainer"] p { color: #4A7A9C !important; }

[data-testid="stDownloadButton"] > button {
    background: #1B9FD8 !important;
    border: none !important;
    color: #fff !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 14px rgba(27,159,216,0.28) !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #1790C5 !important;
    box-shadow: 0 4px 24px rgba(27,159,216,0.46) !important;
    transform: translateY(-1px) !important;
}

/* ══════════════════════════════════════════════════════
   COMPONENTES REUTILIZABLES NEXA
══════════════════════════════════════════════════════ */

/* Cabecera de página */
.nx-page-header {
    padding: 4px 0 20px 0;
    border-bottom: 1px solid rgba(27,159,216,0.1);
    margin-bottom: 22px;
}
.nx-page-title { font-size: 22px; font-weight: 700; color: #E0F0FF; line-height: 1.3; }
.nx-page-sub   { font-size: 13px; color: #2A4A6A; margin-top: 5px; line-height: 1.55; }
.nx-page-sub strong { color: #4A7A9C !important; }

/* Barra de pasos */
.nx-steps { display: flex; align-items: center; justify-content: center; padding: 8px 0 24px 0; }
.nx-step  { display: flex; flex-direction: column; align-items: center; gap: 5px; min-width: 90px; }
.nx-circle {
    width: 40px; height: 40px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 15px; font-weight: 700;
}
.nx-circle.done   { background: #1B9FD8; color: #fff; box-shadow: 0 0 14px rgba(27,159,216,.45); }
.nx-circle.active { background: #1B9FD8; color: #fff; box-shadow: 0 0 22px rgba(27,159,216,.7); }
.nx-circle.idle   { background: #07111C; color: #0F2A42; border: 2px solid #0D2035; }
.nx-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .7px; color: #0F2035; }
.nx-label.active, .nx-label.done { color: #1B9FD8; }
.nx-line { flex:1; height:2px; max-width:68px; border-radius:2px; margin-bottom:18px; }
.nx-line.done { background: #1B9FD8; }
.nx-line.idle { background: #0D2035; }

/* Encabezados de sección */
.nx-section {
    font-size: 11px; font-weight: 700; color: #1B9FD8;
    text-transform: uppercase; letter-spacing: 1.3px;
    margin: 22px 0 10px 0;
    display: flex; align-items: center; gap: 10px;
}
.nx-section::after {
    content: ''; flex:1; height:1px;
    background: linear-gradient(90deg, rgba(27,159,216,.35) 0%, transparent 100%);
}

/* Estado vacío */
.nx-empty {
    text-align: center; padding: 44px 24px;
    background: #050C18; border-radius: 14px;
    border: 2px dashed rgba(27,159,216,0.16); margin-top: 12px;
}
.nx-empty-icon { font-size: 48px; margin-bottom: 12px; }
.nx-empty-text { font-size: 15px; color: #1E3A55; }
.nx-empty-sub  { font-size: 12px; color: #0F2035; margin-top: 8px; }

/* Tarjeta de éxito */
.nx-success-card {
    background: linear-gradient(135deg, #04101C 0%, #061422 100%);
    border: 1px solid rgba(27,159,216,0.28);
    border-radius: 14px; padding: 28px; text-align: center; margin: 14px 0;
}
.nx-success-icon  { font-size: 48px; margin-bottom: 10px; }
.nx-success-title { font-size: 20px; font-weight: 700; color: #1B9FD8; margin-bottom: 6px; }
.nx-success-sub   { font-size: 13px; color: #1E3A58; }
.nx-success-sub strong { color: #4A8EB0 !important; }

/* Sección inferior (nombre + botón) */
.nx-export-bar {
    border-top: 1px solid rgba(27,159,216,0.12);
    background: linear-gradient(0deg, rgba(4,13,26,0.9) 0%, transparent 100%);
    padding: 18px 0 4px 0;
    margin-top: 18px;
}
.nx-export-label {
    font-size: 11px; font-weight: 700; color: #0F2A42;
    text-transform: uppercase; letter-spacing: 1.2px;
    margin-bottom: 10px;
}

/* Placeholder "Próximamente" */
.nx-coming-soon {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; padding: 80px 24px; text-align: center;
}
.nx-cs-icon  { font-size: 54px; margin-bottom: 18px; }
.nx-cs-title { font-size: 20px; font-weight: 700; color: #1B3A5A; margin-bottom: 8px; }
.nx-cs-sub   { font-size: 13px; color: #0F2035; }
.nx-cs-badge {
    display: inline-block; margin-top: 16px;
    font-size: 11px; font-weight: 700; color: #1B9FD8;
    background: rgba(27,159,216,0.08); border: 1px solid rgba(27,159,216,0.2);
    padding: 4px 14px; border-radius: 20px; letter-spacing: 0.8px;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# FITZ (PyMuPDF) — Import global
# ==========================================
try:
    import fitz
    FITZ_OK = True
except ImportError:
    FITZ_OK = False


# ==========================================
# HELPERS — Miniaturas y tarjetas
# ==========================================

def _gen_thumbs(pdf_bytes: bytes, scale: float = 1.2):
    """Genera lista de imágenes base64 (una por página) usando fitz."""
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
    """
    Devuelve (thumbs_b64_list, n_pages) cacheados en session_state.
    Regenera sólo si el PDF cambió. Resetea sel_key si se indica.
    """
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
    """
    Retorna HTML de una tarjeta miniatura con borde condicional.
    mode='delete' → borde rojo al seleccionar
    mode='split'  → borde amarillo al seleccionar
    """
    if selected and mode == "delete":
        border  = "2px solid rgba(220,50,50,0.85)"
        overlay = ('<div style="position:absolute;inset:0;background:rgba(220,50,50,0.22);'
                   'border-radius:7px;pointer-events:none;"></div>')
        pin = ('<div style="position:absolute;top:5px;right:5px;background:rgba(220,50,50,0.9);'
               'color:#fff;width:20px;height:20px;border-radius:50%;display:flex;'
               'align-items:center;justify-content:center;font-size:12px;'
               'font-weight:700;pointer-events:none;">✕</div>')
    elif selected and mode == "split":
        border  = "2px solid rgba(255,210,0,0.9)"
        overlay = ('<div style="position:absolute;inset:0;background:rgba(255,210,0,0.1);'
                   'border-radius:7px;pointer-events:none;"></div>')
        pin = ('<div style="position:absolute;top:5px;right:5px;background:rgba(255,210,0,0.9);'
               'color:#000;width:20px;height:20px;border-radius:50%;display:flex;'
               'align-items:center;justify-content:center;font-size:11px;'
               'font-weight:700;pointer-events:none;">✂</div>')
    else:
        border, overlay, pin = "2px solid rgba(27,159,216,0.25)", "", ""

    badge = ""
    if badge_label:
        badge = (f'<div style="position:absolute;top:5px;left:5px;background:{badge_bg};'
                 f'color:{badge_fg};padding:1px 6px;border-radius:7px;font-size:9px;'
                 f'font-weight:700;pointer-events:none;">{badge_label}</div>')

    return (f'<div style="border:{border};border-radius:10px;padding:6px;background:#0A1626;'
            f'position:relative;margin-bottom:2px;">'
            f'{overlay}{pin}{badge}'
            f'<img src="data:image/png;base64,{b64}" '
            f'style="width:100%;border-radius:6px;display:block;" draggable="false"/>'
            f'<div style="text-align:center;font-size:11px;color:#2E5878;'
            f'margin-top:5px;font-weight:600;">Pág. {page_num}</div>'
            f'</div>')


def _parse_fabric_color(color_str):
    """Parse CSS/Fabric.js color string to fitz (r, g, b) tuple, or None for transparent."""
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
    """Extract list of fitz.Point from a Fabric.js freedraw path object."""
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


# ==========================================
# LOGO Y DETECCIÓN DE RUTA
# ==========================================
ruta_logo = None
if os.path.exists("logo.png"):    ruta_logo = "logo.png"
elif os.path.exists("logo.jpg"):  ruta_logo = "logo.jpg"
elif os.path.exists("logo.jpeg"): ruta_logo = "logo.jpeg"

if ruta_logo:
    try:
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([2.5, 1, 2.5])
        with col2:
            st.image(ruta_logo, use_container_width=True)
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    except Exception:
        pass


# ==========================================
# SISTEMA DE SEGURIDAD (EL CADENERO)
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div style="background:#07111C;border:1px solid rgba(27,159,216,0.2);
                    border-radius:14px;padding:32px 28px;text-align:center;">
            <div style="font-size:36px;margin-bottom:12px;">🔒</div>
            <div style="font-size:18px;font-weight:700;color:#C8E4F0;margin-bottom:6px;">
                Acceso Restringido
            </div>
            <div style="font-size:13px;color:#2A4A6A;margin-bottom:20px;">
                Ingresa tu código de acceso para continuar.
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        password = st.text_input("Código de acceso:", type="password",
                                  placeholder="••••••••••••",
                                  label_visibility="collapsed")
        if st.button("Entrar", type="primary", use_container_width=True):
            try:
                claves_validas = list(st.secrets["accesos"].values())
                if password in claves_validas:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Código incorrecto o inactivo.")
            except Exception:
                st.warning("⚠️ La bóveda de contraseñas no está configurada en Streamlit.")
    st.stop()


# ==========================================
# BARRA DE NAVEGACIÓN HORIZONTAL
# ==========================================
tabs = st.tabs([
    "🗂️ Nexíficar Masivo",
    "📄 Nexíficar PDFs",
    "✂️ Dividir PDF",
    "🗜️ Comprimir PDF",
    "🔗 Merge PDF",
    "✏️ Editar PDF",
    "🗑️ Eliminar Páginas",
])


# ==========================================
# TAB 0 — NEXÍFICAR MASIVAMENTE
# ==========================================
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
    archivos_zip  = st.file_uploader("Archivos ZIP con los PDFs (puedes seleccionar varios)",
                                      type=["zip"], accept_multiple_files=True)
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


# ==========================================
# TAB 1 — NEXÍFICAR PDFs
# ==========================================
with tabs[1]:

    try:
        import fitz
        FITZ_OK = True
    except ImportError:
        FITZ_OK = False

    # ── Session state ──────────────────────────────────────────────────────
    for _k, _v in [("nx_done", False), ("nx_buffer", None),
                   ("nx_nombre", "Documento_Unificado.pdf"),
                   ("nx_order", []), ("nx_files_sig", ""),
                   ("nx_editor_ver", 0)]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    # ── Helpers ────────────────────────────────────────────────────────────
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

    # ── Cabecera ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">📄🔗📄 Nexíficar PDFs</div>
        <div class="nx-page-sub">Sube varios PDFs sueltos y únelos en <strong>un solo archivo</strong>,
        con previsualización de miniatura y reordenamiento por número de posición.</div>
    </div>
    """, unsafe_allow_html=True)

    # ── File uploader ──────────────────────────────────────────────────────
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
        # Build metadata map
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

        # Sync order
        files_sig = hashlib.md5(
            "".join(sorted(file_info_map.keys())).encode()
        ).hexdigest()[:8]

        if st.session_state.nx_files_sig != files_sig:
            st.session_state.nx_files_sig = files_sig
            st.session_state.nx_order     = list(file_info_map.keys())
            st.session_state.nx_done      = False
            st.session_state.nx_buffer    = None

        st.session_state.nx_order = [n for n in st.session_state.nx_order if n in file_info_map]
        orden = st.session_state.nx_order

        # ── PASO 2 ────────────────────────────────────────────────────────
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
                                   "Orden", min_value=1, max_value=len(orden),
                                   step=1, width="small"),
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

            # Vista previa de miniaturas
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
                                    st.image(base64.b64decode(fi["thumb"]),
                                             use_container_width=True)
                                else:
                                    st.markdown(
                                        '<div style="background:#07111C;border-radius:8px;'
                                        'height:100px;display:flex;align-items:center;'
                                        'justify-content:center;font-size:32px;">📄</div>',
                                        unsafe_allow_html=True
                                    )
                                short = (fi["name"][:22] + "…") if len(fi["name"]) > 22 else fi["name"]
                                st.markdown(
                                    f'<div style="font-size:12px;font-weight:600;color:#C8E4F0;'
                                    f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'
                                    f'margin:5px 0 2px 0;">'
                                    f'<span style="background:#1B9FD8;color:#fff;'
                                    f'border-radius:50%;padding:1px 7px;margin-right:5px;'
                                    f'font-size:11px;font-weight:700;">{idx + 1}</span>'
                                    f'{short}</div>',
                                    unsafe_allow_html=True
                                )
                                st.caption(f"📄 {fi['pages']} pág. · 💾 {fi['size']}")

            # Sección inferior: nombre + Nexíficar
            st.markdown('<div class="nx-export-bar">', unsafe_allow_html=True)
            st.markdown('<div class="nx-export-label">💾 Nombre del PDF final y exportación</div>',
                        unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            col_name, col_btn = st.columns([3, 1])
            with col_name:
                nombre_final = st.text_input(
                    "Nombre del archivo unificado:",
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
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error al unir los archivos: {e}")

        # ── PASO 3: éxito + descarga ──────────────────────────────────────
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


# ==========================================
# TAB 2 — DIVIDIR PDF
# ==========================================
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

            # ── Selector de modo (3 opciones estilo iLovePDF) ─────────────
            st.markdown('<div class="nx-section">⚙️ Elige el modo de división</div>',
                        unsafe_allow_html=True)

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
                bdr = "2px solid #1B9FD8" if active else "1px solid rgba(27,159,216,0.15)"
                bg  = "rgba(27,159,216,0.1)" if active else "#07111C"
                col.markdown(
                    f'<div style="background:{bg};border:{bdr};border-radius:12px;'
                    f'padding:18px 14px;text-align:center;min-height:120px;">'
                    f'<div style="font-size:30px;margin-bottom:8px;">{icon}</div>'
                    f'<div style="font-size:13px;font-weight:700;color:#C8E4F0;margin-bottom:6px;">{title}</div>'
                    f'<div style="font-size:11px;color:#2A4A6A;line-height:1.4;">{desc}</div>'
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

            # ── Layout principal: grid + panel lateral ─────────────────────
            grid_col, panel_col = st.columns([4, 1])
            mode = st.session_state.sp_mode
            _SP_SEC = ["#1B9FD8","#27AE60","#E74C3C","#F39C12","#9B59B6","#1ABC9C","#E91E63","#FF9800"]

            # ── Panel lateral ──────────────────────────────────────────────
            with panel_col:
                st.markdown(
                    '<div style="background:#07111C;border:1px solid rgba(27,159,216,0.18);'
                    'border-radius:12px;padding:18px 14px;">',
                    unsafe_allow_html=True
                )
                _mode_names = {"A": ("📄","Extraer todo"), "B": ("☑️","Por páginas"), "C": ("🔢","Partes iguales")}
                _ic, _lb = _mode_names[mode]
                st.markdown(
                    f'<div style="font-size:11px;font-weight:700;color:#1B9FD8;text-transform:uppercase;'
                    f'letter-spacing:1px;margin-bottom:10px;">Modo activo</div>'
                    f'<div style="font-size:22px;margin-bottom:4px;">{_ic}</div>'
                    f'<div style="font-size:13px;font-weight:700;color:#C8E4F0;margin-bottom:14px;">{_lb}</div>',
                    unsafe_allow_html=True
                )

                if mode == "B":
                    n_sel_b = len(st.session_state.sp_sel_b)
                    st.markdown(
                        f'<div style="font-size:28px;font-weight:700;color:#1B9FD8;text-align:center;'
                        f'margin-bottom:4px;">{n_sel_b}</div>'
                        f'<div style="font-size:11px;color:#2A4A6A;text-align:center;'
                        f'margin-bottom:14px;">páginas seleccionadas</div>',
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
                                if not part:
                                    continue
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
                        f'<div style="font-size:12px;color:#2A4A6A;margin-top:6px;">'
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
                        st.session_state.sp_result        = None
                        st.session_state.sp_sel_b         = set()
                        st.session_state.sp_mode          = "A"
                        st.session_state.sp_pages_per_part = 2
                        st.session_state.sp_sig           = ""
                        st.session_state.sp_thumbs        = []
                        st.session_state.sp_n             = 0
                        st.rerun()

            # ── Grid de miniaturas (6 columnas) ───────────────────────────
            with grid_col:
                st.markdown(
                    '<div class="nx-section">🖼️ Páginas del PDF</div>',
                    unsafe_allow_html=True
                )
                COLS = 6
                for row_start in range(0, n_pages, COLS):
                    row_cols = st.columns(COLS)
                    for ci, pi in enumerate(range(row_start, min(row_start + COLS, n_pages))):
                        with row_cols[ci]:
                            if mode == "B":
                                is_sel = pi in st.session_state.sp_sel_b
                                bdr_c  = "2px solid #27AE60" if is_sel else "2px solid rgba(27,159,216,0.25)"
                                ov     = ('<div style="position:absolute;inset:0;'
                                          'background:rgba(39,174,96,0.18);border-radius:7px;'
                                          'pointer-events:none;"></div>') if is_sel else ""
                                chk    = ('<div style="position:absolute;top:5px;left:5px;'
                                          'background:#27AE60;color:#fff;width:18px;height:18px;'
                                          'border-radius:4px;display:flex;align-items:center;'
                                          'justify-content:center;font-size:11px;font-weight:700;'
                                          'pointer-events:none;">✓</div>') if is_sel else (
                                          '<div style="position:absolute;top:5px;left:5px;'
                                          'background:rgba(255,255,255,0.08);'
                                          'border:2px solid rgba(27,159,216,0.3);'
                                          'width:18px;height:18px;border-radius:4px;'
                                          'pointer-events:none;"></div>'
                                )
                                st.markdown(
                                    f'<div style="border:{bdr_c};border-radius:10px;padding:5px;'
                                    f'background:#0A1626;position:relative;margin-bottom:2px;">'
                                    f'{ov}{chk}'
                                    f'<img src="data:image/png;base64,{thumbs[pi]}" '
                                    f'style="width:100%;border-radius:5px;display:block;" draggable="false"/>'
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
                                st.markdown(
                                    _thumb_card(thumbs[pi], pi + 1),
                                    unsafe_allow_html=True
                                )


# ==========================================
# TAB 3 — COMPRIMIR PDF (próximamente)
# ==========================================
with tabs[3]:
    st.markdown("""
    <div class="nx-coming-soon">
        <div class="nx-cs-icon">🗜️</div>
        <div class="nx-cs-title">Comprimir PDF</div>
        <div class="nx-cs-sub">Reduce el tamaño de tus PDFs sin perder calidad visible.</div>
        <div class="nx-cs-badge">Próximamente</div>
    </div>""", unsafe_allow_html=True)


# ==========================================
# TAB 4 — MERGE PDF (próximamente)
# ==========================================
with tabs[4]:
    st.markdown("""
    <div class="nx-coming-soon">
        <div class="nx-cs-icon">🔗</div>
        <div class="nx-cs-title">Merge PDF</div>
        <div class="nx-cs-sub">Combina PDFs con opciones avanzadas de intercalado y portada.</div>
        <div class="nx-cs-badge">Próximamente</div>
    </div>""", unsafe_allow_html=True)


# ==========================================
# TAB 5 — EDITAR PDF (Canvas Visual)
# ==========================================
with tabs[5]:

    # ── Importaciones opcionales ───────────────────────────────────────────
    try:
        from streamlit_drawable_canvas import st_canvas as _st_canvas
        _CANVAS_PKG = True
    except ImportError:
        _CANVAS_PKG = False

    try:
        from PIL import Image as _PILImg
        _PIL_PKG = True
    except ImportError:
        _PIL_PKG = False

    for _k, _v in [
        ("ed_sig", ""), ("ed_thumbs", []), ("ed_n", 0),
        ("ed_text_elems", []),
        ("ed_cur_page", 0),
        ("ed_draw_mode", "rect"),
        ("ed_stroke_c", "#E74C3C"),
        ("ed_fill_hex", "#FFCC00"),
        ("ed_has_fill", False),
        ("ed_stroke_w", 2),
        ("ed_font_name", "Helvetica"),
        ("ed_font_sz", 14),
        ("ed_txt_color", "#000000"),
        ("ed_result", None),
    ]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">✏️ Editar PDF</div>
        <div class="nx-page-sub">Editor visual interactivo: dibuja <strong>rectángulos</strong>,
        <strong>elipses</strong>, <strong>trazos libres</strong> y <strong>líneas</strong>
        directamente sobre el PDF. Agrega texto en posición exacta y descarga el resultado.</div>
    </div>
    """, unsafe_allow_html=True)

    ed_file = st.file_uploader("Sube el PDF a editar", type=["pdf"], key="ed_uploader")

    if not ed_file:
        st.session_state.ed_result     = None
        st.session_state.ed_text_elems = []
        st.session_state.ed_cur_page   = 0
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">✏️</div>
            <div class="nx-empty-text">Sube un PDF para empezar a editarlo</div>
            <div class="nx-empty-sub">Dibuja sobre el PDF y agrega texto en posición exacta</div>
        </div>""", unsafe_allow_html=True)
    else:
        if not FITZ_OK:
            st.error("⚠️ PyMuPDF no está instalado.")
        elif not _CANVAS_PKG:
            st.error("⚠️ `streamlit-drawable-canvas` no está instalado. "
                     "Asegúrate de que está en requirements.txt y redespliega la app.")
        elif not _PIL_PKG:
            st.error("⚠️ `Pillow` no está instalado.")
        else:
            ed_bytes = ed_file.read()
            thumbs_ed, n_pages_ed = _ensure_thumbs("ed", ed_bytes)

            cur_page = max(0, min(st.session_state.ed_cur_page, n_pages_ed - 1))
            st.session_state.ed_cur_page = cur_page

            # ── Render página actual como imagen PIL para el canvas ─────────
            CANVAS_W = 680
            _doc_ed = fitz.open(stream=ed_bytes, filetype="pdf")
            _pg_ed  = _doc_ed[cur_page]
            pw_r    = _pg_ed.rect.width
            ph_r    = _pg_ed.rect.height
            canvas_h = int(ph_r * CANVAS_W / pw_r)
            _pix_bg  = _pg_ed.get_pixmap(
                matrix=fitz.Matrix(CANVAS_W / pw_r, canvas_h / ph_r), alpha=False
            )
            _doc_ed.close()
            _bg_pil = _PILImg.frombytes("RGB", [_pix_bg.width, _pix_bg.height], _pix_bg.samples)

            # ── Layout 3 columnas ──────────────────────────────────────────
            nav_col, canvas_col, props_col = st.columns([1, 3, 1.5])

            # ── Columna izquierda: miniaturas de páginas ───────────────────
            with nav_col:
                st.markdown(
                    '<div style="font-size:11px;font-weight:700;color:#1B9FD8;'
                    'text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Páginas</div>',
                    unsafe_allow_html=True
                )
                for pi in range(n_pages_ed):
                    _is_cur = pi == cur_page
                    _n_txt  = sum(1 for e in st.session_state.ed_text_elems if e["page"] == pi)
                    _cv_raw = st.session_state.get(f"ed_canvas_{pi}")
                    _n_shp  = len(_cv_raw.get("objects", [])) if isinstance(_cv_raw, dict) else 0
                    _n_el   = _n_txt + _n_shp
                    _bdr_pn = "2px solid #1B9FD8" if _is_cur else "1px solid rgba(27,159,216,0.15)"
                    _bg_pn  = "rgba(27,159,216,0.12)" if _is_cur else "#0A1626"
                    _bdg_pn = (f'<div style="position:absolute;bottom:4px;right:4px;background:#1B9FD8;'
                                f'color:#fff;font-size:9px;font-weight:700;padding:1px 5px;'
                                f'border-radius:5px;">{_n_el}</div>') if _n_el > 0 else ""
                    st.markdown(
                        f'<div style="border:{_bdr_pn};border-radius:8px;padding:4px;'
                        f'background:{_bg_pn};position:relative;margin-bottom:4px;">'
                        f'<img src="data:image/png;base64,{thumbs_ed[pi]}" '
                        f'style="width:100%;border-radius:4px;display:block;"/>'
                        f'{_bdg_pn}'
                        f'<div style="text-align:center;font-size:10px;color:#2E5878;'
                        f'margin-top:3px;font-weight:600;">{pi+1}</div></div>',
                        unsafe_allow_html=True
                    )
                    if not _is_cur:
                        if st.button("Ver", key=f"ed_nav_{pi}", use_container_width=True):
                            st.session_state.ed_cur_page = pi
                            st.rerun()

            # ── Columna derecha: herramientas + propiedades + guardar ──────
            with props_col:
                st.markdown(
                    '<div style="font-size:11px;font-weight:700;color:#1B9FD8;'
                    'text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Herramienta</div>',
                    unsafe_allow_html=True
                )
                _tools_ed = [
                    ("rect",      "□  Rectángulo"),
                    ("circle",    "○  Elipse"),
                    ("freedraw",  "✏️  Trazo libre"),
                    ("line",      "↗  Línea"),
                    ("transform", "✥  Mover / Redim"),
                    ("text",      "T   Texto"),
                ]
                for _tm, _tlbl in _tools_ed:
                    _is_act_t = st.session_state.ed_draw_mode == _tm
                    if st.button(
                        _tlbl, key=f"ed_tool_{_tm}",
                        use_container_width=True,
                        type="primary" if _is_act_t else "secondary"
                    ):
                        st.session_state.ed_draw_mode = _tm
                        st.rerun()

                st.markdown("<br>", unsafe_allow_html=True)
                _dm = st.session_state.ed_draw_mode

                # ── Propiedades según herramienta ────────────────────────
                if _dm == "text":
                    st.markdown(
                        '<div style="font-size:11px;font-weight:700;color:#4A7A9C;'
                        'margin-bottom:8px;">Configurar texto</div>', unsafe_allow_html=True
                    )
                    _fsel = st.selectbox(
                        "Fuente", ["Helvetica", "Times", "Courier"],
                        index=["Helvetica","Times","Courier"].index(st.session_state.ed_font_name),
                        key="ed_fsel_t"
                    )
                    _fsz  = st.number_input("Tamaño pt", 6, 120, st.session_state.ed_font_sz, key="ed_fsz_t")
                    _tcol = st.color_picker("Color texto", st.session_state.ed_txt_color, key="ed_tcol_t")
                    st.session_state.ed_font_name = _fsel
                    st.session_state.ed_font_sz   = _fsz
                    st.session_state.ed_txt_color = _tcol

                    st.markdown(
                        '<div style="font-size:11px;font-weight:700;color:#4A7A9C;'
                        'margin:10px 0 6px 0;">Posición en el PDF</div>', unsafe_allow_html=True
                    )
                    _tx = st.number_input("X →", 0, int(pw_r), int(pw_r)//6, key="ed_tx_t")
                    _ty = st.number_input("Y ↓", 0, int(ph_r), int(ph_r)//2, key="ed_ty_t")
                    _tv = st.text_area("Texto", height=70, key="ed_tv_t",
                                        placeholder="Escribe el texto…")
                    if st.button("➕ Agregar texto", type="primary",
                                  use_container_width=True, key="ed_add_txt",
                                  disabled=not _tv.strip()):
                        _fmap_ed = {"Helvetica": "helv", "Times": "tiro", "Courier": "cour"}
                        st.session_state.ed_text_elems.append({
                            "page": cur_page, "x": _tx, "y": _ty,
                            "text": _tv, "fontname": _fmap_ed[_fsel],
                            "fontsize": _fsz, "color": _tcol,
                        })
                        st.rerun()

                elif _dm in ("rect", "circle", "freedraw", "line"):
                    st.markdown(
                        '<div style="font-size:11px;font-weight:700;color:#4A7A9C;'
                        'margin-bottom:8px;">Apariencia</div>', unsafe_allow_html=True
                    )
                    st.session_state.ed_stroke_c = st.color_picker(
                        "Color borde", st.session_state.ed_stroke_c, key="ed_sc_t"
                    )
                    st.session_state.ed_stroke_w = st.number_input(
                        "Grosor", 1, 20, st.session_state.ed_stroke_w, key="ed_sw_t"
                    )
                    st.session_state.ed_has_fill = st.checkbox(
                        "Con relleno", value=st.session_state.ed_has_fill, key="ed_hf_t"
                    )
                    if st.session_state.ed_has_fill:
                        st.session_state.ed_fill_hex = st.color_picker(
                            "Color relleno", st.session_state.ed_fill_hex, key="ed_fc_t"
                        )
                else:
                    st.markdown(
                        '<div style="font-size:12px;color:#2A4A6A;padding:8px 0;">'
                        'Selecciona elementos en el canvas para moverlos y redimensionarlos.</div>',
                        unsafe_allow_html=True
                    )

                st.markdown("---")

                # ── Lista de textos ──────────────────────────────────────
                _n_texts = len(st.session_state.ed_text_elems)
                if _n_texts > 0:
                    st.markdown(
                        f'<div style="font-size:11px;font-weight:700;color:#1B9FD8;'
                        f'margin-bottom:8px;">Textos ({_n_texts})</div>', unsafe_allow_html=True
                    )
                    for _ti, _te in enumerate(st.session_state.ed_text_elems):
                        _te_lbl = f'{_te["text"][:10]}{"…" if len(_te["text"])>10 else ""}'
                        _te_pg  = _te["page"] + 1
                        _te_bg  = "rgba(27,159,216,0.08)" if _te["page"] == cur_page else "rgba(27,159,216,0.02)"
                        st.markdown(
                            f'<div style="background:{_te_bg};border:1px solid rgba(27,159,216,0.12);'
                            f'border-radius:6px;padding:7px 10px;margin-bottom:4px;">'
                            f'<span style="background:rgba(27,159,216,0.2);color:#1B9FD8;'
                            f'padding:1px 5px;border-radius:3px;font-size:10px;font-weight:700;">T</span> '
                            f'<span style="font-size:12px;color:#C8E4F0;font-weight:600;">{_te_lbl}</span>'
                            f'<div style="font-size:10px;color:#2A4A6A;">'
                            f'Pág.{_te_pg} · {_te["fontsize"]}pt</div></div>',
                            unsafe_allow_html=True
                        )
                        if st.button("✕", key=f"ed_del_t_{_ti}", use_container_width=True):
                            st.session_state.ed_text_elems.pop(_ti)
                            st.rerun()
                    st.markdown("---")

                # ── Guardar cambios ──────────────────────────────────────
                _has_shapes = any(
                    len((st.session_state.get(f"ed_canvas_{_pi}") or {}).get("objects", [])) > 0
                    for _pi in range(n_pages_ed)
                )
                _has_content = _has_shapes or _n_texts > 0

                if st.button("💾 Guardar cambios", type="primary",
                              use_container_width=True, key="ed_save3",
                              disabled=not _has_content):
                    with st.spinner("Generando PDF…"):
                        try:
                            _doc_sv = fitz.open(stream=ed_bytes, filetype="pdf")

                            # Aplicar formas del canvas por página
                            for _pi in range(n_pages_ed):
                                _cv_r = st.session_state.get(f"ed_canvas_{_pi}")
                                if not isinstance(_cv_r, dict):
                                    continue
                                _objs = _cv_r.get("objects", [])
                                if not _objs:
                                    continue
                                _pg_sv   = _doc_sv[_pi]
                                _pw_sv   = _pg_sv.rect.width
                                _ph_sv   = _pg_sv.rect.height
                                _ch_sv   = int(_ph_sv * CANVAS_W / _pw_sv)
                                _sxp     = _pw_sv / CANVAS_W
                                _syp     = _ph_sv / _ch_sv

                                for _obj in _objs:
                                    _otype = _obj.get("type", "")
                                    _left  = _obj.get("left", 0) * _sxp
                                    _top   = _obj.get("top",  0) * _syp
                                    _sc_x  = _obj.get("scaleX", 1.0)
                                    _sc_y  = _obj.get("scaleY", 1.0)
                                    _s_clr = _parse_fabric_color(_obj.get("stroke", "#000000"))
                                    _f_clr = _parse_fabric_color(_obj.get("fill", ""))
                                    _lw    = max(0.5, _obj.get("strokeWidth", 2) * _sxp)

                                    if _otype == "rect":
                                        _w2 = _obj.get("width", 0) * _sc_x * _sxp
                                        _h2 = _obj.get("height", 0) * _sc_y * _syp
                                        _pg_sv.draw_rect(
                                            fitz.Rect(_left, _top, _left+_w2, _top+_h2),
                                            color=_s_clr, fill=_f_clr, width=_lw
                                        )
                                    elif _otype == "circle":
                                        _r  = _obj.get("radius", 0)
                                        _w2 = _r * 2 * _sc_x * _sxp
                                        _h2 = _r * 2 * _sc_y * _syp
                                        _pg_sv.draw_oval(
                                            fitz.Rect(_left, _top, _left+_w2, _top+_h2),
                                            color=_s_clr, fill=_f_clr, width=_lw
                                        )
                                    elif _otype == "line":
                                        _bw = _obj.get("width", 0)
                                        _bh = _obj.get("height", 0)
                                        _cx = _left + _bw * _sxp / 2
                                        _cy = _top  + _bh * _syp / 2
                                        _p1 = fitz.Point(
                                            _cx + _obj.get("x1", 0) * _sxp,
                                            _cy + _obj.get("y1", 0) * _syp
                                        )
                                        _p2 = fitz.Point(
                                            _cx + _obj.get("x2", 0) * _sxp,
                                            _cy + _obj.get("y2", 0) * _syp
                                        )
                                        _pg_sv.draw_line(_p1, _p2, color=_s_clr, width=_lw)
                                    elif _otype == "path":
                                        _pts = _extract_path_pts(_obj, _sxp, _syp)
                                        if len(_pts) >= 2:
                                            _pg_sv.draw_polyline(_pts, color=_s_clr, width=_lw)

                            # Aplicar textos
                            for _te in st.session_state.ed_text_elems:
                                _pg_sv2 = _doc_sv[_te["page"]]
                                _tc = _te["color"]
                                _pg_sv2.insert_text(
                                    fitz.Point(_te["x"], _te["y"]),
                                    _te["text"], fontname=_te["fontname"],
                                    fontsize=_te["fontsize"],
                                    color=(int(_tc[1:3],16)/255,
                                           int(_tc[3:5],16)/255,
                                           int(_tc[5:7],16)/255)
                                )

                            _buf_sv = BytesIO()
                            _doc_sv.save(_buf_sv)
                            _doc_sv.close()
                            _buf_sv.seek(0)
                            st.session_state.ed_result = _buf_sv.getvalue()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

                if st.session_state.ed_result:
                    st.download_button(
                        label="⬇️ Descargar PDF editado",
                        data=st.session_state.ed_result,
                        file_name=ed_file.name.replace(".pdf", "_editado.pdf"),
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🔄 Nueva edición", use_container_width=True, key="ed_rst3"):
                        st.session_state.ed_text_elems = []
                        st.session_state.ed_result     = None
                        st.session_state.ed_cur_page   = 0
                        for _pi in range(n_pages_ed):
                            _ck = f"ed_canvas_{_pi}"
                            if _ck in st.session_state:
                                del st.session_state[_ck]
                        st.rerun()

            # ── Canvas principal (columna central) ─────────────────────────
            with canvas_col:
                _dm2 = st.session_state.ed_draw_mode
                # En modo texto, el canvas está en "transform" (sólo mover existentes)
                _canvas_mode = "transform" if _dm2 == "text" else _dm2

                # Calcular fill RGBA para el canvas
                if _dm2 in ("rect", "circle") and st.session_state.ed_has_fill:
                    _fh = st.session_state.ed_fill_hex
                    _fr_v, _fg_v, _fb_v = int(_fh[1:3],16), int(_fh[3:5],16), int(_fh[5:7],16)
                    _fill_str = f"rgba({_fr_v},{_fg_v},{_fb_v},0.3)"
                else:
                    _fill_str = "rgba(0,0,0,0)"

                _stroke_str = (st.session_state.ed_stroke_c
                               if _dm2 not in ("text", "transform") else "#1B9FD8")
                _sw_val     = (st.session_state.ed_stroke_w
                               if _dm2 not in ("text", "transform") else 1)

                _canvas_res = _st_canvas(
                    fill_color=_fill_str,
                    stroke_width=_sw_val,
                    stroke_color=_stroke_str,
                    background_image=_bg_pil,
                    update_streamlit=True,
                    height=canvas_h,
                    width=CANVAS_W,
                    drawing_mode=_canvas_mode,
                    point_display_radius=0,
                    display_toolbar=False,
                    key=f"ed_canvas_{cur_page}",
                )

                # Guardar estado del canvas en session_state (para acceso entre páginas)
                if _canvas_res is not None and _canvas_res.json_data is not None:
                    st.session_state[f"ed_canvas_{cur_page}"] = _canvas_res.json_data

                # Tip contextual
                _ed_tips = {
                    "rect":      "💡 Clic y arrastra para dibujar un rectángulo",
                    "circle":    "💡 Clic y arrastra para dibujar una elipse",
                    "freedraw":  "💡 Dibuja trazos libres directamente sobre el PDF",
                    "line":      "💡 Clic y arrastra para trazar una línea",
                    "transform": "💡 Clic para seleccionar · arrastra para mover · handles para redimensionar",
                    "text":      "💡 Configura el texto en el panel derecho y pulsa ➕ Agregar texto",
                }
                st.caption(_ed_tips.get(_dm2, ""))


# ==========================================
# TAB 6 — ELIMINAR PÁGINAS
# ==========================================
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
            ep_bytes = ep_file.read()
            thumbs_ep, n_pages_ep = _ensure_thumbs("ep", ep_bytes, "ep_sel")
            n_sel = len(st.session_state.ep_sel)

            # ── Resultado listo: mostrar descarga + reinicio ───────────────
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
                if st.button("🔄 Eliminar páginas de otro PDF", use_container_width=True,
                             key="ep_reset"):
                    st.session_state.ep_result = None
                    st.session_state.ep_sel    = set()
                    st.session_state.ep_sig    = ""
                    st.session_state.ep_thumbs = []
                    st.session_state.ep_n      = 0
                    st.rerun()

            else:
                # ── Selector de rango ──────────────────────────────────────
                st.markdown(
                    '<div class="nx-section">🎯 Selección rápida por rango</div>',
                    unsafe_allow_html=True
                )
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

                # ── Barra de estado + acciones globales ────────────────────
                st.markdown(
                    '<div class="nx-section">🗑️ Selecciona páginas para eliminar</div>',
                    unsafe_allow_html=True
                )
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

                # ── Grid de miniaturas (6 columnas) ────────────────────────
                COLS = 6
                for row_start in range(0, n_pages_ep, COLS):
                    cols_ep = st.columns(COLS)
                    for ci, pi in enumerate(range(row_start, min(row_start + COLS, n_pages_ep))):
                        is_sel = pi in st.session_state.ep_sel
                        with cols_ep[ci]:
                            st.markdown(
                                _thumb_card(thumbs_ep[pi], pi + 1,
                                            selected=is_sel, mode="delete"),
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

                # ── Acción ─────────────────────────────────────────────────
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
