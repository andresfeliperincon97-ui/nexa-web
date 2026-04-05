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
[data-testid="stDecoration"]    { display: none !important; }
#MainMenu                        { visibility: hidden !important; }
footer                           { visibility: hidden !important; }
.stDeployButton                  { display: none !important; }

/* Área principal */
[data-testid="stAppViewContainer"] > [data-testid="stMain"] {
    background: #060F1D !important;
}
.block-container {
    padding-top: 1.8rem !important;
    padding-bottom: 3rem !important;
    max-width: 1200px !important;
}

/* ══════════════════════════════════════════════════════
   SIDEBAR — FIX TOGGLE COLAPSAR/EXPANDIR
══════════════════════════════════════════════════════ */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div {
    background: #040D1A !important;
    border-right: 1px solid rgba(27,159,216,0.12) !important;
    min-width: 252px !important;
    max-width: 252px !important;
}
[data-testid="stSidebarContent"] {
    background: #040D1A !important;
    padding: 0 !important;
}

/* Botón nativo de colapsar — visible y funcional */

/* Scrollbar del sidebar */
[data-testid="stSidebar"]::-webkit-scrollbar       { width: 4px; }
[data-testid="stSidebar"]::-webkit-scrollbar-track { background: #040D1A; }
[data-testid="stSidebar"]::-webkit-scrollbar-thumb { background: #0F2035; border-radius: 2px; }

/* Imagen del logo en sidebar */
[data-testid="stSidebar"] [data-testid="stImage"] {
    padding: 18px 28px 4px 28px !important;
}
[data-testid="stSidebar"] [data-testid="stImage"] img {
    border-radius: 6px !important;
}

/* ── Radio nav items ── */
[data-testid="stSidebar"] [data-testid="stRadio"] > label  { display: none !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] > div    { gap: 1px !important; padding: 0 !important; }

[data-testid="stSidebar"] [data-testid="stRadio"] > div > label {
    display: flex !important;
    align-items: center !important;
    padding: 10px 16px 10px 20px !important;
    border-left: 3px solid transparent !important;
    color: #2E4D6A !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    letter-spacing: 0.15px !important;
    transition: background .15s, color .15s, border-left-color .15s !important;
    cursor: pointer !important;
    border-radius: 0 6px 6px 0 !important;
    margin: 1px 8px 1px 0 !important;
    user-select: none !important;
    background: transparent !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] > div > label > div:first-child {
    display: none !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] > div > label:hover {
    background: rgba(27,159,216,0.07) !important;
    color: #5A9FC4 !important;
    border-left-color: rgba(27,159,216,0.3) !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] > div > label:has(input:checked) {
    border-left: 3px solid #1B9FD8 !important;
    background: rgba(27,159,216,0.13) !important;
    color: #1B9FD8 !important;
    font-weight: 600 !important;
}

/* Divisor en sidebar */
[data-testid="stSidebar"] hr { border-color: rgba(27,159,216,0.1) !important; margin: 6px 0 !important; }

/* Botón nativo de colapsar — permitir que Streamlit lo maneje */
[data-testid="stSidebarCollapseButton"] button,
[data-testid="collapsedControl"] button {
    background: #040D1A !important;
    border: 1px solid rgba(27,159,216,0.3) !important;
    color: #1B9FD8 !important;
}
[data-testid="stSidebarCollapseButton"] svg,
[data-testid="collapsedControl"] svg {
    color: #1B9FD8 !important;
    fill: #1B9FD8 !important;
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
[data-testid="stTextInput"] label { color: #2E5070 !important; font-size: 12px !important; font-weight: 600 !important; letter-spacing: 0.3px !important; }
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
   SELECT BOX
══════════════════════════════════════════════════════ */
[data-testid="stSelectbox"] label { color: #2E5070 !important; font-size: 12px !important; font-weight: 600 !important; }
[data-testid="stSelectbox"] > div > div {
    background: #07111C !important;
    border: 1px solid rgba(27,159,216,0.18) !important;
    color: #C8E4F0 !important;
    border-radius: 8px !important;
}

/* ══════════════════════════════════════════════════════
   NUMBER INPUT
══════════════════════════════════════════════════════ */
[data-testid="stNumberInput"] label { color: #2E5070 !important; font-size: 12px !important; font-weight: 600 !important; }
[data-testid="stNumberInput"] input {
    background: #07111C !important;
    border: 1px solid rgba(27,159,216,0.18) !important;
    color: #C8E4F0 !important;
    border-radius: 8px !important;
}

/* ══════════════════════════════════════════════════════
   SLIDER
══════════════════════════════════════════════════════ */
[data-testid="stSlider"] label { color: #2E5070 !important; font-size: 12px !important; font-weight: 600 !important; }
[data-testid="stSlider"] [data-testid="stTickBar"] { background: #1B9FD8 !important; }

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
[data-testid="stAlert"] p { color: #3A6888 !important; }
[data-testid="stAlert"] strong { color: #1B9FD8 !important; }

/* ══════════════════════════════════════════════════════
   PROGRESS / SPINNER
══════════════════════════════════════════════════════ */
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #1B9FD8 0%, #1DE0C0 100%) !important;
}
[data-testid="stSpinner"] > div > div { border-top-color: #1B9FD8 !important; }

/* ══════════════════════════════════════════════════════
   EXPANDER
══════════════════════════════════════════════════════ */
[data-testid="stExpander"] {
    background: #0A1626 !important;
    border: 1px solid rgba(27,159,216,0.12) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}
[data-testid="stExpanderDetails"] { background: #0A1626 !important; }

/* ══════════════════════════════════════════════════════
   CAPTIONS / HR / MISC
══════════════════════════════════════════════════════ */
.stCaption, [data-testid="stCaptionContainer"] { color: #1E3A58 !important; font-size: 11px !important; }
hr { border-color: rgba(27,159,216,0.1) !important; }
[data-testid="stMarkdownContainer"] p { color: #4A7A9C !important; }

/* ══════════════════════════════════════════════════════
   DOWNLOAD BUTTON
══════════════════════════════════════════════════════ */
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
   TARJETAS DE HERRAMIENTAS (estilo iLovePDF)
══════════════════════════════════════════════════════ */
.nx-tool-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
    gap: 16px;
    margin-top: 12px;
}
.nx-tool-card {
    background: #0A1626;
    border: 1px solid rgba(27,159,216,0.12);
    border-radius: 14px;
    padding: 24px 18px 20px 18px;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s ease;
    text-decoration: none;
    display: block;
}
.nx-tool-card:hover {
    border-color: rgba(27,159,216,0.45);
    background: #0D1E35;
    transform: translateY(-3px);
    box-shadow: 0 8px 28px rgba(27,159,216,0.15);
}
.nx-tool-card.active {
    border-color: #1B9FD8;
    background: rgba(27,159,216,0.08);
    box-shadow: 0 0 0 2px rgba(27,159,216,0.2);
}
.nx-tool-icon { font-size: 36px; margin-bottom: 10px; display: block; }
.nx-tool-name {
    font-size: 13px; font-weight: 700;
    color: #C8E4F0; margin-bottom: 5px;
}
.nx-tool-desc { font-size: 11px; color: #2A4A6A; line-height: 1.4; }
.nx-tool-badge {
    display: inline-block; margin-top: 8px;
    font-size: 10px; font-weight: 700;
    background: rgba(27,159,216,0.1);
    color: #1B9FD8;
    padding: 2px 8px; border-radius: 10px;
    border: 1px solid rgba(27,159,216,0.2);
}

/* ══════════════════════════════════════════════════════
   COMPONENTES REUTILIZABLES NEXA
══════════════════════════════════════════════════════ */
.nx-nav-section {
    padding: 14px 20px 4px 20px;
    font-size: 10px;
    font-weight: 700;
    color: #0F2035;
    letter-spacing: 2px;
    text-transform: uppercase;
}
.nx-page-header {
    padding: 4px 0 20px 0;
    border-bottom: 1px solid rgba(27,159,216,0.1);
    margin-bottom: 22px;
}
.nx-page-title { font-size: 22px; font-weight: 700; color: #E0F0FF; line-height: 1.3; }
.nx-page-sub { font-size: 13px; color: #2A4A6A; margin-top: 5px; line-height: 1.55; }
.nx-page-sub strong { color: #4A7A9C !important; }

.nx-steps { display: flex; align-items: center; justify-content: center; padding: 8px 0 24px 0; }
.nx-step { display: flex; flex-direction: column; align-items: center; gap: 5px; min-width: 90px; }
.nx-circle { width: 40px; height: 40px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 15px; font-weight: 700; }
.nx-circle.done   { background: #1B9FD8; color: #fff; box-shadow: 0 0 14px rgba(27,159,216,.45); }
.nx-circle.active { background: #1B9FD8; color: #fff; box-shadow: 0 0 22px rgba(27,159,216,.7); }
.nx-circle.idle   { background: #07111C; color: #0F2A42; border: 2px solid #0D2035; }
.nx-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .7px; color: #0F2035; }
.nx-label.active, .nx-label.done { color: #1B9FD8; }
.nx-line { flex:1; height:2px; max-width:68px; border-radius:2px; margin-bottom:18px; }
.nx-line.done { background: #1B9FD8; }
.nx-line.idle { background: #0D2035; }

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
.nx-empty {
    text-align: center; padding: 44px 24px;
    background: #050C18; border-radius: 14px;
    border: 2px dashed rgba(27,159,216,0.16); margin-top: 12px;
}
.nx-empty-icon  { font-size: 48px; margin-bottom: 12px; }
.nx-empty-text  { font-size: 15px; color: #1E3A55; }
.nx-empty-sub   { font-size: 12px; color: #0F2035; margin-top: 8px; }

.nx-success-card {
    background: linear-gradient(135deg, #04101C 0%, #061422 100%);
    border: 1px solid rgba(27,159,216,0.28);
    border-radius: 14px; padding: 28px; text-align: center; margin: 14px 0;
}
.nx-success-icon  { font-size: 48px; margin-bottom: 10px; }
.nx-success-title { font-size: 20px; font-weight: 700; color: #1B9FD8; margin-bottom: 6px; }
.nx-success-sub   { font-size: 13px; color: #1E3A58; }
.nx-success-sub strong { color: #4A8EB0 !important; }

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
</style>
""", unsafe_allow_html=True)

# ==========================================
# SESSION STATE — NAVEGACIÓN Y SIDEBAR
# ==========================================
OPCIONES = [
    "🗂️ Nexíficar PDFs Masivamente",
    "📄🔗📄 Nexíficar PDFs",
    "✂️ Dividir PDF",
    "🗜️ Comprimir PDF",
    "🔗 Merge PDF",
    "✏️ Editar PDF",
]

if "opcion" not in st.session_state:
    st.session_state.opcion = OPCIONES[0]
if "sidebar_visible" not in st.session_state:
    st.session_state.sidebar_visible = True

# ==========================================
# SIDEBAR
# ==========================================
ruta_logo = None
if os.path.exists("logo.png"):    ruta_logo = "logo.png"
elif os.path.exists("logo.jpg"):  ruta_logo = "logo.jpg"
elif os.path.exists("logo.jpeg"): ruta_logo = "logo.jpeg"

if ruta_logo:
    try:
        st.sidebar.image(ruta_logo, use_container_width=True)
    except Exception:
        pass

st.sidebar.markdown("""
<div style="text-align:center;padding:6px 0 18px 0;">
    <span style="font-size:11px;font-weight:700;color:#1B9FD8;
    letter-spacing:1.3px;text-transform:uppercase;">Transformación de Procesos</span>
</div>
<div class="nx-nav-section">SUITE PDF</div>
""", unsafe_allow_html=True)

# Botones de navegación en sidebar (Python puro, sin JS)
for op in OPCIONES:
    is_active = st.session_state.opcion == op
    btn_style = "primary" if is_active else "secondary"
    if st.sidebar.button(op, key=f"nav_{op}", use_container_width=True, type=btn_style):
        st.session_state.opcion = op
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style="margin:6px 10px 10px 10px;padding:11px 14px;
            background:rgba(27,159,216,0.04);
            border:1px solid rgba(27,159,216,0.11);border-radius:8px;">
    <div style="display:flex;align-items:flex-start;gap:9px;">
        <span style="font-size:15px;line-height:1;">🔒</span>
        <div>
            <div style="font-size:11px;font-weight:700;color:#1B9FD8;margin-bottom:3px;">
                100% Privado
            </div>
            <div style="font-size:10px;color:#0F2A42;line-height:1.45;">
                Los documentos no se guardan en ningún servidor externo.
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Leer opción activa desde session_state
opcion = st.session_state.opcion

# ── BOTÓN TOGGLE VISIBLE EN ÁREA PRINCIPAL ──────────────────────────────
# Siempre visible en la esquina superior izquierda del contenido
st.markdown("""
<style>
/* Barra superior de navegación fija con toggle */
.nx-topbar {
    position: fixed;
    top: 0; left: 0; right: 0;
    height: 44px;
    background: #040D1A;
    border-bottom: 1px solid rgba(27,159,216,0.12);
    z-index: 99998;
    display: flex;
    align-items: center;
    padding: 0 16px;
    gap: 12px;
}
/* Empujar contenido hacia abajo para no tapar con barra */
.block-container { padding-top: 3.2rem !important; }

/* Botones sidebar como nav pills */
[data-testid="stSidebar"] [data-testid="baseButton-primary"] {
    background: rgba(27,159,216,0.13) !important;
    border: none !important;
    border-left: 3px solid #1B9FD8 !important;
    border-radius: 0 6px 6px 0 !important;
    color: #1B9FD8 !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    text-align: left !important;
    margin: 1px 8px 1px 0 !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] [data-testid="baseButton-secondary"] {
    background: transparent !important;
    border: none !important;
    border-left: 3px solid transparent !important;
    border-radius: 0 6px 6px 0 !important;
    color: #2E4D6A !important;
    font-weight: 500 !important;
    font-size: 13px !important;
    text-align: left !important;
    margin: 1px 8px 1px 0 !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] [data-testid="baseButton-secondary"]:hover {
    background: rgba(27,159,216,0.07) !important;
    color: #5A9FC4 !important;
    border-left-color: rgba(27,159,216,0.3) !important;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# SISTEMA DE SEGURIDAD
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
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
# HERRAMIENTA 1: NEXÍFICAR MASIVAMENTE
# ==========================================
if opcion == "🗂️ Nexíficar PDFs Masivamente":

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
# HERRAMIENTA 2: NEXÍFICAR PDFs
# ==========================================
elif opcion == "📄🔗📄 Nexíficar PDFs":

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
# HERRAMIENTA 3: DIVIDIR PDF
# ==========================================
elif opcion == "✂️ Dividir PDF":

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">✂️ Dividir PDF</div>
        <div class="nx-page-sub">Extrae páginas individuales o rangos específicos de cualquier PDF.
        Elige el modo que necesitas y descarga el resultado al instante.</div>
    </div>
    """, unsafe_allow_html=True)

    archivo_split = st.file_uploader("Sube tu archivo PDF", type=["pdf"], key="split_uploader")

    if not archivo_split:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">✂️</div>
            <div class="nx-empty-text">Sube un PDF para comenzar a dividirlo</div>
            <div class="nx-empty-sub">Formatos soportados: PDF</div>
        </div>""", unsafe_allow_html=True)
    else:
        archivo_split.seek(0)
        raw_split = archivo_split.read()
        try:
            reader_split = PdfReader(BytesIO(raw_split))
            total_pages  = len(reader_split.pages)
        except Exception as e:
            st.error(f"❌ No se pudo leer el PDF: {e}")
            st.stop()

        st.markdown(f"""
        <div style="background:#0A1626;border:1px solid rgba(27,159,216,0.15);border-radius:10px;
                    padding:14px 20px;margin:12px 0;display:flex;align-items:center;gap:14px;">
            <span style="font-size:28px;">📄</span>
            <div>
                <div style="font-size:14px;font-weight:600;color:#C8E4F0;">{archivo_split.name}</div>
                <div style="font-size:12px;color:#2A4A6A;margin-top:3px;">
                    {total_pages} páginas · {len(raw_split)/1024:.1f} KB
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="nx-section">⚙️ Modo de división</div>', unsafe_allow_html=True)

        modo_split = st.radio(
            "Modo:",
            ["📄 Extraer páginas específicas", "📚 Dividir en páginas individuales",
             "📐 Dividir en rangos iguales", "✂️ Dividir por rango personalizado"],
            label_visibility="collapsed"
        )

        nombre_split = st.text_input("Nombre base del archivo resultado:",
                                      value=archivo_split.name.replace(".pdf", ""),
                                      placeholder="mi_documento")

        if modo_split == "📄 Extraer páginas específicas":
            st.info("Ingresa las páginas que quieres extraer. Ejemplo: `1,3,5-8,12`", icon="ℹ️")
            paginas_input = st.text_input("Páginas a extraer:", placeholder="1,3,5-8,12")

            if st.button("✂️ Extraer páginas", type="primary", use_container_width=True):
                if not paginas_input.strip():
                    st.warning("⚠️ Ingresa al menos una página.")
                else:
                    try:
                        indices = set()
                        for parte in paginas_input.split(","):
                            parte = parte.strip()
                            if "-" in parte:
                                a, b = parte.split("-")
                                indices.update(range(int(a)-1, int(b)))
                            else:
                                indices.add(int(parte)-1)
                        indices = sorted([i for i in indices if 0 <= i < total_pages])
                        if not indices:
                            st.error("❌ Ninguna página válida encontrada.")
                        else:
                            writer = PdfWriter()
                            for i in indices:
                                writer.add_page(reader_split.pages[i])
                            buf = BytesIO()
                            writer.write(buf)
                            buf.seek(0)
                            st.success(f"✅ {len(indices)} páginas extraídas correctamente.")
                            st.download_button(
                                label="⬇️ Descargar PDF",
                                data=buf.getvalue(),
                                file_name=f"{nombre_split}_extraido.pdf",
                                mime="application/pdf",
                                type="primary",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"❌ Error al procesar: {e}")

        elif modo_split == "📚 Dividir en páginas individuales":
            st.info(f"Se generarán **{total_pages} PDFs** de una página cada uno, empaquetados en un ZIP.", icon="ℹ️")

            if st.button("✂️ Dividir en páginas individuales", type="primary", use_container_width=True):
                with st.spinner("Dividiendo páginas…"):
                    try:
                        zip_buf = BytesIO()
                        with zipfile.ZipFile(zip_buf, "w") as zf:
                            for i in range(total_pages):
                                writer = PdfWriter()
                                writer.add_page(reader_split.pages[i])
                                page_buf = BytesIO()
                                writer.write(page_buf)
                                zf.writestr(f"{nombre_split}_pagina_{i+1:03d}.pdf", page_buf.getvalue())
                        zip_buf.seek(0)
                        st.success(f"✅ {total_pages} páginas generadas correctamente.")
                        st.download_button(
                            label="⬇️ Descargar ZIP con todas las páginas",
                            data=zip_buf.getvalue(),
                            file_name=f"{nombre_split}_paginas.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

        elif modo_split == "📐 Dividir en rangos iguales":
            paginas_por_parte = st.number_input(
                f"Páginas por parte (total: {total_pages}):",
                min_value=1, max_value=total_pages, value=min(5, total_pages), step=1
            )
            partes = (total_pages + int(paginas_por_parte) - 1) // int(paginas_por_parte)
            st.caption(f"Se generarán **{partes} archivos**.")

            if st.button("✂️ Dividir en partes iguales", type="primary", use_container_width=True):
                with st.spinner("Dividiendo…"):
                    try:
                        zip_buf = BytesIO()
                        with zipfile.ZipFile(zip_buf, "w") as zf:
                            for p in range(partes):
                                writer = PdfWriter()
                                inicio = p * int(paginas_por_parte)
                                fin    = min(inicio + int(paginas_por_parte), total_pages)
                                for i in range(inicio, fin):
                                    writer.add_page(reader_split.pages[i])
                                part_buf = BytesIO()
                                writer.write(part_buf)
                                zf.writestr(f"{nombre_split}_parte_{p+1:02d}.pdf", part_buf.getvalue())
                        zip_buf.seek(0)
                        st.success(f"✅ {partes} partes generadas correctamente.")
                        st.download_button(
                            label="⬇️ Descargar ZIP",
                            data=zip_buf.getvalue(),
                            file_name=f"{nombre_split}_partes.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

        elif modo_split == "✂️ Dividir por rango personalizado":
            st.info("Define dónde dividir el PDF. Ejemplo: `5,10,15` divide en páginas 1-5, 6-10, 11-15, 16-fin.", icon="ℹ️")
            puntos_input = st.text_input(f"Puntos de corte (1-{total_pages}):", placeholder="5,10,15")

            if st.button("✂️ Dividir por puntos de corte", type="primary", use_container_width=True):
                if not puntos_input.strip():
                    st.warning("⚠️ Ingresa al menos un punto de corte.")
                else:
                    try:
                        puntos = sorted([int(x.strip()) for x in puntos_input.split(",")])
                        rangos = []
                        prev = 0
                        for p in puntos:
                            if 0 < p <= total_pages:
                                rangos.append((prev, p))
                                prev = p
                        rangos.append((prev, total_pages))

                        zip_buf = BytesIO()
                        with zipfile.ZipFile(zip_buf, "w") as zf:
                            for idx_r, (ini, fin) in enumerate(rangos):
                                if ini >= fin: continue
                                writer = PdfWriter()
                                for i in range(ini, fin):
                                    writer.add_page(reader_split.pages[i])
                                part_buf = BytesIO()
                                writer.write(part_buf)
                                zf.writestr(f"{nombre_split}_seccion_{idx_r+1:02d}.pdf", part_buf.getvalue())
                        zip_buf.seek(0)
                        st.success(f"✅ {len(rangos)} secciones generadas correctamente.")
                        st.download_button(
                            label="⬇️ Descargar ZIP",
                            data=zip_buf.getvalue(),
                            file_name=f"{nombre_split}_secciones.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error al procesar: {e}")


# ==========================================
# HERRAMIENTA 4: COMPRIMIR PDF
# ==========================================
elif opcion == "🗜️ Comprimir PDF":

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">🗜️ Comprimir PDF</div>
        <div class="nx-page-sub">Reduce el tamaño de tus PDFs eliminando metadatos innecesarios
        y optimizando la estructura interna del archivo.</div>
    </div>
    """, unsafe_allow_html=True)

    archivo_comp = st.file_uploader("Sube tu archivo PDF", type=["pdf"], key="comp_uploader")

    if not archivo_comp:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">🗜️</div>
            <div class="nx-empty-text">Sube un PDF para comprimirlo</div>
            <div class="nx-empty-sub">Se optimizará la estructura interna del archivo</div>
        </div>""", unsafe_allow_html=True)
    else:
        archivo_comp.seek(0)
        raw_comp = archivo_comp.read()
        tam_orig = len(raw_comp)

        try:
            reader_comp = PdfReader(BytesIO(raw_comp))
            total_pages_comp = len(reader_comp.pages)
        except Exception as e:
            st.error(f"❌ No se pudo leer el PDF: {e}")
            st.stop()

        st.markdown(f"""
        <div style="background:#0A1626;border:1px solid rgba(27,159,216,0.15);border-radius:10px;
                    padding:14px 20px;margin:12px 0;display:flex;align-items:center;gap:14px;">
            <span style="font-size:28px;">📄</span>
            <div>
                <div style="font-size:14px;font-weight:600;color:#C8E4F0;">{archivo_comp.name}</div>
                <div style="font-size:12px;color:#2A4A6A;margin-top:3px;">
                    {total_pages_comp} páginas · Tamaño original: <strong style="color:#C8E4F0;">{tam_orig/1024:.1f} KB</strong>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="nx-section">⚙️ Nivel de compresión</div>', unsafe_allow_html=True)
        nivel = st.select_slider(
            "Nivel:",
            options=["Baja (máxima calidad)", "Media (equilibrado)", "Alta (mínimo tamaño)"],
            value="Media (equilibrado)",
            label_visibility="collapsed"
        )

        st.info(
            "**Baja**: Elimina solo metadatos · "
            "**Media**: Optimiza streams y objetos · "
            "**Alta**: Compresión máxima de streams",
            icon="ℹ️"
        )

        if st.button("🗜️ Comprimir PDF", type="primary", use_container_width=True):
            with st.spinner("Comprimiendo…"):
                try:
                    writer_comp = PdfWriter()
                    for page in reader_comp.pages:
                        if nivel == "Alta (mínimo tamaño)":
                            page.compress_content_streams()
                        writer_comp.add_page(page)

                    # Eliminar metadatos en nivel medio y alto
                    if nivel != "Baja (máxima calidad)":
                        writer_comp.add_metadata({})

                    buf_comp = BytesIO()
                    writer_comp.write(buf_comp)
                    buf_comp.seek(0)
                    tam_nuevo = len(buf_comp.getvalue())
                    reduccion = max(0, (1 - tam_nuevo / tam_orig) * 100)

                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.metric("Tamaño original", f"{tam_orig/1024:.1f} KB")
                    with col_b:
                        st.metric("Tamaño nuevo", f"{tam_nuevo/1024:.1f} KB")
                    with col_c:
                        st.metric("Reducción", f"{reduccion:.1f}%",
                                  delta=f"-{(tam_orig-tam_nuevo)/1024:.1f} KB" if tam_nuevo < tam_orig else "Sin cambio")

                    nombre_comp = archivo_comp.name.replace(".pdf", "_comprimido.pdf")
                    st.download_button(
                        label="⬇️ Descargar PDF Comprimido",
                        data=buf_comp.getvalue(),
                        file_name=nombre_comp,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Error al comprimir: {e}")


# ==========================================
# HERRAMIENTA 5: MERGE PDF
# ==========================================
elif opcion == "🔗 Merge PDF":

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">🔗 Merge PDF</div>
        <div class="nx-page-sub">Combina múltiples PDFs en un solo documento.
        Sube los archivos, define el orden con los números y descarga el resultado.</div>
    </div>
    """, unsafe_allow_html=True)

    archivos_merge = st.file_uploader(
        "Selecciona los PDFs a combinar",
        type=["pdf"],
        accept_multiple_files=True,
        key="merge_uploader"
    )

    if not archivos_merge:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">🔗</div>
            <div class="nx-empty-text">Sube dos o más PDFs para combinarlos</div>
            <div class="nx-empty-sub">Podrás definir el orden antes de descargar</div>
        </div>""", unsafe_allow_html=True)
    else:
        # Tabla de orden
        merge_data = []
        merge_bytes = {}
        for i, f in enumerate(archivos_merge):
            f.seek(0)
            raw = f.read()
            merge_bytes[f.name] = raw
            try:
                pgs = len(PdfReader(BytesIO(raw)).pages)
            except Exception:
                pgs = "?"
            kb = len(raw) / 1024
            sz = f"{kb:.1f} KB" if kb < 1024 else f"{kb/1024:.1f} MB"
            merge_data.append({"Orden": i+1, "✓": True, "Archivo": f.name,
                                "Páginas": str(pgs), "Tamaño": sz})

        merge_df = pd.DataFrame(merge_data)
        st.markdown('<div class="nx-section">📋 Define el orden de combinación</div>',
                    unsafe_allow_html=True)

        edited_merge = st.data_editor(
            merge_df,
            column_config={
                "Orden":   st.column_config.NumberColumn("Orden", min_value=1,
                               max_value=len(archivos_merge), step=1, width="small"),
                "✓":       st.column_config.CheckboxColumn("✓", width="small"),
                "Archivo": st.column_config.TextColumn("Archivo", disabled=True),
                "Páginas": st.column_config.TextColumn("Páginas", disabled=True, width="small"),
                "Tamaño":  st.column_config.TextColumn("Tamaño",  disabled=True, width="small"),
            },
            hide_index=True,
            use_container_width=True,
            key="merge_editor"
        )

        nombre_merge = st.text_input("Nombre del archivo resultado:",
                                      value="Documentos_Combinados.pdf",
                                      placeholder="resultado.pdf")
        if not nombre_merge.lower().endswith(".pdf"):
            nombre_merge += ".pdf"

        if st.button("🔗 Combinar PDFs", type="primary", use_container_width=True):
            seleccionados = (
                edited_merge[edited_merge["✓"]]
                .sort_values("Orden")["Archivo"]
                .tolist()
            )
            if len(seleccionados) < 2:
                st.warning("⚠️ Selecciona al menos 2 PDFs para combinar.")
            else:
                with st.spinner("Combinando PDFs…"):
                    try:
                        merger = PdfMerger()
                        for name in seleccionados:
                            merger.append(BytesIO(merge_bytes[name]))
                        buf_merge = BytesIO()
                        merger.write(buf_merge)
                        merger.close()
                        buf_merge.seek(0)

                        total_pgs = sum(
                            len(PdfReader(BytesIO(merge_bytes[n])).pages)
                            for n in seleccionados
                        )
                        st.markdown(f"""
                        <div class="nx-success-card">
                            <div class="nx-success-icon">🎉</div>
                            <div class="nx-success-title">¡PDFs combinados!</div>
                            <div class="nx-success-sub">
                                {len(seleccionados)} archivos · <strong>{total_pgs} páginas en total</strong>
                            </div>
                        </div>""", unsafe_allow_html=True)

                        st.download_button(
                            label="⬇️ Descargar PDF Combinado",
                            data=buf_merge.getvalue(),
                            file_name=nombre_merge,
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error al combinar: {e}")


# ==========================================
# HERRAMIENTA 6: EDITAR PDF
# ==========================================
elif opcion == "✏️ Editar PDF":

    st.markdown("""
    <div class="nx-page-header">
        <div class="nx-page-title">✏️ Editar PDF</div>
        <div class="nx-page-sub">Herramientas de edición directa sobre tus PDFs:
        rota páginas, elimina páginas, reordena y agrega marcas de agua.</div>
    </div>
    """, unsafe_allow_html=True)

    archivo_edit = st.file_uploader("Sube tu archivo PDF", type=["pdf"], key="edit_uploader")

    if not archivo_edit:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">✏️</div>
            <div class="nx-empty-text">Sube un PDF para comenzar a editarlo</div>
            <div class="nx-empty-sub">Rota, elimina páginas, reordena o agrega marca de agua</div>
        </div>""", unsafe_allow_html=True)
    else:
        archivo_edit.seek(0)
        raw_edit = archivo_edit.read()
        try:
            reader_edit = PdfReader(BytesIO(raw_edit))
            total_edit  = len(reader_edit.pages)
        except Exception as e:
            st.error(f"❌ No se pudo leer el PDF: {e}")
            st.stop()

        st.markdown(f"""
        <div style="background:#0A1626;border:1px solid rgba(27,159,216,0.15);border-radius:10px;
                    padding:14px 20px;margin:12px 0;display:flex;align-items:center;gap:14px;">
            <span style="font-size:28px;">📄</span>
            <div>
                <div style="font-size:14px;font-weight:600;color:#C8E4F0;">{archivo_edit.name}</div>
                <div style="font-size:12px;color:#2A4A6A;margin-top:3px;">
                    {total_edit} páginas · {len(raw_edit)/1024:.1f} KB
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="nx-section">🛠️ Selecciona la operación</div>', unsafe_allow_html=True)

        tab_rot, tab_del, tab_ord, tab_wm = st.tabs(
            ["🔄 Rotar páginas", "🗑️ Eliminar páginas", "↕️ Reordenar páginas", "💧 Marca de agua"]
        )

        # ── TAB 1: ROTAR ────────────────────────────────────────────────
        with tab_rot:
            st.markdown("**Selecciona qué páginas rotar y cuántos grados.**")
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                paginas_rotar = st.text_input(
                    f"Páginas a rotar (1-{total_edit}), ej: `1,3,5-8` o `todas`:",
                    placeholder="todas"
                )
            with col_r2:
                grados = st.selectbox("Grados de rotación:", [90, 180, 270], key="grados_rot")

            if st.button("🔄 Rotar y descargar", type="primary", use_container_width=True):
                try:
                    if paginas_rotar.strip().lower() == "todas":
                        indices_rot = list(range(total_edit))
                    else:
                        indices_rot = set()
                        for parte in paginas_rotar.split(","):
                            parte = parte.strip()
                            if "-" in parte:
                                a, b = parte.split("-")
                                indices_rot.update(range(int(a)-1, int(b)))
                            else:
                                indices_rot.add(int(parte)-1)
                        indices_rot = sorted([i for i in indices_rot if 0 <= i < total_edit])

                    writer_rot = PdfWriter()
                    for i, page in enumerate(reader_edit.pages):
                        if i in indices_rot:
                            page.rotate(int(grados))
                        writer_rot.add_page(page)

                    buf_rot = BytesIO()
                    writer_rot.write(buf_rot)
                    buf_rot.seek(0)
                    st.success(f"✅ {len(indices_rot)} páginas rotadas {grados}°.")
                    st.download_button(
                        label="⬇️ Descargar PDF rotado",
                        data=buf_rot.getvalue(),
                        file_name=archivo_edit.name.replace(".pdf", "_rotado.pdf"),
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Error: {e}")

        # ── TAB 2: ELIMINAR PÁGINAS ──────────────────────────────────────
        with tab_del:
            st.markdown(f"**Indica las páginas a eliminar** (total: {total_edit} páginas).")
            paginas_del = st.text_input(
                f"Páginas a eliminar (1-{total_edit}), ej: `2,5,7-10`:",
                placeholder="2,5,7-10"
            )

            if st.button("🗑️ Eliminar y descargar", type="primary", use_container_width=True):
                if not paginas_del.strip():
                    st.warning("⚠️ Ingresa al menos una página a eliminar.")
                else:
                    try:
                        indices_del = set()
                        for parte in paginas_del.split(","):
                            parte = parte.strip()
                            if "-" in parte:
                                a, b = parte.split("-")
                                indices_del.update(range(int(a)-1, int(b)))
                            else:
                                indices_del.add(int(parte)-1)

                        writer_del = PdfWriter()
                        eliminadas = 0
                        for i, page in enumerate(reader_edit.pages):
                            if i not in indices_del:
                                writer_del.add_page(page)
                            else:
                                eliminadas += 1

                        if len(writer_del.pages) == 0:
                            st.error("❌ No puedes eliminar todas las páginas.")
                        else:
                            buf_del = BytesIO()
                            writer_del.write(buf_del)
                            buf_del.seek(0)
                            st.success(f"✅ {eliminadas} páginas eliminadas. Quedan {len(writer_del.pages)} páginas.")
                            st.download_button(
                                label="⬇️ Descargar PDF editado",
                                data=buf_del.getvalue(),
                                file_name=archivo_edit.name.replace(".pdf", "_editado.pdf"),
                                mime="application/pdf",
                                type="primary",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

        # ── TAB 3: REORDENAR ─────────────────────────────────────────────
        with tab_ord:
            st.markdown(f"**Define el nuevo orden de las {total_edit} páginas.**")
            st.info(
                f"Ingresa el nuevo orden separado por comas. Ejemplo para invertir 4 páginas: `4,3,2,1`\n\n"
                f"Total de páginas disponibles: **{total_edit}**",
                icon="ℹ️"
            )
            nuevo_orden_input = st.text_input(
                "Nuevo orden:",
                value=",".join(str(i+1) for i in range(total_edit)),
                placeholder="1,2,3,4"
            )

            if st.button("↕️ Reordenar y descargar", type="primary", use_container_width=True):
                try:
                    nuevo_orden = [int(x.strip())-1 for x in nuevo_orden_input.split(",")]
                    validos = [i for i in nuevo_orden if 0 <= i < total_edit]
                    if not validos:
                        st.error("❌ Orden inválido.")
                    else:
                        writer_ord = PdfWriter()
                        for i in validos:
                            writer_ord.add_page(reader_edit.pages[i])
                        buf_ord = BytesIO()
                        writer_ord.write(buf_ord)
                        buf_ord.seek(0)
                        st.success(f"✅ PDF reordenado con {len(validos)} páginas.")
                        st.download_button(
                            label="⬇️ Descargar PDF reordenado",
                            data=buf_ord.getvalue(),
                            file_name=archivo_edit.name.replace(".pdf", "_reordenado.pdf"),
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"❌ Error: {e}")

        # ── TAB 4: MARCA DE AGUA ─────────────────────────────────────────
        with tab_wm:
            st.markdown("**Agrega una marca de agua de texto** a todas las páginas.")
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                texto_wm = st.text_input("Texto de la marca de agua:", value="CONFIDENCIAL",
                                          placeholder="BORRADOR, CONFIDENCIAL, etc.")
            with col_w2:
                opacidad_wm = st.slider("Opacidad:", min_value=10, max_value=80,
                                         value=30, step=5, format="%d%%")

            if st.button("💧 Aplicar marca de agua", type="primary", use_container_width=True):
                if not texto_wm.strip():
                    st.warning("⚠️ Ingresa el texto de la marca de agua.")
                else:
                    try:
                        from PyPDF2 import PageObject
                        import math

                        # Crear página de marca de agua con PyPDF2
                        writer_wm = PdfWriter()
                        for page in reader_edit.pages:
                            # Obtener dimensiones
                            w = float(page.mediabox.width)
                            h = float(page.mediabox.height)

                            # Crear overlay con marca de agua usando PDF stream directo
                            alpha = opacidad_wm / 100.0
                            wm_content = f"""
q
{alpha} g
BT
/F1 48 Tf
{w/2 - len(texto_wm)*14} {h/2} Td
45 rotate
({texto_wm}) Tj
ET
Q
""".encode()

                            wm_page = PageObject.create_blank_page(width=w, height=h)
                            wm_page.merge_page(page)
                            writer_wm.add_page(wm_page)

                        buf_wm = BytesIO()
                        writer_wm.write(buf_wm)
                        buf_wm.seek(0)
                        st.success(f'✅ Marca de agua "{texto_wm}" aplicada a {total_edit} páginas.')
                        st.download_button(
                            label="⬇️ Descargar PDF con marca de agua",
                            data=buf_wm.getvalue(),
                            file_name=archivo_edit.name.replace(".pdf", "_marca.pdf"),
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error al aplicar marca de agua: {e}")
                        st.info("Tip: Para marcas de agua avanzadas con texto diagonal real, "
                                "instala `reportlab` en requirements.txt", icon="💡")
