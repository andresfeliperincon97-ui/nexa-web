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
   SIDEBAR
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
/* Ocultar el círculo del radio */
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
   COMPONENTES REUTILIZABLES NEXA
══════════════════════════════════════════════════════ */

/* Etiquetas de sección en sidebar */
.nx-nav-section {
    padding: 14px 20px 4px 20px;
    font-size: 10px;
    font-weight: 700;
    color: #0F2035;
    letter-spacing: 2px;
    text-transform: uppercase;
}

/* Cabecera de página */
.nx-page-header {
    padding: 4px 0 20px 0;
    border-bottom: 1px solid rgba(27,159,216,0.1);
    margin-bottom: 22px;
}
.nx-page-title {
    font-size: 22px;
    font-weight: 700;
    color: #E0F0FF;
    line-height: 1.3;
}
.nx-page-sub {
    font-size: 13px;
    color: #2A4A6A;
    margin-top: 5px;
    line-height: 1.55;
}
.nx-page-sub strong { color: #4A7A9C !important; }

/* Barra de pasos */
.nx-steps {
    display: flex; align-items: center; justify-content: center;
    padding: 8px 0 24px 0;
}
.nx-step { display: flex; flex-direction: column; align-items: center; gap: 5px; min-width: 90px; }
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
.nx-empty-icon  { font-size: 48px; margin-bottom: 12px; }
.nx-empty-text  { font-size: 15px; color: #1E3A55; }
.nx-empty-sub   { font-size: 12px; color: #0F2035; margin-top: 8px; }

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
</style>
""", unsafe_allow_html=True)


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

opcion = st.sidebar.radio(
    "",
    ("🗂️ Nexíficar PDFs Masivamente", "📄🔗📄 Nexíficar PDFs"),
    label_visibility="collapsed"
)

st.sidebar.markdown("""
<div class="nx-nav-section" style="margin-top:12px;">EDICIÓN</div>
<div style="padding:9px 16px 9px 20px;color:#0F2035;font-size:13px;font-weight:500;
            display:flex;align-items:center;gap:8px;">
    ✏️ Editor PDF
    <span style="font-size:10px;background:#060F1D;color:#0F2A42;
    padding:2px 8px;border-radius:10px;font-weight:600;
    border:1px solid rgba(27,159,216,0.08);">Próximamente</span>
</div>
""", unsafe_allow_html=True)

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


# ==========================================
# SISTEMA DE SEGURIDAD (EL CADENERO)
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

    # ── Cabecera de página ─────────────────────────────────────────────────
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

    # ── Estado vacío ───────────────────────────────────────────────────────
    if not archivos_subidos:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">📂</div>
            <div class="nx-empty-text">Usa el selector de arriba para cargar tus PDFs</div>
            <div class="nx-empty-sub">Puedes seleccionar múltiples archivos a la vez</div>
        </div>""", unsafe_allow_html=True)

    # ── Pasos 2 / 3 ───────────────────────────────────────────────────────
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

            # Tabla de orden editable
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

            # ── Sección inferior: nombre + botón Nexíficar ─────────────────
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

        # ── PASO 3: éxito + descarga ───────────────────────────────────────
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
