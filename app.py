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
# HEADER PRINCIPAL (LOGO)
# ==========================================
ruta_logo = None
if os.path.exists("logo.png"):   ruta_logo = "logo.png"
elif os.path.exists("logo.jpg"): ruta_logo = "logo.jpg"
elif os.path.exists("logo.jpeg"): ruta_logo = "logo.jpeg"

if ruta_logo:
    try:
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            st.image(ruta_logo, use_container_width=True)
        st.markdown("---")
    except Exception:
        pass

# ==========================================
# SISTEMA DE SEGURIDAD (EL CADENERO)
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<h3 style='text-align: center;'>🔒 Acceso Restringido</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Por favor, ingresa tu código de acceso para entrar a la plataforma.</p>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        password = st.text_input("Contraseña / ID de acceso:", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            try:
                claves_validas = list(st.secrets["accesos"].values())
                if password in claves_validas:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Código incorrecto o inactivo. Intenta de nuevo.")
            except Exception:
                st.warning("⚠️ La bóveda de contraseñas no ha sido configurada correctamente en Streamlit.")
    st.stop()


# ==========================================
# MENÚ LATERAL
# ==========================================
st.sidebar.title("🛠️ Automatizaciones NEXA")
st.sidebar.markdown("Elige el proceso que necesitas:")
opcion = st.sidebar.radio(
    "",
    ("🗂️ Nexíficar PDFs Masivamente", "📄🔗📄 Nexíficar PDFs")
)

st.sidebar.markdown("---")
st.sidebar.info("🔒 **100% Privado:** Los documentos procesados aquí no se guardan en ningún servidor externo.")


# ==========================================
# HERRAMIENTA 1: NEXÍFICAR MASIVAMENTE
# ==========================================
if opcion == "🗂️ Nexíficar PDFs Masivamente":
    st.title("🗂️ Nexíficar PDFs Masivamente")
    st.markdown("Ensambla cientos de expedientes al mismo tiempo usando tu **Plantilla de Excel** y archivos **ZIP**, o simplemente utilízalo para **renombrar** tus documentos de forma automática.")

    st.markdown("---")
    archivo_excel = st.file_uploader("📊 1. Sube tu Plantilla de Excel de Mapeo", type=["xlsx"])
    archivos_zip  = st.file_uploader("🗂️ 2. Sube tus archivos ZIP (Puedes seleccionar varios)", type=["zip"], accept_multiple_files=True)
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

    if st.button("Nexíficar Documentos Masivamente", type="primary", use_container_width=True):
        if not archivo_excel or not archivos_zip:
            st.warning("⚠️ Por favor, sube el Excel y al menos un archivo ZIP para comenzar.")
        else:
            with st.spinner('Nexíficando documentos mágicamente… Esto puede tomar unos segundos.'):
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
                                                    for i in range(len(reader.pages)): paginas_pos[p_final].append(reader.pages[i])
                                                else:
                                                    for i in p_spec:
                                                        if i < len(reader.pages): paginas_pos[p_final].append(reader.pages[i])
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
                   ("nx_order", []), ("nx_files_sig", "")]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    # ── CSS ────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    /* ======== STEP BAR ======== */
    .nx-steps {
        display: flex; align-items: center; justify-content: center;
        padding: 10px 0 28px 0;
    }
    .nx-step {
        display: flex; flex-direction: column; align-items: center;
        gap: 6px; min-width: 95px;
    }
    .nx-circle {
        width: 44px; height: 44px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 16px; font-weight: 700; transition: all .3s;
    }
    .nx-circle.done   { background: #1D9E75; color: #fff; box-shadow: 0 0 16px rgba(29,158,117,.6); }
    .nx-circle.active { background: #1D9E75; color: #fff; box-shadow: 0 0 24px rgba(29,158,117,.8); }
    .nx-circle.idle   { background: #0A1826; color: #3A6A8C; border: 2px solid #1A3A5C; }
    .nx-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .6px; color: #3A6A8C; }
    .nx-label.active, .nx-label.done { color: #1D9E75; }
    .nx-line {
        flex: 1; height: 3px; max-width: 72px; border-radius: 2px;
        margin-bottom: 20px; transition: background .3s;
    }
    .nx-line.done { background: #1D9E75; }
    .nx-line.idle { background: #1A3A5C; }

    /* ======== SECTION HEADERS ======== */
    .nx-section {
        font-size: 12px; font-weight: 700; color: #1D9E75;
        text-transform: uppercase; letter-spacing: 1.1px;
        margin: 24px 0 12px 0; display: flex; align-items: center; gap: 10px;
    }
    .nx-section::after {
        content: ''; flex: 1; height: 1px;
        background: linear-gradient(90deg, #1D9E75 0%, transparent 100%);
    }

    /* ======== NEXÍFICAR BUTTON ======== */
    button[kind="primary"] {
        background: linear-gradient(135deg, #1D9E75 0%, #14835D 100%) !important;
        border: none !important; color: #fff !important;
        border-radius: 10px !important; font-size: 17px !important;
        font-weight: 700 !important; letter-spacing: .4px !important;
        box-shadow: 0 4px 20px rgba(29,158,117,.45) !important;
        transition: all .2s !important;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #22B587 0%, #1AAD6E 100%) !important;
        box-shadow: 0 6px 28px rgba(29,158,117,.65) !important;
        transform: translateY(-2px) !important;
    }
    button[kind="primary"]:active { transform: translateY(0px) !important; }

    /* ======== EMPTY STATE ======== */
    .nx-empty {
        text-align: center; padding: 48px 24px;
        background: #071420; border-radius: 16px;
        border: 2px dashed #1A3A5C; margin-top: 14px;
    }
    .nx-empty-icon  { font-size: 54px; margin-bottom: 14px; }
    .nx-empty-text  { font-size: 16px; color: #3A6A8C; }
    .nx-empty-sub   { font-size: 13px; color: #1A3A5C; margin-top: 8px; }

    /* ======== SUCCESS CARD ======== */
    .nx-success-card {
        background: linear-gradient(135deg, #071F14 0%, #0A2B1C 100%);
        border: 2px solid #1D9E75; border-radius: 16px;
        padding: 28px; text-align: center; margin: 16px 0;
    }
    .nx-success-icon  { font-size: 52px; margin-bottom: 10px; }
    .nx-success-title { font-size: 20px; font-weight: 700; color: #1D9E75; margin-bottom: 6px; }
    .nx-success-sub   { font-size: 14px; color: #4ABFA0; }
    </style>
    """, unsafe_allow_html=True)

    # ── Helpers ────────────────────────────────────────────────────────────
    def _render_steps(step):
        cfg = [("1", "Subir PDFs"), ("2", "Ordenar"), ("3", "Unificar")]
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
        """Returns base64-encoded PNG of the first page, or empty string on failure."""
        if not FITZ_OK:
            return ""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(0.9, 0.9), alpha=False)
            return base64.b64encode(pix.tobytes("png")).decode()
        except Exception:
            return ""

    # ── Title ──────────────────────────────────────────────────────────────
    st.title("📄🔗📄 Nexíficar PDFs")
    st.markdown("Sube varios PDFs sueltos y únelos en **un solo archivo**, con previsualización de la primera página y reordenamiento con botones ↑ ↓ sobre cada tarjeta.")

    # ── File uploader ──────────────────────────────────────────────────────
    st.markdown('<div class="nx-section">📂 Paso 1 — Subir PDFs</div>', unsafe_allow_html=True)
    archivos_subidos = st.file_uploader(
        "Selecciona o arrastra tus archivos PDF aquí",
        type=["pdf"],
        accept_multiple_files=True
    )

    # Reset when files are cleared
    if not archivos_subidos:
        st.session_state.nx_done      = False
        st.session_state.nx_buffer    = None
        st.session_state.nx_order     = []
        st.session_state.nx_files_sig = ""

    step = 1 if not archivos_subidos else (3 if st.session_state.nx_done else 2)
    st.markdown(_render_steps(step), unsafe_allow_html=True)

    # ── Empty state ────────────────────────────────────────────────────────
    if not archivos_subidos:
        st.markdown("""
        <div class="nx-empty">
            <div class="nx-empty-icon">📂</div>
            <div class="nx-empty-text">Usa el selector de arriba para cargar tus PDFs</div>
            <div class="nx-empty-sub">Puedes seleccionar múltiples archivos a la vez</div>
        </div>""", unsafe_allow_html=True)

    # ── Pasos 2 / 3 ───────────────────────────────────────────────────────
    else:
        # Build per-file metadata + thumbnails (keyed by filename)
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
                "arch":  arch,
                "name":  arch.name,
                "raw":   raw,
                "pages": pages,
                "size":  sz,
                "thumb": _get_thumb(raw),
            }

        # Sync order with uploaded file set (reset when files change)
        files_sig = hashlib.md5(
            "".join(sorted(file_info_map.keys())).encode()
        ).hexdigest()[:8]

        if st.session_state.nx_files_sig != files_sig:
            st.session_state.nx_files_sig = files_sig
            st.session_state.nx_order     = list(file_info_map.keys())
            st.session_state.nx_done      = False
            st.session_state.nx_buffer    = None

        # Remove stale names in case a file was removed from the uploader
        st.session_state.nx_order = [n for n in st.session_state.nx_order
                                      if n in file_info_map]
        orden = st.session_state.nx_order

        if not st.session_state.nx_done:
            # ── Card grid with ↑ ↓ ✕ buttons (Paso 2) ───────────────────
            st.markdown(
                '<div class="nx-section">↕️ Paso 2 — Ordena los PDFs con los botones ↑ ↓</div>',
                unsafe_allow_html=True
            )

            CARDS_PER_ROW = 4
            n_total = len(orden)

            for row_start in range(0, n_total, CARDS_PER_ROW):
                row_slice = orden[row_start : row_start + CARDS_PER_ROW]
                cols = st.columns(CARDS_PER_ROW)
                for col_offset, name in enumerate(row_slice):
                    fi  = file_info_map[name]
                    idx = row_start + col_offset
                    with cols[col_offset]:
                        with st.container(border=True):
                            # Thumbnail or placeholder
                            if fi["thumb"]:
                                st.image(
                                    base64.b64decode(fi["thumb"]),
                                    use_container_width=True
                                )
                            else:
                                st.markdown(
                                    '<div style="background:#0D1E30;border-radius:8px;'
                                    'height:110px;display:flex;align-items:center;'
                                    'justify-content:center;font-size:38px;">📄</div>',
                                    unsafe_allow_html=True
                                )

                            # Order badge + truncated filename
                            short = (fi["name"][:22] + "…") if len(fi["name"]) > 22 else fi["name"]
                            st.markdown(
                                f'<div style="font-size:12px;font-weight:600;color:#C8E8DF;'
                                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'
                                f'margin:6px 0 2px 0;">'
                                f'<span style="background:#1D9E75;color:#fff;border-radius:50%;'
                                f'padding:1px 7px;margin-right:5px;font-size:11px;font-weight:700;">'
                                f'{idx + 1}</span>{short}</div>',
                                unsafe_allow_html=True
                            )
                            st.caption(f"📄 {fi['pages']} pág. · 💾 {fi['size']}")

                            # Control buttons: ↑  ↓  ✕
                            b_up, b_dn, b_del = st.columns(3)
                            with b_up:
                                if st.button("↑", key=f"up_{idx}",
                                             disabled=(idx == 0),
                                             use_container_width=True):
                                    orden[idx], orden[idx - 1] = orden[idx - 1], orden[idx]
                                    st.rerun()
                            with b_dn:
                                if st.button("↓", key=f"dn_{idx}",
                                             disabled=(idx == n_total - 1),
                                             use_container_width=True):
                                    orden[idx], orden[idx + 1] = orden[idx + 1], orden[idx]
                                    st.rerun()
                            with b_del:
                                if st.button("✕", key=f"dl_{idx}",
                                             use_container_width=True):
                                    orden.pop(idx)
                                    st.rerun()

            # ── File name ─────────────────────────────────────────────────
            st.markdown('<div class="nx-section">💾 Nombre del PDF final</div>',
                        unsafe_allow_html=True)
            nombre_final = st.text_input(
                "Nombre del archivo unificado:",
                "Documento_Unificado.pdf",
                label_visibility="collapsed"
            )
            if not nombre_final.lower().endswith(".pdf"):
                nombre_final += ".pdf"

            st.markdown("<br>", unsafe_allow_html=True)

            # ── NEXÍFICAR BUTTON ──────────────────────────────────────────
            if st.button(f"🔗 Nexíficar {len(orden)} PDFs", type="primary", use_container_width=True):
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

        # ── Paso 3: éxito + descarga ──────────────────────────────────────
        else:
            total = len(orden)
            st.markdown(f"""
            <div class="nx-success-card">
                <div class="nx-success-icon">🎉</div>
                <div class="nx-success-title">¡Nexíficación completada!</div>
                <div class="nx-success-sub">{total} PDF{'s' if total != 1 else ''} unidos en
                <strong>{st.session_state.nx_nombre}</strong></div>
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
                st.session_state.nx_done      = False
                st.session_state.nx_buffer    = None
                st.session_state.nx_order     = []
                st.session_state.nx_files_sig = ""
                st.rerun()
