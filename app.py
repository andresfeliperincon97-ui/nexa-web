import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="NEXA - Transformación de Procesos", page_icon="⚙️", layout="wide")

# ==========================================
# HEADER PRINCIPAL (LOGO CENTRADO Y GRANDE)
# ==========================================
# Buscador inteligente de logos (A prueba de extensiones)
ruta_logo = None
if os.path.exists("logo.png"):
    ruta_logo = "logo.png"
elif os.path.exists("logo.jpg"):
    ruta_logo = "logo.jpg"
elif os.path.exists("logo.jpeg"):
    ruta_logo = "logo.jpeg"

if ruta_logo:
    try:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(ruta_logo, use_container_width=True)
        st.markdown("---")
    except Exception:
        pass 

# ==========================================
# MENÚ LATERAL
# ==========================================
st.sidebar.title("🛠️ Automatizaciones NEXA")
st.sidebar.markdown("Elige el proceso que necesitas:")
opcion = st.sidebar.radio(
    "",
    ("🗂️ Nexificar PDFs Masivamente", "📄🔗📄 Nexificar PDFs")
)

st.sidebar.markdown("---")
st.sidebar.info("🔒 **100% Privado:** Los documentos procesados aquí no se guardan en ningún servidor externo.")

# ==========================================
# HERRAMIENTA 1: NEXIFICAR MASIVAMENTE
# ==========================================
if opcion == "🗂️ Nexificar PDFs Masivamente":
    st.title("🗂️ Nexificar PDFs Masivamente")
    st.markdown("Ensambla cientos de expedientes al mismo tiempo usando tu **Plantilla de Excel** y archivos **ZIP**, o simplemente utilízalo para **renombrar** tus documentos de forma automática.")

    st.markdown("---")
    archivo_excel = st.file_uploader("📊 1. Sube tu Plantilla de Excel de Mapeo", type=["xlsx"])
    archivos_zip = st.file_uploader("🗂️ 2. Sube tus archivos ZIP (Puedes seleccionar varios)", type=["zip"], accept_multiple_files=True)
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

    if st.button("Nexificar Documentos Masivamente", type="primary", use_container_width=True):
        if not archivo_excel or not archivos_zip:
            st.warning("⚠️ Por favor, sube el Excel y al menos un archivo ZIP para comenzar.")
        else:
            with st.spinner('Nexificando documentos mágicamente... Esto puede tomar unos segundos.'):
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
                            barra = st.progress(0)
                            exitos = 0
                            errores = []
                            
                            for idx, row in df.iterrows():
                                nombre_salida = str(row['Nombre_Salida']).strip()
                                if pd.isna(nombre_salida) or nombre_salida == 'nan' or not nombre_salida: continue
                                if not nombre_salida.lower().endswith('.pdf'): nombre_salida += '.pdf'
                                
                                ruta_final = os.path.join(ruta_salida, nombre_salida)
                                max_pos = 0
                                docs_a_procesar = []
                                
                                for col_arch in columnas_archivo:
                                    num_index = col_arch.split('_')[1]
                                    col_inst = f'Instrucciones_{num_index}'
                                    if col_inst in df.columns:
                                        nombre_doc = str(row.get(col_arch, '')).strip()
                                        instrucciones = str(row.get(col_inst, '')).strip()
                                        if nombre_doc and nombre_doc != 'nan':
                                            parsed = parse_paginas(instrucciones)
                                            for _, pos in parsed: max_pos = max(max_pos, pos)
                                            docs_a_procesar.append((nombre_doc, parsed))
                                
                                if max_pos > 0:
                                    paginas_pos = [[] for _ in range(max_pos + 1)]
                                    error_fila = False
                                    
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
                                st.success(f"🎉 ¡Proceso finalizado! Se nexificaron {exitos} documentos con éxito.")
                                
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
# HERRAMIENTA 2: NEXIFICAR PDFs
# ==========================================
elif opcion == "📄🔗📄 Nexificar PDFs":
    st.title("📄🔗📄 Nexificar PDFs")
    st.markdown("Sube varios PDFs sueltos y únelos en **un solo archivo**, eligiendo el orden exacto.")

    st.markdown("---")
    archivos_subidos = st.file_uploader("📄 1. Sube todos los PDFs que quieras unir (Selecciona varios a la vez)", type=["pdf"], accept_multiple_files=True)
    
    if archivos_subidos:
        nombres_archivos = [archivo.name for archivo in archivos_subidos]
        diccionario_archivos = {archivo.name: archivo for archivo in archivos_subidos}
        
        st.markdown("### 2. Selecciona el orden")
        st.info("💡 Haz clic en la caja de abajo y selecciona los archivos **en el orden en el que quieres que se unan**.")
        
        orden_seleccionado = st.multiselect("Orden final de los documentos:", nombres_archivos)
        
        st.markdown("### 3. Nombre del archivo final")
        nombre_final = st.text_input("¿Cómo quieres que se llame el PDF unificado?", "Documento_Unificado.pdf")
        if not nombre_final.lower().endswith(".pdf"):
            nombre_final += ".pdf"
            
        st.markdown("---")
        
        if st.button("Nexificar PDFs Ahora", type="primary", use_container_width=True):
            if not orden_seleccionado:
                st.warning("⚠️ Debes seleccionar al menos un documento para unir.")
            else:
                with st.spinner("Nexificando documentos conservando la calidad original..."):
                    try:
                        fusionador = PdfMerger()
                        
                        for nombre in orden_seleccionado:
                            archivo_actual = diccionario_archivos[nombre]
                            archivo_actual.seek(0)
                            fusionador.append(archivo_actual)
                        
                        buffer_salida = BytesIO()
                        fusionador.write(buffer_salida)
                        fusionador.close()
                        
                        buffer_salida.seek(0)
                        
                        st.success(f"✅ ¡Documento '{nombre_final}' creado con éxito!")
                        
                        st.download_button(
                            label="⬇️ Descargar PDF Unificado",
                            data=buffer_salida,
                            file_name=nombre_final,
                            mime="application/pdf",
                            type="primary"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ Ocurrió un error al unir los archivos: {e}")
