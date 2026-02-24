import streamlit as st
import pandas as pd

import pandas as pd
import zipfile
import os
import shutil
from PyPDF2 import PdfReader, PdfWriter
from google.colab import files

print("🚀 INICIANDO NEXA: MOTOR UNIVERSAL 🚀")
print("Por favor, sube tu Plantilla de Excel (.xlsx) y tu archivo(s) ZIP...")

# 1. Limpieza inicial para que no haya basura de pruebas anteriores
carpetas_limpiar = ['/content/origen', '/content/salida', '/content/NEXA_Resultados.zip']
for ruta in carpetas_limpiar:
    if os.path.exists(ruta):
        shutil.rmtree(ruta) if os.path.isdir(ruta) else os.remove(ruta)

os.makedirs('/content/origen', exist_ok=True)
os.makedirs('/content/salida', exist_ok=True)

# 2. Subida de archivos (¡Aquí subes tu plantilla_nexa y tu ZIP!)
archivos_subidos = files.upload()

# 3. Identificar inteligentemente cuál es el Excel y cuáles son los ZIPs
archivo_excel = None
archivos_zip = []

for nombre_archivo in archivos_subidos.keys():
    if nombre_archivo.endswith('.xlsx'):
        archivo_excel = f'/content/{nombre_archivo}'
        print(f"✅ Excel detectado: {nombre_archivo}")
    elif nombre_archivo.endswith('.zip'):
        archivos_zip.append(f'/content/{nombre_archivo}')
        print(f"✅ ZIP detectado: {nombre_archivo}")

if not archivo_excel or not archivos_zip:
    print("❌ ERROR: Debes subir un archivo Excel (.xlsx) y al menos un archivo .zip")
else:
    print("\n--- DESCOMPRIMIENDO ZIPs ---")
    for zip_path in archivos_zip:
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                z.extractall('/content/origen')
        except Exception as e:
            print(f"⚠️ Error al descomprimir {os.path.basename(zip_path)}: {e}")

    # --- FUNCIONES DEL MOTOR ---
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

    # --- PROCESAMIENTO ---
    print("\n--- INICIANDO ENSAMBLAJE DE DOCUMENTOS ---")
    try:
        df = pd.read_excel(archivo_excel)
        columnas_archivo = [col for col in df.columns if str(col).startswith('Archivo_')]
        exitos = 0

        for idx, row in df.iterrows():
            if 'Nombre_Salida' not in df.columns:
                print("❌ ERROR: El Excel debe tener una columna llamada 'Nombre_Salida'.")
                break

            nombre_salida = str(row['Nombre_Salida']).strip()
            if pd.isna(nombre_salida) or nombre_salida == 'nan' or not nombre_salida: continue
            if not nombre_salida.lower().endswith('.pdf'): nombre_salida += '.pdf'

            ruta_final = os.path.join('/content/salida', nombre_salida)
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
                    r_doc = buscar_archivo_en_dir(n_doc, '/content/origen')
                    if not r_doc:
                        print(f"⚠️ Advertencia: No se encontró '{n_doc}' para crear '{nombre_salida}'.")
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
                        print(f"⚠️ Error leyendo '{n_doc}': {e}")
                        error_fila = True
                        break

                if not error_fila:
                    writer = PdfWriter()
                    for pos_idx in range(1, max_pos + 1):
                        for p_obj in paginas_pos[pos_idx]: writer.add_page(p_obj)
                    if len(writer.pages) > 0:
                        with open(ruta_final, "wb") as f: writer.write(f)
                        print(f"✅ Creado: {nombre_salida}")
                        exitos += 1

        # --- COMPRESIÓN Y DESCARGA ---
        if exitos > 0:
            print("\n--- PREPARANDO DESCARGA ---")
            zip_final = '/content/NEXA_Resultados.zip'
            with zipfile.ZipFile(zip_final, 'w') as z:
                for r, _, archs in os.walk('/content/salida'):
                    for a in archs: z.write(os.path.join(r, a), a)

            print(f"🎉 ¡Proceso finalizado! Se generaron {exitos} documentos.")
            files.download(zip_final)
        else:
            print("\n⚠️ No se generó ningún documento. Revisa tu Excel y que los PDFs existan en los ZIPs.")

    except Exception as e:
        print(f"❌ Error crítico procesando el Excel: {e}")
