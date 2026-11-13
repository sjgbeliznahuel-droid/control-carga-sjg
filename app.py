import datetime
import io
import os
import re
import time
import unicodedata
import json
import pandas as pd
import numpy as np
import requests
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

st.set_page_config(
    page_title="Control de Fichadas - SJG", page_icon="🏭", layout="wide"
)

# Inicializar estado de sesión
if "mostrar_reporte" not in st.session_state:
    st.session_state["mostrar_reporte"] = False
if "cruce_cierre_ejecutado" not in st.session_state:
    st.session_state["cruce_cierre_ejecutado"] = False

# --- CONFIGURACIÓN GOOGLE SHEETS & DRIVE ---
SHEET_ID = "1gGkGjO-ZX7auYdlKAuI9rLEN_KoVah3vdnz7WofcHq0"
URL_DESCARGA_SHEET = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxZOKYHQHzLR-gZionm3Ub1z8Rlrz5l6AtR1D8gLKKIeAqhSGMf62vMfsa4F2s1gcweFA/exec"
DRIVE_FOLDER_ID = "1fMm6XYmw-dwqSqhijX6tHAe2VdVMtA2-"
NOTIF_FOLDER_ID = "1GJGTwWRwmrWT-Zgkp4nQcWmm3kk-eR8y"
CERT_FOLDER_ID = "1DCkt7I45-dq2hTJYxUKaE-p-jLlbiXqg" 

@st.cache_resource
def conectar_drive():
    credenciales_dict = json.loads(st.secrets["credenciales"])
    creds = service_account.Credentials.from_service_account_info(
        credenciales_dict, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

drive_service = conectar_drive()

def normalizar_ot(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if not s or s.lower() in ["nan", "none", "null"]: return ""
    s = re.sub(r'\.0+$', '', s)
    try:
        val_float = float(s)
        if val_float.is_integer():
            return str(int(val_float))
    except Exception:
        pass
    return s

def obtener_archivos_drive():
    query = f"'{DRIVE_FOLDER_ID}' in parents and trashed=false"
    resultados = drive_service.files().list(q=query, orderBy="name desc", pageSize=1000, fields="files(id, name)").execute()
    return resultados.get('files', [])

def obtener_archivos_cert_drive():
    try:
        query = f"'{CERT_FOLDER_ID}' in parents and trashed=false"
        resultados = drive_service.files().list(
            q=query, orderBy="modifiedTime desc", pageSize=1000, fields="files(id, name, mimeType)"
        ).execute()
        archivos = resultados.get('files', [])
        archivos_validos = []
        for f in archivos:
            nombre = f.get('name', '')
            mime = f.get('mimeType', '')
            if nombre.lower().endswith(('.xlsx', '.xltx', '.csv')) or mime == 'application/vnd.google-apps.spreadsheet':
                archivos_validos.append(f)
        return archivos_validos
    except Exception as e:
        st.error(f"⚠️ Error de acceso a la carpeta de Certificación en Drive: {e}")
        return []

def obtener_archivos_notif_drive():
    try:
        query = f"'{NOTIF_FOLDER_ID}' in parents and trashed=false"
        resultados = drive_service.files().list(
            q=query, orderBy="modifiedTime desc", pageSize=1000, fields="files(id, name, mimeType)"
        ).execute()
        archivos = resultados.get('files', [])
        archivos_validos = []
        for f in archivos:
            nombre = f.get('name', '')
            mime = f.get('mimeType', '')
            if nombre.lower().endswith(('.xlsx', '.xltx', '.csv')) or mime == 'application/vnd.google-apps.spreadsheet':
                archivos_validos.append(f)
        return archivos_validos
    except Exception as e:
        st.error(f"⚠️ Error de acceso a la carpeta de Drive: {e}")
        return []

def descargar_archivo_drive(file_id):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    terminado = False
    while not terminado:
        _, terminado = downloader.next_chunk()
    fh.seek(0)
    return fh

def descargar_archivo_notif_drive(file_id, mime_type):
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    terminado = False
    while not terminado:
        _, terminado = downloader.next_chunk()
    fh.seek(0)
    return fh

def detectar_estructura_fichada(df_raw):
    header_row = 0 
    col_leg, col_nom, col_ing, col_egr, col_nov = -1, -1, -1, -1, -1
    if df_raw.empty:
        return header_row, 2, 3, 6, 7, 12

    limite_busqueda = min(20, len(df_raw))
    for idx in range(limite_busqueda):
        row_str = " ".join([str(val).upper() for val in df_raw.iloc[idx].values if pd.notna(val)])
        if "APELLIDO" in row_str or "NOMBRE" in row_str or "LEGAJO" in row_str or "ENTRADA" in row_str:
            header_row = idx
            break

    if header_row < len(df_raw):
        headers = [str(v).upper().strip() for v in df_raw.iloc[header_row].values]
        for idx, val in enumerate(headers):
            if "LEGAJO" in val or "LEG" in val: col_leg = idx
            elif "NOMBRE" in val or "APELLIDO" in val: col_nom = idx
            elif "ING" in val or "ENTRADA" in val or "1RA" in val or "HS.ING" in val or "MARCA" in val: col_ing = idx
            elif "SAL" in val or "EGRESO" in val or "ULTIMA" in val or "HS. SAL" in val: col_egr = idx
            elif "OBRA" in val or "NOVEDAD" in val or "OBSERVAC" in val: col_nov = idx

    if col_leg == -1: col_leg = 2   
    if col_nom == -1: col_nom = 3   
    if col_ing == -1: col_ing = 6   
    if col_egr == -1: col_egr = 7   
    if col_nov == -1: col_nov = 12  
    return header_row + 1, col_leg, col_nom, col_ing, col_egr, col_nov

def cargar_dataframe_fichada(fh_bytes):
    xl = pd.ExcelFile(fh_bytes, engine="openpyxl")
    dfs = []
    hojas_a_leer = xl.sheet_names[1:] if len(xl.sheet_names) > 1 else xl.sheet_names
    for sheet in hojas_a_leer:
        df_s = pd.read_excel(fh_bytes, sheet_name=sheet, header=None)
        if not df_s.empty:
            df_s = df_s.dropna(how="all").dropna(axis=1, how="all")
            if not df_s.empty:
                dfs.append(df_s)
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    return pd.DataFrame()

def convertir_a_horas_decimal(val):
    if pd.isna(val) or val == "": return None
    s_val = str(val).upper().replace("\xa0", " ").strip()
    if s_val in ["--:--", "NAN", "NONE", "NULL", "X", "-"]: return None
    try:
        if isinstance(val, datetime.time): return val.hour + val.minute / 60.0
        if isinstance(val, (pd.Timestamp, datetime.datetime)): return val.hour + val.minute / 60.0
        s_val = s_val.replace(",", ".")
        match = re.search(r"(\d+[:\.]?\d*)", s_val)
        if not match: return None
        clean_val = match.group(1)
        if ":" in clean_val:
            parts = clean_val.split(":")
            return float(parts[0]) + (float(parts[1]) / 60.0 if len(parts) > 1 and parts[1] else 0.0)
        return float(clean_val)
    except Exception:
        return None

def obtener_notificaciones_drive(file_id, mime_type=None):
    if not file_id: return {}, {}
    try:
        fh = descargar_archivo_notif_drive(file_id, mime_type)
        df_raw = pd.read_excel(fh, header=None)
        header_row = 0
        for idx in range(min(15, len(df_raw))):
            row_vals = [str(v).upper() for v in df_raw.iloc[idx].values if pd.notna(v)]
            row_str = " ".join(row_vals)
            if any(k in row_str for k in ["OT", "SOLICITUD", "ORDEN", "NOTIF"]):
                header_row = idx
                break
        raw_headers = [str(v).strip() if pd.notna(v) else f"Col_{i}" for i, v in enumerate(df_raw.iloc[header_row].values)]
        df_notif = df_raw.iloc[header_row + 1:].copy()
        df_notif.columns = raw_headers
        
        ot_col = None
        for c in df_notif.columns:
            c_str = str(c).strip().upper()
            if c_str in ["OT", "ORDEN", "N° OT", "Nº OT", "N° ORDEN", "ORDEN DE TRABAJO", "OT PROGRAMADA"]:
                ot_col = c
                break
        if not ot_col:
            for c in df_notif.columns:
                c_str = str(c).strip().upper()
                if re.search(r'\b(OT|ORDEN)\b', c_str):
                    ot_col = c
                    break

        sol_col = next((c for c in df_notif.columns if any(k in str(c).upper() for k in ["SOLICITUD", "SOLIC", "N° SOL"])), None)
        est_col = next((c for c in df_notif.columns if any(k in str(c).upper() for k in ["ESTADO", "STATUS"])), None)
        fecha_col = df_notif.columns[3] if len(df_notif.columns) > 3 else next((c for c in df_notif.columns if any(k in str(c).upper() for k in ["FECHA", "NOTIF"])), None)

        mapa_notif_acumulado = {}
        if ot_col:
            for _, row in df_notif.iterrows():
                ot_key = normalizar_ot(row[ot_col])
                if not ot_key: continue
                val_sol = str(row[sol_col]).strip() if sol_col and pd.notna(row[sol_col]) and str(row[sol_col]).lower() != "nan" else ""
                val_est = str(row[est_col]).strip() if est_col and pd.notna(row[est_col]) and str(row[est_col]).lower() != "nan" else ""
                val_fecha = ""
                if fecha_col and pd.notna(row[fecha_col]) and str(row[fecha_col]).lower() != "nan":
                    v_f = row[fecha_col]
                    if isinstance(v_f, (pd.Timestamp, datetime.date, datetime.datetime)):
                        val_fecha = v_f.strftime("%d/%m/%Y")
                    else:
                        str_vf = str(v_f).strip()
                        match_dt = re.search(r'(\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4})', str_vf)
                        val_fecha = match_dt.group(1) if match_dt else (str_vf.split()[0] if " " in str_vf else str_vf)

                if ot_key not in mapa_notif_acumulado:
                    mapa_notif_acumulado[ot_key] = {"solicitudes": [], "notificaciones": [], "estados": []}
                if val_sol and val_sol not in mapa_notif_acumulado[ot_key]["solicitudes"]:
                    mapa_notif_acumulado[ot_key]["solicitudes"].append(val_sol)
                if val_fecha and val_fecha not in mapa_notif_acumulado[ot_key]["notificaciones"]:
                    mapa_notif_acumulado[ot_key]["notificaciones"].append(val_fecha)
                if val_est and val_est not in mapa_notif_acumulado[ot_key]["estados"]:
                    mapa_notif_acumulado[ot_key]["estados"].append(val_est)

        mapa_notif_final = {
            k_ot: {
                "solicitud": " - ".join(v_data["solicitudes"]),
                "notificacion": " - ".join(v_data["notificaciones"]),
                "estado": " - ".join(v_data["estados"])
            } for k_ot, v_data in mapa_notif_acumulado.items()
        }
        diag = {"ot_col": ot_col, "sol_col": sol_col, "est_col": est_col, "fecha_col": fecha_col, "total_keys": len(mapa_notif_final)}
        return mapa_notif_final, diag
    except Exception as e:
        st.error(f"Error al procesar la planilla de notificaciones: {e}")
        return {}, {}

def limpiar_legajo(val):
    if pd.isna(val): return None
    try: return int(float(str(val).strip()))
    except (ValueError, TypeError): return None

def normalizar_texto(val):
    if pd.isna(val): return ""
    s = str(val).lower().strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return " ".join("".join(c for c in s if c.isalnum() or c.isspace()).split())

# --- CONFIGURACIONES DE REGLAS Y CONTRATOS ---
EXCLUIDOS_SECTOR = {
    60003127: ("Villaseca Urrutia, Silvina", "Seguridad e Higiene"),
    60017753: ("Miño, Dario Oscar", "Seguridad e Higiene"),
    60016617: ("Ferreyra, Valentin", "Seguridad e Higiene"),
    50027446: ("Romano, Nicolas", "Delegados"),
    50012920: ("Chavez, Fernando", "Delegados"),
    50027313: ("Carrizo, Pedro Paz", "Delegados"),
    50012909: ("Andenmatten, Marcelo", "Delegados"),
    50029799: ("Sinner, Damian", "Personal de Pañol"),
    60008058: ("Nuñez, Gabriel", "Personal de Pañol"),
    60002019: ("Gonzalez, Hector", "Personal de Pañol"),
    60008120: ("Fourcade, Hernan", "Personal de Pañol"),
    60022214: ("Franco, Martin Jesus", "Personal de Pañol"),
    50034415: ("Brozman, Ariel", "Personal de Pañol"),
    60008497: ("Diaz Paez, Alejandro", "Coordinación"),
    60012942: ("Cometti, Juliana", "Coordinación"),
    60003811: ("Gauna, Mauricio", "Coordinación"),
    60008266: ("Lascano, Luis Alejandro", "Coordinación"),
}

CONTRATO_MMYHH_LEGAJOS = {
    60005508: "BUSTOS CELEDONIO JOSE",
    60014119: "CARBALLO NASCHEL MANUEL",
    60021316: "COLANDRE VALENTIN",
    60016432: "CORTES RODRIGO MAXIMILIANO",
    60019145: "CORTEZ CLAUDIO JOSE",
    60014170: "DI VIRGILIO LEANDRO GERMAN",
    60014126: "FARIAS RODRIGO DAMIAN",
    50003371: "SAAVEDRA ALBERTO OSCAR",
    60018676: "SIMEAND DIEGO EXEQUIEL",
    60022855: "ZABALA JUAN CRUZ"
}

def filtrar_por_contrato(legajo_val, nombre_val, tipo_contrato):
    if tipo_contrato == "Todos los Contratos":
        return True
    clean_l = limpiar_legajo(legajo_val)
    es_mmyhh = False
    if clean_l is not None and clean_l in CONTRATO_MMYHH_LEGAJOS:
        es_mmyhh = True
    elif nombre_val and pd.notna(nombre_val):
        n_norm = normalizar_texto(nombre_val)
        for m_nom in CONTRATO_MMYHH_LEGAJOS.values():
            m_norm = normalizar_texto(m_nom)
            tokens_m = set(m_norm.split())
            tokens_n = set(n_norm.split())
            if len(tokens_m) > 0 and len(tokens_m.intersection(tokens_n)) >= min(2, len(tokens_m)):
                es_mmyhh = True
                break
            if m_norm in n_norm or n_norm in m_norm:
                es_mmyhh = True
                break

    if tipo_contrato == "MMYHH":
        return es_mmyhh
    elif tipo_contrato == "MOA":
        return not es_mmyhh
    return True

def obtener_excel_nube(url):
    response = requests.get(url)
    response.raise_for_status()
    return response.content

def obtener_dia_mes(texto_o_fecha):
    if isinstance(texto_o_fecha, (pd.Timestamp, datetime.date, datetime.datetime)):
        return texto_o_fecha.day, texto_o_fecha.month
    match = re.search(r"(\d{1,2})[-/\.](\d{1,2})", str(texto_o_fecha))
    if match:
        try:
            d, m = int(match.group(1)), int(match.group(2))
            if 1 <= d <= 31 and 1 <= m <= 12: return d, m
        except Exception: pass
    return None, None

def extraer_horas(val):
    if pd.isna(val): return 0.0
    match = re.search(r"(\d+(?:[\.,]\d+)?)", str(val).upper().strip())
    if match: return float(match.group(1).replace(",", "."))
    return 0.0

SJG_LOGO_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 130" width="100%" height="75" style="max-width: 380px;"><g fill="#0A0A0A"><text x="5" y="82" font-family="\'Montserrat\', \'Arial Black\', sans-serif" font-weight="900" font-size="78" letter-spacing="-1.5">SJG</text><text x="8" y="116" font-family="\'Montserrat\', \'Arial\', sans-serif" font-weight="800" font-size="17" letter-spacing="3.2">MONTAJES INDUSTRIALES</text><polygon points="230,126 260,126 282,98 230,126" /><polygon points="318,24 294,126 268,126 290,98 318,24" /><polygon points="352,70 302,126 392,126" /></g></svg>'

col_logo, col_titulo, col_refrescar = st.columns([1.2, 2.3, 0.9])
with col_logo: st.markdown(SJG_LOGO_SVG, unsafe_allow_html=True)
with col_titulo:
    st.title("Sistema de Control y Validación de Personal")
    st.caption("Conciliación automatizada de Partes de Carga vs. Fichadas en Puerta (Cloud - En Vivo)")
with col_refrescar:
    st.write("")
    if st.button("🔄 Refrescar App", help="Limpia la caché y fuerza la recarga de datos"):
        st.cache_data.clear()
        st.session_state["mostrar_reporte"] = False
        st.session_state["cruce_cierre_ejecutado"] = False
        st.rerun()

st.divider()

MESES_NOMBRES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

t_rrhh, t_carga, t_reglas, t_reporte, t_cierre, t_quincena, t_detalle, t_cert = st.tabs([
    "📥 Carga de Fichadas/RRHH", "📂 1. Carga", "⚙️ 2. Reglas", "📊 3. Reporte", "📋 Cierre de Período", "⏱️ 4. Horas y Período Personalizado", "🔍 5. Control Horarios", "📁 6. Certificación de Horas"
])

with t_rrhh:
    st.header("Carga Diaria de Fichadas (Portería / RRHH)")
    st.info("ℹ️ Para cargar una nueva fichada, arrastrá el Excel exportado directamente a tu carpeta 'FICHADAS' en Google Drive.")

with t_carga:
    c1, c2 = st.columns(2)
    with c2:
        st.subheader("Base de Fichadas Guardadas (Drive)")
        try:
            archivos_drive = obtener_archivos_drive()
            archivos_guardados = [f['name'] for f in archivos_drive if f['name'].lower().endswith(('.xlsx', '.xltx'))]
            mapa_archivos = {f['name']: f['id'] for f in archivos_drive}
            af_selected = archivos_guardados[0] if archivos_guardados else None
        except Exception:
            af_selected = None
    with c1:
        st.subheader("Planilla de Horas (Carga)")
        st.success("☁️ Sincronizado con Google Sheets")
        try:
            carga_bytes = obtener_excel_nube(URL_DESCARGA_SHEET)
            xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
            hojas_disponibles = xl_carga.sheet_names
            nh = st.selectbox("Seleccionar Fecha / Pestaña:", options=hojas_disponibles)
        except Exception as e:
            carga_bytes, nh = None, None

with t_reglas:
    st.subheader("⚙️ Configuración de Personal Excluido / No Facturable")
    df_reglas = pd.DataFrame([{"Legajo": leg, "Nombre": datos[0], "Sector / Área": datos[1]} for leg, datos in EXCLUIDOS_SECTOR.items()])
    st.dataframe(df_reglas, use_container_width=True)

with t_reporte:
    st.info("Utilice las solapas específicas de Cierre, Horas o Certificación para análisis avanzados.")

with t_cierre:
    st.info("Gestión de OTs y Cierre de Período.")

with t_quincena:
    st.info("Recuento y Selección de Horas.")

with t_detalle:
    st.info("Control Detallado.")

# --- PESTAÑA: 📁 6. CERTIFICACIÓN DE HORAS (CRUCE ROBUSTO POR DICCIONARIO) ---
with t_cert:
    st.header("📁 Certificación y Cruce Previo a Facturación (MOA vs Drive)")
    st.info("ℹ️ Este módulo compara las horas de Google Sheets (MOA) con el archivo de certificación de Drive (Col Q: Horas, Col S: Certificación 'X') mediante un cruce por Legajo y Nombre, restando automáticamente las certificadas y no certificadas para dar con el limpio de **Pendientes de Carga**.")

    if carga_bytes:
        col_cert1, col_cert2, col_cert3 = st.columns([1.2, 1.2, 1.2])
        with col_cert1:
            contrato_sel_cert = st.selectbox("🏗️ Seleccionar Contrato:", options=["MOA", "MMYHH", "Todos los Contratos"], index=0, key="sel_contrato_cert")
        with col_cert2:
            dia_desde_cert = st.selectbox("📅 Día Desde:", options=list(range(1, 32)), index=0, key="sel_dia_desde_cert")
        with col_cert3:
            dia_hasta_cert = st.selectbox("📅 Día Hasta:", options=list(range(1, 32)), index=30, key="sel_dia_hasta_cert")

        archivos_cert_drive = obtener_archivos_cert_drive()
        mapa_cert_files = {f['name']: f for f in archivos_cert_drive}
        opciones_cert = ["-- Seleccionar archivo de certificación --"] + list(mapa_cert_files.keys())
        cert_sel_name = st.selectbox("📂 Archivo de Certificación (Drive):", options=opciones_cert, key="sel_cert_drive_file")

        cert_file_id = mapa_cert_files.get(cert_sel_name, {}).get('id') if cert_sel_name != "-- Seleccionar archivo de certificación --" else None
        cert_mime_type = mapa_cert_files.get(cert_sel_name, {}).get('mimeType') if cert_sel_name != "-- Seleccionar archivo de certificación --" else None

        if cert_file_id is not None:
            if st.button("🚀 Ejecutar Cruce y Auditoría MOA", type="primary"):
                with st.spinner("Procesando planillas de Google Sheets (MOA) y cruzando con el archivo de certificación..."):
                    try:
                        # 1. Leer archivo de certificación de Drive
                        fh_cert = descargar_archivo_notif_drive(cert_file_id, cert_mime_type)
                        df_cert_raw = pd.read_excel(fh_cert)
                        
                        cols = df_cert_raw.columns
                        col_horas_idx = 16 if len(cols) > 16 else None  # Columna Q
                        col_cert_idx = 18 if len(cols) > 18 else None   # Columna S
                        col_fecha_cert = next((c for c in cols if any(k in str(c).upper() for k in ["FECHA", "DIA", "DÍA"])), cols[3] if len(cols) > 3 else None)
                        col_legajo_cert = next((c for c in cols if "LEGAJO" in str(c).upper()), cols[1] if len(cols) > 1 else None)
                        col_nombre_cert = next((c for c in cols if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE", "PERSONAL"])), cols[2] if len(cols) > 2 else None)

                        # 2. Cargar planillas diarias de Google Sheets para el rango de días seleccionado
                        xl_sheets = pd.ExcelFile(io.BytesIO(carga_bytes))
                        dias_map_cert = {int(m.group(1)): s for s in xl_sheets.sheet_names if (m := re.search(r"(\d{1,2})[-/\.](\d{1,2})", s))}
                        
                        registros_sheets = []
                        for d_num, s_name in dias_map_cert.items():
                            if dia_desde_cert <= d_num <= dia_hasta_cert:
                                df_s_raw = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=s_name, header=None)
                                h_idx = next((idx for idx, row in df_s_raw.iterrows() if any(k in " ".join([str(v).upper() for v in row.values if pd.notna(v)]) for k in ["LEGAJO", "OBRA", "APELLIDO"])), 0)
                                df_s = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=s_name, skiprows=h_idx)
                                
                                l_col = next((c for c in df_s.columns if "LEGAJO" in str(c).upper()), None)
                                n_col = next((c for c in df_s.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE"])), None)
                                t_col = next((c for c in df_s.columns if "TRABAJO" in str(c).upper()), None)

                                if l_col and t_col:
                                    for _, r in df_s.iterrows():
                                        l_val = limpiar_legajo(r[l_col])
                                        n_val = str(r[n_col]).strip() if n_col and pd.notna(r[n_col]) else ""
                                        
                                        if not filtrar_por_contrato(l_val, n_val, contrato_sel_cert): 
                                            continue
                                            
                                        hs_hojas = extraer_horas(r[t_col])
                                        if hs_hojas > 0:
                                            registros_sheets.append({
                                                "Legajo": l_val if l_val is not None else 0,
                                                "Nombre_Norm": normalizar_texto(n_val),
                                                "Nombre_Original": n_val,
                                                "HorasSheets": hs_hojas
                                            })

                        df_sheets_proc = pd.DataFrame(registros_sheets)
                        df_sheets_agrupado = pd.DataFrame(columns=["Legajo", "Nombre_Norm", "Nombre", "Total Horas Sheets"])
                        if not df_sheets_proc.empty:
                            df_sheets_agrupado = df_sheets_proc.groupby(["Legajo", "Nombre_Norm"], as_index=False).agg({
                                "Nombre_Original": "first",
                                "HorasSheets": "sum"
                            }).rename(columns={"Nombre_Original": "Nombre", "HorasSheets": "Total Horas Sheets"})

                        # 3. Procesar archivo de Drive con filtro de contrato y fecha
                        resultados_cert = []
                        for _, row in df_cert_raw.iterrows():
                            if col_fecha_cert and col_fecha_cert in df_cert_raw.columns and pd.notna(row[col_fecha_cert]):
                                d_val, _ = obtener_dia_mes(row[col_fecha_cert])
                                if d_val is not None and not (dia_desde_cert <= d_val <= dia_hasta_cert):
                                    continue

                            l_val = limpiar_legajo(row[col_legajo_cert]) if col_legajo_cert in df_cert_raw.columns else None
                            n_val = str(row[col_nombre_cert]).strip() if col_nombre_cert and col_nombre_cert in df_cert_raw.columns and pd.notna(row[col_nombre_cert]) else ""

                            if not filtrar_por_contrato(l_val, n_val, contrato_sel_cert):
                                continue

                            val_horas_raw = row.iloc[col_horas_idx] if col_horas_idx is not None and col_horas_idx < len(row) else 0
                            hs = extraer_horas(val_horas_raw)

                            val_cert_raw = str(row.iloc[col_cert_idx]).strip().upper() if col_cert_idx is not None and col_cert_idx < len(row) and pd.notna(row.iloc[col_cert_idx]) else ""
                            es_certificado = True if "X" in val_cert_raw else False

                            resultados_cert.append({
                                "Legajo": l_val if l_val is not None else 0,
                                "Nombre_Norm": normalizar_texto(n_val),
                                "Nombre": n_val,
                                "Horas Certificadas": hs if es_certificado else 0.0,
                                "Horas No Certificadas": hs if not es_certificado else 0.0,
                                "Total Horas Drive": hs
                            })

                        df_cert_proc = pd.DataFrame(resultados_cert)
                        df_cert_resumen = pd.DataFrame(columns=["Legajo", "Nombre_Norm", "Horas Certificadas", "Horas No Certificadas", "Total Horas Drive"])
                        if not df_cert_proc.empty:
                            df_cert_resumen = df_cert_proc.groupby(["Legajo", "Nombre_Norm"], as_index=False)[
                                ["Horas Certificadas", "Horas No Certificadas", "Total Horas Drive"]
                            ].sum()

                        # 4. Cruce robusto mediante diccionarios de búsqueda (Legajo y Nombre)
                        cert_dict_legajo = {}
                        cert_dict_nombre = {}
                        for _, r_c in df_cert_resumen.iterrows():
                            leg_c = int(r_c["Legajo"]) if pd.notna(r_c["Legajo"]) else 0
                            nom_c_norm = str(r_c["Nombre_Norm"]).strip()
                            h_cert = r_c["Horas Certificadas"]
                            h_nocert = r_c["Horas No Certificadas"]
                            if leg_c > 0:
                                cert_dict_legajo[leg_c] = {"Certificadas": h_cert, "NoCertificadas": h_nocert}
                            if nom_c_norm:
                                cert_dict_nombre[nom_c_norm] = {"Certificadas": h_cert, "NoCertificadas": h_nocert}

                        filas_cruce = []
                        for _, r_s in df_sheets_agrupado.iterrows():
                            leg_s = int(r_s["Legajo"]) if pd.notna(r_s["Legajo"]) else 0
                            nom_s_norm = str(r_s["Nombre_Norm"]).strip()
                            nom_s_orig = r_s["Nombre"]
                            tot_sheets = r_s["Total Horas Sheets"]

                            h_cert = 0.0
                            h_nocert = 0.0

                            if leg_s > 0 and leg_s in cert_dict_legajo:
                                h_cert = cert_dict_legajo[leg_s]["Certificadas"]
                                h_nocert = cert_dict_legajo[leg_s]["NoCertificadas"]
                            elif nom_s_norm in cert_dict_nombre:
                                h_cert = cert_dict_nombre[nom_s_norm]["Certificadas"]
                                h_nocert = cert_dict_nombre[nom_s_norm]["NoCertificadas"]

                            pend_carga = max(0.0, tot_sheets - (h_cert + h_nocert))

                            filas_cruce.append({
                                "Legajo": leg_s if leg_s > 0 else "SIN LEGAJO",
                                "Nombre": nom_s_orig,
                                "Total Horas Sheets": tot_sheets,
                                "Horas Certificadas": h_cert,
                                "Horas No Certificadas": h_nocert,
                                "⚠️ Pendiente de Carga / Fuera de Drive": pend_carga
                            })

                        df_final_cruce = pd.DataFrame(filas_cruce)

                        if not df_final_cruce.empty:
                            tot_cert = df_final_cruce["Horas Certificadas"].sum()
                            tot_no_cert = df_final_cruce["Horas No Certificadas"].sum()
                            tot_pendiente_carga = df_final_cruce["⚠️ Pendiente de Carga / Fuera de Drive"].sum()
                            tot_gral_sheets = df_final_cruce["Total Horas Sheets"].sum()

                            k_c1, k_c2, k_c3, k_c4 = st.columns(4)
                            k_c1.metric("🟢 Certificadas", f"{tot_cert:.1f} hs")
                            k_c2.metric("🟡 No Certificadas", f"{tot_no_cert:.1f} hs")
                            k_c3.metric("⚠️ Pendiente de Carga", f"{tot_pendiente_carga:.1f} hs")
                            k_c4.metric("📊 Total Google Sheets", f"{tot_gral_sheets:.1f} hs")

                            st.divider()
                            st.subheader(f"📋 Auditoría de Certificación y Carga ({contrato_sel_cert} - Días {dia_desde_cert:02d} al {dia_hasta_cert:02d})")
                            
                            columnas_vista = ["Legajo", "Nombre", "Total Horas Sheets", "Horas Certificadas", "Horas No Certificadas", "⚠️ Pendiente de Carga / Fuera de Drive"]
                            st.dataframe(df_final_cruce[columnas_vista], use_container_width=True)

                            out_cert_excel = io.BytesIO()
                            with pd.ExcelWriter(out_cert_excel, engine="openpyxl") as writer:
                                df_final_cruce[columnas_vista].to_excel(writer, index=False, sheet_name=f"Certificacion_{contrato_sel_cert}")
                            out_cert_excel.seek(0)

                            st.download_button(
                                label=f"📥 Descargar Auditoría MOA ({contrato_sel_cert} - Días {dia_desde_cert:02d} al {dia_hasta_cert:02d})",
                                data=out_cert_excel,
                                file_name=f"Auditoria_Certificacion_{contrato_sel_cert}_Dias_{dia_desde_cert:02d}_al_{dia_hasta_cert:02d}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("No se encontraron registros para los filtros seleccionados.")
                    except Exception as e_c:
                        st.error(f"Error al procesar el cruce MOA: {e_c}")
        else:
            st.info("👆 Seleccioná un archivo de certificación en Google Drive para procesar el cruce de MOA.")
    else:
        st.warning("Conectando con Google Sheets...")
