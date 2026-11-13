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

@st.cache_data(show_spinner=False, ttl=43200)
def procesar_fichada_drive_cache(file_id):
    fichadas_horas = {}
    try:
        fh_loc = descargar_archivo_drive(file_id)
        dff_loc = cargar_dataframe_fichada(fh_loc)
        start_row_q, c_leg_q, c_nom_q, c_ing_q, c_egr_q, c_nov_q = detectar_estructura_fichada(dff_loc)
        for idx_fl in range(start_row_q, len(dff_loc)):
            lv_q = dff_loc.iloc[idx_fl, c_leg_q] if c_leg_q < len(dff_loc.columns) else None
            if str(lv_q).startswith('#'): lv_q = None
            l_f_clean = limpiar_legajo(lv_q)
            if l_f_clean is not None:
                ing_q = dff_loc.iloc[idx_fl, c_ing_q] if c_ing_q < len(dff_loc.columns) else ""
                sal_q = dff_loc.iloc[idx_fl, c_egr_q] if c_egr_q < len(dff_loc.columns) else ""
                if str(ing_q).startswith('#'): ing_q = ""
                if str(sal_q).startswith('#'): sal_q = ""
                horas_calculadas = calcular_horas_fichada(ing_q, sal_q)
                fichadas_horas[l_f_clean] = fichadas_horas.get(l_f_clean, 0.0) + horas_calculadas
    except Exception:
        pass
    return fichadas_horas

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

def extraer_fecha_contenido(df_raw):
    for idx in range(min(10, len(df_raw))):
        for val in df_raw.iloc[idx].dropna():
            if isinstance(val, (pd.Timestamp, datetime.date, datetime.datetime)):
                return val.day, val.month
            match = re.search(r"(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{2,4})", str(val))
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

def calcular_horas_fichada(ing, sal):
    i = convertir_a_horas_decimal(ing)
    s = convertir_a_horas_decimal(sal)
    if i is None or s is None:
        return 0.0
    return s - i if s >= i else (24.0 - i) + s

def formatear_horario(val):
    dec = convertir_a_horas_decimal(val)
    if dec is None:
        return "--:--"
    horas = int(dec)
    minutos = int(round((dec - horas) * 60))
    if minutos >= 60:
        horas += 1
        minutos -= 60
    return f"{horas:02d}:{minutos:02d}"

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
    st.success("La aplicación detectará automáticamente los nuevos archivos cuando los subas a Drive.")

with t_carga:
    c1, c2 = st.columns(2)
    with c2:
        st.subheader("Base de Fichadas Guardadas (Drive)")
        try:
            archivos_drive = obtener_archivos_drive()
            archivos_guardados = [f['name'] for f in archivos_drive if f['name'].lower().endswith(('.xlsx', '.xltx'))]
            mapa_archivos = {f['name']: f['id'] for f in archivos_drive}
            
            if not archivos_guardados:
                st.warning("Aún no hay fichadas en Drive.")
                af_selected = None
            else:
                fichadas_por_mes = {}
                for f in archivos_guardados:
                    match = re.search(r"fichada[_\s]*(\d{1,2})[-/\.](\d{1,2})", f, re.IGNORECASE)
                    if match:
                        dia, mes = int(match.group(1)), int(match.group(2))
                        if mes not in fichadas_por_mes: fichadas_por_mes[mes] = {}
                        fichadas_por_mes[mes][dia] = f

                if not fichadas_por_mes:
                    opciones = ["-- Seleccione una fichada --"] + archivos_guardados
                    af_selected = st.selectbox("Elegir Fichada:", options=opciones)
                    if af_selected == "-- Seleccione una fichada --": af_selected = None
                else:
                    col_mes, col_dia = st.columns(2)
                    with col_mes:
                        meses_ordenados = sorted(fichadas_por_mes.keys(), reverse=True)
                        mes_elegido_num = st.selectbox("Mes de Fichada:", options=meses_ordenados, format_func=lambda m: f"📅 {MESES_NOMBRES.get(m, 'Mes')}")
                    with col_dia:
                        dias_del_mes = sorted(fichadas_por_mes[mes_elegido_num].keys(), reverse=True)
                        opciones_dia = ["-- Seleccionar Día --"] + [f"Día {d:02d}" for d in dias_del_mes]
                        dia_elegido_str = st.selectbox("Día de Fichada:", options=opciones_dia)

                    if dia_elegido_str != "-- Seleccionar Día --":
                        dia_elegido_num = int(dia_elegido_str.replace("Día ", ""))
                        af_selected = fichadas_por_mes[mes_elegido_num][dia_elegido_num]
                        st.info(f"☁️ Archivo vinculado desde Drive: **{af_selected}**")
                    else:
                        af_selected = None
        except Exception as e:
            st.error("Error al leer archivos de Drive.")
            af_selected = None

    with c1:
        st.subheader("Planilla de Horas (Carga)")
        st.success("☁️ Sincronizado con Google Sheets")
        try:
            with st.spinner("Sincronizando con Google Sheets..."):
                carga_bytes = obtener_excel_nube(URL_DESCARGA_SHEET)
                xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
                hojas_disponibles = xl_carga.sheet_names
            nh = st.selectbox("Seleccionar Fecha / Pestaña:", options=hojas_disponibles)
        except Exception as e:
            st.error(f"No se pudo conectar a Google Sheets. Error: {e}")
            carga_bytes, nh = None, None

with t_reglas:
    st.subheader("⚙️ Configuración de Personal Excluido / No Facturable")
    df_reglas = pd.DataFrame([{"Legajo": leg, "Nombre": datos[0], "Sector / Área": datos[1]} for leg, datos in EXCLUIDOS_SECTOR.items()])
    st.dataframe(df_reglas, use_container_width=True)

with t_reporte:
    if carga_bytes and af_selected:
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            if st.button("🚀 Ejecutar Cruce Estricto por Fecha", type="primary"):
                st.session_state["mostrar_reporte"] = True
                st.rerun()
        with col_btn2:
            if st.button("🔄 Sincronizar Todo (Drive y Sheets)"):
                st.session_state["mostrar_reporte"] = True
                st.rerun()

        if st.session_state["mostrar_reporte"]:
            try:
                af_selected_id = mapa_archivos.get(af_selected)
                with st.spinner("Descargando fichadas desde Drive (Múltiples hojas)..."):
                    fh_fichada = descargar_archivo_drive(af_selected_id)

                d_pestana, m_pestana = obtener_dia_mes(nh)
                dff_raw = cargar_dataframe_fichada(fh_fichada)
                fh_fichada.seek(0) 
                
                d_fichada, m_fichada = obtener_dia_mes(af_selected)
                if d_fichada is None: d_fichada, m_fichada = extraer_fecha_contenido(dff_raw)

                if d_pestana and d_fichada and (d_pestana, m_pestana) != (d_fichada, m_fichada):
                    st.error(f"⛔ **Error de Validación:** La Pestaña `{nh}` es del **{d_pestana:02d}/{m_pestana:02d}**, pero la Fichada es del **{d_fichada:02d}/{m_fichada:02d}**.")
                    st.stop()

                dfc_raw = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, header=None)
                header_row_idx = 0
                for idx, row in dfc_raw.iterrows():
                    row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
                    if "LEGAJO" in row_str or "OBRA" in row_str or "APELLIDO" in row_str:
                        header_row_idx = idx
                        break

                dfc = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, skiprows=header_row_idx)
                legajo_col, nombre_col, obra_col = None, None, None
                for col in dfc.columns:
                    col_str = str(col).strip().upper()
                    if "LEGAJO" in col_str: legajo_col = col
                    elif "APELLIDO" in col_str or "NOMBRE" in col_str: nombre_col = col
                    elif "OBRA" in col_str: obra_col = col

                if legajo_col is None and len(dfc.columns) > 2: legajo_col = dfc.columns[2]
                if obra_col is None: obra_col = dfc.columns[12] if len(dfc.columns) > 12 else dfc.columns[-1]

                dff = dff_raw.copy()
                start_row_f, c_leg, c_nom, c_ing, c_egr, c_nov = detectar_estructura_fichada(dff)

                keywords_novedades = {
                    "COMPENSATORIO": "Compensatorio", "COMPENSATORIA": "Compensatorio",
                    "ENFERMO": "Enfermo / Lic. Médica", "ENFERMEDAD": "Enfermo / Lic. Médica",
                    "ENFEMRO": "Enfermo / Lic. Médica", "ENFEMEDAD": "Enfermo / Lic. Médica",
                    "MEDICO": "Enfermo / Lic. Médica", "MÉDICO": "Enfermo / Lic. Médica",
                    "SUSPENDIDO": "Suspendido", "SUSPENSION": "Suspendido", "SUSPENSIÓN": "Suspendido",
                    "AUSENTE": "Ausente / Justificado", "AUSENCIA": "Ausente / Justificado",
                    "VACACIONES": "Vacaciones", "VACACION": "Vacaciones",
                    "SUR": "SUR / Franco Compensatorio", "LICENCIA": "Licencia / Permiso",
                }

                all_fichadas_list, fd_by_legajo, fd_by_name = [], {}, {}
                for i in range(start_row_f, len(dff)):
                    lv = dff.iloc[i, c_leg] if c_leg < len(dff.columns) else None
                    nv = dff.iloc[i, c_nom] if c_nom < len(dff.columns) else None
                    if str(lv).startswith('#'): lv = None
                    if str(nv).startswith('#'): nv = None

                    nom_str = str(nv).strip() if pd.notna(nv) else ""
                    for kw in keywords_novedades.keys():
                        nom_str = re.sub(rf"\s*-\s*{kw}.*", "", nom_str, flags=re.IGNORECASE).strip()

                    col_m_val = str(dff.iloc[i, c_nov]).upper() if c_nov < len(dff.columns) and pd.notna(dff.iloc[i, c_nov]) else ""
                    row_text_sin_nombre = " ".join([str(val).upper() for idx_v, val in enumerate(dff.iloc[i].values) if idx_v != c_nom and pd.notna(val)])
                    
                    novedad_detectada = None
                    for kw, label in keywords_novedades.items():
                        if kw in col_m_val or kw in row_text_sin_nombre:
                            novedad_detectada = label
                            break

                    clean_lv = limpiar_legajo(lv)
                    val_nombre_norm = normalizar_texto(nom_str)

                    info_dict = {
                        "legajo": clean_lv, "nombre": nom_str, "nombre_norm": val_nombre_norm,
                        "estado": novedad_detectada if novedad_detectada else "Presente", "es_novedad": novedad_detectada is not None,
                    }

                    if val_nombre_norm or clean_lv is not None: all_fichadas_list.append(info_dict)
                    if clean_lv is not None: fd_by_legajo[clean_lv] = info_dict
                    if val_nombre_norm: fd_by_name[val_nombre_norm] = info_dict

                def buscar_fichada(val_legajo, val_nombre):
                    clean_l = limpiar_legajo(val_legajo)
                    if clean_l is not None and clean_l in fd_by_legajo: return fd_by_legajo[clean_l], clean_l
                    val_nombre_norm = normalizar_texto(val_nombre)
                    if val_nombre_norm:
                        if val_nombre_norm in fd_by_name: return (fd_by_name[val_nombre_norm], fd_by_name[val_nombre_norm]["legajo"])
                        tokens_c = set(val_nombre_norm.split())
                        if len(tokens_c) >= 2:
                            for n_key, item in fd_by_name.items():
                                if len(tokens_c.intersection(set(n_key.split()))) >= 2:
                                    return item, item["legajo"]
                    return None, clean_l

                fichadas_emparejadas, cargados_map = set(), {}
                for _, row_c in dfc.iterrows():
                    l_val = row_c[legajo_col] if legajo_col in dfc.columns else None
                    n_val = row_c[nombre_col] if nombre_col and nombre_col in dfc.columns else None
                    info_match, clean_l = buscar_fichada(l_val, n_val)
                    o_val = str(row_c[obra_col]).strip() if obra_col and pd.notna(row_c[obra_col]) else "Sin Obra"

                    if clean_l is not None: cargados_map[clean_l] = o_val
                    if info_match:
                        if info_match["legajo"] is not None:
                            cargados_map[info_match["legajo"]] = o_val
                            fichadas_emparejadas.add(info_match["legajo"])
                        if info_match["nombre_norm"]: cargados_map[info_match["nombre_norm"]] = o_val

                no_cargados, excluidos_list, novedades_list = [], [], []
                fichadas_procesadas = set()

                for info_f in all_fichadas_list:
                    legajo_f, nombre_norm_f = info_f["legajo"], info_f["nombre_norm"]
                    f_id = legajo_f if legajo_f is not None else nombre_norm_f
                    if f_id in fichadas_procesadas or not f_id: continue
                    fichadas_procesadas.add(f_id)

                    nombre_fichada, estado_fichada, es_nov = info_f["nombre"], info_f["estado"], info_f["es_novedad"]
                    es_cargado = ((legajo_f is not None and legajo_f in cargados_map) or (legajo_f is not None and legajo_f in fichadas_emparejadas) or (nombre_norm_f and nombre_norm_f in cargados_map))

                    if legajo_f is not None and legajo_f in EXCLUIDOS_SECTOR:
                        excluidos_list.append({"LEGAJO": legajo_f, "Nombre y Apellido": (nombre_fichada or EXCLUIDOS_SECTOR[legajo_f][0]), "Sector": EXCLUIDOS_SECTOR[legajo_f][1], "Estado Fichada": estado_fichada, "Observación": "Sector no facturable"})
                    elif not es_cargado:
                        if es_nov: novedades_list.append({"LEGAJO": legajo_f if legajo_f else "SIN LEGAJO", "Nombre y Apellido": nombre_fichada, "Novedad / Estado": estado_fichada, "Observación": "Excluido por Novedad"})
                        else: no_cargados.append({"LEGAJO": legajo_f if legajo_f else "SIN LEGAJO", "Nombre y Apellido (Fichada)": nombre_fichada, "OBRA (Columna M)": "NO CARGADO EN PLANILLA", "Estado": "Fichó en planta pero NO fue cargado"})
                    else:
                        if es_nov: novedades_list.append({"LEGAJO": legajo_f if legajo_f else "SIN LEGAJO", "Nombre y Apellido": nombre_fichada, "Novedad / Estado": estado_fichada, "Observación": "Cargado registrando novedad"})

                df_no_cargados, df_excluidos, df_novedades = pd.DataFrame(no_cargados), pd.DataFrame(excluidos_list), pd.DataFrame(novedades_list)

                v1, v2, v3, v4, v5 = st.tabs(["📊 Vista General", "⚠️ Faltan Cargar (Reales)", "🚫 Sectores No Facturables", "📋 Novedades / Licencias", "🟡 Ausentes / No Ficharon"])
                with v2:
                    st.subheader("Personal que Fichó en Planta pero NO fue Cargado")
                    if not df_no_cargados.empty:
                        st.warning(f"⚠️ Se detectaron **{len(df_no_cargados)} personas** que ingresaron a trabajar y **NO figuran cargadas**.")
                        df_edit = df_no_cargados.copy()
                        df_edit["Asignar Novedad"] = "Seleccionar..."
                        edited_df = st.data_editor(df_edit, column_config={"Asignar Novedad": st.column_config.SelectboxColumn("📝 Asignar Novedad Manual", width="medium", options=["Seleccionar...", "COMPENSATORIO", "ENFERMO", "VACACIONES", "AUSENTE", "SUSPENDIDO", "SUR", "LICENCIA"])}, disabled=["LEGAJO", "Nombre y Apellido (Fichada)", "OBRA (Columna M)", "Estado"], use_container_width=True, key="editor_novedades_faltantes")

                        if st.button("☁️ Guardar Novedades Seleccionadas en Drive", type="primary"):
                            cambios = edited_df[edited_df["Asignar Novedad"].notna() & (edited_df["Asignar Novedad"] != "Seleccionar...") & (edited_df["Asignar Novedad"].astype(str).str.strip() != "")]
                            if not cambios.empty:
                                try:
                                    with st.spinner("Actualizando Excel en Google Drive..."):
                                        fh_fichada.seek(0)
                                        wb_fichada = load_workbook(fh_fichada)
                                        hojas_fichada = wb_fichada.sheetnames
                                        col_leg_ws = c_leg + 1
                                        col_nom_ws = c_nom + 1
                                        col_nov_ws = c_nov + 1

                                        for index, row_edit in cambios.iterrows():
                                            leg_buscar = limpiar_legajo(row_edit["LEGAJO"])
                                            nom_buscar = str(row_edit["Nombre y Apellido (Fichada)"]).strip()
                                            nov_asignar = str(row_edit["Asignar Novedad"]).strip()
                                            actualizado = False
                                            for h_name in hojas_fichada:
                                                if actualizado: break
                                                ws_temp = wb_fichada[h_name]
                                                for r_idx in range(1, ws_temp.max_row + 1):
                                                    cell_val = ws_temp.cell(row=r_idx, column=col_leg_ws).value
                                                    cell_name = str(ws_temp.cell(row=r_idx, column=col_nom_ws).value or "").strip()
                                                    cell_leg = limpiar_legajo(cell_val)
                                                    if not cell_name and cell_leg is None: continue
                                                    match_found = (cell_leg == leg_buscar) if (cell_leg is not None and leg_buscar is not None) else (nom_buscar and nom_buscar.lower() in cell_name.lower())
                                                    if match_found:
                                                        ws_temp.cell(row=r_idx, column=col_nov_ws).value = nov_asignar
                                                        actualizado = True
                                                        break

                                        out_bytes = io.BytesIO()
                                        wb_fichada.save(out_bytes)
                                        out_bytes.seek(0)
                                        media_update = MediaIoBaseUpload(out_bytes, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
                                        drive_service.files().update(fileId=af_selected_id, media_body=media_update).execute()

                                    st.success("¡Archivo actualizado en Google Drive con éxito!")
                                    time.sleep(2) 
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error al escribir en Drive: {e}")
                            else: st.info("Elegí alguna novedad del menú antes de Guardar.")
                    else: st.success("🎉 ¡Todo el personal activo facturable está cargado!")

                with v1:
                    st.subheader("Control General de Planilla de Carga vs Fichadas")
                    def color_rows(row):
                        try:
                            l_val = row[legajo_col] if legajo_col in dfc.columns else None
                            n_val = row[nombre_col] if nombre_col and nombre_col in dfc.columns else None
                            info_match, clean_l = buscar_fichada(l_val, n_val)
                            if (clean_l is not None and clean_l in EXCLUIDOS_SECTOR) or (info_match and info_match["legajo"] in EXCLUIDOS_SECTOR): return ["background-color: #E2EFDA"] * len(row)
                            if info_match: return ["background-color: #D9E1F2"] * len(row) if info_match["es_novedad"] else ["background-color: #C6EFCE"] * len(row)
                            return ["background-color: #FFEB9C"] * len(row)
                        except Exception: return ["background-color: #FFC7CE"] * len(row)
                    st.dataframe(dfc.style.apply(color_rows, axis=1), use_container_width=True)

                with v3:
                    st.subheader("Personal de Sectores No Facturables")
                    st.dataframe(df_excluidos, use_container_width=True) if not df_excluidos.empty else st.write("Sin registros.")
                with v4:
                    st.subheader("Novedades y Licencias")
                    st.dataframe(df_novedades, use_container_width=True) if not df_novedades.empty else st.write("Sin registros.")
                with v5:
                    st.subheader("Personal Cargado que NO Registró Fichada")
                    ausentes_list = [row_c for _, row_c in dfc.iterrows() if not buscar_fichada(row_c.get(legajo_col), row_c.get(nombre_col))[0]]
                    df_ausentes = pd.DataFrame(ausentes_list) if ausentes_list else pd.DataFrame()
                    if not df_ausentes.empty:
                        st.info(f"**{len(df_ausentes)} registros** cargados sin fichada:")
                        st.dataframe(df_ausentes, use_container_width=True)
                    else: st.success("🎉 Todo el personal cargado registró su fichada.")

                wb = Workbook()
                ws1 = wb.active
                ws1.title = f"Control {nh}"
                gf, yf, rf, bf, exf = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFEB9C"), PatternFill("solid", fgColor="FFC7CE"), PatternFill("solid", fgColor="D9E1F2"), PatternFill("solid", fgColor="E2EFDA")

                for r_idx, row in enumerate(dataframe_to_rows(dfc, index=False, header=True), 1):
                    ws1.append(row)
                    if r_idx == 1: continue
                    try:
                        info_match, clean_l = buscar_fichada(row[list(dfc.columns).index(legajo_col)] if legajo_col in dfc.columns else None, row[list(dfc.columns).index(nombre_col)] if nombre_col and nombre_col in dfc.columns else None)
                        fc = exf if (clean_l is not None and clean_l in EXCLUIDOS_SECTOR) or (info_match and info_match["legajo"] in EXCLUIDOS_SECTOR) else bf if info_match and info_match["es_novedad"] else gf if info_match else yf
                        for c in range(1, len(row) + 1): ws1.cell(row=r_idx, column=c).fill = fc
                    except Exception:
                        for c in range(1, len(row) + 1): ws1.cell(row=r_idx, column=c).fill = rf

                if not df_no_cargados.empty:
                    ws2 = wb.create_sheet("Fichados Sin Carga")
                    for r_idx, row in enumerate(dataframe_to_rows(df_no_cargados, index=False, header=True), 1):
                        ws2.append(row)
                        if r_idx > 1:
                            for c in range(1, len(row) + 1): ws2.cell(row=r_idx, column=c).fill = PatternFill("solid", fgColor="FCE4D6")

                if not df_excluidos.empty:
                    ws3 = wb.create_sheet("Sectores No Facturables")
                    for r_idx, row in enumerate(dataframe_to_rows(df_excluidos, index=False, header=True), 1):
                        ws3.append(row)
                        if r_idx > 1:
                            for c in range(1, len(row) + 1): ws3.cell(row=r_idx, column=c).fill = exf

                out = io.BytesIO()
                wb.save(out)
                out.seek(0)
                st.divider()
                st.success("¡Cruce completado con éxito!")
                st.download_button("📥 Descargar Excel Formateado", data=out, file_name=f"Control_{nh}.xlsx")

            except Exception as e: st.error(f"Error al procesar: {e}")
    else: st.info("Seleccione una fecha de Google Sheets y un archivo de fichada en Drive para cruzar.")

# --- PESTAÑA: 📋 CIERRE DE PERÍODO (GESTIÓN DE OTs) ---
with t_cierre:
    st.header("📋 Notificación y Cierre de Período (Gestión de OTs)")
    if carga_bytes:
        xl_carga_cierre = pd.ExcelFile(io.BytesIO(carga_bytes))
        dias_disponibles_map = {int(m.group(1)): s for s in xl_carga_cierre.sheet_names if (m := re.search(r"(\d{1,2})[-/\.](\d{1,2})", s))}
        dias_ordenados = sorted(dias_disponibles_map.keys())

        if dias_ordenados:
            col_f1, col_f2, col_f3 = st.columns([1, 1, 1.2])
            with col_f1: 
                desde_dia_sel = st.selectbox("📅 Fecha Desde:", options=dias_ordenados, format_func=lambda d: f"Día {d:02d} ({dias_disponibles_map[d]})", index=0)
            with col_f2:
                dias_hasta_opciones = [d for d in dias_ordenados if d >= desde_dia_sel]
                hasta_dia_sel = st.selectbox("📅 Fecha Hasta:", options=dias_hasta_opciones, format_func=lambda d: f"Día {d:02d} ({dias_disponibles_map[d]})", index=len(dias_hasta_opciones)-1 if dias_hasta_opciones else 0)

            with col_f3:
                archivos_notif = obtener_archivos_notif_drive()
                mapa_notif_files = {f['name']: f for f in archivos_notif}
                opciones_notif = ["-- Sin planilla de notificaciones --"] + list(mapa_notif_files.keys())
                notif_sel_name = st.selectbox("📑 Planilla Notificaciones (Drive):", options=opciones_notif)
                
                if notif_sel_name != "-- Sin planilla de notificaciones --":
                    selected_file_info = mapa_notif_files.get(notif_sel_name, {})
                    notif_file_id = selected_file_info.get('id')
                    notif_mime_type = selected_file_info.get('mimeType')
                else:
                    notif_file_id = None
                    notif_mime_type = None

            st.write("")
            contrato_sel_cierre = st.radio(
                "🏗️ Filtrar Cierre por Contrato:",
                options=["Todos los Contratos", "MMYHH", "MOA"],
                horizontal=True,
                key="radio_contrato_cierre"
            )
            
            col_btn_ej, _ = st.columns([1, 3])
            with col_btn_ej:
                btn_ejecutar_cierre = st.button("🚀 Ejecutar Cruce de OTs", type="primary")

            if btn_ejecutar_cierre:
                st.session_state["cruce_cierre_ejecutado"] = True

            if st.session_state["cruce_cierre_ejecutado"]:
                with st.spinner("Procesando OTs y cruzando notificaciones desde Drive..."):
                    mapa_notif_drive, diag_info = obtener_notificaciones_drive(notif_file_id, notif_mime_type) if notif_file_id else ({}, {})
                    ots_dict = {}
                    for d in dias_ordenados:
                        if d < desde_dia_sel or d > hasta_dia_sel: continue
                        s_name = dias_disponibles_map[d]
                        dfc_raw = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=s_name, header=None)
                        header_row_idx = next((idx for idx, row in dfc_raw.iterrows() if any(k in " ".join([str(v).upper() for v in row.values if pd.notna(v)]) for k in ["LEGAJO", "OBRA", "APELLIDO", "OT"])), 0)
                        dfc_ot = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=s_name, skiprows=header_row_idx)

                        ot_col = next((c for c in dfc_ot.columns if "OT" in str(c).upper()), None)
                        sol_col = next((c for c in dfc_ot.columns if "SOLICITUD" in str(c).upper()), None)
                        fecha_col = next((c for c in dfc_ot.columns if "FECHA" in str(c).upper()), None)
                        notif_col = next((c for c in dfc_ot.columns if "NOTIFICA" in str(c).upper()), None)
                        sector_col = next((c for c in dfc_ot.columns if any(k in str(c).upper() for k in ["SECTOR", "OBRA", "LUGAR", "AREA", "ÁREA"])), None)
                        resp_col = next((c for c in dfc_ot.columns if "RESPONSABLE" in str(c).upper()), None)
                        legajo_col_cierre = next((c for c in dfc_ot.columns if "LEGAJO" in str(c).upper()), None)
                        nombre_col_cierre = next((c for c in dfc_ot.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE", "PERSONAL"])), None)

                        if ot_col:
                            for _, row_r in dfc_ot.iterrows():
                                ot_str = normalizar_ot(row_r[ot_col])
                                if not ot_str: continue
                                l_val = limpiar_legajo(row_r[legajo_col_cierre]) if legajo_col_cierre else None
                                n_val = str(row_r[nombre_col_cierre]).strip() if nombre_col_cierre and pd.notna(row_r[nombre_col_cierre]) else ""
                                if not filtrar_por_contrato(l_val, n_val, contrato_sel_cierre):
                                    continue

                                sol_val = str(row_r[sol_col]).strip() if sol_col and pd.notna(row_r[sol_col]) and str(row_r[sol_col]).lower() != "nan" else ""
                                fecha_val = str(row_r[fecha_col]).strip() if fecha_col and pd.notna(row_r[fecha_col]) and str(row_r[fecha_col]).lower() != "nan" else s_name
                                notif_val = str(row_r[notif_col]).strip() if notif_col and pd.notna(row_r[notif_col]) and str(row_r[notif_col]).lower() != "nan" else ""
                                est_val = ""

                                if ot_str in mapa_notif_drive:
                                    data_d = mapa_notif_drive[ot_str]
                                    if not sol_val and data_d.get("solicitud"): sol_val = data_d["solicitud"]
                                    if not notif_val and data_d.get("notificacion"): notif_val = data_d["notificacion"]
                                    if data_d.get("estado"): est_val = data_d["estado"]

                                sector_val = str(row_r[sector_col]).strip() if sector_col and pd.notna(row_r[sector_col]) and str(row_r[sector_col]).lower() != "nan" else ""
                                resp_val = str(row_r[resp_col]).strip() if resp_col and pd.notna(row_r[resp_col]) and str(row_r[resp_col]).lower() != "nan" else ""

                                if ot_str not in ots_dict:
                                    ots_dict[ot_str] = {
                                        "OT Programada": ot_str, "N° Solicitud": sol_val, "Estado Solicitud": est_val,
                                        "Fecha": fecha_val, "Fecha notificación final": notif_val,
                                        "Sector / Obra": sector_val, "Responsable del Trabajo": resp_val, 
                                        "Días Presente": [f"Día {d:02d}"], "Apariciones": 1
                                    }
                                else:
                                    ots_dict[ot_str]["Apariciones"] += 1
                                    if f"Día {d:02d}" not in ots_dict[ot_str]["Días Presente"]: ots_dict[ot_str]["Días Presente"].append(f"Día {d:02d}")
                                    if not ots_dict[ot_str]["N° Solicitud"] and sol_val: ots_dict[ot_str]["N° Solicitud"] = sol_val
                                    if not ots_dict[ot_str]["Estado Solicitud"] and est_val: ots_dict[ot_str]["Estado Solicitud"] = est_val

                    df_ots_unicas = pd.DataFrame(list(ots_dict.values()))

                if not df_ots_unicas.empty:
                    df_ots_unicas["Días Presente"] = df_ots_unicas["Días Presente"].apply(lambda x: ", ".join(x))
                    kpi_container = st.container()
                    st.divider()
                    st.write("📝 **Completá o modificá los datos directamente en la tabla:**")
                    edited_ots_df = st.data_editor(df_ots_unicas, use_container_width=True, key="editor_ots_cierre")
                    
                    def evaluar_estado_ot(row):
                        est = str(row["Estado Solicitud"]).strip().upper()
                        if any(k in est for k in ["RECHAZAD", "RECHAZO", "CANCELAD", "ANULAD"]): return "RECHAZADA"
                        if any(k in est for k in ["PENDIENTE", "PEND", "REVISION", "REVISIÓN", "CREADA", "EN PROCESO"]): return "PENDIENTE"
                        if any(k in est for k in ["APROBAD", "LIBERAD", "COMPLETAD", "FINALIZAD", "NOTIFICAD", "CERRAD", "OK"]): return "APROBADA"
                        f_notif = str(row["Fecha notificación final"]).strip()
                        if f_notif and f_notif.lower() not in ["nan", "none", "", "-"]: return "APROBADA"
                        return "PENDIENTE"

                    estados_eval = edited_ots_df.apply(evaluar_estado_ot, axis=1)
                    tot_ots, con_solic = len(edited_ots_df), len(edited_ots_df[edited_ots_df["N° Solicitud"].astype(str).str.strip() != ""])
                    ots_aprobadas, ots_rechazadas, ots_pendientes = int((estados_eval == "APROBADA").sum()), int((estados_eval == "RECHAZADA").sum()), int((estados_eval == "PENDIENTE").sum())

                    with kpi_container:
                        k_c1, k_c2, k_c3, k_c4, k_c5 = st.columns(5)
                        k_c1.metric("📌 Total OTs Únicas", tot_ots)
                        k_c2.metric("📋 OTs con Solicitud", con_solic)
                        k_c3.metric("🟢 OTs Aprobadas", ots_aprobadas)
                        k_c4.metric("🔴 OTs Rechazadas", ots_rechazadas)
                        k_c5.metric("⏳ OTs Pendientes", ots_pendientes)

                    st.write("")
                    col_b_guardar, col_b_descargar = st.columns([1, 1])
                    with col_b_guardar:
                        if st.button("💾 Guardar y Sincronizar Google Sheets", type="primary"):
                            try:
                                cambios_ots = {
                                    str(r["OT Programada"]).strip(): {
                                        "Solicitud": str(r["N° Solicitud"]).strip() if pd.notna(r["N° Solicitud"]) and str(r["N° Solicitud"]).lower() != "nan" else "", 
                                        "Estado": str(r["Estado Solicitud"]).strip() if pd.notna(r["Estado Solicitud"]) and str(r["Estado Solicitud"]).lower() != "nan" else "",
                                        "Fecha": str(r["Fecha"]).strip() if pd.notna(r["Fecha"]) and str(r["Fecha"]).lower() != "nan" else "",
                                        "FechaNotificacion": str(r["Fecha notificación final"]).strip() if pd.notna(r["Fecha notificación final"]) and str(r["Fecha notificación final"]).lower() != "nan" else "",
                                        "Sector": str(r["Sector / Obra"]).strip() if pd.notna(r["Sector / Obra"]) and str(r["Sector / Obra"]).lower() != "nan" else "",
                                        "Responsable": str(r["Responsable del Trabajo"]).strip() if pd.notna(r["Responsable del Trabajo"]) and str(r["Responsable del Trabajo"]).lower() != "nan" else ""
                                    } for _, r in edited_ots_df.iterrows()
                                }
                                with st.spinner("Sincronizando con Google Sheets..."):
                                    res = requests.post(WEBAPP_URL, json={"desdeDia": desde_dia_sel, "hastaDia": hasta_dia_sel, "cambios": cambios_ots})
                                if res.status_code == 200: st.success("🟢 ¡Sincronizado con éxito!")
                                else: st.error(f"Error {res.status_code}")
                            except Exception as e: st.error(f"Error al enviar datos: {e}")

                    with col_b_descargar:
                        out_cierre = io.BytesIO()
                        with pd.ExcelWriter(out_cierre, engine="openpyxl") as writer:
                            edited_ots_df.to_excel(writer, index=False, sheet_name="Cierre_OTs")
                        out_cierre.seek(0)
                        sufijo_arch = f"_{contrato_sel_cierre}" if contrato_sel_cierre != "Todos los Contratos" else ""
                        st.download_button("📥 Descargar Cierre de OTs (Excel)", data=out_cierre, file_name=f"Cierre_OTs{sufijo_arch}_Dia_{desde_dia_sel:02d}_al_{hasta_dia_sel:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else: 
                    st.warning("No se encontraron OTs cargadas para los filtros seleccionados.")
        else: st.info("No se detectaron días cargados en la planilla.")
    else: st.info("Conectando con Google Sheets...")

# --- PESTAÑA: ⏱️ RECUENTO Y SELECCIÓN PERSONALIZADA DE HORAS ---
with t_quincena:
    st.header("⏱️ Recuento y Selección Personalizada de Horas")
    st.info("ℹ️ Seleccioná el rango de días, el contrato a auditar y opcionalmente una **Planilla de Notificaciones de Drive**.")
    if carga_bytes:
        xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
        dias_disponibles_map_q = {int(m.group(1)): s for s in xl_carga.sheet_names if (m := re.search(r"(\d{1,2})[-/\.](\d{1,2})", s))}
        dias_ordenados_q = sorted(dias_disponibles_map_q.keys())

        if dias_ordenados_q:
            col_q1, col_q2, col_q3, col_q4 = st.columns([1, 1, 1.2, 1.2])
            with col_q1: desde_dia_q = st.selectbox("📅 Desde:", options=dias_ordenados_q, format_func=lambda d: f"Día {d:02d}", index=0, key="desde_q")
            with col_q2:
                dias_hasta_opciones_q = [d for d in dias_ordenados_q if d >= desde_dia_q]
                hasta_dia_q = st.selectbox("📅 Hasta:", options=dias_hasta_opciones_q, format_func=lambda d: f"Día {d:02d}", index=len(dias_hasta_opciones_q)-1 if dias_hasta_opciones_q else 0, key="hasta_q")
            
            with col_q3:
                archivos_notif_q = obtener_archivos_notif_drive()
                mapa_notif_files_q = {f['name']: f for f in archivos_notif_q}
                opciones_notif_q = ["-- Sin notificaciones (Drive) --"] + list(mapa_notif_files_q.keys())
                notif_sel_name_q = st.selectbox("📑 Planilla Notificaciones:", options=opciones_notif_q, key="notif_q_sel")
                notif_file_id_q = mapa_notif_files_q.get(notif_sel_name_q, {}).get('id') if notif_sel_name_q != "-- Sin notificaciones (Drive) --" else None
                notif_mime_type_q = mapa_notif_files_q.get(notif_sel_name_q, {}).get('mimeType') if notif_sel_name_q != "-- Sin notificaciones (Drive) --" else None

            with col_q4:
                st.write("")
                contrato_sel_horas = st.radio("🏗️ Filtrar Recuento por Contrato:", options=["Todos los Contratos", "MMYHH", "MOA"], horizontal=True, key="radio_contrato_horas")
                btn_calc_q = st.button("📊 Calcular Resumen de Horas", type="primary")

            if btn_calc_q or st.session_state.get("ver_quincena", False):
                st.session_state["ver_quincena"] = True
                try:
                    mapa_notif_drive_q, _ = obtener_notificaciones_drive(notif_file_id_q, notif_mime_type_q) if notif_file_id_q else ({}, {})
                    registros_quincena = []

                    for sheet in xl_carga.sheet_names:
                        match = re.search(r"(\d{1,2})[-/\.](\d{1,2})", sheet)
                        if not match: continue
                        dia, mes = int(match.group(1)), int(match.group(2))
                        if dia < desde_dia_q or dia > hasta_dia_q: continue

                        dfc_raw = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=sheet, header=None)
                        header_row_idx = next((idx for idx, row in dfc_raw.iterrows() if any(k in " ".join([str(v).upper() for v in row.values if pd.notna(v)]) for k in ["LEGAJO", "OBRA", "APELLIDO", "OT"])), 0)
                        dfc_q = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=sheet, skiprows=header_row_idx)

                        legajo_col = next((c for c in dfc_q.columns if "LEGAJO" in str(c).upper()), None)
                        nombre_col = next((c for c in dfc_q.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE", "PERSONAL"])), None)
                        trabajo_col = next((c for c in dfc_q.columns if "TRABAJO" in str(c).upper()), None)
                        tipo_col = next((c for c in dfc_q.columns if any(k in str(c).upper() for k in ["TIPO DE HORA", "TIPO"])), None)
                        ot_col_q = next((c for c in dfc_q.columns if "OT" in str(c).upper()), None)

                        if legajo_col and trabajo_col:
                            for _, row in dfc_q.iterrows():
                                leg_clean = limpiar_legajo(row[legajo_col])
                                nom = str(row[nombre_col]).strip() if nombre_col and pd.notna(row[nombre_col]) else (f"Legajo {leg_clean}" if leg_clean else "Sin Nombre")
                                if not filtrar_por_contrato(leg_clean, nom, contrato_sel_horas): continue
                                hs = extraer_horas(row[trabajo_col])
                                tipo_cat = "Hs. Normales"
                                if tipo_col and pd.notna(row[tipo_col]):
                                    tipo_raw = str(row[tipo_col]).strip().upper()
                                    tipo_cat = "Hs. 50%" if "0.5" in tipo_raw or "50" in tipo_raw else "Hs. 100%" if "1" in tipo_raw or "100" in tipo_raw else "Hs. Normales"

                                estado_ot = "APROBADA"
                                ot_str = normalizar_ot(row[ot_col_q]) if ot_col_q and pd.notna(row[ot_col_q]) else ""
                                if ot_str and ot_str in mapa_notif_drive_q:
                                    est_drive = mapa_notif_drive_q[ot_str].get("estado", "").upper()
                                    if any(k in est_drive for k in ["RECHAZAD", "RECHAZO", "CANCELAD", "ANULAD"]): estado_ot = "RECHAZADA"
                                    elif any(k in est_drive for k in ["PENDIENTE", "PEND", "REVISION", "REVISIÓN", "CREADA", "EN PROCESO"]): estado_ot = "PENDIENTE"

                                if leg_clean is not None or nom != "Sin Nombre":
                                    registros_quincena.append({"Legajo": leg_clean if leg_clean else 0, "Nombre": nom, "Horas": hs, "Tipo": tipo_cat, "EstadoOT": estado_ot})

                    df_q_all = pd.DataFrame(registros_quincena)
                    if not df_q_all.empty:
                        piv_q = df_q_all.pivot_table(index=["Legajo", "Nombre"], columns=["Tipo", "EstadoOT"], values="Horas", aggfunc="sum", fill_value=0.0)
                        piv_q.columns = [f"{col[0]} ({col[1]})" if isinstance(col, tuple) else str(col) for col in piv_q.columns]
                        piv_q = piv_q.reset_index()

                        for col_h in ["Hs. Normales (APROBADA)", "Hs. Normales (PENDIENTE)", "Hs. Normales (RECHAZADA)", "Hs. 50% (APROBADA)", "Hs. 50% (PENDIENTE)", "Hs. 50% (RECHAZADA)", "Hs. 100% (APROBADA)", "Hs. 100% (PENDIENTE)", "Hs. 100% (RECHAZADA)"]:
                            if col_h not in piv_q.columns: piv_q[col_h] = 0.0

                        piv_q["Hs. Aprobadas Total"] = piv_q.get("Hs. Normales (APROBADA)", 0) + piv_q.get("Hs. 50% (APROBADA)", 0) + piv_q.get("Hs. 100% (APROBADA)", 0)
                        piv_q["Hs. Pendientes Total"] = piv_q.get("Hs. Normales (PENDIENTE)", 0) + piv_q.get("Hs. 50% (PENDIENTE)", 0) + piv_q.get("Hs. 100% (PENDIENTE)", 0)
                        piv_q["Total Hs. Cargadas"] = piv_q["Hs. Aprobadas Total"] + piv_q["Hs. Pendientes Total"] + piv_q.get("Hs. Normales (RECHAZADA)", 0) + piv_q.get("Hs. 50% (RECHAZADA)", 0) + piv_q.get("Hs. 100% (RECHAZADA)", 0)

                        st.dataframe(piv_q, use_container_width=True)
                except Exception as e_q: st.error(f"Error al calcular: {e_q}")
        else: st.info("No se detectaron días cargados en la planilla.")
    else: st.info("Conectando con Google Sheets...")

# --- PESTAÑA: 🔍 5. CONTROL DETALLADO DE HORARIOS Y FICHADAS ---
with t_detalle:
    st.header("🔍 Control Detallado de Horarios y Carga por Persona")
    if carga_bytes and af_selected:
        st.info(f"📅 Auditando Pestaña de Carga **{nh}** vs Archivo de Fichada **{af_selected}**")
        st.write("")
        if st.button("🚀 Procesar Auditoría Detallada de Horarios", type="primary", key="btn_procesar_detalles_horario"):
            try:
                af_selected_id = mapa_archivos.get(af_selected)
                fh_fichada_det = descargar_archivo_drive(af_selected_id)
                dff_det = cargar_dataframe_fichada(fh_fichada_det)
                start_row_d, c_leg_d, c_nom_d, c_ing_d, c_egr_d, c_nov_d = detectar_estructura_fichada(dff_det)

                dfc_raw_d = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, header=None)
                header_row_idx_d = next((idx for idx, row in dfc_raw_d.iterrows() if any(k in " ".join([str(val).upper() for val in row.values if pd.notna(val)]) for k in ["LEGAJO", "OBRA", "APELLIDO"])), 0)
                dfc_det = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, skiprows=header_row_idx_d)

                legajo_col_d = next((c for c in dfc_det.columns if "LEGAJO" in str(c).upper()), None)
                nombre_col_d = next((c for c in dfc_det.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE"])), None)
                trabajo_col_d = next((c for c in dfc_det.columns if "TRABAJO" in str(c).upper()), None)

                mapa_horas_cargadas = {}
                for _, row_c in dfc_det.iterrows():
                    leg_c = limpiar_legajo(row_c[legajo_col_d]) if legajo_col_d else None
                    nom_c = str(row_c[nombre_col_d]).strip() if nombre_col_d and pd.notna(row_c[nombre_col_d]) else ""
                    hs_c = extraer_horas(row_c[trabajo_col_d]) if trabajo_col_d else 0.0
                    if leg_c is not None: mapa_horas_cargadas[leg_c] = mapa_horas_cargadas.get(leg_c, 0.0) + hs_c
                    if nom_c: mapa_horas_cargadas[normalizar_texto(nom_c)] = mapa_horas_cargadas.get(normalizar_texto(nom_c), 0.0) + hs_c

                filas_detalle = []
                for i in range(start_row_d, len(dff_det)):
                    lv = dff_det.iloc[i, c_leg_d] if c_leg_d < len(dff_det.columns) else None
                    nv = dff_det.iloc[i, c_nom_d] if c_nom_d < len(dff_det.columns) else None
                    ing = dff_det.iloc[i, c_ing_d] if c_ing_d < len(dff_det.columns) else ""
                    sal = dff_det.iloc[i, c_egr_d] if c_egr_d < len(dff_det.columns) else ""
                    leg_clean = limpiar_legajo(lv)
                    nom_str = str(nv).strip() if pd.notna(nv) else ""
                    if not nom_str and leg_clean is None: continue

                    hs_reloj = calcular_horas_fichada(ing, sal)
                    hs_cargadas = mapa_horas_cargadas.get(leg_clean, 0.0) if leg_clean is not None else mapa_horas_cargadas.get(normalizar_texto(nom_str), 0.0)

                    estado_cargas = "🟢 Carga Completa" if abs(hs_cargadas - hs_reloj) <= 0.25 and hs_reloj > 0 else "🔴 Sin Horas Cargadas" if hs_cargadas == 0.0 and hs_reloj > 0 else "⚠️ Discrepancia"
                    filas_detalle.append({"Legajo": str(leg_clean) if leg_clean else "SIN LEGAJO", "Nombre y Apellido": nom_str, "Horas Reloj Fichadas": round(hs_reloj, 1), "Horas Cargadas": round(hs_cargadas, 1), "Estado": estado_cargas})

                st.dataframe(pd.DataFrame(filas_detalle), use_container_width=True)
            except Exception as e_det: st.error(f"Error: {e_det}")

# --- PESTAÑA: 📁 6. CERTIFICACIÓN DE HORAS (CRUCE DESDE/HASTA + CONTRATO + PENDIENTE DE CARGA) ---
with t_cert:
    st.header("📁 Certificación y Cruce Previo a Facturación")
    st.info("ℹ️ Seleccioná el archivo de certificación en Drive. La app leerá las horas, la certificación (Columna S), el rango de fechas y cruzará contra Google Sheets para mostrarte las **Horas Pendientes de Carga / Fuera de Drive**.")

    if carga_bytes:
        col_cert1, col_cert2, col_cert3 = st.columns([1.2, 1.2, 1.2])
        with col_cert1:
            contrato_sel_cert = st.selectbox("🏗️ Seleccionar Contrato:", options=["MMYHH", "MOA"], index=1, key="sel_contrato_cert")
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
            if st.button("🚀 Procesar Cruce y Certificación de Horas", type="primary"):
                with st.spinner("Procesando datos de certificación y Google Sheets por rango de fechas..."):
                    try:
                        # 1. Leer archivo de certificación de Drive (Col Q: Horas, Col S: Certificación 'X')
                        fh_cert = descargar_archivo_notif_drive(cert_file_id, cert_mime_type)
                        df_cert_raw = pd.read_excel(fh_cert)
                        
                        cols = df_cert_raw.columns
                        col_horas_idx = 16 if len(cols) > 16 else None  # Columna Q
                        col_cert_idx = 18 if len(cols) > 18 else None   # Columna S
                        col_fecha_cert = next((c for c in cols if any(k in str(c).upper() for k in ["FECHA", "DIA", "DÍA"])), cols[3] if len(cols) > 3 else None)
                        col_legajo_cert = next((c for c in cols if "LEGAJO" in str(c).upper()), cols[1] if len(cols) > 1 else None)
                        col_nombre_cert = next((c for c in cols if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE", "PERSONAL"])), cols[2] if len(cols) > 2 else None)

                        # 2. Cargar planillas de Google Sheets para el rango de días seleccionado
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
                                        if not filtrar_por_contrato(l_val, n_val, contrato_sel_cert): continue
                                        hs_hojas = extraer_horas(r[t_col])
                                        if l_val is not None:
                                            registros_sheets.append({"Legajo": l_val, "Nombre": n_val, "HorasSheets": hs_hojas})

                        df_sheets_proc = pd.DataFrame(registros_sheets)
                        df_sheets_agrupado = pd.DataFrame(columns=["Legajo", "Nombre", "Total Horas Sheets"])
                        if not df_sheets_proc.empty:
                            df_sheets_agrupado = df_sheets_proc.groupby(["Legajo", "Nombre"], as_index=False)["HorasSheets"].sum()
                            df_sheets_agrupado.rename(columns={"HorasSheets": "Total Horas Sheets"}, inplace=True)

                        # 3. Procesar archivo de Drive con filtro de fecha
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

                            if l_val is not None or n_val:
                                resultados_cert.append({
                                    "Legajo": l_val if l_val else 0,
                                    "Nombre": n_val,
                                    "Horas Certificadas": hs if es_certificado else 0.0,
                                    "Horas No Certificadas": hs if not es_certificado else 0.0,
                                    "Total Horas Drive": hs
                                })

                        df_cert_proc = pd.DataFrame(resultados_cert)

                        if not df_cert_proc.empty or not df_sheets_agrupado.empty:
                            df_cert_resumen = df_cert_proc.groupby(["Legajo", "Nombre"], as_index=False)[
                                ["Horas Certificadas", "Horas No Certificadas", "Total Horas Drive"]
                            ].sum() if not df_cert_proc.empty else pd.DataFrame(columns=["Legajo", "Nombre", "Horas Certificadas", "Horas No Certificadas", "Total Horas Drive"])

                            # Realizar el merge completo con Google Sheets para capturar operarios de ambos lados
                            df_final_cruce = pd.merge(df_sheets_agrupado, df_cert_resumen, on=["Legajo", "Nombre"], how="outer").fillna(0.0)
                            
                            # Calcular la brecha: Horas en Google Sheets menos lo que figura en Drive (Certificadas + No Certificadas)
                            df_final_cruce["⚠️ Pendiente de Carga / Fuera de Drive"] = df_final_cruce["Total Horas Sheets"] - (df_final_cruce["Horas Certificadas"] + df_final_cruce["Horas No Certificadas"])
                            df_final_cruce["⚠️ Pendiente de Carga / Fuera de Drive"] = df_final_cruce["⚠️ Pendiente de Carga / Fuera de Drive"].apply(lambda x: max(0.0, x))

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
                            st.subheader(f"📋 Cruce y Auditoría de Certificación ({contrato_sel_cert} - Días {dia_desde_cert:02d} al {dia_hasta_cert:02d})")
                            
                            columnas_vista = ["Legajo", "Nombre", "Total Horas Sheets", "Horas Certificadas", "Horas No Certificadas", "⚠️ Pendiente de Carga / Fuera de Drive"]
                            st.dataframe(df_final_cruce[columnas_vista], use_container_width=True)

                            out_cert_excel = io.BytesIO()
                            with pd.ExcelWriter(out_cert_excel, engine="openpyxl") as writer:
                                df_final_cruce[columnas_vista].to_excel(writer, index=False, sheet_name=f"Certificacion_{contrato_sel_cert}")
                            out_cert_excel.seek(0)

                            st.download_button(
                                label=f"📥 Descargar Reporte de Auditoría ({contrato_sel_cert} - Días {dia_desde_cert:02d} al {dia_hasta_cert:02d})",
                                data=out_cert_excel,
                                file_name=f"Reporte_Auditoria_Certificacion_{contrato_sel_cert}_Dias_{dia_desde_cert:02d}_al_{dia_hasta_cert:02d}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("No se encontraron registros para los filtros seleccionados.")
                    except Exception as e_c:
                        st.error(f"Error al procesar el cruce de certificación: {e_c}")
        else:
            st.info("👆 Seleccioná un archivo de la carpeta de certificación en Google Drive para comenzar.")
    else:
        st.warning("Cargando conexión con las planillas...")
