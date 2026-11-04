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


# 🛡️ CAPA 2 MEJORADA: Detector dinámico con Fallback Inteligente a columnas G y H
def detectar_estructura_fichada(df_raw):
    header_row = 1 
    col_leg, col_nom, col_ing, col_egr, col_nov = -1, -1, -1, -1, -1

    for idx in range(min(15, len(df_raw))):
        row_str = " ".join([str(val).upper() for val in df_raw.iloc[idx].values if pd.notna(val)])
        if "APELLIDO" in row_str or "NOMBRE" in row_str:
            header_row = idx
            break

    # 1. Intentamos detectar por texto (Esto hace que los Días 1 al 8 vuelvan a funcionar perfecto)
    if header_row < len(df_raw):
        headers = [str(v).upper().strip() for v in df_raw.iloc[header_row].values]
        for idx, val in enumerate(headers):
            if "LEGAJO" in val or "LEG" in val: col_leg = idx
            elif "NOMBRE" in val or "APELLIDO" in val: col_nom = idx
            elif any(k in val for k in ["ING", "ENTRADA", "1RA", "HS.ING"]): col_ing = idx
            elif any(k in val for k in ["SAL", "EGRESO", "ULTIMA", "HS. SAL"]): col_egr = idx
            elif any(k in val for k in ["OBRA", "NOVEDAD", "OBSERVAC", "ESTADO"]): col_nov = idx

    # 2. FALLBACK DÍA 9: Si los textos fallan o vienen vacíos, forzamos G(6) y H(7) como indicaste
    if col_ing == -1 or col_egr == -1:
        if len(df_raw.columns) > 7:
            col_ing = 6  # Columna G
            col_egr = 7  # Columna H
        else:
            col_ing = col_nom + 1 if col_nom != -1 else 4
            col_egr = col_ing + 1
            
    # Resguardos finales
    if col_leg == -1: col_leg = max(0, col_nom - 1) if col_nom > 0 else 2
    if col_nov == -1: col_nov = 12

    return header_row + 1, col_leg, col_nom, col_ing, col_egr, col_nov


# 🛡️ CAPA 1 MEJORADA: Une múltiples hojas pero ignora la Hoja 1 si existe la Hoja 2
def cargar_dataframe_fichada(fh_bytes):
    xl = pd.ExcelFile(fh_bytes)
    hojas_validas = []
    
    for sheet in xl.sheet_names:
        df_temp = pd.read_excel(fh_bytes, sheet_name=sheet, nrows=15, header=None)
        if df_temp.empty: continue
        texto_hoja = " ".join([str(val).upper() for val in df_temp.values.flatten() if pd.notna(val)])
        if any(k in texto_hoja for k in ["APELLIDO", "LEGAJO", "NOMBRE"]):
            hojas_validas.append(sheet)
            
    # Si el archivo tiene "Hoja 2" y "Hoja 1", la Hoja 1 suele ser basura o resumen. La omitimos.
    if len(hojas_validas) > 1 and any(s in hojas_validas for s in ["Hoja2", "Hoja 2", "Sheet2"]):
        if hojas_validas[0] in ["Hoja1", "Hoja 1", "Sheet1"]:
            hojas_validas.pop(0)
            
    if not hojas_validas:
        hojas_validas = [xl.sheet_names[-1]]
        
    dfs = []
    for idx, sheet in enumerate(hojas_validas):
        df_s = pd.read_excel(fh_bytes, sheet_name=sheet, header=None)
        if df_s.empty: continue
        
        if idx > 0:
            # Quitamos los títulos de la Hoja 3 para que se pegue impecable debajo de la Hoja 2
            start_r, _, _, _, _, _ = detectar_estructura_fichada(df_s)
            df_s = df_s.iloc[start_r:]
            
        dfs.append(df_s)
        
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# 🛡️ CAPA 3: Extracción de números puros mediante Regex
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

def limpiar_legajo(val):
    if pd.isna(val): return None
    try: return int(float(str(val).strip()))
    except (ValueError, TypeError): return None

def obtener_excel_nube(url):
    response = requests.get(url)
    response.raise_for_status()
    return response.content

def normalizar_texto(val):
    if pd.isna(val): return ""
    s = str(val).lower().strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return " ".join("".join(c for c in s if c.isalnum() or c.isspace()).split())

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

t_rrhh, t_carga, t_reglas, t_reporte, t_cierre, t_quincena, t_detalle = st.tabs([
    "📥 Carga de Fichadas/RRHH", "📂 1. Carga", "⚙️ 2. Reglas", "📊 3. Reporte", "📋 Cierre de Período", "⏱️ 4. Horas y Período Personalizado", "🔍 5. Control Horarios y Fichadas"
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
                                        
                                        # NUEVO: Busca por TODAS las hojas del excel original
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
                                                    
                                                    if not cell_name and cell_leg is None:
                                                        continue

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
            col_btn_ej, _ = st.columns([1, 3])
            with col_btn_ej:
                btn_ejecutar_cierre = st.button("🚀 Ejecutar Cruce de OTs", type="primary")

            if btn_ejecutar_cierre:
                st.session_state["cruce_cierre_ejecutado"] = True

            if st.session_state["cruce_cierre_ejecutado"]:
                with st.spinner("Procesando OTs y cruzando notificaciones desde Drive..."):
                    mapa_notif_drive, diag_info = obtener_notificaciones_drive(notif_file_id, notif_mime_type) if notif_file_id else ({}, {})

                    if mapa_notif_drive:
                        st.info(f"💡 Se detectaron **{len(mapa_notif_drive)} OTs válidas** en Drive (Columna OT: `{diag_info.get('ot_col')}`, Solicitud: `{diag_info.get('sol_col')}`, Fecha Col. D: `{diag_info.get('fecha_col')}`).")
                    elif notif_file_id:
                        st.warning("⚠️ No se pudieron asociar OTs desde la planilla seleccionada. Verificá el formato del archivo.")

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

                        if ot_col:
                            for _, row_r in dfc_ot.iterrows():
                                ot_str = normalizar_ot(row_r[ot_col])
                                if not ot_str: continue

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
                                        "OT Programada": ot_str, 
                                        "N° Solicitud": sol_val, 
                                        "Estado Solicitud": est_val,
                                        "Fecha": fecha_val,
                                        "Fecha notificación final": notif_val,
                                        "Sector / Obra": sector_val,
                                        "Responsable del Trabajo": resp_val, 
                                        "Días Presente": [f"Día {d:02d}"], 
                                        "Apariciones": 1
                                    }
                                else:
                                    ots_dict[ot_str]["Apariciones"] += 1
                                    if f"Día {d:02d}" not in ots_dict[ot_str]["Días Presente"]: ots_dict[ot_str]["Días Presente"].append(f"Día {d:02d}")
                                    if not ots_dict[ot_str]["N° Solicitud"] and sol_val: ots_dict[ot_str]["N° Solicitud"] = sol_val
                                    if not ots_dict[ot_str]["Estado Solicitud"] and est_val: ots_dict[ot_str]["Estado Solicitud"] = est_val
                                    if not ots_dict[ot_str]["Fecha"] and fecha_val: ots_dict[ot_str]["Fecha"] = fecha_val
                                    if not ots_dict[ot_str]["Fecha notificación final"] and notif_val: ots_dict[ot_str]["Fecha notificación final"] = notif_val
                                    if not ots_dict[ot_str]["Sector / Obra"] and sector_val: ots_dict[ot_str]["Sector / Obra"] = sector_val
                                    if not ots_dict[ot_str]["Responsable del Trabajo"] and resp_val: ots_dict[ot_str]["Responsable del Trabajo"] = resp_val

                    df_ots_unicas = pd.DataFrame(list(ots_dict.values()))

                if not df_ots_unicas.empty:
                    df_ots_unicas["Días Presente"] = df_ots_unicas["Días Presente"].apply(lambda x: ", ".join(x))
                    
                    kpi_container = st.container()
                    st.divider()
                    st.write("📝 **Completá o modificá los datos directamente en la tabla:**")
                    
                    edited_ots_df = st.data_editor(
                        df_ots_unicas, 
                        column_config={
                            "OT Programada": st.column_config.TextColumn("OT Programada", disabled=True), 
                            "N° Solicitud": st.column_config.TextColumn("N° Solicitud"),
                            "Estado Solicitud": st.column_config.TextColumn("Estado Solicitud"),
                            "Fecha": st.column_config.TextColumn("Fecha"),
                            "Fecha notificación final": st.column_config.TextColumn("Fecha notificación final"),
                            "Sector / Obra": st.column_config.TextColumn("Sector / Obra"),
                            "Responsable del Trabajo": st.column_config.TextColumn("Responsable del Trabajo"), 
                            "Días Presente": st.column_config.TextColumn("Días Presente", disabled=True), 
                            "Apariciones": st.column_config.NumberColumn("Apariciones", disabled=True)
                        }, 
                        use_container_width=True, 
                        key="editor_ots_cierre"
                    )

                    def evaluar_estado_ot(row):
                        est = str(row["Estado Solicitud"]).strip().upper()
                        if any(k in est for k in ["RECHAZAD", "RECHAZO", "CANCELAD", "ANULAD"]): return "RECHAZADA"
                        if any(k in est for k in ["PENDIENTE", "PEND", "REVISION", "REVISIÓN", "CREADA", "EN PROCESO"]): return "PENDIENTE"
                        if any(k in est for k in ["APROBAD", "LIBERAD", "COMPLETAD", "FINALIZAD", "NOTIFICAD", "CERRAD", "OK"]): return "APROBADA"
                        f_notif = str(row["Fecha notificación final"]).strip()
                        if f_notif and f_notif.lower() not in ["nan", "none", "", "-"]: return "APROBADA"
                        return "PENDIENTE"

                    estados_eval = edited_ots_df.apply(evaluar_estado_ot, axis=1)

                    tot_ots = len(edited_ots_df)
                    con_solic = len(edited_ots_df[edited_ots_df["N° Solicitud"].astype(str).str.strip() != ""])
                    ots_aprobadas = int((estados_eval == "APROBADA").sum())
                    ots_rechazadas = int((estados_eval == "RECHAZADA").sum())
                    ots_pendientes = int((estados_eval == "PENDIENTE").sum())

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
                        st.download_button(
                            label="📥 Descargar Cierre de OTs (Excel)",
                            data=out_cierre,
                            file_name=f"Cierre_OTs_Dia_{desde_dia_sel:02d}_al_{hasta_dia_sel:02d}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else: st.warning("No se encontraron OTs cargadas.")
        else: st.info("No se detectaron días cargados en la planilla.")
    else: st.info("Conectando con Google Sheets...")

# --- PESTAÑA: ⏱️ RECUENTO Y SELECCIÓN PERSONALIZADA DE HORAS ---
with t_quincena:
    st.header("⏱️ Recuento y Selección Personalizada de Horas")
    if carga_bytes:
        xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
        dias_disponibles_map_q = {int(m.group(1)): s for s in xl_carga.sheet_names if (m := re.search(r"(\d{1,2})[-/\.](\d{1,2})", s))}
        dias_ordenados_q = sorted(dias_disponibles_map_q.keys())

        if dias_ordenados_q:
            col_q1, col_q2, col_q3 = st.columns([1, 1, 1.2])
            with col_q1: desde_dia_q = st.selectbox("📅 Fecha Desde:", options=dias_ordenados_q, format_func=lambda d: f"Día {d:02d} ({dias_disponibles_map_q[d]})", index=0, key="desde_q")
            with col_q2:
                dias_hasta_opciones_q = [d for d in dias_ordenados_q if d >= desde_dia_q]
                hasta_dia_q = st.selectbox("📅 Fecha Hasta:", options=dias_hasta_opciones_q, format_func=lambda d: f"Día {d:02d} ({dias_disponibles_map_q[d]})", index=len(dias_hasta_opciones_q)-1 if dias_hasta_opciones_q else 0, key="hasta_q")
            with col_q3:
                st.write("")
                btn_calc_q = st.button("📊 Calcular Resumen de Horas", type="primary")

            if btn_calc_q or st.session_state.get("ver_quincena", False):
                st.session_state["ver_quincena"] = True
                try:
                    registros_quincena = []

                    for sheet in xl_carga.sheet_names:
                        match = re.search(r"(\d{1,2})[-/\.](\d{1,2})", sheet)
                        if not match: continue
                        dia, mes = int(match.group(1)), int(match.group(2))
                        if dia < desde_dia_q or dia > hasta_dia_q: continue

                        dfc_raw = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=sheet, header=None)
                        header_row_idx = next((idx for idx, row in dfc_raw.iterrows() if any(k in " ".join([str(v).upper() for v in row.values if pd.notna(v)]) for k in ["LEGAJO", "OBRA", "APELLIDO"])), 0)
                        dfc_q = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=sheet, skiprows=header_row_idx)

                        legajo_col = next((c for c in dfc_q.columns if "LEGAJO" in str(c).upper()), None)
                        nombre_col = next((c for c in dfc_q.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE", "PERSONAL"])), None)
                        trabajo_col = next((c for c in dfc_q.columns if "TRABAJO" in str(c).upper()), None)
                        tipo_col = next((c for c in dfc_q.columns if any(k in str(c).upper() for k in ["TIPO DE HORA", "TIPO"])), None)

                        if legajo_col and trabajo_col and tipo_col:
                            for _, row in dfc_q.iterrows():
                                leg_clean = limpiar_legajo(row[legajo_col])
                                nom = str(row[nombre_col]).strip() if nombre_col and pd.notna(row[nombre_col]) else (f"Legajo {leg_clean}" if leg_clean else "Sin Nombre")
                                hs = extraer_horas(row[trabajo_col])
                                tipo_raw = str(row[tipo_col]).strip().upper() if pd.notna(row[tipo_col]) else "HS. NORMALES"
                                tipo_cat = "Hs. 50%" if "0.5" in tipo_raw or "50" in tipo_raw else "Hs. 100%" if "1" in tipo_raw or "100" in tipo_raw else "Hs. Normales"

                                if leg_clean is not None or nom != "Sin Nombre":
                                    registros_quincena.append({"Legajo": leg_clean if leg_clean else 0, "Nombre": nom, "Horas": hs, "Tipo": tipo_cat})

                    df_q_all = pd.DataFrame(registros_quincena)
                    if not df_q_all.empty:
                        piv_q = df_q_all.pivot_table(index=["Legajo", "Nombre"], columns="Tipo", values="Horas", aggfunc="sum", fill_value=0.0).reset_index()
                        for col_h in ["Hs. Normales", "Hs. 50%", "Hs. 100%"]:
                            if col_h not in piv_q.columns: piv_q[col_h] = 0.0
                        piv_q["Total Hs. Cargadas"] = piv_q["Hs. Normales"] + piv_q["Hs. 50%"] + piv_q["Hs. 100%"]

                        fichadas_horas_dict = {}
                        with st.spinner("Analizando fichadas en Google Drive..."):
                            archivos_drive = obtener_archivos_drive()
                            mapa_archivos_q = {f['name']: f['id'] for f in archivos_drive if f['name'].lower().endswith(('.xlsx', '.xltx'))}
                            
                            for f_loc, f_id in mapa_archivos_q.items():
                                match_f = re.search(r"fichada[_\s]*(\d{1,2})[-/\.](\d{1,2})", f_loc, re.IGNORECASE)
                                if match_f:
                                    d_f = int(match_f.group(1))
                                    if d_f < desde_dia_q or d_f > hasta_dia_q: continue
                                    try:
                                        fh_loc = descargar_archivo_drive(f_id)
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

                                                fichadas_horas_dict[l_f_clean] = fichadas_horas_dict.get(l_f_clean, 0.0) + calcular_horas_fichada(ing_q, sal_q)
                                    except Exception: pass

                        df_q_final = pd.merge(piv_q, pd.DataFrame(list(fichadas_horas_dict.items()), columns=["Legajo", "Total Hs. Fichadas (Reloj)"]), on="Legajo", how="left").fillna(0.0)
                        df_q_final["Diferencia (Cargado vs Fichado)"] = df_q_final["Total Hs. Cargadas"] - df_q_final["Total Hs. Fichadas (Reloj)"]

                        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                        kpi1.metric("📊 Total Hs. Normales", f"{df_q_final['Hs. Normales'].sum():.1f} hs")
                        kpi2.metric("⚡ Total Hs. 50%", f"{df_q_final['Hs. 50%'].sum():.1f} hs")
                        kpi3.metric("🔥 Total Hs. 100%", f"{df_q_final['Hs. 100%'].sum():.1f} hs")
                        kpi4.metric("⏱️ Total Horas Cargadas", f"{df_q_final['Total Hs. Cargadas'].sum():.1f} hs")
                        
                        st.divider()
                        st_q_styled = df_q_final[["Legajo", "Nombre", "Hs. Normales", "Hs. 50%", "Hs. 100%", "Total Hs. Cargadas", "Total Hs. Fichadas (Reloj)", "Diferencia (Cargado vs Fichado)"]].style
                        def func_color_q(val):
                            if val > 0: return "background-color: #FFEB9C; color: #856404;"
                            if val < 0: return "background-color: #F8D7DA; color: #721C24;"
                            return "background-color: #C6EFCE; color: #155724;"

                        if hasattr(st_q_styled, "map"):
                            st_q_styled = st_q_styled.map(func_color_q, subset=["Diferencia (Cargado vs Fichado)"])
                        else:
                            st_q_styled = st_q_styled.applymap(func_color_q, subset=["Diferencia (Cargado vs Fichado)"])

                        st.dataframe(st_q_styled, use_container_width=True)

                        out_q = io.BytesIO()
                        with pd.ExcelWriter(out_q, engine="openpyxl") as writer: df_q_final.to_excel(writer, index=False, sheet_name="Resumen Horas")
                        out_q.seek(0)
                        st.download_button("📥 Descargar Resumen de Horas (Excel)", data=out_q, file_name=f"Resumen_Horas_Dia_{desde_dia_q:02d}_al_{hasta_dia_q:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else: st.warning("No se encontraron registros de horas cargadas para el rango seleccionado.")
                except Exception as e_q: st.error(f"Error al calcular: {e_q}")
        else: st.info("No se detectaron pestañas válidas de días en la planilla.")
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
                with st.spinner("Leyendo registros de fichadas y horarios desde Drive..."):
                    fh_fichada_det = descargar_archivo_drive(af_selected_id)
                    dff_det = cargar_dataframe_fichada(fh_fichada_det)

                start_row_d, c_leg_d, c_nom_d, c_ing_d, c_egr_d, c_nov_d = detectar_estructura_fichada(dff_det)

                with st.spinner("Procesando datos de la planilla de carga..."):
                    dfc_raw_d = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, header=None)
                    header_row_idx_d = 0
                    for idx, row in dfc_raw_d.iterrows():
                        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
                        if "LEGAJO" in row_str or "OBRA" in row_str or "APELLIDO" in row_str:
                            header_row_idx_d = idx
                            break
                    dfc_det = pd.read_excel(io.BytesIO(carga_bytes), sheet_name=nh, skiprows=header_row_idx_d)

                    legajo_col_d = next((c for c in dfc_det.columns if "LEGAJO" in str(c).upper()), None)
                    nombre_col_d = next((c for c in dfc_det.columns if any(k in str(c).upper() for k in ["APELLIDO", "NOMBRE"])), None)
                    trabajo_col_d = next((c for c in dfc_det.columns if "TRABAJO" in str(c).upper()), None)

                    mapa_horas_cargadas = {}
                    for _, row_c in dfc_det.iterrows():
                        leg_c = limpiar_legajo(row_c[legajo_col_d]) if legajo_col_d else None
                        nom_c = str(row_c[nombre_col_d]).strip() if nombre_col_d and pd.notna(row_c[nombre_col_d]) else ""
                        hs_c = extraer_horas(row_c[trabajo_col_d]) if trabajo_col_d else 0.0

                        if leg_c is not None:
                            mapa_horas_cargadas[leg_c] = mapa_horas_cargadas.get(leg_c, 0.0) + hs_c
                        if nom_c:
                            mapa_horas_cargadas[normalizar_texto(nom_c)] = mapa_horas_cargadas.get(normalizar_texto(nom_c), 0.0) + hs_c

                filas_detalle = []
                
                for i in range(start_row_d, len(dff_det)):
                    lv = dff_det.iloc[i, c_leg_d] if c_leg_d < len(dff_det.columns) else None
                    nv = dff_det.iloc[i, c_nom_d] if c_nom_d < len(dff_det.columns) else None
                    ing = dff_det.iloc[i, c_ing_d] if c_ing_d < len(dff_det.columns) else ""
                    sal = dff_det.iloc[i, c_egr_d] if c_egr_d < len(dff_det.columns) else ""

                    if str(lv).startswith('#'): lv = None
                    if str(nv).startswith('#'): nv = None
                    if str(ing).startswith('#'): ing = ""
                    if str(sal).startswith('#'): sal = ""

                    leg_clean = limpiar_legajo(lv)
                    nom_str = str(nv).strip() if pd.notna(nv) else ""
                    if not nom_str and leg_clean is None: continue

                    ing_str = formatear_horario(ing)
                    sal_str = formatear_horario(sal)
                    hs_reloj = calcular_horas_fichada(ing, sal)

                    hs_cargadas = 0.0
                    if leg_clean is not None and leg_clean in mapa_horas_cargadas:
                        hs_cargadas = mapa_horas_cargadas[leg_clean]
                    elif nom_str and normalizar_texto(nom_str) in mapa_horas_cargadas:
                        hs_cargadas = mapa_horas_cargadas[normalizar_texto(nom_str)]

                    if leg_clean is not None and leg_clean in EXCLUIDOS_SECTOR:
                        estado_cargas = "🚫 Sector Excluido"
                    elif hs_cargadas == 0.0 and hs_reloj > 0:
                        estado_cargas = "🔴 Sin Horas Cargadas"
                    elif abs(hs_cargadas - hs_reloj) <= 0.25 and hs_reloj > 0:
                        estado_cargas = "🟢 Carga Completa / Correcta"
                    elif hs_cargadas > 0:
                        estado_cargas = f"⚠️ Discrepancia (Cargado: {hs_cargadas:.1f}h / Fichado: {hs_reloj:.1f}h)"
                    else:
                        estado_cargas = "⚪ Sin Registro"

                    filas_detalle.append({
                        "Legajo": str(leg_clean) if leg_clean else "SIN LEGAJO",
                        "Nombre y Apellido": nom_str,
                        "Horario Desde (Ingreso)": ing_str,
                        "Horario Hasta (Salida)": sal_str,
                        "Horas Reloj Fichadas": round(hs_reloj, 1),
                        "Horas Cargadas en Planilla": round(hs_cargadas, 1),
                        "Estado de Validación": estado_cargas
                    })

                df_resultado_det = pd.DataFrame(filas_detalle)

                if not df_resultado_det.empty:
                    c_k1, c_k2, c_k3, c_k4 = st.columns(4)
                    c_k1.metric("👥 Total Personal Fichado", len(df_resultado_det))
                    c_k2.metric("🟢 Carga Correcta", len(df_resultado_det[df_resultado_det["Estado de Validación"].str.contains("🟢")]))
                    c_k3.metric("🔴 Faltan Cargar", len(df_resultado_det[df_resultado_det["Estado de Validación"].str.contains("🔴")]))
                    c_k4.metric("⚠️ Discrepancias de Horas", len(df_resultado_det[df_resultado_det["Estado de Validación"].str.contains("⚠️")]))

                    st.divider()

                    col_f_txt, col_f_sel = st.columns([1.5, 1])
                    with col_f_txt:
                        texto_busqueda = st.text_input("🔍 Buscar por Nombre o Legajo:", value="", placeholder="Ej: Kriger, 50028177...").strip().lower()
                    with col_f_sel:
                        filtro_estado = st.selectbox(
                            "🔎 Filtrar por Estado de Carga:",
                            options=["Todos los Registros", "🔴 Faltan Cargar (Sin Horas Cargadas)", "⚠️ Discrepancias de Horas", "🟢 Carga Completa / Correcta", "🚫 Sector Excluido"]
                        )

                    df_mostrar = df_resultado_det.copy()

                    if texto_busqueda:
                        df_mostrar = df_mostrar[
                            df_mostrar["Nombre y Apellido"].astype(str).str.lower().str.contains(texto_busqueda) |
                            df_mostrar["Legajo"].astype(str).str.lower().str.contains(texto_busqueda)
                        ]

                    if filtro_estado != "Todos los Registros":
                        tag_emoji = filtro_estado.split()[0]
                        df_mostrar = df_mostrar[df_mostrar["Estado de Validación"].str.contains(tag_emoji, na=False)]

                    st_styled_det = df_mostrar.style
                    def func_color_det(val):
                        sval = str(val)
                        if "🟢" in sval: return "background-color: #C6EFCE; color: #155724;"
                        if "🔴" in sval: return "background-color: #F8D7DA; color: #721C24;"
                        if "⚠️" in sval: return "background-color: #FFEB9C; color: #856404;"
                        if "🚫" in sval: return "background-color: #E2EFDA;"
                        return ""

                    if hasattr(st_styled_det, "map"):
                        st_styled_det = st_styled_det.map(func_color_det, subset=["Estado de Validación"])
                    else:
                        st_styled_det = st_styled_det.applymap(func_color_det, subset=["Estado de Validación"])

                    st.dataframe(
                        st_styled_det,
                        column_config={
                            "Horas Reloj Fichadas": st.column_config.NumberColumn(format="%.1f hs"),
                            "Horas Cargadas en Planilla": st.column_config.NumberColumn(format="%.1f hs")
                        },
                        use_container_width=True
                    )

                    out_det = io.BytesIO()
                    with pd.ExcelWriter(out_det, engine="openpyxl") as writer:
                        df_resultado_det.to_excel(writer, index=False, sheet_name="Control_Horarios")
                    out_det.seek(0)

                    st.download_button(
                        "📥 Descargar Control Detallado de Horarios (Excel)",
                        data=out_det,
                        file_name=f"Control_Detallado_Horarios_{nh}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No se encontraron registros de fichadas para procesar.")
            except Exception as e_det:
                st.error(f"Error al procesar el detalle de horarios: {e_det}")
    else:
        st.info("Seleccioná un día y un archivo de fichada en la solapa '1. Carga' para ejecutar la auditoría.")
