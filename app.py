import datetime
import io
import os
import re
import unicodedata
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(
    page_title="Control de Fichadas - SJG", page_icon="🏭", layout="wide"
)

# Inicializar estado de sesión
if "mostrar_reporte" not in st.session_state:
  st.session_state["mostrar_reporte"] = False

# --- CONFIGURACIÓN GOOGLE SHEETS & APPS SCRIPT WEB APP ---
SHEET_ID = "1gGkGjO-ZX7auYdlKAuI9rLEN_KoVah3vdnz7WofcHq0"
URL_DESCARGA_SHEET = (
    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
)
WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxZOKYHQHzLR-gZionm3Ub1z8Rlrz5l6AtR1D8gLKKIeAqhSGMf62vMfsa4F2s1gcweFA/exec"

# --- CREAR CARPETA BASE DE DATOS SI NO EXISTE ---
BASE_FICHADAS = r"C:\Users\SJG Coordinacion\Desktop\Fichadas_Crudas"
if not os.path.exists(BASE_FICHADAS):
  os.makedirs(BASE_FICHADAS)

# --- LOGO VECTORIAL (SJG MONTAJES INDUSTRIALES) ---
SJG_LOGO_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 130" width="100%" height="75" style="max-width: 380px;"><g fill="#0A0A0A"><text x="5" y="82" font-family="\'Montserrat\', \'Arial Black\', sans-serif" font-weight="900" font-size="78" letter-spacing="-1.5">SJG</text><text x="8" y="116" font-family="\'Montserrat\', \'Arial\', sans-serif" font-weight="800" font-size="17" letter-spacing="3.2">MONTAJES INDUSTRIALES</text><polygon points="230,126 260,126 282,98 230,126" /><polygon points="318,24 294,126 268,126 290,98 318,24" /><polygon points="352,70 302,126 392,126" /></g></svg>'

col_logo, col_titulo = st.columns([1.2, 3])

with col_logo:
  st.markdown(SJG_LOGO_SVG, unsafe_allow_html=True)

with col_titulo:
  st.title("Sistema de Control y Validación de Personal")
  st.caption(
      "Conciliación automatizada de Partes de Carga vs. Fichadas en Puerta"
  )

st.divider()

MESES_NOMBRES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

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
  if pd.isna(val):
    return None
  try:
    return int(float(str(val).strip()))
  except (ValueError, TypeError):
    return None


@st.cache_data(show_spinner=False)
def obtener_excel_nube(url):
  response = requests.get(url)
  response.raise_for_status()
  return response.content


def normalizar_texto(val):
  if pd.isna(val):
    return ""
  s = str(val).lower().strip()
  s = "".join(
      c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
  )
  s = "".join(c for c in s if c.isalnum() or c.isspace())
  return " ".join(s.split())


def obtener_dia_mes(texto_o_fecha):
  if isinstance(
      texto_o_fecha, (pd.Timestamp, datetime.date, datetime.datetime)
  ):
    return texto_o_fecha.day, texto_o_fecha.month
  patron = r"(\d{1,2})[-/\.](\d{1,2})"
  match = re.search(patron, str(texto_o_fecha))
  if match:
    try:
      d, m = int(match.group(1)), int(match.group(2))
      if 1 <= d <= 31 and 1 <= m <= 12:
        return d, m
    except Exception:
      pass
  return None, None


def extraer_fecha_contenido(df_raw):
  patron = r"(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{2,4})"
  for idx in range(min(10, len(df_raw))):
    row = df_raw.iloc[idx]
    for val in row.dropna():
      if isinstance(
          val, (pd.Timestamp, datetime.date, datetime.datetime)
      ):
        return val.day, val.month
      match = re.search(patron, str(val))
      if match:
        try:
          d, m = int(match.group(1)), int(match.group(2))
          if 1 <= d <= 31 and 1 <= m <= 12:
            return d, m
        except Exception:
          pass
  return None, None


def extraer_horas(val):
  if pd.isna(val):
    return 0.0
  val_str = str(val).upper().strip()
  match = re.search(r"(\d+(?:[\.,]\d+)?)", val_str)
  if match:
    return float(match.group(1).replace(",", "."))
  return 0.0


def calcular_horas_fichada(ing, sal):
  try:
    i = float(ing)
    s = float(sal)
    if s >= i:
      return s - i
    else:
      return (24.0 - i) + s
  except Exception:
    return 0.0


t_rrhh, t_carga, t_reglas, t_reporte, t_cierre, t_quincena = st.tabs([
    "📥 Carga de Fichadas/RRHH",
    "📂 1. Carga",
    "⚙️ 2. Reglas",
    "📊 3. Reporte",
    "📋 Cierre de Período",
    "⏱️ 4. Horas y Quincenas",
])

with t_rrhh:
  st.header("Carga Diaria de Fichadas (Portería / RRHH)")
  st.info(
      "ℹ️ Sube el reporte exportado del reloj de puerta para guardarlo en la"
      " base."
  )
  archivo_nuevo = st.file_uploader(
      "Subir Archivo de Fichada Diaria", type=["xlsx", "xltx"], key="uploader_rrhh"
  )

  if archivo_nuevo:
    if st.button("💾 Guardar Fichada en Base de Datos", type="primary"):
      try:
        dff_temp = pd.read_excel(
            archivo_nuevo, sheet_name="Hoja2", header=None
        )
        d, m = extraer_fecha_contenido(dff_temp)
        if d and m:
          nombre_final = f"Fichada_{d:02d}-{m:02d}.xlsx"
        else:
          hoy = datetime.datetime.now()
          nombre_final = f"Fichada_{hoy.day:02d}-{hoy.month:02d}.xlsx"

        ruta_guardado = os.path.join(BASE_FICHADAS, nombre_final)
        with open(ruta_guardado, "wb") as f:
          f.write(archivo_nuevo.getbuffer())

        logo_pequeno = SJG_LOGO_SVG.replace('height="75"', 'height="40"').replace(
            "max-width: 380px;", "max-width: 150px;"
        )
        mensaje_exito = f"""
                <div style="display: flex; align-items: center; gap: 20px; background-color: #E2EFDA; padding: 15px 25px; border-radius: 8px; border-left: 6px solid #28a745; margin-top: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <div style="flex-shrink: 0;">{logo_pequeno}</div>
                    <div>
                        <h4 style="margin: 0; color: #155724; font-size: 1.15rem;">✅ ¡Guardado con éxito!</h4>
                        <p style="margin: 5px 0 0 0; color: #1e5c2b; font-size: 0.95rem;">
                            El archivo <b>{nombre_final}</b> se almacenó correctamente.
                        </p>
                    </div>
                </div>
                """
        st.markdown(mensaje_exito, unsafe_allow_html=True)
      except Exception as e:
        st.error(f"Error al guardar el archivo: {e}")

with t_carga:
  c1, c2 = st.columns(2)

  with c2:
    st.subheader("Base de Fichadas Guardadas")
    archivos_guardados = (
        [
            f
            for f in os.listdir(BASE_FICHADAS)
            if f.endswith((".xlsx", ".xltx"))
        ]
        if os.path.exists(BASE_FICHADAS)
        else []
    )

    if len(archivos_guardados) == 0:
      st.warning("Aún no hay fichadas guardadas.")
      af_selected = None
    else:
      fichadas_por_mes = {}
      for f in archivos_guardados:
        match = re.search(r"Fichada_(\d{1,2})-(\d{1,2})", f)
        if match:
          dia, mes = int(match.group(1)), int(match.group(2))
          if mes not in fichadas_por_mes:
            fichadas_por_mes[mes] = {}
          fichadas_por_mes[mes][dia] = f

      if not fichadas_por_mes:
        opciones = ["-- Seleccione una fichada --"] + archivos_guardados
        af_selected = st.selectbox("Elegir Fichada:", options=opciones)
        if af_selected == "-- Seleccione una fichada --":
          af_selected = None
      else:
        col_mes, col_dia = st.columns(2)
        with col_mes:
          meses_ordenados = sorted(fichadas_por_mes.keys(), reverse=True)
          mes_elegido_num = st.selectbox(
              "Mes de Fichada:",
              options=meses_ordenados,
              format_func=lambda m: f"📅 {MESES_NOMBRES.get(m, 'Mes')}",
          )
        with col_dia:
          dias_del_mes = sorted(
              fichadas_por_mes[mes_elegido_num].keys(), reverse=True
          )
          opciones_dia = ["-- Seleccionar Día --"] + [
              f"Día {d:02d}" for d in dias_del_mes
          ]
          dia_elegido_str = st.selectbox(
              "Día de Fichada:", options=opciones_dia
          )

        if dia_elegido_str != "-- Seleccionar Día --":
          dia_elegido_num = int(dia_elegido_str.replace("Día ", ""))
          af_selected = fichadas_por_mes[mes_elegido_num][dia_elegido_num]
          st.info(f"📂 Archivo vinculado: **{af_selected}**")
        else:
          af_selected = None

  with c1:
    st.subheader("Planilla de Horas (Carga)")
    st.success("☁️ Sincronizado con Google Sheets")
    try:
      with st.spinner("Sincronizando con Google Sheets..."):
        carga_bytes = obtener_excel_nube(URL_DESCARGA_SHEET)
        xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
        hojas_disponibles = xl_carga.sheet_names

      nh = st.selectbox(
          "Seleccionar Fecha / Pestaña:", options=hojas_disponibles
      )
    except Exception as e:
      st.error(f"No se pudo conectar a Google Sheets. Error: {e}")
      carga_bytes, nh = None, None

with t_reglas:
  st.subheader("⚙️ Configuración de Personal Excluido / No Facturable")
  df_reglas = pd.DataFrame([
      {"Legajo": leg, "Nombre": datos[0], "Sector / Área": datos[1]}
      for leg, datos in EXCLUIDOS_SECTOR.items()
  ])
  st.dataframe(df_reglas, use_container_width=True)

with t_reporte:
  if carga_bytes and af_selected:

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
      if st.button("🚀 Ejecutar Cruce Estricto por Fecha", type="primary"):
        st.session_state["mostrar_reporte"] = True

    with col_btn2:
      if st.button("🔄 Sincronizar Cambios de Google Sheets"):
        obtener_excel_nube.clear()
        st.session_state["mostrar_reporte"] = True
        if hasattr(st, "rerun"):
          st.rerun()
        else:
          st.experimental_rerun()

    if st.session_state["mostrar_reporte"]:
      try:
        ruta_fichada = os.path.join(BASE_FICHADAS, af_selected)
        d_pestana, m_pestana = obtener_dia_mes(nh)
        d_fichada, m_fichada = obtener_dia_mes(af_selected)

        dff_raw = pd.read_excel(ruta_fichada, sheet_name="Hoja2", header=None)
        if d_fichada is None:
          d_fichada, m_fichada = extraer_fecha_contenido(dff_raw)

        if d_pestana and d_fichada:
          if (d_pestana, m_pestana) != (d_fichada, m_fichada):
            st.error(
                f"⛔ **Error de Validación:** La Pestaña `{nh}` es del"
                f" **{d_pestana:02d}/{m_pestana:02d}**, pero la Fichada es del"
                f" **{d_fichada:02d}/{m_fichada:02d}**."
            )
            st.stop()

        dfc_raw = pd.read_excel(
            io.BytesIO(carga_bytes), sheet_name=nh, header=None
        )
        header_row_idx = 0
        for idx, row in dfc_raw.iterrows():
          row_str = " ".join(
              [str(val).upper() for val in row.values if pd.notna(val)]
          )
          if (
              "LEGAJO" in row_str
              or "OBRA" in row_str
              or "APELLIDO" in row_str
          ):
            header_row_idx = idx
            break

        dfc = pd.read_excel(
            io.BytesIO(carga_bytes), sheet_name=nh, skiprows=header_row_idx
        )

        legajo_col, nombre_col, obra_col = None, None, None
        for col in dfc.columns:
          col_str = str(col).strip().upper()
          if "LEGAJO" in col_str:
            legajo_col = col
          elif "APELLIDO" in col_str or "NOMBRE" in col_str:
            nombre_col = col
          elif "OBRA" in col_str:
            obra_col = col

        if legajo_col is None and len(dfc.columns) > 2:
          legajo_col = dfc.columns[2]
        if obra_col is None:
          obra_col = (
              dfc.columns[12] if len(dfc.columns) > 12 else dfc.columns[-1]
          )

        dff = pd.read_excel(ruta_fichada, sheet_name="Hoja2")

        keywords_novedades = {
            "COMPENSATORIO": "Compensatorio",
            "COMPENSATORIA": "Compensatorio",
            "ENFERMO": "Enfermo / Lic. Médica",
            "ENFERMEDAD": "Enfermo / Lic. Médica",
            "ENFEMRO": "Enfermo / Lic. Médica",
            "ENFEMEDAD": "Enfermo / Lic. Médica",
            "MEDICO": "Enfermo / Lic. Médica",
            "MÉDICO": "Enfermo / Lic. Médica",
            "SUSPENDIDO": "Suspendido",
            "SUSPENSION": "Suspendido",
            "SUSPENSIÓN": "Suspendido",
            "AUSENTE": "Ausente / Justificado",
            "AUSENCIA": "Ausente / Justificado",
            "VACACIONES": "Vacaciones",
            "VACACION": "Vacaciones",
            "SUR": "SUR / Franco Compensatorio",
            "LICENCIA": "Licencia / Permiso",
        }

        all_fichadas_list = []
        fd_by_legajo, fd_by_name = {}, {}

        for i in range(2, len(dff)):
          lv, nv = dff.iloc[i, 2], dff.iloc[i, 3]
          row_text = " ".join([
              str(val).upper() for val in dff.iloc[i].values if pd.notna(val)
          ])

          novedad_detectada = None
          for kw, label in keywords_novedades.items():
            if kw in row_text:
              novedad_detectada = label
              break

          clean_lv = limpiar_legajo(lv)
          nom_str = str(nv).strip() if pd.notna(nv) else ""
          val_nombre_norm = normalizar_texto(nom_str)

          info_dict = {
              "legajo": clean_lv,
              "nombre": nom_str,
              "nombre_norm": val_nombre_norm,
              "estado": novedad_detectada if novedad_detectada else "Presente",
              "es_novedad": novedad_detectada is not None,
          }

          if val_nombre_norm or clean_lv is not None:
            all_fichadas_list.append(info_dict)
          if clean_lv is not None:
            fd_by_legajo[clean_lv] = info_dict
          if val_nombre_norm:
            fd_by_name[val_nombre_norm] = info_dict

        def buscar_fichada(val_legajo, val_nombre):
          clean_l = limpiar_legajo(val_legajo)
          if clean_l is not None and clean_l in fd_by_legajo:
            return fd_by_legajo[clean_l], clean_l

          val_nombre_norm = normalizar_texto(val_nombre)
          if val_nombre_norm:
            if val_nombre_norm in fd_by_name:
              return (
                  fd_by_name[val_nombre_norm],
                  fd_by_name[val_nombre_norm]["legajo"],
              )

            tokens_c = set(val_nombre_norm.split())
            if len(tokens_c) >= 2:
              for n_key, item in fd_by_name.items():
                tokens_f = set(n_key.split())
                if len(tokens_c.intersection(tokens_f)) >= 2:
                  return item, item["legajo"]

          return None, clean_l

        fichadas_emparejadas, cargados_map = set(), {}
        for idx_c, row_c in dfc.iterrows():
          l_val = row_c[legajo_col] if legajo_col in dfc.columns else None
          n_val = (
              row_c[nombre_col]
              if nombre_col and nombre_col in dfc.columns
              else None
          )
          info_match, clean_l = buscar_fichada(l_val, n_val)
          o_val = (
              str(row_c[obra_col]).strip()
              if obra_col and pd.notna(row_c[obra_col])
              else "Sin Obra"
          )

          if clean_l is not None:
            cargados_map[clean_l] = o_val
          if info_match:
            if info_match["legajo"] is not None:
              cargados_map[info_match["legajo"]] = o_val
              fichadas_emparejadas.add(info_match["legajo"])
            if info_match["nombre_norm"]:
              cargados_map[info_match["nombre_norm"]] = o_val

        no_cargados, excluidos_list, novedades_list = [], [], []
        fichadas_procesadas = set()

        for info_f in all_fichadas_list:
          legajo_f = info_f["legajo"]
          nombre_norm_f = info_f["nombre_norm"]
          f_id = legajo_f if legajo_f is not None else nombre_norm_f

          if f_id in fichadas_procesadas or not f_id:
            continue
          fichadas_procesadas.add(f_id)

          nombre_fichada = info_f["nombre"]
          estado_fichada = info_f["estado"]
          es_nov = info_f["es_novedad"]

          es_cargado = (
              (legajo_f is not None and legajo_f in cargados_map)
              or (legajo_f is not None and legajo_f in fichadas_emparejadas)
              or (nombre_norm_f and nombre_norm_f in cargados_map)
          )

          if legajo_f is not None and legajo_f in EXCLUIDOS_SECTOR:
            excluidos_list.append({
                "LEGAJO": legajo_f,
                "Nombre y Apellido": (
                    nombre_fichada or EXCLUIDOS_SECTOR[legajo_f][0]
                ),
                "Sector": EXCLUIDOS_SECTOR[legajo_f][1],
                "Estado Fichada": estado_fichada,
                "Observación": "Sector no facturable",
            })
          elif not es_cargado:
            if es_nov:
              novedades_list.append({
                  "LEGAJO": legajo_f if legajo_f else "SIN LEGAJO",
                  "Nombre y Apellido": nombre_fichada,
                  "Novedad / Estado": estado_fichada,
                  "Observación": "Excluido por Novedad",
              })
            else:
              no_cargados.append({
                  "LEGAJO": legajo_f if legajo_f else "SIN LEGAJO",
                  "Nombre y Apellido (Fichada)": nombre_fichada,
                  "OBRA (Columna M)": "NO CARGADO EN PLANILLA",
                  "Estado": "Fichó en planta pero NO fue cargado",
              })
          else:
            if es_nov:
              novedades_list.append({
                  "LEGAJO": legajo_f if legajo_f else "SIN LEGAJO",
                  "Nombre y Apellido": nombre_fichada,
                  "Novedad / Estado": estado_fichada,
                  "Observación": "Cargado registrando novedad",
              })

        df_no_cargados = pd.DataFrame(no_cargados)
        df_excluidos = pd.DataFrame(excluidos_list)
        df_novedades = pd.DataFrame(novedades_list)

        v1, v2, v3, v4, v5 = st.tabs([
            "📊 Vista General",
            "⚠️ Faltan Cargar (Reales)",
            "🚫 Sectores No Facturables",
            "📋 Novedades / Licencias",
            "🟡 Ausentes / No Ficharon",
        ])

        with v2:
          st.subheader(
              "Personal que Fichó en Planta pero NO fue Cargado (Reales a"
              " Cargar)"
          )

          if not df_no_cargados.empty:
            st.warning(
                "⚠️ Se detectaron **"
                f"{len(df_no_cargados)} personas** que ingresaron a trabajar y"
                " **NO figuran cargadas**. Podes cargarles la novedad"
                " manualmente aquí abajo:"
            )

            df_edit = df_no_cargados.copy()
            df_edit["Asignar Novedad"] = "Seleccionar..."
            opciones_novedad = [
                "Seleccionar...",
                "COMPENSATORIO",
                "ENFERMO",
                "VACACIONES",
                "AUSENTE",
                "SUSPENDIDO",
                "SUR",
                "LICENCIA",
            ]

            edited_df = st.data_editor(
                df_edit,
                column_config={
                    "Asignar Novedad": st.column_config.SelectboxColumn(
                        "📝 Asignar Novedad Manual",
                        help=(
                            "Elegí la novedad para corregir el archivo"
                            " original de fichadas."
                        ),
                        width="medium",
                        options=opciones_novedad,
                    )
                },
                disabled=[
                    "LEGAJO",
                    "Nombre y Apellido (Fichada)",
                    "OBRA (Columna M)",
                    "Estado",
                ],
                use_container_width=True,
                key="editor_novedades_faltantes",
            )

            if st.button(
                "💾 Guardar Novedades Seleccionadas en la Fichada Base",
                type="primary",
            ):

              cambios = edited_df[
                  edited_df["Asignar Novedad"].notna()
                  & (edited_df["Asignar Novedad"] != "Seleccionar...")
                  & (
                      edited_df["Asignar Novedad"].astype(str).str.strip() != ""
                  )
              ]

              if not cambios.empty:
                try:
                  wb_fichada = load_workbook(ruta_fichada)
                  ws_fichada = wb_fichada["Hoja2"]

                  for index, row_edit in cambios.iterrows():
                    leg_buscar = limpiar_legajo(row_edit["LEGAJO"])
                    nom_buscar = str(
                        row_edit["Nombre y Apellido (Fichada)"]
                    ).strip()
                    nov_asignar = str(row_edit["Asignar Novedad"]).strip()

                    for r_idx in range(1, ws_fichada.max_row + 1):
                      cell_val = ws_fichada.cell(row=r_idx, column=3).value
                      cell_name = str(
                          ws_fichada.cell(row=r_idx, column=4).value or ""
                      ).strip()
                      cell_leg = limpiar_legajo(cell_val)

                      match_found = False
                      if cell_leg is not None and leg_buscar is not None:
                        if cell_leg == leg_buscar:
                          match_found = True
                      elif nom_buscar and nom_buscar.lower() in cell_name.lower():
                        match_found = True

                      if match_found:
                        ws_fichada.cell(row=r_idx, column=13).value = nov_asignar
                        if nov_asignar not in cell_name:
                          ws_fichada.cell(row=r_idx, column=4).value = (
                              f"{cell_name} - {nov_asignar}"
                          )
                        break

                  wb_fichada.save(ruta_fichada)

                  if hasattr(st, "rerun"):
                    st.rerun()
                  else:
                    st.experimental_rerun()

                except PermissionError:
                  st.error(
                      "⛔ **ERROR DE PERMISOS:** Tenés el archivo Excel de la"
                      " fichada abierto en otra ventana. ¡Cerralo y volvé a"
                      " apretar Guardar!"
                  )
                except Exception as e:
                  st.error(f"Error técnico al escribir en el Excel: {e}")
              else:
                st.info(
                    "💡 Por favor, elegí alguna novedad del menú desplegable"
                    " en la tabla antes de apretar Guardar."
                )
          else:
            st.success(
                "🎉 ¡Todo en orden! Todo el personal activo facturable que fichó"
                " en planta se encuentra cargado."
            )

        with v1:
          st.subheader("Control General de Planilla de Carga vs Fichadas")
          st.caption(
              "Verde = Presente | Amarillo = Ausente | Azul = Novedad | Verde"
              " Suave = Sector No Facturable"
          )

          def color_rows(row):
            try:
              l_val = row[legajo_col] if legajo_col in dfc.columns else None
              n_val = (
                  row[nombre_col]
                  if nombre_col and nombre_col in dfc.columns
                  else None
              )
              info_match, clean_l = buscar_fichada(l_val, n_val)
              if (
                  clean_l is not None and clean_l in EXCLUIDOS_SECTOR
              ) or (info_match and info_match["legajo"] in EXCLUIDOS_SECTOR):
                return ["background-color: #E2EFDA"] * len(row)
              if info_match:
                return (
                    ["background-color: #D9E1F2"] * len(row)
                    if info_match["es_novedad"]
                    else ["background-color: #C6EFCE"] * len(row)
                )
              return ["background-color: #FFEB9C"] * len(row)
            except Exception:
              return ["background-color: #FFC7CE"] * len(row)

          st.dataframe(
              dfc.style.apply(color_rows, axis=1), use_container_width=True
          )

        with v3:
          st.subheader("Personal de Sectores No Facturables")
          if not df_excluidos.empty:
            st.dataframe(df_excluidos, use_container_width=True)
          else:
            st.write(
                "No se registraron movimientos para personal de sectores"
                " excluidos."
            )

        with v4:
          st.subheader(
              "Novedades y Licencias (Compensatorios, Enfermos, Suspendidos,"
              " Ausentes)"
          )
          if not df_novedades.empty:
            st.dataframe(df_novedades, use_container_width=True)
          else:
            st.write("No se registraron personas con estado de novedad o licencia.")

        with v5:
          st.subheader("Personal Cargado en Planilla que NO Registró Fichada")
          ausentes_list = []
          for idx_c, row_c in dfc.iterrows():
            l_val = row_c[legajo_col] if legajo_col in dfc.columns else None
            n_val = (
                row_c[nombre_col]
                if nombre_col and nombre_col in dfc.columns
                else None
            )
            info_match, clean_l = buscar_fichada(l_val, n_val)
            if not info_match:
              ausentes_list.append(row_c)
          df_ausentes = (
              pd.DataFrame(ausentes_list) if ausentes_list else pd.DataFrame()
          )

          if not df_ausentes.empty:
            st.info(
                f"Se encontraron **{len(df_ausentes)} registros** cargados sin"
                " fichada:"
            )
            st.dataframe(df_ausentes, use_container_width=True)
          else:
            st.success("🎉 Todo el personal cargado registró su fichada.")

        wb = Workbook()
        ws1 = wb.active
        ws1.title = f"Control {nh}"

        gf = PatternFill(
            start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
        )
        yf = PatternFill(
            start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
        )
        rf = PatternFill(
            start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
        )
        bf = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        exf = PatternFill(
            start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
        )

        for r_idx, row in enumerate(
            dataframe_to_rows(dfc, index=False, header=True), 1
        ):
          ws1.append(row)
          if r_idx == 1:
            continue
          try:
            l_val = (
                row[list(dfc.columns).index(legajo_col)]
                if legajo_col in dfc.columns
                else None
            )
            n_val = (
                row[list(dfc.columns).index(nombre_col)]
                if nombre_col and nombre_col in dfc.columns
                else None
            )
            info_match, clean_l = buscar_fichada(l_val, n_val)

            if (
                clean_l is not None and clean_l in EXCLUIDOS_SECTOR
            ) or (info_match and info_match["legajo"] in EXCLUIDOS_SECTOR):
              fc = exf
            elif info_match:
              fc = bf if info_match["es_novedad"] else gf
            else:
              fc = yf
            for c in range(1, len(row) + 1):
              ws1.cell(row=r_idx, column=c).fill = fc
          except Exception:
            for c in range(1, len(row) + 1):
              ws1.cell(row=r_idx, column=c).fill = rf

        if not df_no_cargados.empty:
          ws2 = wb.create_sheet(title="Fichados Sin Carga")
          orange_fill = PatternFill(
              start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
          )
          for r_idx, row in enumerate(
              dataframe_to_rows(df_no_cargados, index=False, header=True), 1
          ):
            ws2.append(row)
            if r_idx > 1:
              for c in range(1, len(row) + 1):
                ws2.cell(row=r_idx, column=c).fill = orange_fill

        if not df_excluidos.empty:
          ws3 = wb.create_sheet(title="Sectores No Facturables")
          ex_fill = PatternFill(
              start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
          )
          for r_idx, row in enumerate(
              dataframe_to_rows(df_excluidos, index=False, header=True), 1
          ):
            ws3.append(row)
            if r_idx > 1:
              for c in range(1, len(row) + 1):
                ws3.cell(row=r_idx, column=c).fill = ex_fill

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        st.divider()
        st.success("¡Cruce completado y procesado con éxito!")
        st.download_button(
            "📥 Descargar Excel con Formato",
            data=out,
            file_name=f"Control_{nh}.xlsx",
        )

      except Exception as e:
        st.error(f"Error al procesar: {e}")
  else:
    st.info(
        "Por favor, seleccione una fecha de carga en el desplegable de arriba y"
        " elija un archivo de fichada antes de cruzar."
    )

# --- PESTAÑA: 📋 CIERRE DE PERÍODO (CONECTADO A APPS SCRIPT WEB APP) ---
with t_cierre:
  st.header("📋 Notificación y Cierre de Período (Gestión de OTs)")
  st.caption(
      "Consolidado de OTs únicas en el rango de fechas seleccionado para cargar"
      " N° de Solicitud y Responsable del Trabajo."
  )

  if carga_bytes:
    xl_carga_cierre = pd.ExcelFile(io.BytesIO(carga_bytes))

    dias_disponibles_map = {}
    for sheet in xl_carga_cierre.sheet_names:
      match = re.search(r"(\d{1,2})[-/\.](\d{1,2})", sheet)
      if match:
        dia_num = int(match.group(1))
        dias_disponibles_map[dia_num] = sheet

    dias_ordenados = sorted(dias_disponibles_map.keys())

    if dias_ordenados:
      col_f1, col_f2 = st.columns(2)
      with col_f1:
        desde_dia_sel = st.selectbox(
            "📅 Seleccionar Fecha Desde:",
            options=dias_ordenados,
            format_func=lambda d: (
                f"Desde el Día {d:02d} (Pestaña {dias_disponibles_map[d]})"
            ),
            index=0,
        )

      with col_f2:
        dias_hasta_opciones = [d for d in dias_ordenados if d >= desde_dia_sel]
        hasta_dia_sel = st.selectbox(
            "📅 Seleccionar Fecha Hasta:",
            options=dias_hasta_opciones,
            format_func=lambda d: (
                f"Hasta el Día {d:02d} (Pestaña {dias_disponibles_map[d]})"
            ),
            index=len(dias_hasta_opciones) - 1 if dias_hasta_opciones else 0,
        )

      ots_dict = {}
      for d in dias_ordenados:
        if d < desde_dia_sel or d > hasta_dia_sel:
          continue

        s_name = dias_disponibles_map[d]
        dfc_raw = pd.read_excel(
            io.BytesIO(carga_bytes), sheet_name=s_name, header=None
        )
        header_row_idx = 0
        for idx, row in dfc_raw.iterrows():
          row_str = " ".join(
              [str(v).upper() for v in row.values if pd.notna(v)]
          )
          if (
              "LEGAJO" in row_str
              or "OBRA" in row_str
              or "APELLIDO" in row_str
              or "OT" in row_str
          ):
            header_row_idx = idx
            break

        dfc_ot = pd.read_excel(
            io.BytesIO(carga_bytes), sheet_name=s_name, skiprows=header_row_idx
        )

        ot_col = next(
            (c for c in dfc_ot.columns if "OT" in str(c).upper()), None
        )
        sol_col = next(
            (c for c in dfc_ot.columns if "SOLICITUD" in str(c).upper()), None
        )
        resp_col = next(
            (c for c in dfc_ot.columns if "RESPONSABLE" in str(c).upper()), None
        )

        if ot_col:
          for idx_r, row_r in dfc_ot.iterrows():
            ot_val = row_r[ot_col]
            if pd.isna(ot_val) or str(ot_val).strip() == "":
              continue

            try:
              ot_str = str(int(float(str(ot_val).strip())))
            except Exception:
              ot_str = str(ot_val).strip()

            sol_val = (
                str(row_r[sol_col]).strip()
                if sol_col
                and pd.notna(row_r[sol_col])
                and str(row_r[sol_col]).lower() != "nan"
                else ""
            )
            resp_val = (
                str(row_r[resp_col]).strip()
                if resp_col
                and pd.notna(row_r[resp_col])
                and str(row_r[resp_col]).lower() != "nan"
                else ""
            )

            if ot_str not in ots_dict:
              ots_dict[ot_str] = {
                  "OT Programada": ot_str,
                  "N° Solicitud": sol_val,
                  "Responsable del Trabajo": resp_val,
                  "Días Presente": [f"Día {d:02d}"],
                  "Apariciones": 1,
              }
            else:
              ots_dict[ot_str]["Apariciones"] += 1
              d_label = f"Día {d:02d}"
              if d_label not in ots_dict[ot_str]["Días Presente"]:
                ots_dict[ot_str]["Días Presente"].append(d_label)
              if not ots_dict[ot_str]["N° Solicitud"] and sol_val:
                ots_dict[ot_str]["N° Solicitud"] = sol_val
              if not ots_dict[ot_str]["Responsable del Trabajo"] and resp_val:
                ots_dict[ot_str]["Responsable del Trabajo"] = resp_val

      df_ots_unicas = pd.DataFrame(list(ots_dict.values()))

      if not df_ots_unicas.empty:
        df_ots_unicas["Días Presente"] = df_ots_unicas["Días Presente"].apply(
            lambda x: ", ".join(x)
        )

        k_c1, k_c2, k_c3 = st.columns(3)
        k_c1.metric("📌 Total OTs Únicas", len(df_ots_unicas))
        k_c2.metric(
            "✅ OTs con Solicitud Informada",
            len(
                df_ots_unicas[
                    df_ots_unicas["N° Solicitud"].astype(str).str.strip() != ""
                ]
            ),
        )
        k_c3.metric(
            "⏳ OTs Pendientes de Cierre",
            len(
                df_ots_unicas[
                    df_ots_unicas["N° Solicitud"].astype(str).str.strip() == ""
                ]
            ),
        )

        st.divider()
        st.write(
            "📝 **Completá, borrá o pegá datos en N° de Solicitud y Responsable del Trabajo**"
            " usando copiar/pegar (`Ctrl+C` / `Ctrl+V`) o suprimir (`Supr`):"
        )

        edited_ots_df = st.data_editor(
            df_ots_unicas,
            column_config={
                "OT Programada": st.column_config.TextColumn(
                    "OT Programada", disabled=True
                ),
                "N° Solicitud": st.column_config.TextColumn(
                    "N° Solicitud",
                    help="Ingresá o pegá el número de solicitud asignado.",
                ),
                "Responsable del Trabajo": st.column_config.TextColumn(
                    "Responsable del Trabajo",
                    help="Nombre y apellido del responsable.",
                ),
                "Días Presente": st.column_config.TextColumn(
                    "Días en Planilla", disabled=True
                ),
                "Apariciones": st.column_config.NumberColumn(
                    "Líneas Cargadas", disabled=True
                ),
            },
            use_container_width=True,
            key="editor_ots_cierre",
        )

        if st.button(
            "💾 Guardar Cierre de Período y Sincronizar Google Sheets", type="primary"
        ):
          try:
            cambios_ots = {}
            for _, r_ot in edited_ots_df.iterrows():
              ot_k = str(r_ot["OT Programada"]).strip()

              sol_v = (
                  str(r_ot["N° Solicitud"]).strip()
                  if pd.notna(r_ot["N° Solicitud"])
                  and str(r_ot["N° Solicitud"]).lower() != "nan"
                  else ""
              )
              resp_v = (
                  str(r_ot["Responsable del Trabajo"]).strip()
                  if pd.notna(r_ot["Responsable del Trabajo"])
                  and str(r_ot["Responsable del Trabajo"]).lower() != "nan"
                  else ""
              )

              cambios_ots[ot_k] = {"Solicitud": sol_v, "Responsable": resp_v}

            payload = {
                "desdeDia": desde_dia_sel,
                "hastaDia": hasta_dia_sel,
                "cambios": cambios_ots,
            }

            with st.spinner("Sincronizando directamente con Google Sheets..."):
              res = requests.post(WEBAPP_URL, json=payload)

            if res.status_code == 200:
              obtener_excel_nube.clear()
              st.success(
                  "🟢 ¡Sincronizado con éxito directamente en la planilla de"
                  " Google Sheets!"
              )
            else:
              st.error(
                  f"Error al conectar con Google Sheets: STATUS {res.status_code}"
              )

          except Exception as e_cierre:
            st.error(f"Error al enviar datos: {e_cierre}")
      else:
        st.warning(
            "No se encontraron OTs cargadas en el rango de fechas seleccionado."
        )
    else:
      st.info("No se detectaron días cargados en la planilla.")
  else:
    st.info("Conectando con Google Sheets para cargar datos...")

# --- PESTAÑA ⏱️ 4. HORAS Y QUINCENAS ---
with t_quincena:
  st.header("⏱️ Recuento y Validación Quincenal de Horas")
  st.caption(
      "Consolidado de Horas Normales, 50% y 100% cargadas en Google Sheets"
      " vs. Horas Reales Fichadas en Reloj"
  )

  if carga_bytes:
    col_q1, col_q2 = st.columns([2, 1])
    with col_q1:
      quincena_sel = st.radio(
          "Seleccionar Período de Análisis:",
          options=[
              "1ª Quincena (Días 1 al 15)",
              "2ª Quincena (Días 16 al 31)",
              "Todo el Mes",
          ],
          horizontal=True,
      )

    with col_q2:
      st.markdown("### ")
      btn_calc_q = st.button(
          "📊 Calcular Resumen Quincenal", type="primary"
      )

    if btn_calc_q or st.session_state.get("ver_quincena", False):
      st.session_state["ver_quincena"] = True

      try:
        xl_carga = pd.ExcelFile(io.BytesIO(carga_bytes))
        registros_quincena = []

        for sheet in xl_carga.sheet_names:
          match = re.search(r"(\d{1,2})[-/\.](\d{1,2})", sheet)
          if not match:
            continue
          dia, mes = int(match.group(1)), int(match.group(2))

          if quincena_sel == "1ª Quincena (Días 1 al 15)" and dia > 15:
            continue
          elif quincena_sel == "2ª Quincena (Días 16 al 31)" and dia <= 15:
            continue

          dfc_raw = pd.read_excel(
              io.BytesIO(carga_bytes), sheet_name=sheet, header=None
          )
          header_row_idx = 0
          for idx, row in dfc_raw.iterrows():
            row_str = " ".join(
                [str(v).upper() for v in row.values if pd.notna(v)]
            )
            if (
                "LEGAJO" in row_str
                or "OBRA" in row_str
                or "APELLIDO" in row_str
            ):
              header_row_idx = idx
              break

          dfc_q = pd.read_excel(
              io.BytesIO(carga_bytes), sheet_name=sheet, skiprows=header_row_idx
          )

          legajo_col = next(
              (c for c in dfc_q.columns if "LEGAJO" in str(c).upper()), None
          )
          nombre_col = next(
              (
                  c
                  for c in dfc_q.columns
                  if "APELLIDO" in str(c).upper()
                  or "NOMBRE" in str(c).upper()
                  or "PERSONAL" in str(c).upper()
              ),
              None,
          )
          trabajo_col = next(
              (c for c in dfc_q.columns if "TRABAJO" in str(c).upper()), None
          )
          tipo_col = next(
              (
                  c
                  for c in dfc_q.columns
                  if "TIPO DE HORA" in str(c).upper()
                  or "TIPO" in str(c).upper()
              ),
              None,
          )

          if legajo_col and trabajo_col and tipo_col:
            for idx_row, row in dfc_q.iterrows():
              leg = row[legajo_col]
              leg_clean = limpiar_legajo(leg)

              nom = (
                  str(row[nombre_col]).strip()
                  if nombre_col and pd.notna(row[nombre_col])
                  else (
                      f"Legajo {leg_clean}"
                      if leg_clean
                      else str(leg or "Sin Nombre")
                  )
              )
              hs = extraer_horas(row[trabajo_col])
              tipo_raw = (
                  str(row[tipo_col]).strip().upper()
                  if pd.notna(row[tipo_col])
                  else "HS. NORMALES"
              )

              if "0.5" in tipo_raw or "50" in tipo_raw:
                tipo_cat = "Hs. 50%"
              elif "1" in tipo_raw or "100" in tipo_raw:
                tipo_cat = "Hs. 100%"
              else:
                tipo_cat = "Hs. Normales"

              if leg_clean is not None or nom != "Sin Nombre":
                registros_quincena.append({
                    "Legajo": leg_clean if leg_clean else 0,
                    "Nombre": nom,
                    "Horas": hs,
                    "Tipo": tipo_cat,
                })

        df_q_all = pd.DataFrame(registros_quincena)

        if not df_q_all.empty:
          piv_q = df_q_all.pivot_table(
              index=["Legajo", "Nombre"],
              columns="Tipo",
              values="Horas",
              aggfunc="sum",
              fill_value=0.0,
          ).reset_index()

          for col_h in ["Hs. Normales", "Hs. 50%", "Hs. 100%"]:
            if col_h not in piv_q.columns:
              piv_q[col_h] = 0.0

          piv_q["Total Hs. Cargadas"] = (
              piv_q["Hs. Normales"] + piv_q["Hs. 50%"] + piv_q["Hs. 100%"]
          )

          fichadas_horas_dict = {}
          archivos_locales = (
              [
                  f
                  for f in os.listdir(BASE_FICHADAS)
                  if f.endswith((".xlsx", ".xltx"))
              ]
              if os.path.exists(BASE_FICHADAS)
              else []
          )

          for f_loc in archivos_locales:
            match_f = re.search(r"Fichada_(\d{1,2})-(\d{1,2})", f_loc)
            if match_f:
              d_f, m_f = int(match_f.group(1)), int(match_f.group(2))
              if quincena_sel == "1ª Quincena (Días 1 al 15)" and d_f > 15:
                continue
              elif quincena_sel == "2ª Quincena (Días 16 al 31)" and d_f <= 15:
                continue

              try:
                dff_loc = pd.read_excel(
                    os.path.join(BASE_FICHADAS, f_loc), sheet_name="Hoja2"
                )
                for idx_fl in range(2, len(dff_loc)):
                  l_f = dff_loc.iloc[idx_fl, 2]
                  ing_f = dff_loc.iloc[idx_fl, 4]
                  sal_f = dff_loc.iloc[idx_fl, 5]
                  l_f_clean = limpiar_legajo(l_f)
                  if l_f_clean is not None:
                    h_trab = calcular_horas_fichada(ing_f, sal_f)
                    fichadas_horas_dict[l_f_clean] = (
                        fichadas_horas_dict.get(l_f_clean, 0.0) + h_trab
                    )
              except Exception:
                pass

          df_fich_tot = pd.DataFrame(
              list(fichadas_horas_dict.items()),
              columns=["Legajo", "Total Hs. Fichadas (Reloj)"],
          )

          df_q_final = pd.merge(piv_q, df_fich_tot, on="Legajo", how="left")
          df_q_final["Total Hs. Fichadas (Reloj)"] = df_q_final[
              "Total Hs. Fichadas (Reloj)"
          ].fillna(0.0)
          df_q_final["Diferencia (Cargado vs Fichado)"] = (
              df_q_final["Total Hs. Cargadas"]
              - df_q_final["Total Hs. Fichadas (Reloj)"]
          )

          kpi1, kpi2, kpi3, kpi4 = st.columns(4)
          kpi1.metric("📊 Total Hs. Normales", f"{df_q_final['Hs. Normales'].sum():.1f} hs")
          kpi2.metric("⚡ Total Hs. 50%", f"{df_q_final['Hs. 50%'].sum():.1f} hs")
          kpi3.metric("🔥 Total Hs. 100%", f"{df_q_final['Hs. 100%'].sum():.1f} hs")
          kpi4.metric("⏱️ Total Horas Cargadas", f"{df_q_final['Total Hs. Cargadas'].sum():.1f} hs")

          st.divider()

          cols_orden = [
              "Legajo",
              "Nombre",
              "Hs. Normales",
              "Hs. 50%",
              "Hs. 100%",
              "Total Hs. Cargadas",
              "Total Hs. Fichadas (Reloj)",
              "Diferencia (Cargado vs Fichado)",
          ]
          df_q_display = df_q_final[cols_orden]

          def highlight_diff(val):
            if val > 0:
              return "background-color: #FFEB9C; color: #856404;"
            elif val < 0:
              return "background-color: #F8D7DA; color: #721C24;"
            return "background-color: #C6EFCE; color: #155724;"

          st.dataframe(
              df_q_display.style.applymap(
                  highlight_diff, subset=["Diferencia (Cargado vs Fichado)"]
              ),
              use_container_width=True,
          )

          out_q = io.BytesIO()
          with pd.ExcelWriter(out_q, engine="openpyxl") as writer:
            df_q_display.to_excel(
                writer, index=False, sheet_name="Resumen Quincenal"
            )
          out_q.seek(0)

          st.download_button(
              "📥 Descargar Resumen Quincenal (Excel)",
              data=out_q,
              file_name=f"Resumen_Horas_{quincena_sel.replace(' ', '_')}.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          )
        else:
          st.warning("No se encontraron registros de horas cargadas para la quincena seleccionada.")
      except Exception as e_q:
        st.error(f"Error al calcular la quincena: {e_q}")
  else:
    st.info("Conectando con Google Sheets para cargar datos...")
