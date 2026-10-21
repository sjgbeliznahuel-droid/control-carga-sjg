import streamlit as st
import pandas as pd
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Configuración principal de la página
st.set_page_config(page_title="Control de Carga SJG", layout="wide")
st.title("Sistema de Control de Fichadas - SJG Montajes Industriales")

# Función para conectar a Drive sin sobrecargar la app
@st.cache_resource
def conectar_drive():
    credenciales_dict = json.loads(st.secrets["credenciales"])
    creds = service_account.Credentials.from_service_account_info(
        credenciales_dict, scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds)

try:
    drive_service = conectar_drive()
    
    # -------------------------------------------------------------------
    # ¡IMPORTANTE!: Pegá acá el código de tu carpeta de Google Drive
    # -------------------------------------------------------------------
    folder_id = 'ACA_VA_EL_ID_DE_TU_CARPETA' 
    
    query = f"'{folder_id}' in parents and trashed=false"
    resultados = drive_service.files().list(q=query, orderBy="createdTime desc", fields="files(id, name)").execute()
    archivos = resultados.get('files', [])

    if archivos:
        ultimo_archivo = archivos[0]
        st.info(f"Descargando archivo más reciente: {ultimo_archivo['name']}...")
        
        # Descarga interna del archivo a la memoria
        request = drive_service.files().get_media(fileId=ultimo_archivo['id'])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        terminado = False
        while not terminado:
            _, terminado = downloader.next_chunk()
        fh.seek(0)
        
        # Lectura de los datos
        if ultimo_archivo['name'].endswith('.csv'):
            df_fichadas = pd.read_csv(fh)
        else:
            df_fichadas = pd.read_excel(fh)
            
        st.success("¡Datos cargados correctamente desde Google Drive!")
        
        # Vista previa básica de lo que descargó
        st.subheader("Vista previa de los datos importados")
        st.dataframe(df_fichadas)
        
        # -------------------------------------------------------------------
        # ACÁ ABAJO PEGÁ LA PARTE DE TU CÓDIGO VIEJO QUE CRUZABA LOS DATOS
        # La tabla ya se llama "df_fichadas", así que el resto de tu
        # programación va a funcionar directamente.
        # -------------------------------------------------------------------
        
    else:
        st.warning("No se encontró ningún archivo en la carpeta 'fichadas' de Drive.")

except Exception as e:
    st.error(f"Hubo un error al intentar conectarse con Google Drive: {e}")
