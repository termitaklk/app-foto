import logging
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
from httplib2 import Http
from googleapiclient.errors import HttpError
import logging
import os
import time

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)

# Alcances de la API de Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Autenticación de Google Sheets
def authenticate_sheets():
    """Autentica el acceso a la API de Google Sheets."""
    try:
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds
    except Exception as e:
        logging.error(f"Error autenticando con Google Sheets: {e}")
        return None
    
# Nueva función para obtener nombres de hojas
def get_sheet_names(sheet_id):
    """Obtiene los nombres de todas las hojas en la hoja de cálculo."""
    try:
        creds = authenticate_sheets()
        service = build("sheets", "v4", credentials=creds)
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = sheet_metadata.get("sheets", [])
        return [sheet.get("properties", {}).get("title", "") for sheet in sheets]
    except Exception as e:
        raise Exception(f"Error obteniendo nombres de hojas: {e}")

def read_configuration(sheet_id):
    """
    Lee la configuración desde la hoja 'configuracion' de Google Sheets.
    Devuelve base_path, min_files, start_row, create_interval, validate_interval, y valores Día/Noche.
    """
    try:
        creds = authenticate_sheets()
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        # Leer valores de configuración principales desde 'configuracion!A2:F2'
        range_name = "configuracion!A2:F2"
        result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get("values", [])

         # Leer la ruta adicional desde 'configuracion!A4'
        range_base_night = "configuracion!A4"
        result_base_night = sheet.values().get(spreadsheetId=sheet_id, range=range_base_night).execute()
        base_night_values = result_base_night.get("values", [])

        base_night = base_night_values[0][0] if base_night_values and len(base_night_values[0]) > 0 else None


        if not values:
            raise ValueError("No se encontraron valores en la hoja de configuración.")

        # Asignar valores desde la hoja
        base_path = values[0][0] if len(values[0]) > 0 else None
        min_files = int(values[0][1]) if len(values[0]) > 1 and values[0][1].isdigit() else 10
        start_row = int(values[0][2]) if len(values[0]) > 2 and values[0][2].isdigit() else 2
        create_interval_minutes = float(values[0][4]) if len(values[0]) > 4 else 1  # En minutos
        validate_interval_minutes = float(values[0][5]) if len(values[0]) > 5 else 2  # En minutos

        # Convertir los intervalos de minutos a milisegundos
        create_interval_ms = int(create_interval_minutes * 60 * 1000)
        validate_interval_ms = int(validate_interval_minutes * 60 * 1000)

        # Leer valores de Día y Noche desde las columnas H e I (H1:H2 e I1:I2)
        day_night_range = "configuracion!H1:I2"
        day_night_result = sheet.values().get(spreadsheetId=sheet_id, range=day_night_range).execute()
        day_night_values = day_night_result.get("values", [])

        if len(day_night_values) < 2:
            raise ValueError("No se encontraron suficientes valores para 'Día' y 'Noche'.")

        # Extraer valores Día y Noche
        day_label = day_night_values[0][0]  # Encabezado en H1
        night_label = day_night_values[0][1]  # Encabezado en I1
        day_value = day_night_values[1][0]  # Valor en H2
        night_value = day_night_values[1][1]  # Valor en I2

        # Crear un diccionario con los valores de Día y Noche
        day_night_config = {day_label: day_value, night_label: night_value}

        return base_path, base_night, min_files, start_row, create_interval_ms, validate_interval_ms, day_night_config

    except Exception as e:
        raise Exception(f"Error leyendo la configuración: {e}")
    
def create_new_sheet(sheet_id, sheet_name):
    """
    Crea una nueva hoja en el archivo de Google Sheets con los encabezados especificados.

    :param sheet_id: ID del archivo de Google Sheets.
    :param sheet_name: Nombre de la nueva hoja a crear.
    """
    try:
        creds = authenticate_sheets()
        service = build("sheets", "v4", credentials=creds)

        # Crear una nueva hoja
        request_body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name
                        }
                    }
                }
            ]
        }
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body=request_body
        ).execute()

        # Añadir los encabezados a la nueva hoja
        headers = [["ID", "LAST NAME", "PAX", "EMAIL", "ATRIBUTO", "VENDIDO", "REQUISITO", "IDIOMA", "ENVIADO", "PERIODO", "SUNSHINE CRUISE", "CORTESÍA"]]
        range_name = f"'{sheet_name}'!A1:L1"
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=range_name,
            valueInputOption="RAW",
            body={"values": headers},
        ).execute()

        logging.info(f"Hoja '{sheet_name}' creada con encabezados.")
        return response

    except Exception as e:
        logging.error(f"Error creando la hoja '{sheet_name}': {e}")
from datetime import datetime

def check_and_create_today_sheet(sheet_id):
    """
    Comprueba si existe una hoja con el nombre de la fecha actual. Si no existe, la crea.
    
    :param sheet_id: ID del archivo de Google Sheets.
    """
    try:
        # Obtener la lista de hojas existentes
        creds = authenticate_sheets()
        service = build("sheets", "v4", credentials=creds)
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheet_names = [sheet["properties"]["title"] for sheet in sheet_metadata.get("sheets", [])]

        # Obtener la fecha actual en formato día-mes-año
        today_date = datetime.now().strftime("%d-%m-%Y")

        # Verificar si ya existe una hoja con este nombre
        if today_date in sheet_names:
            logging.info(f"La hoja con la fecha '{today_date}' ya existe.")
            return today_date  # Devuelve el nombre de la hoja existente
        else:
            # Crear una nueva hoja con la fecha actual
            logging.info(f"No existe una hoja con la fecha '{today_date}'. Creando...")
            create_new_sheet(sheet_id, today_date)
            return today_date  # Devuelve el nombre de la hoja recién creada

    except Exception as e:
        logging.error(f"Error comprobando o creando la hoja de la fecha actual: {e}")
        return None

def update_google_sheet(sheet_id, range_name, values, sheet_name):
    """Actualiza valores en la hoja seleccionada en Google Sheets."""
    try:
        creds = authenticate_sheets()
        service = build('sheets', 'v4', credentials=creds)
        body = {"values": values}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!{range_name}",
            valueInputOption="RAW",
            body=body
        ).execute()
        return result
    except Exception as e:
        raise Exception(f"Error actualizando Google Sheets: {e}")

# Leer datos de una hoja de cálculo
def read_google_sheet(sheet_id, range_name, max_retries=5, timeout=60):
    """Lee datos desde Google Sheets con reintentos en caso de error."""
    creds = authenticate_sheets()
    if creds is None:
        logging.error("Credenciales no disponibles. No se puede autenticar.")
        return []

    # Construir el servicio de Google Sheets usando credenciales directamente
    service = build("sheets", "v4", credentials=creds)

    for attempt in range(max_retries):
        try:
            logging.info(f"Intentando leer datos (intento {attempt + 1}/{max_retries})...")
            result = service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            return result.get("values", [])
        except HttpError as e:
            logging.error(f"Error HTTP al leer datos: {e}")
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt  # Retroceso exponencial
                logging.info(f"Reintentando en {wait_time} segundos...")
                time.sleep(wait_time)
            else:
                logging.error("Máximo número de reintentos alcanzado.")
                return []
        except Exception as e:
            logging.error(f"Error inesperado: {e}")
            return []
        
# Escribir datos en una hoja de cálculo
def write_google_sheet(sheet_id, range_name, values, sheet_name):
    """Actualiza valores en la hoja seleccionada en Google Sheets."""
    try:
        creds = authenticate_sheets()
        service = build('sheets', 'v4', credentials=creds)
        body = {"values": values}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!{range_name}",
            valueInputOption="RAW",
            body=body
        ).execute()
        return result
    except Exception as e:
        raise Exception(f"Error actualizando Google Sheets: {e}")
