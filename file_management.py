import os
import logging
import fnmatch
from google_sheets_utils import read_google_sheet, write_google_sheet
from google_sheets_utils import read_google_sheet, write_google_sheet, read_configuration, update_google_sheet, get_sheet_names
from tkinter import messagebox
import re
import win32net
import win32netcon
from smb.SMBConnection import SMBConnection
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()
omit_c2_values = ["MADRE E HIJO", "MADRE E HIJA", "PADRE E HIJO", "PADRE E HIJA", "PAREJA"]
# Configuración global
SHEET_ID = os.getenv("SHEET_ID")
START_ROW = 2  # La fila donde comienzan los datos en la hoja de cálculo

def construct_client_folder_path(base_path, sheet_name, id_value, last_name, cantidad, atributo, ):
    """
    Construye la ruta completa a la carpeta del cliente utilizando la misma lógica de creación de carpetas.
    """
    try:
        omit_c2_values = ["MADRE E HIJO", "MADRE E HIJA", "PADRE E HIJO", "PADRE E HIJA", "PAREJA"]
        # Ruta base para la fecha
        date_folder_path = os.path.join(base_path, sheet_name)

        # Crear el nombre de la carpeta siguiendo la lógica de `create_folders`
        if atributo in (omit_c2_values or []):
            folder_name = f"{atributo}-{last_name}-{id_value}".upper()
        elif atributo == "FAMILIA":
            folder_name = f"{atributo}-D{cantidad}-{last_name}-{id_value}".upper()
        else:
            folder_name = f"{cantidad} {atributo}-{last_name}-{id_value}".upper()

        # Ruta completa a la carpeta
        folder_path = os.path.join(date_folder_path, folder_name)

        # Verificar si la carpeta existe
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"La ruta del cliente no existe: {folder_path}")

        return folder_path

    except Exception as e:
        logging.error(f"Error construyendo la ruta de la carpeta: {e}")
        return None

def establish_smb_connection(full_path):
    """Establece la conexión SMB y descompone la ruta."""
    try:
        username = os.getenv("SMB_USERNAME")
        password = os.getenv("SMB_PASSWORD")

        if not username or not password:
            raise ValueError("Usuario o contraseña no encontrados en las variables de entorno")

        if not full_path.startswith("\\\\"):
            raise ValueError("La ruta completa debe comenzar con '\\\\'")

        path_parts = full_path[2:].split("\\", 2)
        server = path_parts[0]
        share = path_parts[1]
        sub_path = path_parts[2] if len(path_parts) > 2 else ""

        conn = SMBConnection(username, password, "client", server, use_ntlm_v2=True)
        conn.connect(server, 445)
        return conn, share, sub_path
    except Exception as e:
        logging.error(f"Error al establecer la conexión SMB: {e}")
        raise

def smb_create_directory_if_not_exists(conn, share, path):
    """Crea un directorio en SMB si no existe."""
    try:
        # Verificar si el directorio ya existe
        parent_path, folder_name = os.path.split(path)
        existing_dirs = [
            f.filename for f in conn.listPath(share, parent_path)
            if f.isDirectory and f.filename not in [".", ".."]
        ]
        if folder_name in existing_dirs:
            return

        # Crear el directorio
        conn.createDirectory(share, path)
    except Exception as e:
        logging.error(f"Error al crear el directorio '{path}': {e}")
        raise

def ensure_full_path_exists(conn, share, full_path):
    """Asegura que toda la jerarquía de carpetas exista."""
    try:
        logging.info(f"Validando la existencia de la ruta completa: {full_path}")
        parts = full_path.split("\\")
        current_path = ""
        for part in parts:
            if not part.strip():  # Evitar rutas vacías
                continue
            current_path = os.path.join(current_path, part) if current_path else part
            smb_create_directory_if_not_exists(conn, share, current_path)
    except ZeroDivisionError as e:
        logging.error(f"Error durante la validación de carpetas: división por cero. Ruta: {full_path}", exc_info=e)
        raise
    except Exception as e:
        logging.error(f"Error asegurando la ruta completa '{full_path}': {e}")
        raise

def create_folders(sheet_id, sheet_name, base_path, start_row=2, omit_c2_values=None, update_progress=None):
    """Crea carpetas basadas en los datos de Google Sheets en una ruta SMB, evitando duplicados y controlando celdas vacías."""
    omit_c2_values = ["MADRE E HIJO", "MADRE E HIJA", "PADRE E HIJO", "PADRE E HIJA", "PAREJA"]
    try:
        # Establecer conexión SMB
        smb_conn, share, sub_path = establish_smb_connection(base_path)

        # Validar y crear la ruta base completa
        date_folder_path = os.path.join(sub_path, sheet_name)
        ensure_full_path_exists(smb_conn, share, date_folder_path)

        # Leer datos de la hoja
        sheet_values = read_google_sheet(sheet_id, f"'{sheet_name}'!A{start_row}:M")  # Hasta columna M
        if not sheet_values:
            logging.warning("No hay datos para procesar en la hoja.")
            if update_progress:
                update_progress(100, 100, "No hay datos para procesar.")
            return

        # Listar carpetas existentes en la ruta
        try:
            existing_folders = [
                f.filename for f in smb_conn.listPath(share, date_folder_path)
                if f.isDirectory and f.filename not in [".", ".."]
            ]
            logging.info(f"Carpetas existentes indexadas: {existing_folders}")
        except Exception as e:
            logging.warning(f"No se pudo listar el directorio o está vacío: {date_folder_path}")
            existing_folders = []

        total_rows = len(sheet_values)

        # Manejo de casos sin filas
        if total_rows == 0:
            logging.warning("No hay filas para procesar en la hoja.")
            if update_progress:
                update_progress(100, 100, "No hay filas para procesar.")
            return

        # Inicializar barra de progreso
        if update_progress:
            update_progress(0, total_rows, "Iniciando creación de carpetas...")

        for idx, row in enumerate(sheet_values, start=1):
            if len(row) < 3:
                logging.warning(f"Fila incompleta en la posición {idx}: {row}")
                continue

            # Validar que la columna J contenga "DIA"
            if len(row) < 10 or row[9].strip().upper() != "DIA":  # Columna J es el índice 9
                logging.info(f"Fila {idx} ignorada porque no es para clientes 'DIA': {row}")
                continue

            try:
                # Asignar valores de columnas, manejando celdas vacías
                id_value = row[0]
                last_name = row[1]
                cantidad = row[2] if len(row) > 2 else "N/A"
                atributo = row[4] if len(row) > 4 else ""
                vendido = row[5] if len(row) > 5 and row[5] else None  # Columna F (VENDIDO)
                requisito = row[6] if len(row) > 6 and row[6] else None  # Columna G (REQUISITO)
                sunshine = row[10].strip().upper() if len(row) > 10 else "NO"  # Columna K (SUNSHINE)
                cortesia = row[11].strip().upper() if len(row) > 11 else "NO"  # Columna L (CORTESÍA)

                # Verificar si ya existe una carpeta con este ID
                folder_found = next((folder for folder in existing_folders if id_value in folder), None)
                if folder_found:
                    logging.info(f"Carpeta existente encontrada para ID '{id_value}': {folder_found}")
                    if update_progress:
                        update_progress(idx, total_rows, f"Carpeta existente: {folder_found}")
                    continue  # Saltar a la siguiente fila, ya que la carpeta existe

                # Construcción del nombre de la carpeta
                additions = []
                if sunshine == "SI":
                    additions.append("SUNSHINE")
                if cortesia == "SI":
                    additions.append("CORTESIA")
                additions_str = "-".join(additions)

                if atributo in (omit_c2_values or []):
                    folder_name = f"{atributo}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{atributo}-{last_name}-{id_value}".upper()
                elif atributo == "FAMILIA":
                    folder_name = f"{atributo}-D{cantidad}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{atributo}-D{cantidad}-{last_name}-{id_value}".upper()
                else:
                    folder_name = f"{cantidad} {atributo}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{cantidad} {atributo}-{last_name}-{id_value}".upper()

                # Ruta completa a la carpeta
                folder_path = os.path.join(date_folder_path, folder_name)

                # Crear la carpeta si no existe
                smb_create_directory_if_not_exists(smb_conn, share, folder_path)

                logging.info(f"Carpeta creada: {folder_path}")
                if update_progress:
                    update_progress(idx, total_rows, f"Creando carpeta: {folder_name}")

            except Exception as row_error:
                logging.error(f"Error procesando fila {idx}: {row_error}")

        # Finalizar barra de progreso
        if update_progress:
            update_progress(total_rows, total_rows, "Proceso completado.")

    except Exception as e:
        logging.error(f"Error creando carpetas: {e}")
        if update_progress:
            update_progress(100, 100, "Error durante el proceso.")

def create_folders_night(sheet_id, sheet_name, base_night, start_row=2, omit_c2_values=None, update_progress=None):
    """Crea carpetas basadas en los datos de Google Sheets usando conexión SMB, evitando duplicados."""
    omit_c2_values = ["MADRE E HIJO", "MADRE E HIJA", "PADRE E HIJO", "PADRE E HIJA", "PAREJA"]
    try:
        # Establecer conexión SMB
        smb_conn, share, sub_path = establish_smb_connection(base_night)

        # Validar y crear la ruta base completa
        date_folder_path = os.path.join(sub_path, sheet_name)
        ensure_full_path_exists(smb_conn, share, date_folder_path)

        # Leer los datos de la hoja
        sheet_values = read_google_sheet(sheet_id, f"'{sheet_name}'!A{start_row}:M")
        if not sheet_values:
            logging.warning("No hay datos para procesar en NIGHT.")
            if update_progress:
                update_progress(100, 100, "No hay datos para procesar.")
            return

        # Crear subcarpeta NIGHT dentro de la base_night_path
        night_folder_path = os.path.join(sub_path, sheet_name, "0-NIGHT")
        smb_create_directory_if_not_exists(smb_conn, share, night_folder_path)

        # Listar carpetas existentes en la ruta
        try:
            existing_folders = [
                f.filename for f in smb_conn.listPath(share, night_folder_path)
                if f.isDirectory and f.filename not in [".", ".."]
            ]
        except Exception as e:
            logging.warning(f"No se pudo listar las carpetas en la ruta NIGHT: {night_folder_path}. {e}")
            existing_folders = []

        total_rows = len(sheet_values)

        # Manejo de casos sin filas
        if total_rows == 0:
            logging.warning("No hay filas para procesar en la hoja.")
            if update_progress:
                update_progress(100, 100, "No hay filas para procesar.")
            return

        # Reiniciar barra de progreso
        if update_progress:
            update_progress(0, total_rows, "Iniciando creación de carpetas NIGHT...")

        # Procesar filas
        for idx, row in enumerate(sheet_values, start=1):
            if len(row) < 3:
                logging.warning(f"Fila incompleta en la posición {idx}: {row}")
                continue

            # Validar que la columna J contenga "DIA"
            if len(row) < 10 or row[9].strip().upper() != "NOCHE":  # Columna J es el índice 9
                logging.info(f"Fila {idx} ignorada porque no es para clientes 'DIA': {row}")
                continue

            try:
                id_value = row[0]
                last_name = row[1]
                cantidad = row[2] if len(row) > 2 else "N/A"
                atributo = row[4] if len(row) > 4 else ""
                vendido = row[5] if len(row) > 5 and row[5] else None  # Columna F (VENDIDO)
                requisito = row[6] if len(row) > 6 and row[6] else None  # Columna G (REQUISITO)
                sunshine = row[10].strip().upper() if len(row) > 10 else "NO"  # Columna K (SUNSHINE)
                cortesia = row[11].strip().upper() if len(row) > 11 else "NO"  # Columna L (CORTESÍA)

                # Verificar si ya existe una carpeta con este ID
                folder_found = next((folder for folder in existing_folders if id_value in folder), None)
                if folder_found:
                    logging.info(f"Carpeta existente encontrada para ID '{id_value}': {folder_found}")
                    if update_progress:
                        update_progress(idx, total_rows, f"Carpeta existente: {folder_found}")
                    continue  # Saltar a la siguiente fila, ya que la carpeta existe

                # Construcción del nombre de la carpeta
                additions = []
                if sunshine == "SI":
                    additions.append("SUNSHINE")
                if cortesia == "SI":
                    additions.append("CORTESIA")
                additions_str = "-".join(additions)

                if atributo in (omit_c2_values or []):
                    folder_name = f"{atributo}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{atributo}-{last_name}-{id_value}".upper()
                elif atributo == "FAMILIA":
                    folder_name = f"{atributo}-D{cantidad}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{atributo}-D{cantidad}-{last_name}-{id_value}".upper()
                else:
                    folder_name = f"{cantidad} {atributo}-{last_name}-{additions_str}-{id_value}".upper() if additions_str else f"{cantidad} {atributo}-{last_name}-{id_value}".upper()

                # Ruta completa a la carpeta
                folder_path = os.path.join(night_folder_path, folder_name)

                smb_create_directory_if_not_exists(smb_conn, share, folder_path)

                if update_progress:
                    try:
                        update_progress(idx, total_rows, f"Creando carpeta: {folder_name}")
                    except ZeroDivisionError as e:
                        logging.error(f"Error al actualizar la barra de progreso en fila {idx}: ", exc_info=e)
                        return

            except Exception as row_error:
                logging.error(f"Error procesando fila {idx}: {row_error}")

        # Finalizar barra de progreso
        if update_progress:
            try:
                update_progress(total_rows, total_rows, "Proceso completado.")
            except ZeroDivisionError as e:
                logging.error("Error al finalizar la barra de progreso: división por cero.", exc_info=e)

    except Exception as e:
        logging.error(f"Error creando carpetas: {e}")
        if update_progress:
            try:
                update_progress(100, 100, "Error durante el proceso.")
            except ZeroDivisionError as e:
                logging.error("Error al manejar el progreso en un proceso fallido: división por cero.", exc_info=e)

def validate_folders(sheet_id, selected_sheet, base_path, min_files, update_progress):
    """Valida las carpetas existentes y evita duplicados verificando contra la hoja en línea."""
    try:
        logging.info("Iniciando el dia")

        # Leer los valores de la hoja seleccionada
        sheet_values = read_google_sheet(sheet_id, f"'{selected_sheet}'!A{START_ROW}:J")
        if not sheet_values:
            logging.warning(f"No hay datos en la hoja '{selected_sheet}' para validar.")
            if update_progress:
                update_progress(100, 100, "No hay datos para validar.")
            return

        # Crear un conjunto de IDs ya presentes en la hoja
        existing_ids = {row[0] for row in sheet_values if len(row) > 0}

        # Construir la ruta de la carpeta de la fecha
        date_folder_path = os.path.join(base_path, selected_sheet)

        if not os.path.exists(date_folder_path):
            logging.error(f"La ruta '{date_folder_path}' no existe.")
            if update_progress:
                update_progress(100, 100, "Ruta no encontrada.")
            return

        # Obtener el número total de carpetas
        folder_list = os.listdir(date_folder_path)
        total_folders = len(folder_list)
        logging.info(f"Total de carpetas encontradas: {total_folders}")

        # Manejo de caso en el que no hay carpetas
        if total_folders == 0:
            logging.warning(f"El directorio '{date_folder_path}' está vacío. No hay carpetas para validar.")
            if update_progress:
                update_progress(100, 100, "No hay carpetas para validar.")
            return

        processed_folders = 0

        if update_progress:
            update_progress(0, total_folders, "Iniciando validación de carpetas...")

        # Recorrer las carpetas existentes en la ruta local
        for folder_name in folder_list:
            folder_path = os.path.join(date_folder_path, folder_name)

            # Ignorar si no es una carpeta
            if not os.path.isdir(folder_path):
                continue

            # Buscar el ID en el nombre completo de la carpeta
            match = re.search(r"F\d+", folder_name)  # Expresión regular para IDs válidos
            if not match:
                logging.warning(f"La carpeta '{folder_name}' no contiene un ID válido.")
                processed_folders += 1
                if update_progress:
                    update_progress(processed_folders, total_folders, f"Saltando carpeta inválida: {folder_name}")
                continue

            folder_id = match.group()  # Extraer el ID encontrado

            # Validar si el ID ya está en la hoja en línea
            if folder_id in existing_ids:
                logging.info(f"El ID '{folder_id}' ya existe en la hoja. Validando la carpeta...")

                # Validar si el nombre contiene '@' o '.com' para determinar si fue vendido
                result_at = "SI" if "@" in folder_name or ".com" in folder_name else "NO"
                logging.info(f"Validación Vendido: '{folder_name}' -> Vendido: {result_at}")

                # Contar archivos en la carpeta
                file_count = sum(1 for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.png')))

                # Validar si cumple con el mínimo de archivos
                result_files = "SI" if file_count >= min_files else "NO"
                logging.info(f"Validación Archivos: '{folder_name}' -> {'Cumple' if result_files == 'SI' else 'No cumple'}")

                # Actualizar la hoja con los resultados
                for i, row in enumerate(sheet_values, start=START_ROW):
                    if row[0] == folder_id:  # Buscar la fila correspondiente al ID
                        update_google_sheet(sheet_id, f'F{i}', [[result_at]], selected_sheet)
                        update_google_sheet(sheet_id, f'G{i}', [[result_files]], selected_sheet)
                        break

                if update_progress:
                    update_progress(processed_folders, total_folders, f"Validando carpeta: {folder_name}")
            else:
                logging.info(f"El ID '{folder_id}' no existe en la hoja. Registrando como nuevo cliente.")

            processed_folders += 1

        if update_progress:
            update_progress(total_folders, total_folders, "Validación completada.")
        logging.info("Validación de carpetas completada.")

    except Exception as e:
        logging.error(f"Error durante la validación de carpetas: {e}")
        if update_progress:
            update_progress(100, 100, "Error durante la validación.")
        messagebox.showerror("Error", f"Error durante la validación de carpetas: {e}")

def validate_folders_night(sheet_id, selected_sheet, base_night, min_files, update_progress):
    """Valida las carpetas existentes y evita duplicados verificando contra la hoja en línea."""
    try:

        # Leer los valores de la hoja seleccionada
        sheet_values = read_google_sheet(sheet_id, f"'{selected_sheet}'!A{START_ROW}:J")
        if not sheet_values:
            logging.warning(f"No hay datos en la hoja '{selected_sheet}' para validar.")
            if update_progress:
                update_progress(100, 100, "No hay datos para validar.")
            return

        # Crear un conjunto de IDs ya presentes en la hoja
        existing_ids = {row[0] for row in sheet_values if len(row) > 0}

        # Construir la ruta de la carpeta de la fecha
        date_folder_path = os.path.join(base_night, selected_sheet)

        # Asegurarte de que base_night no termine con una barra invertida
        if not date_folder_path.endswith("\\"):
            date_folder_path += "\\"

        # Añadir \NIGHT al final de base_night
        base_night_with_night = date_folder_path + "NIGHT"
        
        logging.info(f"El ID iniciosss '{base_night_with_night}'")

        if not os.path.exists(base_night_with_night):
            logging.error(f"La ruta '{base_night_with_night}' no existe.")
            if update_progress:
                update_progress(100, 100, "Ruta no encontrada.")
            return

        # Obtener el número total de carpetas
        folder_list = os.listdir(base_night_with_night)
        total_folders = len(folder_list)
        logging.info(f"Total de carpetas encontradas: {total_folders}")

        # Manejo de caso en el que no hay carpetas
        if total_folders == 0:
            logging.warning(f"El directorio '{base_night_with_night}' está vacío. No hay carpetas para validar.")
            if update_progress:
                update_progress(100, 100, "No hay carpetas para validar.")
            return

        processed_folders = 0

        if update_progress:
            update_progress(0, total_folders, "Iniciando validación de carpetas...")

        # Recorrer las carpetas existentes en la ruta local
        for folder_name in folder_list:
            folder_path = os.path.join(base_night_with_night, folder_name)

            # Ignorar si no es una carpeta
            if not os.path.isdir(folder_path):
                continue

            # Buscar el ID en el nombre completo de la carpeta
            match = re.search(r"F\d+", folder_name)  # Expresión regular para IDs válidos
            if not match:
                logging.warning(f"La carpeta '{folder_name}' no contiene un ID válido.")
                processed_folders += 1
                if update_progress:
                    update_progress(processed_folders, total_folders, f"Saltando carpeta inválida: {folder_name}")
                continue

            folder_id = match.group()  # Extraer el ID encontrado

            # Validar si el ID ya está en la hoja en línea
            if folder_id in existing_ids:
                logging.info(f"El IDss '{folder_id}' ya existe en la hoja. Validando la carpeta...")

                # Validar si el nombre contiene '@' o '.com' para determinar si fue vendido
                result_at = "SI" if "@" in folder_name or ".com" in folder_name else "NO"
                logging.info(f"Validación Vendido: '{folder_name}' -> Vendido: {result_at}")

                # Contar archivos en la carpeta
                file_count = sum(1 for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.png')))

                # Validar si cumple con el mínimo de archivos
                result_files = "SI" if file_count >= min_files else "NO"
                logging.info(f"Validación Archivos: '{folder_name}' -> {'Cumple' if result_files == 'SI' else 'No cumple'}")

                # Actualizar la hoja con los resultados
                for i, row in enumerate(sheet_values, start=START_ROW):
                    if row[0] == folder_id:  # Buscar la fila correspondiente al ID
                        update_google_sheet(sheet_id, f'F{i}', [[result_at]], selected_sheet)
                        update_google_sheet(sheet_id, f'G{i}', [[result_files]], selected_sheet)
                        break

                if update_progress:
                    update_progress(processed_folders, total_folders, f"Validando carpeta: {folder_name}")
            else:
                logging.info(f"El ID '{folder_id}' no existe en la hoja. Registrando como nuevo cliente.")

            processed_folders += 1

        if update_progress:
            update_progress(total_folders, total_folders, "Validación completada.")
        logging.info("Validación de carpetas completada.")

    except Exception as e:
        logging.error(f"Error durante la validación de carpetas: {e}")
        if update_progress:
            update_progress(100, 100, "Error durante la validación.")
        messagebox.showerror("Error", f"Error durante la validación de carpetas: {e}")

