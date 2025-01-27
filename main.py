from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.ttk import Combobox
from datetime import datetime
import threading
from file_management import create_folders, create_folders_night, validate_folders, validate_folders_night
from google_sheets_utils import authenticate_sheets, check_and_create_today_sheet, read_google_sheet, update_google_sheet, write_google_sheet, read_configuration, get_sheet_names
from email_utils import construct_client_folder_path, load_email_template, send_email
import logging
from dotenv import load_dotenv
from googleapiclient.discovery import build
import time
from threading import Thread
from tkinter import Toplevel
from tkinter.ttk import Treeview, Scrollbar
import pandas as pd
from tkinter import filedialog
from datetime import datetime, timedelta
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import sys

# Configuración de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Cargar variables del archivo .env
load_dotenv()

# Configuración global
SHEET_ID = os.getenv("SHEET_ID")
START_ROW = 2

# Variables de interfaz
entry_pax_totales = None
entry_fecha_actual = None
entry_fecha_op = None
entry_clientes_vendidos = None
entry_clientes_nuevos = None
progress_bar = None
progress_info = None
entry_total_clientes = None

# Función para validar fechas
def is_valid_date(date_string):
    """Valida si una cadena está en el formato de fecha día-mes-año."""
    try:
        datetime.strptime(date_string, "%d-%m-%Y")
        return True
    except ValueError:
        return False
    
# Configuración SMTP
SMTP_CONFIG = {
    "sender": os.getenv("EMAIL_SENDER"),
    "password": os.getenv("EMAIL_PASSWORD"),
    "server": os.getenv("SMTP_SERVER"),
    "port": int(os.getenv("SMTP_PORT"))
}

# Variables globales para configuración
global_config = {
    "base_path": None,
    "min_files": 10,
    "start_row": 2,
    "create_interval": 60000,  # Milisegundos
    "validate_interval": 120000,  # Milisegundos
}

def initialize_app(sheet_id):
    """
    Inicialización al cargar la aplicación.
    """
    global global_config

    try:
        # Leer configuración desde la hoja de configuración
        #base_path, min_files, start_row, create_interval, day_night_config = read_configuration(sheet_id), validate_interval = read_configuration(sheet_id)
        base_path, base_night, min_files, start_row, create_interval, validate_interval, day_night_config = read_configuration(sheet_id)


        # Almacenar en la configuración global
        global_config.update({
            "base_path": base_path,
            "min_files": min_files,
            "start_row": start_row,
            "create_interval": create_interval,
            "validate_interval": validate_interval,
            "base_night": base_night
        })

        # Cargar las hojas y garantizar la existencia de la hoja para la fecha actual
        sheet_names, today_sheet = load_sheets(sheet_id)
        if not today_sheet:
            raise ValueError("No se pudo crear o encontrar la hoja para la fecha actual.")

        logging.info(f"Hoja activa inicializada: {today_sheet}")
        return sheet_names, today_sheet

    except Exception as e:
        logging.error(f"Error inicializando la aplicación: {e}")
        raise


def load_sheets(sheet_id):
    """
    Carga la lista de hojas y garantiza que exista la hoja para la fecha actual.
    """
    try:
        # Verificar o crear la hoja de la fecha actual
        today_sheet = check_and_create_today_sheet(sheet_id)

        # Obtener la lista de hojas existentes después de crear la hoja de hoy
        creds = authenticate_sheets()
        service = build("sheets", "v4", credentials=creds)
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheet_names = [sheet["properties"]["title"] for sheet in sheet_metadata.get("sheets", [])]

        return sheet_names, today_sheet

    except Exception as e:
        logging.error(f"Error cargando las hojas: {e}")
        return [], None
def log_configuration(sheet_id):
    """
    Registra en los logs las configuraciones establecidas para la aplicación y las muestra en tiempo real.
    """
    try:
        # Leer configuración
        base_path, min_files, start_row, create_interval, validate_interval = read_configuration(sheet_id)

        # Convertir los intervalos de milisegundos a minutos
        create_interval_minutes = create_interval / (60 * 1000)
        validate_interval_minutes = validate_interval / (60 * 1000)

        # Registrar configuraciones en los logs
        logging.info("Configuración establecida:")
        logging.info(f"Ruta base (base_path): {base_path}")
        logging.info(f"Cantidad mínima de fotos (min_files): {min_files}")
        logging.info(f"Fila inicial para procesar datos (start_row): {start_row}")
        logging.info(f"Intervalo de creación de carpetas (create_interval): {create_interval_minutes} minutos")
        logging.info(f"Intervalo de validación de carpetas (validate_interval): {validate_interval_minutes} minutos")
        logging.info(f"SHEET_ID: {sheet_id}")

    except Exception as e:
        logging.error(f"Error registrando configuraciones: {e}")


def schedule_create_folders(sheet_id, sheet_name, base_path, interval):
    """
    Programa la creación de carpetas a intervalos definidos.
    """
    def task():
        global is_running
        while is_running:
            try:
                logging.info("Ejecutando proceso de creación de carpetas...")
                #update_period_combobox()  # Llama a la función de actualización
                progress_bar["value"] = 0
                progress_info.config(text="Procesando filas...")

                create_folders(sheet_id, sheet_name, base_path, update_progress=update_progress)

                # Reiniciar el cronómetro para la interfaz gráfica
                start_timer("C", interval, window)

                # Actualizar la interfaz con los datos obtenidos
                update_interface(SHEET_ID, selected_sheet)

                # Esperar el intervalo o detenerse si se señala
                if stop_event.wait(interval / 1000):
                    break  # Salir si se señala la detención
            except Exception as e:
                logging.error(f"Error en el proceso de creación de carpetas: {e}")
                break

        logging.info("Proceso de creación de carpetas detenido.")

    # Crear y ejecutar un hilo para la tarea
    thread = Thread(target=task)
    threads.append(thread)
    thread.start()

def schedule_create_folders_night(sheet_id, sheet_name, base_night, interval):
    """
    Programa la creación de carpetas a intervalos definidos.
    """
    def task():
        global is_running
        while is_running:
            try:
                logging.info("Ejecutando proceso de creación de carpetas...")
                #update_period_combobox()  # Llama a la función de actualización
                progress_bar["value"] = 0
                progress_info.config(text="Procesando filas...")

                create_folders_night(sheet_id, sheet_name, base_night, update_progress=update_progress)

                # Reiniciar el cronómetro para la interfaz gráfica
                start_timer("C", interval, window)

                # Actualizar la interfaz con los datos obtenidos
                update_interface(SHEET_ID, selected_sheet)

                # Esperar el intervalo o detenerse si se señala
                if stop_event.wait(interval / 1000):
                    break  # Salir si se señala la detención
            except Exception as e:
                logging.error(f"Error en el proceso de creación de carpetas: {e}")
                break

        logging.info("Proceso de creación de carpetas detenido.")

    # Crear y ejecutar un hilo para la tarea
    thread = Thread(target=task)
    threads.append(thread)
    thread.start()

def schedule_validate_folders(sheet_id, sheet_name, base_path, interval, min_files, update_progress):
    """
    Programa la validación de carpetas a intervalos definidos.
    """
    global selected_sheet
    def task():
        global is_running
        while is_running:
            try:
                logging.info("Ejecutando proceso de validación de carpetas dia...")
                progress_bar["value"] = 0
                progress_info.config(text="Procesando filas...")

                validate_folders(sheet_id, sheet_name, base_path, min_files, update_progress)

                # Iniciar o reiniciar el cronómetro visualmente
                start_timer("V", interval, window)

                # Actualizar la interfaz con los datos obtenidos
                update_interface(SHEET_ID, selected_sheet)
                # Esperar el intervalo o detenerse si se señala
                if stop_event.wait(interval / 1000):
                    break
            except Exception as e:
                logging.error(f"Error en el proceso de validación de carpetas: {e}")
                break

        logging.info("Proceso de validación de carpetas detenido.")

    # Crear y ejecutar un hilo para la tarea
    thread = Thread(target=task)
    threads.append(thread)
    thread.start()

def schedule_validate_folders_night(sheet_id, sheet_name, base_night, interval, min_files, update_progress):
    """
    Programa la validación de carpetas a intervalos definidos.
    """
    global selected_sheet
    def task():
        global is_running
        while is_running:
            try:
                logging.info("Ejecutando proceso de validación de carpetas night...")
                progress_bar["value"] = 0
                progress_info.config(text="Procesando filas...")

                validate_folders_night(sheet_id, sheet_name, base_night, min_files, update_progress)

                # Iniciar o reiniciar el cronómetro visualmente
                start_timer("V", interval, window)

                # Actualizar la interfaz con los datos obtenidos
                update_interface(SHEET_ID, selected_sheet)

                # Esperar el intervalo o detenerse si se señala
                if stop_event.wait(interval / 1000):
                    break
            except Exception as e:
                logging.error(f"Error en el proceso de validación de carpetas: {e}")
                break

        logging.info("Proceso de validación de carpetas detenido.")

    # Crear y ejecutar un hilo para la tarea
    thread = Thread(target=task)
    threads.append(thread)
    thread.start()

def start_process(sheet_id, selected_sheet=None):
    """
    Procesa la hoja seleccionada al presionar el botón e inicia los procesos de creación y validación.
    """
    try:
        global global_config
        # Obtener configuración global
        base_path = global_config.get("base_path")
        base_night = global_config.get("base_night")
        min_files = global_config.get("min_files")
        start_row = global_config.get("start_row")
        create_interval = global_config.get("create_interval")
        validate_interval = global_config.get("validate_interval")

        # Validar configuración
        if not base_path:
            logging.error("La ruta base no está configurada.")
            messagebox.showerror("Error de Configuración", "La ruta base no está configurada.")
            return

        if not os.path.exists(base_path):
            logging.error(f"La ruta base '{base_path}' no existe.")
            messagebox.showerror("Error de Configuración", f"La ruta base '{base_path}' no existe.")
            return
        # Actualizar barra de progreso para inicio
        update_progress(0, 100, "Iniciando procesos...")

        # Obtener nombres de hojas
        sheet_names = get_sheet_names(sheet_id)

        # Seleccionar la hoja activa
        if not selected_sheet:
            valid_sheets = [sheet for sheet in sheet_names if is_valid_date(sheet)]

            if valid_sheets:
                selected_sheet = max(valid_sheets, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
                logging.info(f"Hoja seleccionada automáticamente: {selected_sheet}")
            else:
                logging.error("No se encontraron hojas válidas con formato fecha (día-mes-año).")
                messagebox.showerror("Error", "No se encontraron hojas válidas con formato fecha (día-mes-año).")
                return

        # Iniciar procesos con los intervalos cargados
        schedule_create_folders(sheet_id, selected_sheet, base_path, create_interval)

        schedule_validate_folders(sheet_id, selected_sheet, base_path, validate_interval, min_files, update_progress)

        logging.info("Procesos de creación y validación iniciados correctamente.")

    except Exception as e:
        logging.error(f"Error al iniciar los procesos en `start_process`: {e}")
        messagebox.showerror("Error", f"Error al iniciar los procesos: {e}")

# Actualizar progreso
def update_progress(current, total, message=""):
    """Actualiza la barra de progreso y muestra mensajes."""
    if progress_bar:
        progress_value = int((current / total) * 100)
        progress_bar["value"] = progress_value
        progress_bar.update()
    if progress_info:
        progress_info.config(text=message)
        progress_info.update()

# Actualizar campos de la interfaz
def update_field(sheet_name, total_pax, clientes_vendidos, clientes_nuevos, total_clientes):
    """Actualiza los campos de la interfaz gráfica."""
    bg_color = "#2C3E50"  # Fondo que coincide con el tema

    entry_fecha_actual.config(state=tk.NORMAL)
    entry_fecha_actual.delete(0, tk.END)
    entry_fecha_actual.insert(0, datetime.now().strftime("%d-%m-%Y"))
    entry_fecha_actual.config(state="readonly", readonlybackground=bg_color)

    entry_fecha_op.config(state=tk.NORMAL)
    entry_fecha_op.delete(0, tk.END)
    entry_fecha_op.insert(0, sheet_name)
    entry_fecha_op.config(state="readonly", readonlybackground=bg_color)

    entry_pax_totales.config(state=tk.NORMAL)
    entry_pax_totales.delete(0, tk.END)
    entry_pax_totales.insert(0, str(total_pax))
    entry_pax_totales.config(state="readonly", readonlybackground=bg_color)

    entry_clientes_vendidos.config(state=tk.NORMAL)
    entry_clientes_vendidos.delete(0, tk.END)
    entry_clientes_vendidos.insert(0, str(clientes_vendidos))
    entry_clientes_vendidos.config(state="readonly", readonlybackground=bg_color)

    entry_clientes_nuevos.config(state=tk.NORMAL)
    entry_clientes_nuevos.delete(0, tk.END)
    entry_clientes_nuevos.insert(0, str(clientes_nuevos))
    entry_clientes_nuevos.config(state="readonly", readonlybackground=bg_color)

    entry_total_clientes.config(state=tk.NORMAL)
    entry_total_clientes.delete(0, tk.END)
    entry_total_clientes.insert(0, str(total_clientes))
    entry_total_clientes.config(state="readonly", readonlybackground=bg_color)
    
def format_time(milliseconds):
    """Convierte milisegundos a formato MM:SS."""
    seconds = milliseconds // 1000
    minutes, seconds = divmod(seconds, 60)
    return f"{minutes:02}:{seconds:02}"

# Variable global para manejar el estado de los procesos
is_running = False
threads = []  # Para almacenar los hilos en ejecución y detenerlos

timer_threads = {"C": None, "V": None}  # Diccionario para almacenar los hilos de los cronómetros
timer_flags = {"C": True, "V": True}  # Banderas para controlar la ejecución de los cronómetros
stop_event = threading.Event()

def stop_all_processes():
    """
    Detiene todos los procesos en ejecución y reinicia el estado.
    """
    global threads, is_running, stop_event

    # Señalar detención
    is_running = False
    stop_event.set()  # Señalar a los hilos que deben detenerse

    # Detener los cronómetros
    stop_all_timers()

    # Esperar a que todos los hilos terminen
    for thread in threads:
        if thread.is_alive():
            logging.info(f"Esperando que el hilo {thread.name} termine...")
            thread.join(timeout=5)  # Esperar un tiempo limitado para evitar bloqueo

    # Limpiar lista de hilos y reiniciar la señal
    threads = []
    stop_event.clear()  # Restablecer la señal para futuros usos

    # Restaurar botones de envío de correos
    #toggle_email_buttons("normal")
    toggle_combobox("normal")
    
    # Volver a aplicar restricciones basadas en el tipo de usuario
    disable_buttons_by_user_type(user_type)
    logging.info("Todos los procesos han sido detenidos.")

def start_button_action(button, sheet_id, selected_sheet, combobox_periodo):
    """
    Maneja el estado del botón Start/Stop para iniciar o detener los procesos.
    """
    global is_running, threads, timer_flags

    if not is_running:

        # Cambiar el estado a "corriendo"
        is_running = True
        button.config(bg="red", text="STOP")

        # Deshabilitar botones de envío de correos
        toggle_email_buttons("disabled")
        toggle_combobox("disabled")

        # Restablecer las banderas de los cronómetros
        timer_flags["C"] = True
        timer_flags["V"] = True

        # Determinar la función de creación y validación según el período seleccionado
        if combobox_periodo.upper() == "DIA":
            # Configurar creación de carpetas para "DÍA"
            folder_function = schedule_create_folders
            folder_args = (sheet_id, selected_sheet, global_config["base_path"], global_config["create_interval"])

            # Configurar validación para "DÍA"
            validate_function = schedule_validate_folders
            validate_args = (
                sheet_id,
                selected_sheet,
                global_config["base_path"],  # Ruta base para "DÍA"
                global_config["validate_interval"],
                global_config["min_files"],
                update_progress,
            )

        elif combobox_periodo.upper() == "NOCHE":
            # Configurar creación de carpetas para "NOCHE"
            folder_function = schedule_create_folders_night
            folder_args = (sheet_id, selected_sheet, global_config["base_night"], global_config["create_interval"])

            # Configurar validación para "NOCHE"
            validate_function = schedule_validate_folders_night
            validate_args = (
                sheet_id,
                selected_sheet,
                global_config["base_night"],  # Ruta base para "NOCHE"
                global_config["validate_interval"],
                global_config["min_files"],
                update_progress,
            )

        else:
            logging.error(f"Período desconocido: {combobox_periodo}")
            button.config(bg="blue", text="START")
            is_running = False
            toggle_combobox("normal")
            return

        # Crear hilos separados para creación y validación
        thread_create = Thread(target=folder_function, args=folder_args)
        thread_validate = Thread(target=validate_function, args=validate_args)

        # Revisar si los hilos ya existen y están activos
        if thread_create not in threads or not thread_create.is_alive():
            threads.append(thread_create)
            thread_create.start()

        if thread_validate not in threads or not thread_validate.is_alive():
            threads.append(thread_validate)
            thread_validate.start()

    else:
        # Detener los procesos y los cronómetros
        def stop_processes_in_thread():
            stop_all_processes()

        # Crear y ejecutar el hilo para detener los procesos
        stop_thread = threading.Thread(target=stop_processes_in_thread)
        stop_thread.daemon = True  # Permite que el hilo termine cuando la app se cierra
        stop_thread.start()

        # Cambiar el estado a "detenido"
        button.config(bg="blue", text="START")
        is_running = False
        timer_flags["C"] = False
        timer_flags["V"] = False

        # Habilitar botones de envío de correos
        toggle_combobox("normal")

def start_timer(label_type, interval, window):
    """
    Inicia un cronómetro de cuenta regresiva para un intervalo específico y lo reinicia automáticamente.
    """
    def countdown(remaining_time):
        global timer_flags
        while remaining_time > 0 and timer_flags[label_type]:
            minutes, seconds = divmod(remaining_time, 60)
            timer_text = f"{minutes:02}:{seconds:02}"

            # Actualizar el texto del cronómetro
            if label_type == "C":
                cronometro_c.config(text=timer_text)
            elif label_type == "V":
                cronometro_v.config(text=timer_text)

            # Esperar 1 segundo o detenerse si se señala
            if stop_event.wait(1):  # Usar la señal para interrumpir la espera
                break
            remaining_time -= 1

        # Reiniciar a 00:00 si se detuvo
        if not timer_flags[label_type]:
            reset_timer(label_type)

    # Crear un hilo para el cronómetro
    timer_thread = Thread(target=countdown, args=(interval // 1000,))
    timer_threads[label_type] = timer_thread
    timer_thread.start()
    
def reset_timer(label_type):
    """
    Reinicia un cronómetro específico a 00:00.
    """
    if label_type == "C":
        cronometro_c.config(text="00:00")
    elif label_type == "V":
        cronometro_v.config(text="00:00")
    logging.info(f"Cronómetro {label_type} reiniciado a 00:00.")

def stop_all_timers():
    """
    Detiene todos los cronómetros y los reinicia a 00:00.
    """
    global timer_flags
    # Detener los cronómetros estableciendo las banderas en False
    timer_flags["C"] = False
    timer_flags["V"] = False

    # Reiniciar los cronómetros visualmente
    reset_timer("C")
    reset_timer("V")
    logging.info("Todos los cronómetros han sido detenidos y reiniciados a 00:00.")

def send_email_action(sheet_id, selected_sheet, button):
    """
    Acción para enviar correos basados en la hoja seleccionada.
    """
    def task():
        try:
            # Cambiar el estado del botón
            toggle_email_buttons_state("disabled")
            button.config(state="disabled", text="Enviando Correos")
            toggle_start_button("disabled")
            toggle_combobox("disabled")

            logging.info(f"Enviando correos para la hoja: {selected_sheet}")

            # Leer configuración global
            base_path = global_config.get("base_path")
            if not base_path:
                raise ValueError("La ruta base no está configurada.")

            # Leer datos de la hoja seleccionada
            sheet_data = read_google_sheet(sheet_id, f"'{selected_sheet}'!A2:I")
            if not sheet_data:
                logging.warning(f"No se encontraron datos en la hoja: {selected_sheet}")
                progress_info.config(text="No se encontraron datos para procesar.")
                return

            total_rows = len(sheet_data)
            correos_enviados = []
            correos_no_enviados = []

            for index, row in enumerate(sheet_data, start=2):
                try:
                    # Validar que la fila tenga suficientes columnas
                    if len(row) < 8:
                        motivo = "Fila incompleta"
                        logging.warning(f"Fila incompleta: {row}")
                        correos_no_enviados.append({"row": row, "error": motivo})
                        update_google_sheet(sheet_id, f"I{index}", [["NO"]], selected_sheet)
                        update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                        progress_bar.update()
                        continue

                    # Extraer datos de la fila
                    id_value = row[0]
                    last_name = row[1]
                    cantidad = row[2]
                    email = row[3]
                    atributo = row[4]
                    vendido = row[5]
                    requisito = row[6]
                    idioma = row[7]  # El idioma se extrae directamente de la hoja
                    enviado = row[8].strip().upper() if len(row) > 8 else "NO"

                    # Verificar si el correo ya fue enviado
                    if enviado == "SI":
                        motivo = "Ya enviado"
                        update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                        progress_bar.update()
                        continue

                    # Validar si el correo cumple las condiciones
                    if vendido.strip().upper() != "NO":
                        motivo = "Álbum vendido"
                        correos_no_enviados.append({"row": row, "error": motivo})
                        update_google_sheet(sheet_id, f"I{index}", [["NO"]], selected_sheet)
                        update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                        progress_bar.update()
                        continue

                    if requisito.strip().upper() != "SI":
                        motivo = "Requisito de fotos"
                        correos_no_enviados.append({"row": row, "error": motivo})
                        update_google_sheet(sheet_id, f"I{index}", [["NO"]], selected_sheet)
                        update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                        progress_bar.update()
                        continue

                    # Construir la ruta de la carpeta del cliente
                    client_folder_path = construct_client_folder_path(
                        base_path=base_path,
                        sheet_name=selected_sheet,
                        id_value=id_value,
                        last_name=last_name,
                        cantidad=cantidad,
                        atributo=atributo,
                    )

                    if not client_folder_path:
                        motivo = "Carpeta no encontrada"
                        correos_no_enviados.append({"row": row, "error": motivo})
                        update_google_sheet(sheet_id, f"I{index}", [["NO"]], selected_sheet)
                        update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                        progress_bar.update()
                        continue

                    # Configuración del servidor SMTP
                    smtp_config = {
                        "sender": os.getenv("EMAIL_SENDER"),
                        "password": os.getenv("EMAIL_PASSWORD"),
                        "server": os.getenv("SMTP_SERVER"),
                        "port": int(os.getenv("SMTP_PORT", 587)),
                    }

                    # Enviar correo
                    send_email(
                        last_name=last_name,
                        visit_date=selected_sheet,
                        language=idioma,
                        recipients=[email],
                        client_folder_path=client_folder_path,
                        smtp_config=smtp_config,
                    )
                    correos_enviados.append({"row": row, "status": "Enviado"})
                    update_google_sheet(sheet_id, f"I{index}", [["SI"]], selected_sheet)
                    motivo = "Éxito"
                    update_progress(index, total_rows, f"Cliente: {row[1]} - {motivo}")
                    progress_bar.update()

                except Exception as e:
                    motivo = str(e)
                    logging.error(f"Error procesando fila {row}: {motivo}")
                    correos_no_enviados.append({"row": row, "error": motivo})
                    update_google_sheet(sheet_id, f"I{index}", [["NO"]], selected_sheet)
                    update_progress(index, total_rows, f"Cliente: {row[1]} - Falla ({motivo})")
                    progress_bar.update()

            # Mostrar resultados en la interfaz con una tabla
            show_email_summary(correos_enviados, correos_no_enviados, selected_sheet)

        except Exception as e:
            logging.error(f"Error al enviar correos: {e}")
            messagebox.showerror("Error", f"Error al enviar correos: {e}")
        
        finally:
            disable_buttons_by_user_type(user_type)
            button.config(state="normal", text="Enviar Correos Fotografia")
            toggle_combobox("normal")
    
     # Crear y ejecutar el hilo
    thread = threading.Thread(target=task)
    thread.daemon = True  # Permitir que el hilo se cierre con la app
    thread.start()

def show_email_summary(correos_enviados, correos_no_enviados, sheet_name):
    """
    Muestra una ventana con el resumen de correos enviados y no enviados,
    permite exportar a Excel o cerrar la ventana.
    """
    # Crear ventana
    summary_window = tk.Toplevel()
    summary_window.title("Resumen de Correos")
    summary_window.geometry("800x600")

    # Etiqueta de título
    tk.Label(
        summary_window,
        text=f"Resumen de Envíos - {sheet_name}",
        font=("Arial", 14, "bold"),
        pady=10
    ).pack()

    # Frame para los datos
    data_frame = tk.Frame(summary_window)
    data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Tabla de Enviados
    enviados_label = tk.Label(
        data_frame,
        text=f"Correos Enviados: {len(correos_enviados)}",
        font=("Arial", 12, "bold"),
        pady=5,
    )
    enviados_label.grid(row=0, column=0, sticky="w", padx=5)

    enviados_table = tk.Text(data_frame, height=10, width=100)
    enviados_table.grid(row=1, column=0, padx=5, pady=5)
    enviados_table.tag_configure("centered", justify="center")

    for correo in correos_enviados:
        motivo = "N/A"
        row = correo["row"]
        enviados_table.insert(
            tk.END,
            f"Cliente: {row[1]} | Correo: {row[3]} | Motivo: {motivo}\n",
            "centered"
        )

    enviados_table.config(state=tk.DISABLED)

    # Tabla de No Enviados
    no_enviados_label = tk.Label(
        data_frame,
        text=f"Correos No Enviados: {len(correos_no_enviados)}",
        font=("Arial", 12, "bold"),
        pady=5,
    )
    no_enviados_label.grid(row=2, column=0, sticky="w", padx=5)

    no_enviados_table = tk.Text(data_frame, height=10, width=100)
    no_enviados_table.grid(row=3, column=0, padx=5, pady=5)
    no_enviados_table.tag_configure("centered", justify="center")

    for correo in correos_no_enviados:
        motivo = correo.get("error", "Desconocido")
        row = correo["row"]
        no_enviados_table.insert(
            tk.END,
            f"Cliente: {row[1]} | Correo: {row[3]} | Motivo: {motivo}\n",
            "centered"
        )

    no_enviados_table.config(state=tk.DISABLED)

    # Frame para los botones
    button_frame = tk.Frame(summary_window)
    button_frame.pack(fill=tk.X, pady=10)

    def export_and_close():
        try:
            export_summary_to_excel(sheet_name, correos_enviados, correos_no_enviados)
            summary_window.destroy()  # Cerrar la ventana después de exportar
        except Exception as e:
            logging.error(f"Error al exportar el resumen: {e}")
            messagebox.showerror("Error", "No se pudo exportar el resumen.")

    # Crear un Frame para los botones
    button_frame = tk.Frame(summary_window)
    button_frame.pack(pady=10, anchor=tk.CENTER)

    # Botón para exportar a Excel
    export_button = tk.Button(
        button_frame,
        text="Exportar a Excel",
        bg="#4CAF50",
        fg="#ffffff",
        font=("Arial", 10, "bold"),
        command=export_and_close
    )
    export_button.pack(side=tk.LEFT, padx=(10, 5))  # Separación izquierda-derecha

    # Botón para cerrar la ventana
    close_button = tk.Button(
        button_frame,
        text="Cerrar",
        bg="#f44336",
        fg="#ffffff",
        font=("Arial", 10, "bold"),
        command=summary_window.destroy
    )
    close_button.pack(side=tk.LEFT, padx=(5, 10))  # Separación derecha-izquierda

from tkinter import filedialog

from tkinter import filedialog
import pandas as pd
from tkinter import messagebox
import logging

def export_summary_to_excel(sheet_name, correos_enviados, correos_no_enviados):
    """
    Exporta el resumen de correos enviados y no enviados a un archivo Excel en una sola hoja.
    Permite al usuario seleccionar la ubicación para guardar el archivo.
    """
    try:
        # Preparar los datos para el DataFrame
        summary_data = []

        # Agregar datos de correos enviados
        for correo in correos_enviados:
            summary_data.append({
                "Cliente": correo["row"][1],  # Columna de cliente
                "Correo": correo["row"][3],  # Columna de correo
                "Estado": "Enviado",
                "Motivo": "N/A"  # Enviado correctamente
            })

        # Agregar datos de correos no enviados
        for correo in correos_no_enviados:
            summary_data.append({
                "Cliente": correo["row"][1],  # Columna de cliente
                "Correo": correo["row"][3],  # Columna de correo
                "Estado": "No Enviado",
                "Motivo": correo.get("error", "Desconocido")  # Razón del error
            })

        # Crear DataFrame con toda la información
        summary_df = pd.DataFrame(summary_data)

        # Abrir cuadro de diálogo para guardar el archivo
        file_name = f"{sheet_name}_Resumen_Envios.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile=file_name,
            title="Guardar Resumen de Correos"
        )

        if not file_path:  # Si el usuario cancela el guardado
            logging.info("Exportación cancelada por el usuario.")
            return

        # Guardar el archivo Excel con el nombre de la hoja igual al de la hoja seleccionada
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, index=False, sheet_name=sheet_name)

        messagebox.showinfo("Éxito", f"Resumen exportado a {file_path}")
        logging.info(f"Resumen exportado correctamente a {file_path}")

    except Exception as e:
        logging.error(f"Error al exportar el resumen a Excel: {e}")
        messagebox.showerror("Error", f"Error al exportar el resumen a Excel: {e}")

def show_summary(clientes_enviados, clientes_no_enviados, selected_sheet):
    """
    Muestra un resumen de los correos enviados y no enviados.
    """
    resumen_window = tk.Toplevel()
    resumen_window.title("Resumen de Envío de Correos")
    resumen_window.geometry("600x400")

    # Sección de enviados
    tk.Label(resumen_window, text="Correos Enviados", font=("Arial", 12, "bold")).pack(pady=5)
    enviados_frame = tk.Frame(resumen_window, bg="#f0f0f0")
    enviados_frame.pack(fill="both", expand=True, padx=10, pady=5)
    for cliente in clientes_enviados:
        tk.Label(enviados_frame, text=f"ID: {cliente['id']} - Email: {cliente['email']}").pack(anchor="w")

    # Separador
    tk.Label(resumen_window, text="").pack()

    # Sección de no enviados
    tk.Label(resumen_window, text="Correos No Enviados", font=("Arial", 12, "bold")).pack(pady=5)
    no_enviados_frame = tk.Frame(resumen_window, bg="#f0f0f0")
    no_enviados_frame.pack(fill="both", expand=True, padx=10, pady=5)
    for cliente in clientes_no_enviados:
        tk.Label(no_enviados_frame, text=f"ID: {cliente['id']} - Razón: {cliente['razon']}").pack(anchor="w")

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        log_entry = self.format(record)
        self.text_widget.insert(tk.END, log_entry + "\n")
        self.text_widget.see(tk.END)  # Auto-scroll

def setup_log_window(master_window, main_width, main_height, screen_width, screen_height):
    """
    Configura una ventana separada para mostrar los logs en tiempo real.
    Está vinculada a la ventana principal y se cierra junto con ella.
    """
    log_window = tk.Toplevel(master_window)
    log_window.title("Logs de la Aplicación")
    log_window.geometry(f"600x400+{(screen_width // 2) + (main_width // 2)}+{(screen_height // 2) - (main_height // 2)}")

    # Evitar que la ventana de logs se cierre de forma independiente
    log_window.protocol("WM_DELETE_WINDOW", lambda: None)

    # Crear el widget Text
    text_widget = tk.Text(log_window, state="normal", bg="#f0f0f0", wrap="word")
    text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Asociar el Text widget al handler de logs
    text_handler = TextHandler(text_widget)
    text_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logging.getLogger().addHandler(text_handler)

    return log_window

def get_retention_days(sheet_id):
    """
    Obtiene el número de días de retención de logs desde la hoja de configuración (celda G2).
    """
    try:
        # Leer el valor de la celda G2
        result = read_google_sheet(sheet_id, "configuracion!G2:G2")
        if result and result[0]:
            retention_days = int(result[0][0])
            logging.info(f"Días de retención configurados desde la hoja: {retention_days}")
            return retention_days
        else:
            logging.warning("No se encontró un valor válido en configuracion!G2. Usando 3 días por defecto.")
            return 3  # Valor por defecto
    except Exception as e:
        logging.error(f"Error obteniendo días de retención desde la hoja: {e}")
        return 3  # Valor por defecto

class FileAndTextHandler(logging.Handler):
    """
    Manejador de logs personalizado para escribir en un archivo y en un widget Text.
    """
    def __init__(self, text_widget, log_file, retention_days=3):
        super().__init__()
        self.text_widget = text_widget
        self.log_file = log_file
        self.retention_days = retention_days

        # Limpiar logs antiguos al iniciar
        self._clean_old_logs()

    def emit(self, record):
        log_message = self.format(record)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Escribir en el widget Text
        if self.text_widget:
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", f"{timestamp} - {log_message}\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see("end")  # Desplazar automáticamente hacia abajo

        # Escribir en el archivo de logs
        self._write_to_file(f"{timestamp} - {log_message}\n")

        # Limpiar logs antiguos
        self._clean_old_logs()

    def _write_to_file(self, log_message):
        """
        Escribe un mensaje en el archivo de logs.
        """
        try:
            with open(self.log_file, "a", encoding="utf-8") as log_file:
                log_file.write(log_message)
        except Exception as e:
            print(f"Error escribiendo en el archivo de logs: {e}")

    def _clean_old_logs(self):
        """
        Limpia los logs más antiguos que el número de días definido en `retention_days`.
        """
        try:
            cutoff_date = datetime.now() - timedelta(days=self.retention_days)
            cleaned_lines = []

            with open(self.log_file, "r", encoding="utf-8") as log_file:
                lines = log_file.readlines()

            for line in lines:
                try:
                    timestamp_str = line.split(" - ")[0]
                    log_date = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
                    if log_date >= cutoff_date:
                        cleaned_lines.append(line)
                except (ValueError, IndexError):
                    # Si la línea no tiene formato válido, la mantenemos
                    cleaned_lines.append(line)

            # Reescribir el archivo con las líneas filtradas
            with open(self.log_file, "w", encoding="utf-8") as log_file:
                log_file.writelines(cleaned_lines)

        except FileNotFoundError:
            # Si el archivo no existe, no hacemos nada
            pass
        except Exception as e:
            print(f"Error limpiando logs antiguos: {e}")

    def close(self):
        super().close()

from tkinter import Menu, Toplevel, Label

def show_email_templates_window(user_type):
    """
    Muestra una ventana emergente para administrar las plantillas de correo.
    """
    email_window = Toplevel()
    email_window.title("Email Templates")
    email_window.geometry("400x300")
    Label(email_window, text="Aquí puedes administrar las plantillas de correo.", font=("Arial", 12)).pack(pady=20)

import json
from tkinter import Toplevel, Listbox, Text, Scrollbar, Button, Label, messagebox, END

def show_email_templates_window(user_type):
    """
    Muestra una ventana emergente para seleccionar y editar plantillas de correo.
    """
    def load_templates():
        """
        Carga las plantillas desde el archivo JSON.
        """
        try:
            # Ruta absoluta al archivo email_templates.json
            script_directory = os.path.dirname(os.path.abspath(__file__))
            templates_path = os.path.join(script_directory, "email_templates.json")

            # Leer el archivo JSON
            with open(templates_path, "r", encoding="utf-8") as file:
                templates = json.load(file)
            return templates
        except FileNotFoundError:
            messagebox.showerror("Error", f"No se encontró el archivo {templates_path}.")
            return {}
        except json.JSONDecodeError:
            messagebox.showerror("Error", f"El archivo {templates_path} tiene un formato inválido.")
            return {}

    def save_templates():
        """
        Guarda las plantillas actualizadas en el archivo JSON.
        """
        try:
            # Ruta absoluta al archivo email_templates.json
            script_directory = os.path.dirname(os.path.abspath(__file__))
            templates_path = os.path.join(script_directory, "email_templates.json")

            # Guardar el archivo JSON
            with open(templates_path, "w", encoding="utf-8") as file:
                json.dump(templates, file, indent=4, ensure_ascii=False)
            messagebox.showinfo("Éxito", "Los cambios se han guardado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

    def on_language_select(event):
        selected = listbox.get(listbox.curselection())
        bg_color = "#2C3E50"  # Fondo que coincide con el tema
        if selected in templates:
            # Cargar asunto
            subject_entry.config(state="normal")
            subject_entry.delete(0, END)
            subject_entry.insert(0, templates[selected].get("subject", ""))
            if user_type == "Fotografia":
                subject_entry.config(state="readonly", readonlybackground=bg_color)  # Solo lectura para Fotografía
            else:
                subject_entry.config(state="normal")  # Editable para Admin y Comercial
            # Cargar cuerpo
            text_area.delete("1.0", END)
            text_area.insert("1.0", templates[selected].get("body", ""))
        else:
            messagebox.showerror("Error", f"No se encontró la plantilla para el idioma: {selected}")

    def save_changes():
        """
        Guarda los cambios realizados en el asunto y el cuerpo de la plantilla.
        """
        try:
            selected = listbox.get(listbox.curselection())
            if selected in templates:
                # Guardar asunto
                templates[selected]["subject"] = subject_entry.get().strip()
                # Guardar cuerpo
                templates[selected]["body"] = text_area.get("1.0", END).strip()
                save_templates()
            else:
                messagebox.showerror("Error", "Debe seleccionar un idioma antes de guardar.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la plantilla: {e}")

    # Crear la ventana
    email_window = Toplevel()
    email_window.title("Email Templates")
    email_window.geometry("600x500")

    # Cargar las plantillas
    templates = load_templates()

    # Filtrar plantillas según el tipo de usuario
    if user_type == "Admin":
        allowed_templates = list(templates.keys())  # Admin tiene acceso completo
    elif user_type == "Comercial":
        allowed_templates = ["Comercial"]
    elif user_type == "Fotografia":
        allowed_templates = [template for template in templates.keys() if template != "Comercial"]
    else:
        allowed_templates = []  # Usuario no válido

    # Crear lista de idiomas permitidos
    Label(email_window, text="Idiomas disponibles:", font=("Arial", 12)).pack(anchor="w", padx=10, pady=5)
    listbox = Listbox(email_window, selectmode="single", height=10)
    listbox.pack(fill="y", side="left", padx=10, pady=5)
    scrollbar = Scrollbar(email_window, orient="vertical", command=listbox.yview)
    scrollbar.pack(side="left", fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    # Llenar el Listbox con las plantillas permitidas
    for language in allowed_templates:
        listbox.insert(END, language)

    # Campo para editar el asunto
    Label(email_window, text="Asunto de la plantilla:", font=("Arial", 12)).pack(anchor="w", padx=10, pady=5)
    subject_entry = tk.Entry(email_window, font=("Arial", 12))
    subject_entry.pack(fill="x", padx=10, pady=5)

    # Restringir edición del asunto según el tipo de usuario
    if user_type == "Fotografia":
        subject_entry.config(state="disabled")

    # Crear área de texto para editar la plantilla
    Label(email_window, text="Contenido de la plantilla:", font=("Arial", 12)).pack(anchor="w", padx=10, pady=5)
    text_area = Text(email_window, wrap="word", height=15)
    text_area.pack(fill="both", expand=True, padx=10, pady=5)

    # Botón para guardar cambios
    button_frame = Button(email_window, text="Guardar Cambios", command=save_changes, state="disabled")
    button_frame.pack(side="bottom", pady=10)

    # Vincular selección de idioma
    def enable_editor(event):
            selected = listbox.get(listbox.curselection())
            if selected in allowed_templates:
                text_area.config(state="normal")
                button_frame.config(state="normal")
                subject_entry.config(state="normal" if user_type in ["Admin", "Comercial"] else "disabled")
                on_language_select(event)
            else:
                text_area.config(state="disabled")
                button_frame.config(state="disabled")
                subject_entry.config(state="disabled")
                messagebox.showerror("Error", "No tiene permiso para editar esta plantilla.")


    listbox.bind("<<ListboxSelect>>", enable_editor)

    # Mensaje si no hay plantillas disponibles
    if not allowed_templates:
        messagebox.showwarning("Advertencia", "No tiene permiso para editar ninguna plantilla.")
        email_window.destroy()

def show_configuration_window(sheet_id):
    """
    Muestra una ventana emergente con la configuración actual de la aplicación cargada desde la hoja de configuración,
    incluyendo EMAIL_SENDER y SHEET_ID desde el archivo .env, con títulos más descriptivos.
    """
    try:
        # Leer configuración desde la hoja
        config_data = read_google_sheet(sheet_id, "configuracion!A1:G2")

        # Validar que se haya leído la configuración correctamente
        if not config_data or len(config_data) < 2:
            messagebox.showerror("Error", "No se pudo cargar la configuración desde la hoja.")
            return

        headers = config_data[0]  # Encabezados de la hoja
        values = config_data[1]  # Valores configurados

        # Omitir las columnas C2 y D2
        filtered_headers = [header for i, header in enumerate(headers) if i not in (2, 3)]
        filtered_values = [value for i, value in enumerate(values) if i not in (2, 3)]

        # Mapeo de títulos descriptivos para las configuraciones
        titles_mapping = {
            "Ruta Fotografia": "Ruta de Trabajo",
            "Valor Min Foto": "Mínimo de Archivos JPG por Cliente",
            "Crear Carpetas Time": "Tiempo de Creación de Carpetas",
            "Validar Carpetas Time": "Tiempo de Validación de Clientes",
            "EMAIL_SENDER": "Email Fotografía",
            "SHEET_ID": "ID Hoja Google Sheets",
            "Tiempo de Logs": "Tiempo de Duración de LOGS"
        }

        # Añadir EMAIL_SENDER y SHEET_ID desde .env con títulos descriptivos
        filtered_headers.extend(["EMAIL_SENDER", "SHEET_ID"])
        filtered_values.extend([os.getenv("EMAIL_SENDER", "No configurado"), os.getenv("SHEET_ID", "No configurado")])

        # Usar títulos descriptivos
        descriptive_headers = [titles_mapping.get(header, header) for header in filtered_headers]

        # Crear ventana de configuración
        config_window = Toplevel()
        config_window.title("Configuración Actual")
        config_window.geometry("600x400")
        config_window.configure(bg="#f5f5f5")  # Fondo claro

        # Título
        Label(config_window, text="Configuración Actual", font=("Arial", 16, "bold"), bg="#f5f5f5").pack(pady=10)

        # Frame para la tabla
        table_frame = tk.Frame(config_window, bg="#ffffff", bd=2, relief="solid")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Encabezados
        Label(table_frame, text="Configuración", font=("Arial", 12, "bold"), bg="#d9d9d9", anchor="w", padx=10).grid(row=0, column=0, sticky="ew", padx=1, pady=1)
        Label(table_frame, text="Valor", font=("Arial", 12, "bold"), bg="#d9d9d9", anchor="w", padx=10).grid(row=0, column=1, sticky="ew", padx=1, pady=1)

        # Mostrar datos en tabla
        for i, (header, value) in enumerate(zip(descriptive_headers, filtered_values), start=1):
            Label(table_frame, text=header, font=("Arial", 11), bg="#ffffff", anchor="w", padx=10).grid(row=i, column=0, sticky="ew", padx=1, pady=1)
            Label(table_frame, text=value, font=("Arial", 11), bg="#ffffff", anchor="w", padx=10).grid(row=i, column=1, sticky="ew", padx=1, pady=1)

        # Botón para cerrar la ventana
        Button(
            config_window,
            text="Cerrar",
            command=config_window.destroy,
            bg="#ff5e57",
            fg="#ffffff",
            font=("Arial", 10, "bold"),
            relief="groove",
            activebackground="#ff3b30",
            activeforeground="#ffffff"
        ).pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema al cargar la configuración: {e}")

def send_commercial_email_threaded(sheet_id, sheet_name, button):
    """
    Inicia un hilo para enviar correos comerciales y muestra el resumen al terminar.
    Durante el envío, el botón estará deshabilitado y en color rojo.
    """
    def task():
        try:
            # Cambiar estado del botón
            toggle_email_buttons_state("disabled")
            button.config(state="disabled", text="Enviando...")
            toggle_start_button("disabled")
            toggle_combobox("disabled")
            
            # Realizar el envío
            correos_enviados, correos_no_enviados = send_commercial_email(sheet_id, sheet_name)

            # Mostrar el resumen
            show_email_summary(correos_enviados, correos_no_enviados, sheet_name)
        finally:
            # Restaurar estado del botón
            #toggle_email_buttons_state("normal")
            button.config(state="normal", text="Enviar Correos Comerciales")
            #toggle_start_button("normal")
            # Volver a aplicar restricciones basadas en el tipo de usuario
            disable_buttons_by_user_type(user_type)
            toggle_combobox("normal")

    # Crear y comenzar el hilo
    thread = threading.Thread(target=task)
    thread.start()


def send_commercial_email(sheet_id, sheet_name):
    """
    Envía correos a todos los clientes de una hoja específica usando la plantilla 'Comercial'.
    Este proceso no valida más allá del correo.
    """
    correos_enviados = []
    correos_no_enviados = []

    try:
        # Configuración SMTP para este envío
        commercial_smtp_config = {
            "sender": os.getenv("BULK_EMAIL_SENDER", "No configurado"),
            "password": os.getenv("BULK_EMAIL_PASSWORD", "No configurado"),
            "server": os.getenv("BULK_SMTP_SERVER", "smtp.example.com"),
            "port": int(os.getenv("BULK_SMTP_PORT", 587)),
        }

        # Validar configuración SMTP
        if not all(commercial_smtp_config.values()) or "No configurado" in commercial_smtp_config.values():
            messagebox.showerror("Error", "La configuración SMTP para correos comerciales no está completa.")
            return [], []

        # Leer datos de la hoja seleccionada
        sheet_data = read_google_sheet(sheet_id, f"'{sheet_name}'!A2:D")  # Suponemos que el correo está en la columna D
        if not sheet_data:
            messagebox.showerror("Error", f"No se encontraron datos en la hoja: {sheet_name}")
            return [], []

        # Cargar la plantilla "Comercial"
        template = load_email_template("Comercial")
        if not template:
            messagebox.showerror("Error", "No se encontró la plantilla 'Comercial'.")
            return [], []

        # Configuración de la plantilla
        subject = template["subject"]
        body_template = template["body"]

        # Iterar sobre los clientes de la hoja
        for row in sheet_data:
            try:
                # Validar que haya suficiente información en la fila
                if len(row) < 4:
                    correos_no_enviados.append({"row": row, "error": "Fila incompleta"})
                    continue

                # Extraer datos del cliente
                last_name = row[1]
                email = row[3]  # Suponiendo que la columna D contiene el correo electrónico

                # Validar el correo
                if not email or "@" not in email or "." not in email:
                    correos_no_enviados.append({"row": row, "error": "Correo inválido"})
                    continue

                # Personalizar el cuerpo del correo
                body = body_template.format(last_name=last_name)

                # Crear el mensaje
                msg = MIMEMultipart()
                msg["From"] = commercial_smtp_config["sender"]
                msg["To"] = email
                msg["Subject"] = subject
                msg.attach(MIMEText(body, "html"))

                # Enviar el correo
                with smtplib.SMTP(commercial_smtp_config["server"], commercial_smtp_config["port"]) as server:
                    server.starttls()
                    server.login(commercial_smtp_config["sender"], commercial_smtp_config["password"])
                    server.sendmail(commercial_smtp_config["sender"], email, msg.as_string())

                correos_enviados.append({"row": row, "status": "Enviado"})

            except Exception as e:
                correos_no_enviados.append({"row": row, "error": str(e)})

        return correos_enviados, correos_no_enviados

    except Exception as e:
        logging.error(f"Error en el envío de correos comerciales: {e}")
        messagebox.showerror("Error", f"Error en el envío de correos comerciales: {e}")
        return [], []

def toggle_email_buttons(state):
    """
    Cambia el estado de los botones de envío de correo.
    :param state: "disabled" para deshabilitar, "normal" para habilitar.
    """
    try:
        send_email_button.config(state=state)
        commercial_email_button.config(state=state)
    except NameError:
        logging.error("Los botones de envío no están definidos en el alcance global.")

def toggle_combobox(state):
    """
    Cambia el estado del combobox de selección de hoja.
    :param state: "disabled" para deshabilitar, "normal" para habilitar.
    """
    combobox_sheet_select.config(state=state)
    combobox_periodo.config(state=state)

def toggle_start_button(state):
    """
    Cambia el estado del botón de iniciar procesos.
    :param state: "disabled" para deshabilitar, "normal" para habilitar.
    """
    start_button.config(state=state)

def toggle_email_buttons_state(state):
    """
    Cambia el estado de ambos botones de envío de correos.
    :param state: "disabled" para deshabilitar, "normal" para habilitar.
    """
    send_email_button.config(state=state)
    commercial_email_button.config(state=state)

# Variable para almacenar el tipo de usuario actual
current_user_type = None

def login_window(window):
    """
    Ventana de inicio de sesión que bloquea la interacción con la ventana principal hasta que el login sea exitoso.
    """

    def handle_login():
        """
        Maneja el intento de inicio de sesión.
        """
        global user_type, current_user_type

        # Cambiar el botón a rojo mientras se valida
        login_button.config(bg="red", text="Validando...", state="disabled")
        login.update_idletasks()  # Forzar redibujado

        username = username_entry.get()
        password = password_entry.get()

        try:
            # Llamar a la función global verify_credentials
            success, user_type = verify_credentials(username, password)

            if success:
                current_user_type = user_type  # Asignar el tipo de usuario
                login.destroy()  # Cerrar la ventana de login
                load_sheets_and_update_combobox()  # Actualizar combobox con restricciones
                window.attributes("-disabled", False)  # Desbloquear la ventana principal

                # Configurar permisos según el tipo de usuario
                disable_buttons_by_user_type(user_type)

            else:
                logging.warning("[handle_login] Usuario o contraseña incorrectos.")
                messagebox.showerror("Error", "Usuario o contraseña incorrectos.")
                # Restaurar el botón a su estado original
                login_button.config(bg="#4CAF50", text="Iniciar Sesión", state="normal")
        except Exception as e:
            logging.error(f"[handle_login] Error durante el manejo del login: {e}")
            # Restaurar el botón a su estado original
            login_button.config(bg="#4CAF50", text="Iniciar Sesión", state="normal")

    def move_main_window(event):
            """
            Mueve la ventana principal cuando se mueve la ventana de login.
            """
            if not move_main_window.initialized:  # Ignorar el primer evento <Configure>
                move_main_window.initialized = True
                return

            login_x, login_y = login.winfo_x(), login.winfo_y()
            offset_x = (login_width - main_width) // 2
            offset_y = (login_height - main_height) // 2
            new_x = login_x - offset_x
            new_y = login_y - offset_y
            window.geometry(f"+{new_x}+{new_y}")

    # Inicializar la bandera del evento <Configure>
    move_main_window.initialized = False

    def center_window(window, width, height):
        """
        Centra una ventana en la pantalla.
        """
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    # Centrar la ventana principal en la pantalla
    main_width, main_height = 800, 600  # Tamaño de la ventana principal
    center_window(window, main_width, main_height)

    # Crear la ventana de login
    login = ttk.Toplevel(window)
    login.title("Inicio de Sesión")
    login.configure(bg="#f0f0f0")

    # Deshabilitar la ventana principal mientras la ventana de login esté abierta
    window.attributes("-disabled", True)

    # Definir tamaño de la ventana de login y centrarla respecto a la pantalla
    login_width, login_height = 400, 400
    center_window(login, login_width, login_height)

    # Frame para contenido de login
    frame = tk.Frame(login, bg="#ffffff", relief="solid", bd=1)
    frame.place(relx=0.5, rely=0.5, anchor="center", width=350, height=350)

    # Título
    tk.Label(
        frame,
        text="Inicio de Sesión",
        font=("Arial", 18, "bold"),
        bg="#ffffff",
        fg="#333333",
    ).pack(pady=10)

    # Etiqueta y entrada para el usuario
    tk.Label(frame, text="Usuario:", font=("Arial", 14), bg="#ffffff").pack(pady=5)
    username_entry = tk.Entry(frame, font=("Arial", 14), justify="center", relief="solid", bd=1)
    username_entry.pack(pady=5, ipadx=10, ipady=5)

    # Etiqueta y entrada para la contraseña
    tk.Label(frame, text="Contraseña:", font=("Arial", 14), bg="#ffffff").pack(pady=5)
    password_entry = tk.Entry(frame, font=("Arial", 14), show="*", justify="center", relief="solid", bd=1)
    password_entry.pack(pady=5, ipadx=10, ipady=5)

   # Vincular el evento Enter al campo de contraseña
    password_entry.bind("<Return>", lambda event: handle_login())

    # Crear el botón de iniciar sesión
    login_button = tk.Button(
        frame,
        text="Iniciar Sesión",
        font=("Arial", 14, "bold"),
        bg="#4CAF50",
        fg="#ffffff",
        width=20,  # Definir ancho en caracteres
        height=2,  # Definir altura en líneas de texto
        command=handle_login,  # Ejecutar handle_login al hacer clic
    )
    login_button.pack(pady=20)

    def on_close():
        """
        Evitar que se cierre la ventana de login sin autenticarse.
        """
        window.destroy()
        return  # Evitar el cierre de la ventana

    login.protocol("WM_DELETE_WINDOW", on_close)  # Prevenir cierre de login

    # Vincular movimiento de la ventana principal con la ventana de login
    def move_main_window(event):
        x = login.winfo_x() - 50
        y = login.winfo_y() - 50
        window.geometry(f"+{x}+{y}")

    login.transient(window)  # Asociar la ventana de login con la principal
    login.grab_set()  # Bloquear interacción con la ventana principal
    window.wait_window(login)  # Esperar hasta que se cierre la ventana de login

def configure_permissions(user_type):
    """
    Configura los permisos de la aplicación en función del tipo de usuario.
    """
    if user_type == "Fotografia":
        commercial_email_button.config(state=tk.DISABLED)  # Deshabilitar envío comercial
        send_email_button.config(state=tk.NORMAL)  # Habilitar envío fotográfico
    elif user_type == "Comercial":
        send_email_button.config(state=tk.DISABLED)  # Deshabilitar envío fotográfico
        commercial_email_button.config(state=tk.NORMAL)  # Habilitar envío comercial
    else:
        # Si el tipo no es válido, deshabilitar ambos botones por seguridad
        send_email_button.config(state=tk.DISABLED)
        commercial_email_button.config(state=tk.DISABLED)

def disable_buttons_by_user_type(user_type):
    """
    Deshabilita botones en función del tipo de usuario.
    """
    if user_type == "Admin":
        send_email_button.config(state=tk.NORMAL)  # Habilitar botón fotográfico
        commercial_email_button.config(state=tk.NORMAL)  # Habilitar botón comercial
        start_button.config(state=tk.NORMAL)  # Habilitar botón de inicio de procesos
    elif user_type == "Fotografia":
        commercial_email_button.config(state=tk.DISABLED)  # Deshabilitar botón comercial
        send_email_button.config(state=tk.NORMAL)  # Habilitar botón fotográfico
        start_button.config(state=tk.NORMAL)  # Habilitar botón de inicio de procesos
    elif user_type == "Comercial":
        send_email_button.config(state=tk.DISABLED)  # Deshabilitar botón fotográfico
        commercial_email_button.config(state=tk.NORMAL)  # Habilitar botón comercial
        start_button.config(state=tk.DISABLED)  # Deshabilitar botón de inicio de procesos
    else:
        send_email_button.config(state=tk.DISABLED)
        commercial_email_button.config(state=tk.DISABLED)
        start_button.config(state=tk.DISABLED)  # Deshabilitar todos los botones en caso de error

def verify_credentials(username, password):
    """
    Verifica las credenciales ingresadas contra la hoja de usuarios en línea.
    Devuelve un indicador de éxito y el tipo de usuario.
    """
    try:
        # Leer datos de la hoja de usuarios
        user_data = read_google_sheet(SHEET_ID, "Usuarios!A2:C")

        if not user_data:
            logging.warning("[verify_credentials] La hoja de usuarios está vacía o no se pudo leer.")
            return False, None

        # Validar credenciales y registrar cada comparación
        for row in user_data:
            if len(row) < 3:  # Asegurarse de que la fila tenga usuario, contraseña y tipo
                continue
            if row[0] == username and row[1] == password:
                user_type = row[2]  # Fotografia o Comercial
                return True, user_type

        logging.info(f"[verify_credentials] Usuario no encontrado: {username}")
        return False, None
    except Exception as e:
        logging.error(f"[verify_credentials] Error verificando credenciales: {e}")
        return False, None

class CustomCombobox(tk.Frame):
    def __init__(self, parent, values, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.values = values
        self.selected_value = tk.StringVar()

        # Entry para mostrar el valor seleccionado, centrado
        self.entry = ttk.Entry(self, textvariable=self.selected_value, justify="center")
        self.entry.configure(state="readonly")  # Configurar el estado como readonly inicialmente
        self.entry.grid(row=0, column=0, sticky="we")

        # Crear ventana flotante para el Listbox (se inicializa en None)
        self.dropdown_window = None

        # Vincular eventos
        self.entry.bind("<Button-1>", self.toggle_listbox)

    def toggle_listbox(self, event=None):
        """Muestra u oculta la lista desplegable en una ventana flotante."""
        # Verificar si el estado es "disabled"
        if self.entry.cget("state") == "disabled":
            return  # No hacer nada si está deshabilitado

        if self.dropdown_window and self.dropdown_window.winfo_exists():
            self.dropdown_window.destroy()
        else:
            # Crear la ventana flotante
            self.dropdown_window = tk.Toplevel(self)
            self.dropdown_window.wm_overrideredirect(True)  # Sin bordes
            self.dropdown_window.geometry(self.calculate_dropdown_position())

            # Crear el Listbox dentro de la ventana flotante
            listbox = tk.Listbox(
                self.dropdown_window,
                exportselection=False,
                justify="center",  # Centrar el texto en la lista desplegable
                font=("Arial", 10)
            )
            listbox.pack(fill="both", expand=True)

            # Rellenar el Listbox con los valores actuales
            for value in self.values:
                listbox.insert("end", value)

            # Manejar selección
            listbox.bind("<<ListboxSelect>>", lambda event: self.select_item(event, listbox))

    def calculate_dropdown_position(self):
        """Calcula la posición y el tamaño para la ventana flotante."""
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        width = self.entry.winfo_width()  # Ancho del Entry
        height = min(200, len(self.values) * 20)  # Altura máxima de 200 px o según valores
        return f"{width}x{height}+{x}+{y}"

    def select_item(self, event, listbox):
        """Actualiza el valor seleccionado y cierra la ventana flotante."""
        selected_index = listbox.curselection()
        if selected_index:
            self.selected_value.set(listbox.get(selected_index))
        self.dropdown_window.destroy()

    def update_values(self, new_values):
        """Actualiza los valores disponibles en el combobox."""
        self.values = new_values
        if new_values:
            self.selected_value.set(new_values[0])  # Seleccionar el primer valor por defecto
        else:
            self.selected_value.set("")  # Limpiar el valor si la lista está vacía

    def get(self):
        """Devuelve el valor seleccionado actualmente."""
        return self.selected_value.get()

    def config(self, **kwargs):
        """Permite configurar atributos del Entry, como el estado."""
        if "state" in kwargs:
            self.entry.configure(state=kwargs["state"])

def filter_and_sort_sheets(sheet_names, user_type):
    """
    Filtra las hojas para excluir ciertas hojas, aplica restricciones de caducidad
    para usuarios Comerciales, y selecciona la más reciente basada en su nombre.
    :param sheet_names: Lista de nombres de hojas.
    :param user_type: Tipo de usuario (Fotografia, Comercial, Admin).
    :return: Lista de hojas filtradas y la más reciente.
    """
    # Excluir las hojas "Configuracion" y "Usuarios"
    excluded_sheets = {"Configuracion", "Usuarios"}
    valid_sheets = [sheet for sheet in sheet_names if sheet not in excluded_sheets]

    # Intentar convertir los nombres a fechas y ordenar
    def parse_date(sheet_name):
        try:
            return datetime.strptime(sheet_name, "%d-%m-%Y")
        except ValueError:
            return None

    # Filtrar las hojas con nombres válidos de fecha
    valid_sheets_with_dates = [(sheet, parse_date(sheet)) for sheet in valid_sheets]
    valid_sheets_with_dates = [item for item in valid_sheets_with_dates if item[1] is not None]

    # Si el usuario es Comercial, excluir los últimos 10 días
    if user_type == "Comercial":
        today = datetime.now()
        expiry_date = today - timedelta(days=10)
        valid_sheets_with_dates = [
            item for item in valid_sheets_with_dates if item[1] < expiry_date
        ]

    # Ordenar por fecha descendente
    valid_sheets_with_dates.sort(key=lambda x: x[1], reverse=True)

    # Obtener la hoja más reciente
    most_recent_sheet = valid_sheets_with_dates[0][0] if valid_sheets_with_dates else None
    filtered_sheet_names = [item[0] for item in valid_sheets_with_dates]

    return filtered_sheet_names, most_recent_sheet

def load_sheets_and_update_combobox():
    """
    Carga las hojas disponibles desde Google Sheets, aplica restricciones según el tipo de usuario,
    y actualiza las opciones del combobox.
    """
    try:
        sheet_names = get_sheet_names(SHEET_ID)  # Obtener nombres de hojas desde Google Sheets

        # Filtrar hojas según tipo de usuario
        filtered_sheets, most_recent_sheet = filter_and_sort_sheets(sheet_names, current_user_type)

        # Actualizar el combobox con las hojas filtradas
        combobox_sheet_select.update_values(filtered_sheets)

        # Seleccionar la hoja más reciente como valor por defecto
        if most_recent_sheet:
            combobox_sheet_select.selected_value.set(most_recent_sheet)
        else:
            combobox_sheet_select.selected_value.set("No hay hojas disponibles")
    except Exception as e:
        logging.error(f"Error al cargar hojas: {e}")
        combobox_sheet_select.update_values([])  # Pasar una lista vacía si ocurre un error

# Variable global para el Combobox
combobox_periodo = None
total_clientes = None

def get_periodo():
    """
    Obtiene el período seleccionado del Combobox `combobox_periodo`.

    Returns:
        str: Valor seleccionado en el Combobox, en mayúsculas y sin espacios adicionales.
    """
    return combobox_periodo.get().strip().upper()

def fetch_pax_data_and_clientes_vendidos(sheet_id, selected_sheet, combobox_periodo):
    """
    Obtiene los datos de PAX desde la columna 'C', calcula la suma total,
    cuenta los clientes vendidos con 'SI' en la columna 'F',
    y devuelve estos valores junto con el nombre de la hoja.

    Args:
        sheet_id (str): ID de la hoja de cálculo de Google Sheets.
        selected_sheet (str): Nombre de la hoja seleccionada.

    Returns:
        tuple: (sheet_name, total_pax, clientes_vendidos, clientes_nuevos)
            sheet_name (str): Nombre de la hoja seleccionada.
            total_pax (int): Suma total de PAX.
            clientes_vendidos (int): Cantidad de clientes vendidos con 'SI'.
            clientes_nuevos (int): Cantidad de nuevos clientes (aquí puedes implementar su cálculo).
    """
    combo = get_periodo()
    try:
        # Validar el nombre de la hoja seleccionada
        if not selected_sheet or selected_sheet.strip() == "":
            logging.error("El nombre de la hoja seleccionada es inválido.")
            return selected_sheet, 0, 0, 0, 0

        # Rango que incluye columnas 'C' y 'F'
        sheet_range = f"'{selected_sheet}'!C2:J"
        data = read_google_sheet(sheet_id, sheet_range)

        if not data:
            logging.warning(f"No se encontraron datos en el rango {sheet_range}.")
            return selected_sheet, 0, 0, 0, 0

        # Inicializar acumuladores
        total_pax = 0
        clientes_vendidos = 0
        clientes_nuevos = 0  # Puedes implementar la lógica para calcular nuevos clientes si aplica
        total_clientes_del_periodo = 0

        # Procesar las filas
        for index, row in enumerate(data):

            if len(row) > 7 and row[7].strip().upper() == combo.strip().upper():  # Columna 'J' (Período)
                total_clientes_del_periodo += 1  # Contar clientes del período
                try:
                    if len(row) > 0 and row[0]:
                        total_pax += int(row[0])  # Columna 'C' (PAX)
                except ValueError:
                    logging.warning(f"Valor no numérico ignorado en columna 'C': {row[0]}")

                # Contar clientes vendidos (columna 'F')
                if len(row) > 3 and row[3].strip().upper() == "SI":
                    clientes_vendidos += 1

        # Calcular clientes nuevos (no vendidos)
        clientes_nuevos = total_clientes_del_periodo - clientes_vendidos
        # Calcular el total de clientes (vendidos + no vendidos)
        total_clientes = clientes_vendidos + clientes_nuevos

        return selected_sheet, total_pax, clientes_vendidos, clientes_nuevos, total_clientes

    except Exception as e:
        logging.error(f"Error al obtener datos de PAX y clientes vendidos: {e}")
        return selected_sheet, 0, 0, 0, 0

def update_interface(sheet_id, selected_sheet):
    """
    Actualiza la interfaz gráfica con los datos obtenidos de PAX y clientes vendidos.

    Args:
        sheet_id (str): ID de la hoja de cálculo de Google Sheets.
        selected_sheet (str): Nombre de la hoja seleccionada.
    """
    combo = get_periodo()
    # Obtener los datos
    sheet_name, total_pax, clientes_vendidos, clientes_nuevos, total_clientes = fetch_pax_data_and_clientes_vendidos(sheet_id, selected_sheet, combo)

    # Actualizar todos los campos de la interfaz en una sola llamada
    update_field(sheet_name, total_pax, clientes_vendidos, clientes_nuevos, total_clientes)

# Configurar la ventana principal
def setup_window():
    global entry_pax_totales, entry_fecha_actual, entry_fecha_op, entry_clientes_vendidos, entry_clientes_nuevos, progress_bar, progress_info, entry_total_clientes
    global window, cronometro_c, cronometro_v, send_email_button, commercial_email_button, combobox_sheet_select, start_button, combobox_periodo, selected_sheet, combo  # Cronómetros para "C" y "V"

    # Crear ventana principal
    window = ttk.Window(themename="superhero")
    window.title("Gestión Automática - Scape Park")
    main_width, main_height = 800, 600  # Dimensiones de la ventana principal

    # Función personalizada para manejar el cierre de la ventana
    def on_close():
        """Cerrar la ventana y salir completamente de la aplicación."""
        window.destroy()  # Destruye la ventana
        sys.exit()        # Finaliza el programa

    # Vincular la función al evento de cierre
    window.protocol("WM_DELETE_WINDOW", on_close)

    # Etiqueta y Combobox
    label = ttk.Label(window, text="Periodo Actual (Día/Noche):", font=("Arial", 12))
    label.pack(pady=5)

    combobox_periodo = ttk.Combobox(window, state="readonly", values=["DIA", "NOCHE"])
    combobox_periodo.pack(pady=5)

    combobox_periodo.current(0)

    # Simula otro proceso y luego actualiza el Combobox
    #update_period_combobox()  # Llama a la función de actualización

    # Crear barra de menú
    menu_bar = Menu(window)

    # Menú "Utilidades"
    util_menu = Menu(menu_bar, tearoff=0)
    util_menu.add_command(
        label="Email Templates", 
        command=lambda: show_email_templates_window(user_type)  # Pasar user_type
    )
    menu_bar.add_cascade(label="Utilidades", menu=util_menu)

    # Menú "Información"
    info_menu = Menu(menu_bar, tearoff=0)
    info_menu.add_command(label="Acerca de", command=lambda: show_configuration_window(SHEET_ID))
    menu_bar.add_cascade(label="Información", menu=info_menu)

    # Configurar la barra de menú en la ventana
    window.config(menu=menu_bar)

    # Crear ventana de logs
    log_window = tk.Toplevel(window)
    log_window.title("Logs de la Aplicación")
    log_width, log_height = 600, 400  # Dimensiones de la ventana de logs

    # Obtener dimensiones de la pantalla
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Calcular posiciones para centrar ambas ventanas juntas
    total_width = main_width + log_width
    total_height = max(main_height, log_height)
    start_x = (screen_width - total_width) // 2
    start_y = (screen_height - total_height) // 2

    # Configurar geometría de ambas ventanas
    window.geometry(f"{main_width}x{main_height}+{start_x}+{start_y}")
    log_window.geometry(f"{log_width}x{log_height}+{start_x + main_width}+{start_y}")

    # Evitar que la ventana de logs se cierre de forma independiente
    log_window.protocol("WM_DELETE_WINDOW", lambda: None)

    # Configurar la ventana de logs para cerrarse junto con la principal
    window.protocol("WM_DELETE_WINDOW", lambda: [log_window.destroy(), window.destroy()])

    # Variables para rastrear posiciones actuales
    window_last_pos = (start_x, start_y)
    log_last_pos = (start_x + main_width, start_y)

    # Vincular movimiento entre ventanas
    def sync_windows(event):
        """
        Vincula el movimiento de las dos ventanas, evitando parpadeos.
        """
        nonlocal window_last_pos, log_last_pos

        if event.widget == window:
            # Obtener posición actual de la ventana principal
            x, y = window.winfo_x(), window.winfo_y()
            if (x, y) != window_last_pos:  # Solo actualizar si hay un cambio
                log_window.geometry(f"+{x + main_width}+{y}")
                window_last_pos = (x, y)
                log_last_pos = (x + main_width, y)

        elif event.widget == log_window:
            # Obtener posición actual de la ventana de logs
            x, y = log_window.winfo_x(), log_window.winfo_y()
            if (x, y) != log_last_pos:  # Solo actualizar si hay un cambio
                window.geometry(f"+{x - main_width}+{y}")
                log_last_pos = (x, y)
                window_last_pos = (x - main_width, y)

    # Sincronizar minimización y restauración
    def sync_minimize_restore(event):
        """
        Sincroniza la minimización y restauración de ambas ventanas.
        """
        if event.widget == window:
            state = window.state()
            log_window.state(state)
        elif event.widget == log_window:
            state = log_window.state()
            window.state(state)

    # Vincular eventos de movimiento y estado
    window.bind("<Configure>", sync_windows)
    log_window.bind("<Configure>", sync_windows)

    window.bind("<Map>", sync_minimize_restore)  # Restaurar ambas ventanas
    window.bind("<Unmap>", sync_minimize_restore)  # Minimizar ambas ventanas
    log_window.bind("<Map>", sync_minimize_restore)
    log_window.bind("<Unmap>", sync_minimize_restore)

    # Frame principal para la ventana principal
    main_frame = tk.Frame(window, bg="#ffffff", bd=2, relief="ridge")
    main_frame.pack(pady=20, padx=20, fill="both", expand=True)
    
    # Configurar columnas en main_frame
    for i in range(6):  # Aseguramos que haya espacio para 6 columnas
        main_frame.grid_columnconfigure(i, weight=1)

    # Crear widget Text para mostrar los logs en la ventana de logs
    text_widget = tk.Text(log_window, state="normal", bg="#f0f0f0", wrap="word")
    text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Configurar el archivo de logs en la misma carpeta que este script
    script_directory = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(script_directory, "log.txt")

    # Leer días de retención desde la hoja de configuración
    retention_days = get_retention_days(SHEET_ID)

    # Asociar el manejador de logs con limpieza automática
    file_and_text_handler = FileAndTextHandler(text_widget, log_file, retention_days=retention_days)
    file_and_text_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logging.getLogger().addHandler(file_and_text_handler)

    # Etiqueta para seleccionar la hoja
    label_sheet_select = tk.Label(main_frame, text="Seleccionar Hoja:", font=("Arial", 14, "bold"))
    label_sheet_select.grid(row=2, column=4, padx=5, pady=5)

    # Combobox personalizado
    sheet_names = ["Hoja 1", "Hoja 2", "Hoja 3"]
    combobox_sheet_select = CustomCombobox(main_frame, values=sheet_names)
    combobox_sheet_select.grid(row=3, column=4, padx=5, pady=5)

    # Campos de entrada (sin definir bg explícito)
    label_fecha_actual = tk.Label(main_frame, text="Fecha Actual:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_fecha_actual = tk.Entry(main_frame, state="disabled", fg="#ffffff")
    label_fecha_op = tk.Label(main_frame, text="Fecha OP Foto:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_fecha_op = tk.Entry(main_frame, state="disabled", bg="#f7f7f7")
    label_pax_totales = tk.Label(main_frame, text="PAX Totales:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_pax_totales = tk.Entry(main_frame, state="disabled", bg="#f7f7f7")
    label_clientes_vendidos = tk.Label(main_frame, text="Clientes Vendidos:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_clientes_vendidos = tk.Entry(main_frame, state="disabled", bg="#f7f7f7")
    label_clientes_nuevos = tk.Label(main_frame, text="Clientes No Vendidos:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_clientes_nuevos = tk.Entry(main_frame, state="disabled", bg="#f7f7f7")
    label_total_clientes = tk.Label(main_frame, text="Total de Clientes:", bg="#ffffff", font=("Arial", 10, "bold"))
    entry_total_clientes = tk.Entry(main_frame, state="disabled", bg="#f7f7f7")

    # Colocar campos en la ventana
    label_fecha_actual.grid(row=2, column=0, padx=10, pady=5, sticky="w")
    entry_fecha_actual.grid(row=2, column=1, padx=10, pady=5)
    label_fecha_op.grid(row=3, column=0, padx=10, pady=5, sticky="w")
    entry_fecha_op.grid(row=3, column=1, padx=10, pady=5)
    label_pax_totales.grid(row=4, column=0, padx=10, pady=5, sticky="w")
    entry_pax_totales.grid(row=4, column=1, padx=10, pady=5)
    label_clientes_vendidos.grid(row=5, column=0, padx=10, pady=5, sticky="w")
    entry_clientes_vendidos.grid(row=5, column=1, padx=10, pady=5)
    label_clientes_nuevos.grid(row=6, column=0, padx=10, pady=5, sticky="w")
    entry_clientes_nuevos.grid(row=6, column=1, padx=10, pady=5)
    label_total_clientes.grid(row=7, column=0, padx=10, pady=5, sticky="w")
    entry_total_clientes.grid(row=7, column=1, padx=10, pady=5)

    # Barra de progreso
    progress_info = tk.Label(window, text="Esperando...", bg="#f0f0f0")
    progress_info.pack(pady=5)
    progress_bar = ttk.Progressbar(window, orient="horizontal", mode="determinate")
    progress_bar.pack(pady=5, fill="x", padx=10)
    
    # Al cargar la app
    initialize_app(SHEET_ID)

    # Crear un Frame para los botones
    button_frame_app = tk.Frame(window)
    button_frame_app.pack(pady=10, anchor=tk.CENTER)

    # Botón de envío fotográfico
    send_email_button = tk.Button(
        button_frame_app,
        text="Enviar Correos Fotografia",
        command=lambda: send_email_action(SHEET_ID, combobox_sheet_select.get(), send_email_button),
        fg="#ffffff",
        font=("Arial", 10, "bold"),
    )
    #send_email_button.pack(pady=10)
    send_email_button.pack(side=tk.LEFT, padx=(10, 5))  # Separación izquierda-derecha
    

    # Botón de envío comercial
    commercial_email_button = tk.Button(
        button_frame_app,
        text="Enviar Correos Comerciales",
        command=lambda: send_commercial_email_threaded(SHEET_ID, combobox_sheet_select.get(), commercial_email_button),
        fg="#ffffff",
        font=("Arial", 10, "bold"),
    )
    commercial_email_button.pack(side=tk.LEFT, padx=(10, 5)) 


    # Botón de inicio
    start_button = tk.Button(
        button_frame_app,
        text="Iniciar Proceso",
        command=lambda: start_button_action(start_button, SHEET_ID, combobox_sheet_select.get(), combobox_periodo.get()),
        fg="#ffffff",
        font=("Arial", 10, "bold"),
    )
    start_button.pack(side=tk.LEFT, padx=(10, 5)) 
    #start_button.grid(row=0, column=1, pady=10)

    # Etiquetas C y V
    label_c = tk.Label(main_frame, text="Folder Creation Time", font=("Arial", 14, "bold"), bg="#ffffff", fg="blue")
    label_v = tk.Label(main_frame, text="Sales validation time", font=("Arial", 14, "bold"), bg="#ffffff", fg="blue")
    # Cronómetros para C y V
    cronometro_c = tk.Label(main_frame, text="00:00", font=("Arial", 12), bg="#ffffff")
    cronometro_v = tk.Label(main_frame, text="00:00", font=("Arial", 12), bg="#ffffff")

    cronometro_v.grid(row=5, column=4, padx=5)
    label_v.grid(row=4, column=4, padx=5)
    cronometro_c.grid(row=7, column=4, padx=5)
    label_c.grid(row=6, column=4, padx=5)
        
    # Llenar el CustomCombobox con nombres de hojas
    try:
        sheet_names = get_sheet_names(SHEET_ID)
        # Filtrar las hojas en función del tipo de usuario actual
        filtered_sheets, most_recent_sheet = filter_and_sort_sheets(sheet_names, current_user_type)
        
        
        # Actualizar los valores del combobox
        combobox_sheet_select.update_values(filtered_sheets)

        # Seleccionar la hoja más reciente como valor por defecto
        if most_recent_sheet:
            combobox_sheet_select.selected_value.set(most_recent_sheet)
        else:
            combobox_sheet_select.selected_value.set("No hay hojas disponibles")
    except Exception as e:
        logging.error(f"Error al cargar hojas: {e}")
        combobox_sheet_select.update_values([])  # Pasar una lista vacía si ocurre un error

    selected_sheet = combobox_sheet_select.get()  # Obtener nombre de la hoja seleccionada
    if not selected_sheet:
        print("Por favor selecciona una hoja válida.")
        return

    combo = combobox_periodo.get()
    # Actualizar la interfaz con los datos obtenidos
    update_interface(SHEET_ID, selected_sheet)
    
    # Mostrar ventana de login
    login_window(window)
    window.mainloop()

# Ejecutar la aplicación
if __name__ == "__main__":
        setup_window()



