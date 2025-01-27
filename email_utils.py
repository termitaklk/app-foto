import os
import logging
import fnmatch
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
from PIL import Image, ImageDraw, ImageFont
import io
import json
import os
from google_sheets_utils import read_google_sheet
from file_management import construct_client_folder_path
import os
import io
import fnmatch

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Ruta al archivo de plantillas
TEMPLATES_PATH = os.path.join(os.path.dirname(__file__), "email_templates.json")

def send_emails_to_clients(sheet_id, sheet_name, smtp_config):
    """
    Envía correos a los clientes que cumplan las condiciones en la hoja seleccionada.
    Devuelve dos listas: clientes enviados y clientes no enviados.
    """
    try:
        # Leer datos desde la hoja seleccionada
        data = read_google_sheet(sheet_id, f"'{sheet_name}'!A2:H")

        if not data:
            logging.warning("No se encontraron datos en la hoja.")
            return [], []

        clientes_enviados = []
        clientes_no_enviados = []

        for row in data:
            try:
                # Validaciones
                if len(row) < 8:
                    clientes_no_enviados.append({"id": row[0], "razon": "Fila incompleta"})
                    continue

                client_id = row[0]
                last_name = row[1]
                visit_date = row[2]
                email = row[3]
                vendido = row[5].strip().upper()
                requisito = row[6].strip().upper()
                idioma = row[7].strip().capitalize()

                if vendido != "NO" or requisito != "SI":
                    clientes_no_enviados.append({"id": client_id, "razon": "No cumple condiciones (F=NO, G=SI)"})
                    continue

                if not email or "@" not in email or "." not in email:
                    clientes_no_enviados.append({"id": client_id, "razon": f"Correo inválido: {email}"})
                    continue

                if idioma not in ["Español", "Inglés", "Portugués", "Francés", "Ruso"]:
                    clientes_no_enviados.append({"id": client_id, "razon": f"Idioma no soportado: {idioma}"})
                    continue

                # Enviar correo
                send_email(last_name, visit_date, idioma, [email], construct_client_folder_path, smtp_config)
                clientes_enviados.append({"id": client_id, "email": email})

            except Exception as e:
                logging.error(f"Error enviando correo para {row}: {e}")
                clientes_no_enviados.append({"id": row[0], "razon": f"Error durante el envío: {e}"})

        return clientes_enviados, clientes_no_enviados

    except Exception as e:
        logging.error(f"Error procesando clientes para el envío: {e}")
        return [], []
        
def load_email_template(language):
    """Carga la plantilla de correos desde un archivo JSON según el idioma."""
    try:
        # Construir la ruta al archivo JSON en el mismo directorio que este script
        base_dir = os.path.dirname(__file__)  # Directorio actual del archivo
        file_path = os.path.join(base_dir, "email_templates.json")

        # Cargar el archivo JSON
        with open(file_path, "r", encoding="utf-8") as f:
            templates = json.load(f)
        
        # Devolver la plantilla según el idioma
        template = templates.get(language)
        if not template:
            raise ValueError(f"No se encontró una plantilla para el idioma: {language}")
        return template
    except FileNotFoundError:
        logging.error(f"El archivo email_templates.json no se encontró en {file_path}.")
        return None
    except Exception as e:
        logging.error(f"Error cargando plantilla de correo: {e}")
        return None
def get_email_template(language):
    """
    Obtiene una plantilla de correo específica según el idioma.
    """
    templates = load_email_template()
    return templates.get(language, templates.get("Español"))

# Función para comprimir imágenes y añadir marca de agua
def compress_and_watermark_image(image_path, watermark_text="SCAPEPARK", quality=50, max_size=(800, 800), scale=0.15):
    try:
        with Image.open(image_path) as img:
            img.thumbnail(max_size)
            watermark = Image.new("RGBA", img.size, (255, 255, 255, 0))
            draw = ImageDraw.Draw(watermark)
            font_size = int(img.size[0] * scale)
            font = ImageFont.truetype("arial.ttf", font_size)
            
            # Usar `font.getbbox()` para calcular el tamaño del texto
            text_bbox = font.getbbox(watermark_text)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            text_position = ((img.size[0] - text_width) / 2, (img.size[1] - text_height) / 2)
            
            draw.text(text_position, watermark_text, fill=(255, 255, 255, 150), font=font)
            
            # Aplicar la marca de agua
            watermarked = Image.alpha_composite(img.convert("RGBA"), watermark)
            
            # Guardar la imagen procesada en memoria
            img_io = io.BytesIO()
            watermarked.convert("RGB").save(img_io, format="JPEG", quality=quality)
            img_io.seek(0)
            return img_io
    except Exception as e:
        logging.error(f"Error al comprimir o añadir marca de agua: {e}")
        return None


def send_email(last_name, visit_date, language, recipients, client_folder_path, smtp_config):
    """Envía un correo basado en plantillas con archivos adjuntos procesados (compresión y marca de agua)."""
    try:
        # Cargar plantilla según el idioma
        template = load_email_template(language)
        if not template:
            raise ValueError(f"No se encontró una plantilla para el idioma: {language}")

        # Reemplazar marcadores con datos
        subject = template["subject"]
        body = template["body"].format(last_name=last_name, visit_date=visit_date)

        # Preparar el correo
        msg = MIMEMultipart()
        msg["From"] = smtp_config["sender"]
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        # Procesar y adjuntar imágenes desde la carpeta del cliente
        logging.info(f"Buscando archivos adjuntos en: {client_folder_path}")
        attachments = []
        for file_name in os.listdir(client_folder_path):
            if file_name.lower().endswith(("1.jpg", "2.jpg", "3.jpg")):  # Archivos permitidos
                file_path = os.path.join(client_folder_path, file_name)
                processed_image = compress_and_watermark_image(file_path)
                if processed_image:
                    img = MIMEImage(processed_image.read())
                    img.add_header("Content-Disposition", "attachment", filename=file_name)
                    msg.attach(img)
                    attachments.append(file_name)

        if not attachments:
            logging.warning(f"No se encontraron archivos adjuntos en {client_folder_path}")

        # Conectar al servidor SMTP y enviar
        with smtplib.SMTP(smtp_config["server"], smtp_config["port"]) as server:
            server.starttls()
            server.login(smtp_config["sender"], smtp_config["password"])
            server.sendmail(smtp_config["sender"], recipients, msg.as_string())

        logging.info(f"Correo enviado a: {recipients}")
        logging.info(f"Archivos adjuntos enviados: {attachments}")

    except Exception as e:
        logging.error(f"Error enviando correo: {e}")

