import json
import time
import win32com.client as win32
import os
import random
import pandas as pd
from datetime import datetime

# ------------------------------
# Rutas base
# ------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
HTML_DIR = os.path.join(BASE_DIR, "html")
EXCEL_DIR = os.path.join(BASE_DIR, "providers")
PDF_DIR = os.path.join(BASE_DIR, "pdf")

data_path = os.path.join(DATA_DIR, "data.json")
subject_path = os.path.join(DATA_DIR, "subject.json")

# Excel clientes
EXCEL_PATH = os.path.join(EXCEL_DIR, "datos-clientes.xlsx")

# PDF fijo a enviar
PDF_FIJO = os.path.join(
    PDF_DIR,
    "Solicitud de Certificado de retenciones Clientes.pdf"
)

# ------------------------------
# Cargar data JSON
# ------------------------------
with open(data_path, "r", encoding="utf-8") as file:
    data = json.load(file)

asesor = data["asesor"]
correo_asesor = data["correoAsesor"]

# ------------------------------
# Cargar subjects
# ------------------------------
with open(subject_path, "r", encoding="utf-8") as file:
    subjects = json.load(file)["subjects"]

# ------------------------------
# Cargar plantillas HTML
# ------------------------------
html_templates = []
for filename in sorted(os.listdir(HTML_DIR)):
    if filename.endswith(".html"):
        with open(os.path.join(HTML_DIR, filename), "r", encoding="utf-8") as file:
            html_templates.append(file.read())

# ------------------------------
# Leer clientes desde Excel
# ------------------------------
df = pd.read_excel(EXCEL_PATH)

# Normalizar encabezados
df.columns = df.columns.astype(str).str.strip().str.lower()

clientes = []
for _, row in df.iterrows():
    cliente = row.get("clientes")
    correo = row.get("correo")

    if pd.notna(cliente):
        clientes.append({
            "clientes": str(cliente).strip(),
            "correo": str(correo).strip() if pd.notna(correo) else ""
        })

# ------------------------------
# Outlook
# ------------------------------
outlook = win32.Dispatch('Outlook.Application')

# ------------------------------
# Variables de control
# ------------------------------
novedades = []
correos_enviados = 0

# ------------------------------
# Envío de correos
# ------------------------------
for cliente in clientes:
    nombre_cliente = cliente["clientes"]
    correo_cliente = cliente["correo"]

    # ------------------------------
    # Validar correo
    # ------------------------------
    if not correo_cliente or "@" not in correo_cliente:
        novedades.append(f"{nombre_cliente}: no tiene correo válido.")
        continue

    # ------------------------------
    # Validar PDF fijo
    # ------------------------------
    if not os.path.exists(PDF_FIJO):
        novedades.append(
            "No se encontró el PDF fijo: Solicitud de Certificado de retenciones Clientes.pdf"
        )
        break

    # ------------------------------
    # Crear y enviar correo
    # ------------------------------
    html_template = random.choice(html_templates)
    subject = random.choice(subjects)

    html_content = (
        html_template
        .replace("{{ cliente }}", nombre_cliente)
        .replace("{{ asesor }}", asesor)
        .replace("{{ correoAsesor }}", correo_asesor)
    )

    mail = outlook.CreateItem(0)
    mail.To = correo_cliente
    mail.Subject = subject
    mail.HTMLBody = html_content
    mail.Attachments.Add(PDF_FIJO)

    mail.Send()
    correos_enviados += 1

    print(f"✔ Enviado a: {nombre_cliente} -> {correo_cliente}")
    print("📎 PDF adjunto: Solicitud de Certificado de retenciones Clientes.pdf")

    # ------------------------------
    # Límite horario 5:30 PM
    # ------------------------------
    hora_actual = datetime.now()
    hora_limite = hora_actual.replace(hour=17, minute=30, second=0, microsecond=0)

    if hora_actual > hora_limite:
        novedades.append("Proceso detenido por límite horario (5:30 PM).")
        break

    # ------------------------------
    # Espera entre correos
    # ------------------------------
    wait_time = random.uniform(2, 5)
    print(f"⏱ Esperando {wait_time:.2f} minutos...\n")
    time.sleep(wait_time)

# ------------------------------
# Enviar correo de reporte FINAL
# ------------------------------
reporte = []
reporte.append("REPORTE ENVÍO CORREOS SOLICITUD CERTIFICADOS\n")
reporte.append(f"Correos enviados correctamente: {correos_enviados}\n")

if novedades:
    reporte.append("NOVEDADES DETECTADAS:")
    for n in novedades:
        reporte.append(f"- {n}")
else:
    reporte.append("No se presentaron novedades.")

reporte.append("\nEl proceso de envío ha finalizado correctamente.")

reporte_texto = "\n".join(reporte)

mail_reporte = outlook.CreateItem(0)
mail_reporte.To = "abicdev26@gmail.com"
mail_reporte.Subject = "reporte envio correo de retenciones"
mail_reporte.Body = reporte_texto
mail_reporte.Send()

print("📧 Correo de reporte enviado a abicdev26@gmail.com")
print("\n🎉 Proceso finalizado.")
