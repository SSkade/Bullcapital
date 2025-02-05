import os
import win32com.client

# Configuración
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 corresponde a la bandeja de entrada

# Subcarpetas a buscar dentro de la bandeja de entrada
subfolders = ["CORDADA", "LATAM"]

# Obtener el directorio del escritorio del usuario
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Verificar si existe una carpeta llamada "practicante" en el escritorio
practicante_dir = os.path.join(desktop, "practicante")
if os.path.exists(practicante_dir):
    script_dir = practicante_dir
else:
    script_dir = desktop

# Crear las carpetas "CORDADA" y "LATAM" si no existen
for subfolder_name in subfolders:
    output_dir = os.path.join(script_dir, subfolder_name)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # Crear las carpetas "pdfs" y "excels" dentro de cada subcarpeta si no existen
    pdf_dir = os.path.join(output_dir, "pdfs")
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
    excels_dir = os.path.join(output_dir, "excels")
    if not os.path.exists(excels_dir):
        os.makedirs(excels_dir)

# Función para procesar los correos electrónicos en una subcarpeta
def procesar_subcarpeta(subfolder_name):
    subfolder = inbox.Folders[subfolder_name]
    messages = subfolder.Items
    output_dir = os.path.join(script_dir, subfolder_name)  # Directorio específico para cada subcarpeta
    for message in messages:
        if not message.UnRead:
            continue  # Saltar los mensajes que ya están leídos
        if message.Attachments.Count > 0:
            for attachment in message.Attachments:
                # Verificar si el archivo adjunto es un PDF o un Excel
                if attachment.FileName.endswith(".pdf"):
                    # Guardar el archivo PDF en la carpeta "pdfs"
                    save_path = os.path.join(output_dir, "pdfs", attachment.FileName)
                elif attachment.FileName.endswith(".xlsx"):
                    # Guardar el archivo Excel en la carpeta "excels"
                    save_path = os.path.join(output_dir, "excels", attachment.FileName)
                else:
                    continue  # Saltar archivos que no sean PDF o Excel
                # Guardar el archivo adjunto
                attachment.SaveAsFile(save_path)
                print(f"Archivo guardado en {subfolder_name}: {attachment.FileName}")

# Procesar cada subcarpeta
for subfolder_name in subfolders:
    procesar_subcarpeta(subfolder_name)

print("Descarga de archivos adjuntos completada.")