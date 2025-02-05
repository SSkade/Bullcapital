import os
import fitz  # PyMuPDF

# Obtener el directorio del escritorio del usuario
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Verificar si existe una carpeta llamada "practicante" en el escritorio
practicante_dir = os.path.join(desktop, "practicante")
if os.path.exists(practicante_dir):
    script_dir = practicante_dir
else:
    script_dir = desktop

# Subcarpetas principales
main_folders = ["CORDADA", "FINAMERIS", "LATAM"]

# Crear las carpetas principales si no existen
for folder_name in main_folders:
    folder_path = os.path.join(script_dir, folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

# Función para abrir un archivo PDF y devolver el objeto del documento
def open_pdf(file_path):
    try:
        doc = fitz.open(file_path)
        return doc
    except Exception as e:
        print(f"Error al abrir el archivo PDF: {e}")
        return None

# Función para leer el contenido de un archivo PDF
def read_pdf(file_path):
    doc = open_pdf(file_path)
    if doc is None:
        return None
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# Función para convertir un PDF a texto y guardarlo en un archivo .txt
def convert_pdf_to_txt(pdf_path, txt_path):
    text = read_pdf(pdf_path)
    if text is not None:
        with open(txt_path, 'w', encoding='latin-1') as txt_file:
            txt_file.write(text)
        print(f"Archivo TXT guardado en: {txt_path}")
    else:
        print(f"Error al convertir el archivo PDF: {pdf_path}")

# Ejemplo de uso: convertir todos los PDFs en las carpetas principales a archivos TXT
for folder_name in main_folders:
    folder_path = os.path.join(script_dir, folder_name)
    pdf_folder = os.path.join(folder_path, "pdfs")
    txt_folder = os.path.join(folder_path, "txts")
    
    # Crear la carpeta "txts" si no existe
    if not os.path.exists(txt_folder):
        os.makedirs(txt_folder)
    
    # Verificar si la carpeta "pdfs" existe
    if not os.path.exists(pdf_folder):
        print(f"No se encontró la carpeta 'pdfs' en {folder_path}.")
        continue
    
    # Convertir todos los archivos PDF en la carpeta "pdfs" a archivos TXT
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            txt_path = os.path.join(txt_folder, os.path.splitext(pdf_file)[0] + ".txt")
            convert_pdf_to_txt(pdf_path, txt_path)