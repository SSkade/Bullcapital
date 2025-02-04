import os
import fitz  # PyMuPDF

# Directorio del script
script_dir = os.path.dirname(__file__)

# Subcarpetas principales
main_folders = ["CORDADA", "FINAMERIS", "LATAM"]

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
    if doc:
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text

# Iterar sobre las carpetas principales
for folder in main_folders:
    pdf_directory = os.path.join(script_dir, folder, "pdfs")
    output_directory = os.path.join(script_dir, folder, "txts")

    # Crear el directorio de salida si no existe
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    try:
        # Iterar sobre todos los archivos en el directorio de PDFs
        for filename in os.listdir(pdf_directory):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(pdf_directory, filename)
                txt_path = os.path.join(output_directory, filename.replace(".pdf", ".txt"))

                # Leer el PDF y extraer el texto
                text = read_pdf(pdf_path)

                # Guardar el texto en un archivo .txt
                with open(txt_path, "w", encoding="latin-1") as txt_file:
                    txt_file.write(text)
    except FileNotFoundError:
        print(f"Directorio no encontrado: {pdf_directory}")
        continue

print("Extracción completada.")