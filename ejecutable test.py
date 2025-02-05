import os
import subprocess

# Obtener el directorio del escritorio del usuario
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Verificar si existe una carpeta llamada "practicante" en el escritorio
practicante_dir = os.path.join(desktop, "practicante")
if not os.path.exists(practicante_dir):
    raise FileNotFoundError("No se encontró la carpeta 'practicante' en el escritorio.")

# Definir el directorio base como la carpeta "practicante"
base_directory = practicante_dir

# Lista de scripts a ejecutar en el orden especificado
scripts = [
    "Captacion de mail Relavizado LISTO.py",
    "Conversor pdf-txt relativizado LISTO.py",
    "latam relativizado.py",
    "scraping cordada relativizado R.py",
    "Scraping Finameris relativizado.py"
]

# Ejecutar cada script en el orden especificado
for script in scripts:
    script_path = os.path.join(base_directory, script)
    if not os.path.exists(script_path):
        raise FileNotFoundError(f"No se encontró el script: {script_path}")
    
    result = subprocess.run(['python', script_path], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error al ejecutar el script {script}: {result.stderr}")
    else:
        print(f"Script {script} ejecutado correctamente: {result.stdout}")

print("Ejecución de scripts completada.")