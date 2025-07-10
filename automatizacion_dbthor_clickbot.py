import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
from datetime import date # Importamos 'date' para manejar las fechas

# --- CONFIGURACIÓN GENERAL ---
# Asegúrate de que esta ruta a tu chromedriver sea correcta.
PATH_A_TU_CHROMEDRIVER = r"C:\Users\userb\OneDrive\Documents\chromedriver-win64\chromedriver-win64\chromedriver.exe" 
# Ruta donde se guardan los archivos descargados. ¡MODIFICA ESTO A TU CARPETA REAL!
RUTA_DESCARGAS = r"C:\Users\userb\Downloads" 
# Nombre final y estático que tendrá el archivo para que Power BI lo encuentre.
NOMBRE_ARCHIVO_ESTATICO = os.path.join(RUTA_DESCARGAS, 'CRM_Export_Diario.xlsx')

URL_LOGIN = "https://cloud.dbthor.com/BullCapital/HomeAlt.jsp"
TU_USUARIO = "jrojas" 
TU_CONTRASENA = "cafeconron" 

# --- SECUENCIA DE CLICS (Coordenadas para PyAutoGUI) ---
clicks_data = [
    {"x": 33,  "y": 160, "delay_after": 2.5, "comment": "Abrir menú principal (div#thesee)"},
    {"x": 59,  "y": 304, "delay_after": 2.0, "comment": "Clic en 1er elemento de menú (ej. Factoring)"},
    {"x": 64,  "y": 380, "delay_after": 2.0, "comment": "Clic en 2do elemento de menú (ej. Tesorería)"},
    {"x": 78,  "y": 400, "delay_after": 2.0, "comment": "Clic en 3er elemento de menú (ej. Operaciones)"},
    {"x": 99,  "y": 440, "delay_after": 2.0, "comment": "Clic en 4to elemento de menú (ej. Consultas)"},
    {"x": 115, "y": 510, "delay_after": 4.0, "comment": "Clic en 5to elemento de menú (ej. Operaciones Diarias - final)"}
]
# ----------------------------------------------------

driver = None
try:
    # --- 1. INICIALIZACIÓN DE SELENIUM Y LOGIN ---
    print("Iniciando el servicio de ChromeDriver...")
    chrome_service = ChromeService(executable_path=PATH_A_TU_CHROMEDRIVER)
    driver = webdriver.Chrome(service=chrome_service)
    print("Navegador Chrome iniciado.")
    driver.maximize_window()
    original_window = driver.current_window_handle
    
    print(f"Navegando a la página de login: {URL_LOGIN}")
    driver.get(URL_LOGIN)

    print("Ingresando credenciales...")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "Usuario"))).send_keys(TU_USUARIO)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "Contrasena"))).send_keys(TU_CONTRASENA)
    
    print("Haciendo clic en el botón 'Entrar'...")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and normalize-space()='Entrar']"))).click()
    print("Clic en 'Entrar' ejecutado.")
    
    # ... (El resto de tu código de login y cambio de ventana se mantiene igual)
    print("Esperando a que aparezca la segunda ventana/pestaña (máx 20 segundos)...")
    WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(2))
    all_windows = driver.window_handles
    new_window_handle = next((handle for handle in all_windows if handle != original_window), None)
    
    if new_window_handle:
        driver.switch_to.window(new_window_handle)
        print(f"Cambiado exitosamente a la nueva ventana. Título actual: {driver.title}")
    else:
        raise Exception("Nueva ventana no encontrada después del login.")

    print("Esperando a que la URL en la nueva ventana contenga 'MainFrame.jsp'...")
    WebDriverWait(driver, 40).until(EC.url_contains("MainFrame.jsp"))
    print("Página principal cargada.")
    time.sleep(2)

    # --- 4. PREPARACIÓN PARA CLICS CON PYAUTOGUI ---
    print("\n" + "-"*15 + " PREPARACIÓN PARA PYAUTOGUI " + "-"*15)
    print("PyAutoGUI comenzará los clics en 10 segundos. Asegúrate de que la ventana correcta tenga el foco.")
    time.sleep(5) 

    # --- 5. EJECUCIÓN DE SECUENCIA DE CLICS CON PYAUTOGUI ---
    print("--- Iniciando secuencia de clics con PyAutoGUI ---")
    for i, click_info in enumerate(clicks_data):
        pyautogui.moveTo(click_info["x"], click_info["y"], duration=0.25) 
        pyautogui.click(click_info["x"], click_info["y"])
        print(f"Clic {i+1} ejecutado: {click_info['comment']}. Esperando {click_info['delay_after']}s...")
        time.sleep(click_info['delay_after'])
    print("--- Secuencia de clics con PyAutoGUI completada ---\n")
    time.sleep(1)

    # --- 6. INTERACCIÓN CON PÁGINA FINAL USANDO SELENIUM ---
    print("-" * 30)
    print("INTERACTUANDO CON LA PÁGINA FINAL USANDO SELENIUM")
    driver.switch_to.default_content()
    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "centro")))
    print("Cambiado al frame 'centro'.")
    time.sleep(1)

    # Calcular fechas dinámicamente
    fecha_actual = date.today() 
    fecha_inicio_anual = fecha_actual.replace(month=1, day=1) 
    formato_fecha = "%d/%m/%Y" 
    start_date_str = fecha_inicio_anual.strftime(formato_fecha) 
    end_date_str = fecha_actual.strftime(formato_fecha) 
    print(f"Estableciendo rango de fechas de '{start_date_str}' a '{end_date_str}'.")
    
    # Rellenar campos de fecha
    campo_fecha_inicial = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "fechainicial")))
    driver.execute_script(f"arguments[0].value = '{start_date_str}';", campo_fecha_inicial)
    
    campo_fecha_final = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "fechafinal")))
    driver.execute_script(f"arguments[0].value = '{end_date_str}';", campo_fecha_final)
    print("Campos de fecha rellenados.")
    time.sleep(1)

    # --- CLIC PARA EXPORTAR A EXCEL ---
    print("Intentando hacer clic en el icono 'Exportar a Excel'...")
    boton_exportar_excel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Exportar a Excel' and @onclick='excel();']")))
    boton_exportar_excel.click()
    print("Clic en 'Exportar a Excel' realizado.")

    # --- CORREGIDO: ESPERA INTELIGENTE PARA LA DESCARGA (VERSIÓN 2) ---
    print("Esperando a que la descarga comience (máximo 20 segundos)...")
    tiempo_max_espera_inicio = 180
    tiempo_inicio = time.time()
    archivo_temporal_encontrado = False

    while time.time() - tiempo_inicio < tiempo_max_espera_inicio:
        # Busca si ya existe un archivo temporal .crdownload
        if glob.glob(os.path.join(RUTA_DESCARGAS, '*.crdownload')):
            print("Descarga iniciada. Se ha detectado un archivo .crdownload.")
            archivo_temporal_encontrado = True
            break
        time.sleep(1)

    if not archivo_temporal_encontrado:
        raise FileNotFoundError("La descarga no comenzó. No se encontró el archivo .crdownload en el tiempo esperado.")

    # Ahora, esperamos a que el archivo .crdownload desaparezca (lo que significa que la descarga terminó)
    print("Esperando a que la descarga finalice (máximo 180 segundos)...") # TIEMPO AUMENTADO
    tiempo_max_espera_fin = 180  # AUMENTADO A 3 MINUTOS
    tiempo_inicio = time.time()
    descarga_completa = False

    while time.time() - tiempo_inicio < tiempo_max_espera_fin:
        # Si ya NO hay archivos .crdownload, la descarga ha terminado.
        if not glob.glob(os.path.join(RUTA_DESCARGAS, '*.crdownload')):
            print("Descarga finalizada. El archivo .crdownload ha desaparecido.")
            descarga_completa = True
            time.sleep(3)  # Pausa de seguridad para que el sistema libere el archivo completamente
            break
        print("Descarga en progreso...")
        time.sleep(2)

    if not descarga_completa:
        raise TimeoutError("La descarga del archivo del CRM excedió el tiempo límite de 180 segundos.")
    # --- FIN DE LA ESPERA INTELIGENTE ---
    
    
    print("\n" + "="*20 + " PROCESO DE RENOMBRADO DE ARCHIVO " + "="*20)
    # --- 7. ENCONTRAR Y RENOMBRAR EL ARCHIVO DESCARGADO ---
    # 1. Borra el archivo estático del día anterior si existe.
    if os.path.exists(NOMBRE_ARCHIVO_ESTATICO):
        os.remove(NOMBRE_ARCHIVO_ESTATICO)
        print(f"Archivo antiguo '{NOMBRE_ARCHIVO_ESTATICO}' eliminado.")

    # 2. Encuentra el archivo recién descargado usando el patrón.
    patron_busqueda = os.path.join(RUTA_DESCARGAS, 'Trade_ExcelOperacionesDiarias*.xls*')
    lista_de_archivos = glob.glob(patron_busqueda)

    if not lista_de_archivos:
        raise FileNotFoundError("Error crítico: No se encontró el archivo del CRM descargado para renombrar.")
    else:
        # Encuentra el archivo más reciente de la lista.
        archivo_mas_reciente = max(lista_de_archivos, key=os.path.getctime)
        print(f"Archivo nuevo encontrado: '{os.path.basename(archivo_mas_reciente)}'")
        
        # 3. Renombra el archivo nuevo al nombre estático.
        os.rename(archivo_mas_reciente, NOMBRE_ARCHIVO_ESTATICO)
        print(f"Archivo renombrado exitosamente a: '{os.path.basename(NOMBRE_ARCHIVO_ESTATICO)}'")
        print("="*60 + "\n")


except Exception as e:
    print(f"\n==== OCURRIÓ UN ERROR GENERAL DURANTE LA AUTOMATIZACIÓN ====")
    print(str(e))
    # ... (el resto del manejo de errores se mantiene igual)

finally:
    if driver:
        print("Cerrando el navegador...")
        driver.quit()
        print("Navegador cerrado.")
    print("--- Fin del Script ---")
