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
NOMBRE_ARCHIVO_ESTATICO = os.path.join(RUTA_DESCARGAS, 'BD_prorrogas.xls') # Correcto: .xls

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
    {"x": 115, "y": 560, "delay_after": 4.0, "comment": "Clic en 5to elemento de menú (Consulta de Prorrogas)"} 
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

    # Calcular fechas dinámicamente (Lógica conservada)
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

    # --- INTERACCIÓN CON CHECKBOXES ---
    print("Configurando checkboxes de estados...")
    # Desmarcar "Vigente"
    vigente_checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "vigente")))
    if vigente_checkbox.is_selected():
        vigente_checkbox.click()
        print("Checkbox 'Vigente' desmarcado.")
    else:
        print("Checkbox 'Vigente' ya estaba desmarcado.")

    # Marcar "Cancelado"
    cancelado_checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "cancelado")))
    if not cancelado_checkbox.is_selected():
        cancelado_checkbox.click()
        print("Checkbox 'Cancelado' marcado.")
    else:
        print("Checkbox 'Cancelado' ya estaba marcado.")

    # NOTA: No se interactúa con el checkbox 'Pendiente' ya que se nos informó que está desmarcado por defecto.
    time.sleep(1)

    # --- PROCESO DE DESCARGA Y RENOMBRADO ---
    print("\n" + "="*20 + " PROCESO DE DESCARGA Y RENOMBRADO " + "="*20)

    # 1. Borra el archivo estático si existe antes de iniciar la nueva descarga.
    if os.path.exists(NOMBRE_ARCHIVO_ESTATICO):
        try:
            os.remove(NOMBRE_ARCHIVO_ESTATICO)
            print(f"Archivo antiguo '{NOMBRE_ARCHIVO_ESTATICO}' eliminado.")
        except Exception as e:
            print(f"Advertencia: No se pudo eliminar el archivo antiguo '{NOMBRE_ARCHIVO_ESTATICO}': {e}. Intentando continuar.")

    # 2. Obtener un DICCIONARIO de archivos existentes con el patrón esperado ANTES de la descarga
    # Esto nos permitirá identificar el NUEVO archivo después de la descarga
    patron_descarga_final = os.path.join(RUTA_DESCARGAS, 'Trade_ExcelConsultaProrrogas*.xls')
    # Corregido: Usar un diccionario para almacenar la marca de tiempo de cada archivo
    archivos_antes_descarga = {f: os.path.getmtime(f) for f in glob.glob(patron_descarga_final)}
    print(f"Archivos existentes con patrón '{patron_descarga_final}' antes de la descarga: {len(archivos_antes_descarga)}")

    # --- CLIC PARA EXPORTAR A EXCEL ---
    print("Intentando hacer clic en el icono 'Exportar a Excel'...")
    boton_exportar_excel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Exportar a Excel' and @onclick='crearExcel();']")))
    boton_exportar_excel.click()
    print("Clic en 'Exportar a Excel' realizado.")

    # --- ESPERA POR EL NUEVO ARCHIVO (patrón .xls) ---
    print(f"Esperando la aparición del nuevo archivo descargado con patrón '{patron_descarga_final}' (máximo 180 segundos)...")
    tiempo_max_espera_aparicion = 180 
    tiempo_inicio_espera = time.time()
    archivo_descargado_reciente = None

    while time.time() - tiempo_inicio_espera < tiempo_max_espera_aparicion:
        archivos_despues_descarga = glob.glob(patron_descarga_final)
        
        # Filtrar archivos que son NUEVOS o que tienen una marca de tiempo significativamente más reciente
        nuevos_candidatos = [
            f for f in archivos_despues_descarga 
            # Corregido: Usar .get() en el diccionario, con un valor por defecto de 0 si el archivo no existía antes
            if f not in archivos_antes_descarga or os.path.getmtime(f) > archivos_antes_descarga.get(f, 0)
        ]

        if nuevos_candidatos:
            # Tomar el más reciente de los nuevos candidatos
            archivo_descargado_reciente = max(nuevos_candidatos, key=os.path.getmtime)
            # Dar un pequeño tiempo para asegurar que el archivo esté completamente escrito
            time.sleep(3) 
            print(f"Nuevo archivo detectado: '{os.path.basename(archivo_descargado_reciente)}'")
            break
        time.sleep(1)

    if not archivo_descargado_reciente:
        raise TimeoutError(f"El archivo con patrón '{patron_descarga_final}' no apareció o no fue detectado como nuevo en el directorio de descargas dentro del tiempo límite.")
    
    # --- RENOMBRAR EL ARCHIVO ---
    try:
        os.rename(archivo_descargado_reciente, NOMBRE_ARCHIVO_ESTATICO)
        print(f"Archivo renombrado exitosamente a: '{os.path.basename(NOMBRE_ARCHIVO_ESTATICO)}'")
        print("="*60 + "\n")
    except Exception as rename_e:
        print(f"Error al intentar renombrar el archivo '{os.path.basename(archivo_descargado_reciente)}' a '{os.path.basename(NOMBRE_ARCHIVO_ESTATICO)}': {rename_e}")
        raise 


except Exception as e:
    print(f"\n==== OCURRIÓ UN ERROR GENERAL DURANTE LA AUTOMATIZACIÓN ====")
    print(str(e))
    if driver:
        try:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            screenshot_name = f"error_screenshot_{timestamp}.png"
            driver.save_screenshot(os.path.join(RUTA_DESCARGAS, screenshot_name))
            print(f"Captura de pantalla de error guardada como: {screenshot_name}")
        except Exception as ss_e:
            print(f"No se pudo tomar la captura de pantalla: {ss_e}")

finally:
    if driver:
        print("Cerrando el navegador...")
        driver.quit()
        print("Navegador cerrado.")
    print("--- Fin del Script ---")