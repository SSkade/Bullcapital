import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService  
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
from datetime import date

# --- CONFIGURACIÓN GENERAL ---
PATH_A_TU_CHROMEDRIVER = r"C:\Users\userb\OneDrive\Documents\chromedriver-win64\chromedriver-win64\chromedriver.exe" 
RUTA_DESCARGAS = r"C:\Users\userb\Downloads" 
NOMBRE_ARCHIVO_ESTATICO = os.path.join(RUTA_DESCARGAS, 'cobranza_CRM.xlsx')

URL_LOGIN = "https://cloud.dbthor.com/BullCapital/HomeAlt.jsp"
TU_USUARIO = "jrojas" 
TU_CONTRASENA = "cafeconron" 

clicks_data = [
    {"x": 33,  "y": 160, "delay_after": 2.5, "comment": "Abrir menú principal"},
    {"x": 59,  "y": 304, "delay_after": 2.0, "comment": "Clic en Factoring"},
    {"x": 64,  "y": 380, "delay_after": 2.0, "comment": "Clic en Tesorería"},
    {"x": 78,  "y": 400, "delay_after": 2.0, "comment": "Clic en Operaciones"},
    {"x": 99,  "y": 440, "delay_after": 2.0, "comment": "Clic en Consultas"},
    {"x": 115, "y": 455, "delay_after": 4.0, "comment": "Clic en dcto cancelados"}
]
# ----------------------------------------------------

driver = None
try:
    # --- 1. LOGIN Y NAVEGACIÓN ---
    print("Iniciando automatización...")
    chrome_service = ChromeService(executable_path=PATH_A_TU_CHROMEDRIVER)
    driver = webdriver.Chrome(service=chrome_service)
    driver.maximize_window()
    
    print(f"Navegando a {URL_LOGIN}")
    driver.get(URL_LOGIN)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "Usuario"))).send_keys(TU_USUARIO)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "Contrasena"))).send_keys(TU_CONTRASENA)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and normalize-space()='Entrar']"))).click()
    
    print("Login exitoso. Esperando nueva ventana...")
    WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(2))
    new_window_handle = next(handle for handle in driver.window_handles if handle != driver.current_window_handle)
    driver.switch_to.window(new_window_handle)
    
    WebDriverWait(driver, 40).until(EC.url_contains("MainFrame.jsp"))
    print("Página principal cargada.")
    time.sleep(5)

    # --- 2. NAVEGACIÓN CON PYAUTOGUI ---
    print("Iniciando clics de PyAutoGUI...")
    for click_info in clicks_data:
        pyautogui.moveTo(click_info["x"], click_info["y"], duration=0.25)
        pyautogui.click()
        print(f"Clic realizado: {click_info['comment']}")
        time.sleep(click_info['delay_after'])
    print("Navegación del menú completada.")
    
    # --- 3. INTERACCIÓN CON FORMULARIO ---
    driver.switch_to.default_content()
    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "centro")))
    print("Dentro del frame 'centro'.")
    
    opcion_cancelado = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//select[@name='estadoseleccionado']/option[@value='4']"))
    )
    
    actions = ActionChains(driver)
    actions.double_click(opcion_cancelado).perform()
    print("Doble clic en 'Cancelado' realizado.")
    time.sleep(2)
    
    boton_exportar_excel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@onclick, \"crearExcel('si')\")]")))
    boton_exportar_excel.click()
    print("Clic en 'Exportar a Excel' realizado.")

    # --- 4. ESPERA INTELIGENTE (VERSIÓN MEJORADA) ---
    print("Esperando a que la descarga comience (máximo 20 segundos)...")
    tiempo_max_espera_inicio = 20
    tiempo_inicio = time.time()
    archivo_temporal_encontrado = False
    while time.time() - tiempo_inicio < tiempo_max_espera_inicio:
        if glob.glob(os.path.join(RUTA_DESCARGAS, '*.crdownload')):
            print("Descarga iniciada. Se ha detectado un archivo .crdownload.")
            archivo_temporal_encontrado = True
            break
        time.sleep(1)

    if not archivo_temporal_encontrado:
        raise FileNotFoundError("La descarga no comenzó. No se encontró el archivo .crdownload en el tiempo esperado.")

    # Ahora, esperamos a que el archivo .crdownload desaparezca (lo que significa que la descarga terminó)
    print("Esperando a que la descarga finalice (máximo 180 segundos)...")
    tiempo_max_espera_fin = 180
    tiempo_inicio = time.time()
    descarga_completa = False
    while time.time() - tiempo_inicio < tiempo_max_espera_fin:
        if not glob.glob(os.path.join(RUTA_DESCARGAS, '*.crdownload')):
            print("Descarga finalizada. El archivo .crdownload ha desaparecido.")
            descarga_completa = True
            time.sleep(3)  # Pausa de seguridad para que el sistema libere el archivo completamente
            break
        print("Descarga en progreso...")
        time.sleep(2)

    if not descarga_completa:
        raise TimeoutError("La descarga del archivo del CRM excedió el tiempo límite de 180 segundos.")
    
    # --- 5. RENOMBRADO DE ARCHIVO ---
    print("Iniciando proceso de renombrado...")
    if os.path.exists(NOMBRE_ARCHIVO_ESTATICO):
        os.remove(NOMBRE_ARCHIVO_ESTATICO)
        print(f"Archivo antiguo '{os.path.basename(NOMBRE_ARCHIVO_ESTATICO)}' eliminado.")

    # PATRÓN CORREGIDO: Busca cualquier archivo de Excel que contenga la palabra "cartola".
    patron_busqueda = os.path.join(RUTA_DESCARGAS, '*cartola*.xls*')
    lista_de_archivos = glob.glob(patron_busqueda)

    if not lista_de_archivos:
        raise FileNotFoundError(f"Error: No se encontró ningún archivo con el patrón '{patron_busqueda}'.")
    else:
        archivo_mas_reciente = max(lista_de_archivos, key=os.path.getctime)
        print(f"Archivo nuevo encontrado: '{os.path.basename(archivo_mas_reciente)}'")
        os.rename(archivo_mas_reciente, NOMBRE_ARCHIVO_ESTATICO)
        print(f"Archivo renombrado exitosamente a: '{os.path.basename(NOMBRE_ARCHIVO_ESTATICO)}'")
    
    print("Proceso completado.")

except Exception as e:
    print(f"\n==== OCURRIÓ UN ERROR ====\n{e}")

finally:
    if driver:
        driver.quit()
        print("Navegador cerrado.")
