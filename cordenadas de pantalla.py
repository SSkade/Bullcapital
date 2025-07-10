import pyautogui
import time

print("Mueve el ratón a la posición deseada en la pantalla.")
print("Las coordenadas X, Y se actualizarán continuamente aquí.")
print("Presiona Ctrl-C en esta ventana de terminal para detener el script una vez que tengas las coordenadas.")
print("-" * 30)

try:
    while True:
        # Obtiene la posición actual del ratón
        x, y = pyautogui.position()
        
        # Prepara el string para mostrar las coordenadas
        position_str = f"X: {x:4d}  Y: {y:4d}"
        
        # Imprime las coordenadas. El uso de \b y flush=True
        # es para que se actualice en la misma línea en la terminal.
        print(position_str, end='')
        print('\b' * len(position_str), end='', flush=True)
        
        # Pequeña pausa para no sobrecargar la CPU
        time.sleep(0.1)
except KeyboardInterrupt:
    # Esto se ejecuta cuando presionas Ctrl-C
    print("\n" + "-" * 30)
    print("Script detenido por el usuario.")
    # Muestra la última posición capturada, si es útil
    try:
        final_x, final_y = pyautogui.position()
        print(f"Última posición registrada: X: {final_x:4d}  Y: {final_y:4d}")
    except Exception: # Por si acaso hay algún problema al leer la posición al salir
        pass 
    print("Anota estas coordenadas para usarlas en tu script principal.")