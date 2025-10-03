import pyautogui
import time

print("Pressione Ctrl + C para parar o programa.\n")

try:
    while True:
        x, y = pyautogui.position()  # captura coordenadas
        print(f"Posição atual do mouse: X={x} Y={y}", end="\r")  
        time.sleep(0.1)  # evita consumir 100% da CPU
except KeyboardInterrupt:
    print("\nPrograma encerrado.")