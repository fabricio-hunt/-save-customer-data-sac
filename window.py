"""
Mostra posição do mouse em tempo real.
Passe sobre cada campo e anote as coordenadas.
Pressione Ctrl+C para sair.
"""
import time
import pyautogui

print("Mova o mouse sobre os campos. Ctrl+C para sair.\n")
try:
    while True:
        x, y = pyautogui.position()
        print(f"\r  Mouse: ({x:>5}, {y:>5})   ", end="", flush=True)
        time.sleep(0.1)
except KeyboardInterrupt:
    print("\nFim.")