import pyautogui
import time
import subprocess

from config import *

INTERVALO = 3600

def abrir_siesa():

    print("Abriendo...")

    subprocess.Popen(SIESA_PATH)

    time.sleep(TIEMPO_CARGA)


def login():

    print("Login")

    pyautogui.write(USUARIO)
    pyautogui.press("tab")

    pyautogui.write(PASSWORD)
    pyautogui.press("enter")

    time.sleep(10)


def nav_menu():

    print("Navegando...")

    pyautogui.hotkey("alt", "c")  # Comercial
    time.sleep(2)

    pyautogui.press("v")  # Ventas
    time.sleep(2)

    pyautogui.press("d")  # Estándar
    time.sleep(2)

    pyautogui.press("p")  # Pedidos de venta
    time.sleep(5)

    pyautogui.press("v")  # recepción via archivo
    time.sleep(5)


def rec_archivo():

    print("Recepción vía archivo")

    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("f2")

    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("s")


def ejecutar_bot():

    abrir_siesa()
    login()
    nav_menu()
    rec_archivo()

    print("Finalizando...")


def main():

    while True:

        print("Iniciando")

        ejecutar_bot()

        print(f"Esperando {INTERVALO} segundos para la siguiente ejecución")

        time.sleep(INTERVALO)


if __name__ == "__main__":
    main()