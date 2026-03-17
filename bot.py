import os
import pyautogui
import time
import traceback

from config import *

INTERVALO = 3600

def abrir_siesa():
    print("Abriendo...")
    os.startfile(SIESA_PATH)
    time.sleep(TIEMPO_CARGA)


def login():
    print("Login")
    pyautogui.write(USUARIO)
    pyautogui.press("tab")
    pyautogui.write(PASSWORD)
    pyautogui.press("enter")
    time.sleep(5)


def nav_menu():
    print("Navegando...")
    pyautogui.hotkey("c")
    time.sleep(2)

    pyautogui.press("v")
    time.sleep(2)

    pyautogui.press("d")
    time.sleep(2)

    pyautogui.press("p")
    time.sleep(5)

    pyautogui.press("v")
    time.sleep(5)


def rec_archivo():
    print("Recepción vía archivo")

    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("f2")

    time.sleep(2)

    pyautogui.press("enter")
    pyautogui.press("1")
    pyautogui.press("1")
    pyautogui.press("D")
    pyautogui.press("enter")
    pyautogui.press("0")
    pyautogui.press("S")
    pyautogui.press("enter")
    pyautogui.press("f10")

    time.sleep(4)


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


if _name_ == "_main_":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\nPresiona Enter para salir...")