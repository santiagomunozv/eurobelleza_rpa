import os
import glob
import pyautogui
import time
import traceback
import boto3
from datetime import datetime

from config import *

INTERVALO = 1800

s3_client = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_KEY,
    region_name=AWS_REGION,
)


def descargar_pedidos_s3():
    print("Descargando pedidos de S3...")
    respuesta = s3_client.list_objects_v2(Bucket=AWS_BUCKET, Prefix="pedidos/")

    if "Contents" not in respuesta:
        print("No hay pedidos en S3")
        return

    os.makedirs(SIESA_PEDIDOS_PATH, exist_ok=True)

    for objeto in respuesta["Contents"]:
        clave = objeto["Key"]
        if not clave.upper().endswith(".PE0"):
            continue
        nombre_archivo = os.path.basename(clave)
        ruta_local = os.path.join(SIESA_PEDIDOS_PATH, nombre_archivo)
        s3_client.download_file(AWS_BUCKET, clave, ruta_local)
        print(f"Descargado: {nombre_archivo}")


def obtener_archivos_pe0():
    patron = os.path.join(SIESA_PEDIDOS_PATH, "*.PE0")
    archivos = glob.glob(patron)
    return [os.path.basename(f) for f in sorted(archivos)]


def subir_error_p99(pedido_nombre):
    archivos_p99 = glob.glob(os.path.join(P99_DIR, "*.P99"))
    if not archivos_p99:
        return
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_sin_extension = os.path.splitext(pedido_nombre)[0]
    for ruta_p99 in archivos_p99:
        nombre_p99 = os.path.basename(ruta_p99)
        clave_s3 = f"errores/{timestamp}_{nombre_sin_extension}_{nombre_p99}"
        s3_client.upload_file(ruta_p99, AWS_BUCKET, clave_s3)
        print(f"Error subido a S3: {clave_s3}")
        os.remove(ruta_p99)


def borrar_pedido_s3(nombre_archivo):
    clave_s3 = f"pedidos/{nombre_archivo}"
    s3_client.delete_object(Bucket=AWS_BUCKET, Key=clave_s3)
    print(f"Borrado de S3: {clave_s3}")


def borrar_archivo_local(nombre_archivo):
    ruta = os.path.join(SIESA_PEDIDOS_PATH, nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)


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


def rec_archivo(nombre_archivo):
    print(f"Recepción vía archivo: {nombre_archivo}")

    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("f2")

    time.sleep(2)

    pyautogui.write(nombre_archivo)
    pyautogui.press("enter")
    pyautogui.press("1")
    pyautogui.press("1")
    pyautogui.press("D")
    pyautogui.press("enter")
    pyautogui.press("0")
    pyautogui.press("S")
    pyautogui.press("enter")
    pyautogui.press("f10")

    time.sleep(TIEMPO_PROCESO)


def ejecutar_bot():
    descargar_pedidos_s3()
    archivos_pe0 = obtener_archivos_pe0()

    if not archivos_pe0:
        print("No hay archivos .PE0 para procesar")
        return

    print(f"Procesando {len(archivos_pe0)} archivo(s)")
    abrir_siesa()
    login()
    nav_menu()

    for archivo in archivos_pe0:
        print(f"Procesando: {archivo}")
        rec_archivo(archivo)
        subir_error_p99(archivo)
        borrar_pedido_s3(archivo)
        borrar_archivo_local(archivo)
        print(f"Completado: {archivo}")

    print("Todos los pedidos procesados")


def main():
    while True:
        print("Iniciando")
        ejecutar_bot()
        print(f"Esperando {INTERVALO} segundos para la siguiente ejecución")
        time.sleep(INTERVALO)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\nPresiona Enter para salir...")