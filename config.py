from pathlib import Path

# Acceso a Siesa
SIESA_SHORTCUT_PATH = Path(r"D:\Escritorioo\SIESA.85 - GEB.lnk")
SIESA_WORKING_DIR = Path(r"U:\uno85c")
SIESA_PEDIDOS_PATH = Path(r"U:\uno85c\eurobelleza\trm")
SIESA_P99_PATH = Path(r"U:\uno85c\eurobelleza\prt")
SIESA_WINDOW_TITLE = "UNO8L"

SIESA_USER = "PAGINA"
SIESA_PASSWORD = "PAGINA"

# Secuencias de teclado. Ajustar si el menú cambia.
LOGIN_WAIT_SECONDS = 8
MENU_STEP_WAIT_SECONDS = 2
FILE_PROCESS_WAIT_SECONDS = 6

MENU_SEQUENCE = ["c", "v", "d", "p", "v"]
IMPORT_SEQUENCE_PREFIX = ["enter", "enter", "enter", "enter", "enter", "enter", "enter", "f2"]
IMPORT_SEQUENCE_SUFFIX = ["enter", "1", "1", "D", "99", "0", "S", "enter", "f10"]

# Carpeta de trabajo local del bot
BOT_WORKDIR = Path(r"D:\Escritorioo\eurobelleza_rpa")
DOWNLOADS_DIR = BOT_WORKDIR / "downloads"
ARCHIVE_DIR = BOT_WORKDIR / "archive"
LOGS_DIR = BOT_WORKDIR / "logs"
STATE_FILE = BOT_WORKDIR / "state.json"
LOCK_FILE = BOT_WORKDIR / "run.lock"

# AWS / S3
AWS_ACCESS_KEY = "REEMPLAZAR_ACCESS_KEY"
AWS_SECRET_KEY = "REEMPLAZAR_SECRET_KEY"
AWS_REGION = "us-east-2"
AWS_BUCKET = "eurobelleza-siesa"
S3_PEDIDOS_PREFIX = "pedidos/"
S3_ERRORES_PREFIX = "errores/"
S3_RESULTADOS_PREFIX = "resultados/"

# Mantener en False mientras la policy de Windows no tenga DeleteObject sobre pedidos/
DELETE_SOURCE_OBJECTS = False
