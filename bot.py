import json
import os
import shutil
import socket
import subprocess
import sys
import time
import traceback
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import boto3
import pyautogui
import pygetwindow as gw

from config import (
    ARCHIVE_DIR,
    AWS_ACCESS_KEY,
    AWS_BUCKET,
    AWS_REGION,
    AWS_SECRET_KEY,
    DELETE_SOURCE_OBJECTS,
    DOWNLOADS_DIR,
    FILE_PROCESS_WAIT_SECONDS,
    IMPORT_SEQUENCE_PREFIX,
    IMPORT_SEQUENCE_SUFFIX,
    LOCK_FILE,
    LOGIN_WAIT_SECONDS,
    LOGS_DIR,
    MENU_SEQUENCE,
    MENU_STEP_WAIT_SECONDS,
    S3_ERRORES_PREFIX,
    S3_PEDIDOS_PREFIX,
    S3_RESULTADOS_PREFIX,
    SIESA_PASSWORD,
    SIESA_P99_PATH,
    SIESA_PEDIDOS_PATH,
    SIESA_SHORTCUT_PATH,
    SIESA_USER,
    SIESA_WINDOW_TITLE,
    SIESA_WORKING_DIR,
    STATE_FILE,
)


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2


@dataclass
class PendingOrder:
    file_name: str
    s3_key: str
    local_download_path: Path


@dataclass
class ErrorResult:
    file_name: str
    s3_key: str
    p99_key: str | None
    errors: list[str]


class RpaBot:
    def __init__(self) -> None:
        self.s3_client = boto3.client(
            "s3",
            aws_access_key_id=AWS_ACCESS_KEY,
            aws_secret_access_key=AWS_SECRET_KEY,
            region_name=AWS_REGION,
        )
        self.machine_name = socket.gethostname()
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = LOGS_DIR / f"run_{self.run_id}.log"
        self.state = self._load_state()
        self.run_summary: dict[str, Any] = {
            "run_id": self.run_id,
            "started_at": self._iso_now(),
            "finished_at": None,
            "machine_name": self.machine_name,
            "files_detected": [],
            "files_attempted": [],
            "files_without_error": [],
            "files_with_error": [],
            "fatal_error": None,
            "log_file": str(self.log_path),
        }

    def run(self) -> int:
        self._ensure_directories()
        self._log("Iniciando corrida batch")

        try:
            with self._acquire_lock():
                pending_orders = self._fetch_pending_orders()
                self.run_summary["files_detected"] = [order.file_name for order in pending_orders]

                if not pending_orders:
                    self._log("No hay archivos nuevos por procesar")
                    self._finalize_and_upload_result()
                    return 0

                self._open_siesa()
                self._login()
                self._navigate_to_import_menu()

                for order in pending_orders:
                    self._process_order(order)

                self._finalize_and_upload_result()
                self._persist_state()
                self._log("Corrida finalizada correctamente")
                return 0
        except Exception as exc:  # noqa: BLE001
            self.run_summary["fatal_error"] = str(exc)
            self._log(f"Corrida fallida: {exc}")
            self._finalize_and_upload_result()
            self._persist_state()
            traceback.print_exc()
            return 1

    def _fetch_pending_orders(self) -> list[PendingOrder]:
        self._log("Consultando pedidos pendientes en S3")
        response = self.s3_client.list_objects_v2(Bucket=AWS_BUCKET, Prefix=S3_PEDIDOS_PREFIX)

        orders: list[PendingOrder] = []
        processed_keys = set(self.state.get("processed_keys", {}).keys())

        for item in response.get("Contents", []):
            key = item["Key"]
            if not key.upper().endswith(".PE0"):
                continue
            if key in processed_keys:
                continue

            file_name = Path(key).name
            local_path = DOWNLOADS_DIR / file_name
            self._log(f"Descargando {key} -> {local_path}")
            self.s3_client.download_file(AWS_BUCKET, key, str(local_path))
            orders.append(PendingOrder(file_name=file_name, s3_key=key, local_download_path=local_path))

        return orders

    def _process_order(self, order: PendingOrder) -> None:
        self._log(f"Procesando archivo {order.file_name}")
        self.run_summary["files_attempted"].append(order.file_name)

        target_path = SIESA_PEDIDOS_PATH / order.file_name
        shutil.copy2(order.local_download_path, target_path)

        before_p99 = self._list_p99_files()
        self._import_file(order.file_name)
        time.sleep(FILE_PROCESS_WAIT_SECONDS)
        after_p99 = self._list_p99_files()

        new_p99_files = sorted(after_p99 - before_p99)

        if new_p99_files:
            error_result = self._handle_p99_files(order, new_p99_files)
            self.run_summary["files_with_error"].append({
                "file": error_result.file_name,
                "s3_key": error_result.s3_key,
                "p99_key": error_result.p99_key,
                "errors": error_result.errors,
            })
        else:
            self.run_summary["files_without_error"].append(order.file_name)

        self._mark_processed(order.s3_key, order.file_name)
        self._archive_local_file(order.local_download_path)

        if DELETE_SOURCE_OBJECTS:
            self._delete_source_object(order.s3_key)

        if target_path.exists():
            target_path.unlink()

    def _open_siesa(self) -> None:
        self._log("Abriendo Siesa")
        subprocess.Popen([str(SIESA_SHORTCUT_PATH)], cwd=str(SIESA_WORKING_DIR), shell=True)
        time.sleep(LOGIN_WAIT_SECONDS)
        self._activate_siesa_window()

    def _login(self) -> None:
        self._log("Iniciando sesión")
        self._activate_siesa_window()
        pyautogui.write(SIESA_USER)
        pyautogui.press("tab")
        pyautogui.write(SIESA_PASSWORD)
        pyautogui.press("enter")
        time.sleep(LOGIN_WAIT_SECONDS)

    def _navigate_to_import_menu(self) -> None:
        self._log("Navegando al menú de importación")
        self._activate_siesa_window()
        for key in MENU_SEQUENCE:
            pyautogui.press(key)
            time.sleep(MENU_STEP_WAIT_SECONDS)

    def _import_file(self, file_name: str) -> None:
        self._activate_siesa_window()
        for key in IMPORT_SEQUENCE_PREFIX:
            pyautogui.press(key)
            time.sleep(MENU_STEP_WAIT_SECONDS)

        pyautogui.write(file_name)

        for key in IMPORT_SEQUENCE_SUFFIX:
            pyautogui.press(key)
            time.sleep(MENU_STEP_WAIT_SECONDS)

    def _handle_p99_files(self, order: PendingOrder, p99_files: list[Path]) -> ErrorResult:
        self._log(f"Se detectaron {len(p99_files)} archivo(s) P99 para {order.file_name}")

        all_errors: list[str] = []
        uploaded_key: str | None = None

        for p99_file in p99_files:
            errors = self._parse_p99_errors(p99_file)
            all_errors.extend(errors)

            uploaded_key = (
                f"{S3_ERRORES_PREFIX}{self.run_id}_{Path(order.file_name).stem}_{p99_file.name}"
            )
            self._log(f"Subiendo {p99_file} -> s3://{AWS_BUCKET}/{uploaded_key}")
            self.s3_client.upload_file(str(p99_file), AWS_BUCKET, uploaded_key)
            p99_file.unlink()

        unique_errors = list(dict.fromkeys(all_errors))
        if not unique_errors:
            unique_errors = ["Siesa generó un P99 sin detalle reconocible."]

        return ErrorResult(
            file_name=order.file_name,
            s3_key=order.s3_key,
            p99_key=uploaded_key,
            errors=unique_errors,
        )

    def _parse_p99_errors(self, p99_file: Path) -> list[str]:
        lines = p99_file.read_text(encoding="latin-1", errors="ignore").splitlines()
        errors: list[str] = []

        for line in lines:
            if line.startswith("*"):
                continue

            stripped = line.strip()
            if not stripped:
                continue

            if len(stripped) >= 10 and stripped[:10].isdigit():
                stripped = stripped[10:].strip()

            if stripped:
                errors.append(stripped)

        return errors

    def _finalize_and_upload_result(self) -> None:
        self.run_summary["finished_at"] = self._iso_now()
        result_path = ARCHIVE_DIR / f"run_{self.run_id}.json"
        result_path.write_text(
            json.dumps(self.run_summary, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        s3_key = f"{S3_RESULTADOS_PREFIX}run_{self.run_id}.json"
        self._log(f"Subiendo resultado de corrida a s3://{AWS_BUCKET}/{s3_key}")
        self.s3_client.upload_file(str(result_path), AWS_BUCKET, s3_key)

    def _activate_siesa_window(self) -> None:
        windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
        if not windows:
            raise RuntimeError(f"No se encontró una ventana de Siesa con el título '{SIESA_WINDOW_TITLE}'")

        window = windows[0]
        if window.isMinimized:
            window.restore()
        window.activate()
        time.sleep(1)

    def _list_p99_files(self) -> set[Path]:
        return {path for path in SIESA_P99_PATH.glob("*.P99")}

    def _archive_local_file(self, local_path: Path) -> None:
        archive_name = f"{self.run_id}_{local_path.name}"
        archive_path = ARCHIVE_DIR / archive_name
        shutil.move(str(local_path), archive_path)

    def _delete_source_object(self, s3_key: str) -> None:
        try:
            self.s3_client.delete_object(Bucket=AWS_BUCKET, Key=s3_key)
            self._log(f"Objeto fuente eliminado de S3: {s3_key}")
        except Exception as exc:  # noqa: BLE001
            self._log(f"No se pudo eliminar {s3_key} de S3: {exc}")

    def _mark_processed(self, s3_key: str, file_name: str) -> None:
        self.state.setdefault("processed_keys", {})[s3_key] = {
            "file_name": file_name,
            "run_id": self.run_id,
            "processed_at": self._iso_now(),
        }

    def _ensure_directories(self) -> None:
        for path in [DOWNLOADS_DIR, ARCHIVE_DIR, LOGS_DIR]:
            path.mkdir(parents=True, exist_ok=True)

        if not SIESA_SHORTCUT_PATH.exists():
            raise RuntimeError(f"No se encontró el acceso directo de Siesa: {SIESA_SHORTCUT_PATH}")

        for path in [SIESA_WORKING_DIR, SIESA_PEDIDOS_PATH, SIESA_P99_PATH]:
            if not path.exists():
                raise RuntimeError(f"No existe la ruta requerida de Siesa: {path}")

    def _load_state(self) -> dict[str, Any]:
        if not STATE_FILE.exists():
            return {"processed_keys": {}}

        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {"processed_keys": {}}

    def _persist_state(self) -> None:
        STATE_FILE.write_text(json.dumps(self.state, indent=2, ensure_ascii=False), encoding="utf-8")

    def _log(self, message: str) -> None:
        timestamped = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
        print(timestamped)
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(timestamped + "\n")

    @contextmanager
    def _acquire_lock(self):
        LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
        fd = None

        try:
            fd = os.open(str(LOCK_FILE), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("utf-8"))
            yield
        finally:
            if fd is not None:
                os.close(fd)
            if LOCK_FILE.exists():
                LOCK_FILE.unlink()

    def _iso_now(self) -> str:
        return datetime.now().astimezone().isoformat()


def main() -> int:
    bot = RpaBot()
    return bot.run()


if __name__ == "__main__":
    sys.exit(main())
