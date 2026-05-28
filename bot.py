import json
import re
import os
import shutil
import socket
import subprocess
import sys
import time
import traceback
import ctypes
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import boto3
import pyautogui
import pygetwindow as gw
from PIL import Image, ImageChops, ImageOps, ImageStat

from config import (
    ARCHIVE_DIR,
    AWS_ACCESS_KEY,
    AWS_BUCKET,
    AWS_REGION,
    AWS_SECRET_KEY,
    DOWNLOADS_DIR,
    FILE_PROCESS_WAIT_SECONDS,
    IMPORT_SEQUENCE_PREFIX,
    IMPORT_SCREENSHOT_PATTERN,
    IMPORT_SEQUENCE_SUFFIX,
    LOCK_FILE,
    LOGIN_SCREENSHOT_PATTERN,
    LOGIN_WAIT_SECONDS,
    LOGS_DIR,
    MENU_SEQUENCE,
    MENU_STEP_WAIT_SECONDS,
    P97_GENERATE_WITHOUT_ORDERS,
    P97_LOOKBACK_DAYS,
    P97_MENU_SEQUENCE,
    P97_REPORT_ENABLED,
    P97_REPORT_FILE_NAME,
    P97_REPORT_SEQUENCE,
    SCREENSHOTS_DIR,
    SCREEN_CHECK_INTERVAL_SECONDS,
    SCREEN_MATCH_CONFIDENCE,
    SCREEN_CHECK_TIMEOUT_SECONDS,
    SCREEN_VISUAL_SIMILARITY_THRESHOLD,
    SIESA_OPEN_TIMEOUT_SECONDS,
    SIESA_FORCE_MAXIMIZE,
    SIESA_RESET_WINDOW_LAYOUT,
    S3_ERRORES_PREFIX,
    S3_CONFIRMACIONES_PREFIX,
    S3_PEDIDOS_PREFIX,
    S3_RESULTADOS_PREFIX,
    SIESA_PASSWORD,
    SIESA_P97_PATH,
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
PyAutoGuiFailSafeException = pyautogui.FailSafeException


@dataclass
class PendingOrder:
    file_name: str
    s3_key: str
    local_download_path: Path
    object_version: dict[str, Any]


@dataclass
class ErrorResult:
    file_name: str
    s3_key: str
    p99_key: str | None
    errors: list[str]
    warnings: list[str]


@dataclass
class P99ParseResult:
    errors: list[str]
    warnings: list[str]


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
        self.login_screenshot = self._resolve_screenshot(LOGIN_SCREENSHOT_PATTERN)
        self.import_screenshot = self._resolve_screenshot(IMPORT_SCREENSHOT_PATTERN)
        self.run_summary: dict[str, Any] = {
            "run_id": self.run_id,
            "started_at": self._iso_now(),
            "finished_at": None,
            "machine_name": self.machine_name,
            "files_detected": [],
            "files_attempted": [],
            "files_without_error": [],
            "files_with_warning": [],
            "files_with_error": [],
            "files_unresolved": [],
            "files_consumed_by_siesa": [],
            "p97": {
                "enabled": P97_REPORT_ENABLED,
                "generated": False,
                "date_from": None,
                "date_to": None,
                "local_file": None,
                "s3_key": None,
                "error": None,
            },
            "fatal_error": None,
            "log_file": str(self.log_path),
        }

    def run(self) -> int:
        self._ensure_directories()
        self._log("Iniciando corrida batch")
        siesa_opened = False

        try:
            with self._acquire_lock():
                pending_orders = self._fetch_pending_orders()
                self.run_summary["files_detected"] = [order.file_name for order in pending_orders]

                if not pending_orders:
                    self._log("No hay archivos nuevos por procesar")
                    if P97_GENERATE_WITHOUT_ORDERS:
                        self._generate_p97_from_clean_session()
                    else:
                        self._close_siesa_if_open()
                    self._finalize_and_upload_result()
                    return 0

                self._open_siesa()
                siesa_opened = True
                self._login()
                self._navigate_to_import_menu()

                for order in pending_orders:
                    self._process_order(order)

                if siesa_opened:
                    self._close_siesa_if_open()
                    siesa_opened = False

                self._generate_p97_from_clean_session()
                self._finalize_and_upload_result()
                self._persist_state()
                self._log("Corrida finalizada correctamente")
                return 0
        except Exception as exc:  # noqa: BLE001
            self.run_summary["fatal_error"] = str(exc)
            self._log(f"Corrida fallida: {exc}")
            if siesa_opened:
                try:
                    self._close_siesa_if_open()
                except Exception as close_exc:  # noqa: BLE001
                    self._log(f"No se pudo cerrar Siesa tras la falla: {close_exc}")
            self._finalize_and_upload_result()
            self._persist_state()
            traceback.print_exc()
            return 1

    def _fetch_pending_orders(self) -> list[PendingOrder]:
        self._log("Consultando pedidos pendientes en S3")
        response = self.s3_client.list_objects_v2(Bucket=AWS_BUCKET, Prefix=S3_PEDIDOS_PREFIX)

        orders: list[PendingOrder] = []
        processed_objects = self.state.get("processed_keys", {})

        for item in response.get("Contents", []):
            key = item["Key"]
            if not key.upper().endswith(".PE0"):
                continue
            object_version = self._build_object_version(item)
            if not self._should_process_object(key, object_version, processed_objects):
                continue

            file_name = Path(key).name
            local_path = DOWNLOADS_DIR / file_name
            self._log(f"Descargando {key} -> {local_path}")
            self.s3_client.download_file(AWS_BUCKET, key, str(local_path))
            orders.append(PendingOrder(
                file_name=file_name,
                s3_key=key,
                local_download_path=local_path,
                object_version=object_version,
            ))

        return orders

    def _process_order(self, order: PendingOrder) -> None:
        self._log(f"Procesando archivo {order.file_name}")
        self.run_summary["files_attempted"].append(order.file_name)

        target_path = SIESA_PEDIDOS_PATH / order.file_name
        shutil.copy2(order.local_download_path, target_path)

        before_p99 = self._snapshot_p99_files()
        self._import_file(order.file_name)
        time.sleep(FILE_PROCESS_WAIT_SECONDS)
        after_p99 = self._snapshot_p99_files()

        consumed_by_siesa = not target_path.exists()
        new_p99_files = self._detect_changed_p99_files(before_p99, after_p99)

        if consumed_by_siesa:
            self._log(f"Siesa consumió el archivo {order.file_name} desde trm")
            self.run_summary["files_consumed_by_siesa"].append(order.file_name)
        else:
            self._log(f"Siesa no consumió el archivo {order.file_name}; sigue en trm")

        if new_p99_files:
            error_result = self._handle_p99_files(order, new_p99_files)

            if consumed_by_siesa and not error_result.errors:
                if error_result.warnings:
                    self.run_summary["files_with_warning"].append({
                        "file": error_result.file_name,
                        "s3_key": error_result.s3_key,
                        "p99_key": error_result.p99_key,
                        "warnings": error_result.warnings,
                        "consumed_by_siesa": True,
                    })
                else:
                    self.run_summary["files_with_warning"].append({
                        "file": error_result.file_name,
                        "s3_key": error_result.s3_key,
                        "p99_key": error_result.p99_key,
                        "warnings": ["Siesa consumió el PE0, pero generó un P99 sin detalle reconocible."],
                        "consumed_by_siesa": True,
                    })
            elif consumed_by_siesa and error_result.errors:
                self.run_summary["files_unresolved"].append({
                    "file": error_result.file_name,
                    "s3_key": error_result.s3_key,
                    "p99_key": error_result.p99_key,
                    "reason": "Siesa consumió el PE0, pero el P99 contiene errores reales. Requiere revisión manual.",
                    "errors": error_result.errors,
                    "warnings": error_result.warnings,
                    "consumed_by_siesa": True,
                })
            elif error_result.errors:
                self.run_summary["files_with_error"].append({
                    "file": error_result.file_name,
                    "s3_key": error_result.s3_key,
                    "p99_key": error_result.p99_key,
                    "errors": error_result.errors,
                    "warnings": error_result.warnings,
                    "consumed_by_siesa": False,
                })
            elif error_result.warnings:
                self.run_summary["files_unresolved"].append({
                    "file": error_result.file_name,
                    "s3_key": error_result.s3_key,
                    "p99_key": error_result.p99_key,
                    "reason": "Siesa generó advertencias, pero no consumió el PE0 desde trm.",
                    "warnings": error_result.warnings,
                    "consumed_by_siesa": False,
                })
            else:
                self.run_summary["files_unresolved"].append({
                    "file": error_result.file_name,
                    "s3_key": error_result.s3_key,
                    "p99_key": error_result.p99_key,
                    "reason": "Siesa generó un P99, pero el bot no pudo extraer advertencias ni errores y el PE0 no fue consumido desde trm.",
                    "consumed_by_siesa": False,
                })
        elif consumed_by_siesa:
            self.run_summary["files_without_error"].append(order.file_name)
        else:
            self.run_summary["files_unresolved"].append({
                "file": order.file_name,
                "s3_key": order.s3_key,
                "p99_key": None,
                "reason": "Siesa no consumió el PE0 desde trm y no generó P99 detectable.",
                "consumed_by_siesa": False,
            })

        self._mark_processed(order.s3_key, order.file_name, order.object_version)
        self._archive_local_file(order.local_download_path)

        if target_path.exists():
            self._safe_unlink(target_path, "PE0 en trm")

    def _open_siesa(self) -> None:
        existing_windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
        if existing_windows:
            self._log("Reutilizando ventana existente de Siesa")
            self._focus_window(existing_windows[0], reset_layout=True)
            self._wait_for_screen(self.login_screenshot, "pantalla de login de Siesa")
            return

        self._log("Abriendo Siesa")
        subprocess.Popen([str(SIESA_SHORTCUT_PATH)], cwd=str(SIESA_WORKING_DIR), shell=True)
        self._wait_for_siesa_window()
        self._wait_for_screen(self.login_screenshot, "pantalla de login de Siesa")
        self._activate_siesa_window()

    def _wait_for_siesa_window(self) -> None:
        deadline = time.time() + SIESA_OPEN_TIMEOUT_SECONDS

        while time.time() < deadline:
            windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
            if windows:
                self._focus_window(windows[0], reset_layout=True)
                return
            time.sleep(SCREEN_CHECK_INTERVAL_SECONDS)

        raise RuntimeError(
            f"Siesa no abrió una ventana con el título '{SIESA_WINDOW_TITLE}' "
            f"dentro de {SIESA_OPEN_TIMEOUT_SECONDS} segundos."
        )

    def _login(self) -> None:
        self._log("Iniciando sesión")
        self._ensure_screen_visible(self.login_screenshot, "pantalla de login de Siesa")
        self._activate_siesa_window()
        self._write_login_text(SIESA_USER)
        self._press_key("tab")
        self._write_login_text(SIESA_PASSWORD)
        self._press_key("enter")
        time.sleep(LOGIN_WAIT_SECONDS)
        self._warn_if_screen_still_visible(self.login_screenshot, "pantalla de login de Siesa")

    def _navigate_to_import_menu(self) -> None:
        self._log("Navegando al menú de importación")
        self._activate_siesa_window()
        for key in MENU_SEQUENCE:
            self._press_key(key)
            time.sleep(MENU_STEP_WAIT_SECONDS)
        self._wait_for_screen(self.import_screenshot, "pantalla de importación de pedidos")

    def _import_file(self, file_name: str) -> None:
        self._ensure_screen_visible(self.import_screenshot, "pantalla de importación de pedidos")
        self._activate_siesa_window()
        for key in IMPORT_SEQUENCE_PREFIX:
            self._press_key(key)
            time.sleep(MENU_STEP_WAIT_SECONDS)

        self._write_text(file_name)

        for index, key in enumerate(IMPORT_SEQUENCE_SUFFIX):
            if len(key) > 1 and key.lower() not in {"enter", "tab", "f2", "f10"}:
                self._write_text(key)
            else:
                self._press_key(key)

            if key == "S" and index + 1 < len(IMPORT_SEQUENCE_SUFFIX) and IMPORT_SEQUENCE_SUFFIX[index + 1] == "enter":
                time.sleep(LOGIN_WAIT_SECONDS)
            else:
                time.sleep(MENU_STEP_WAIT_SECONDS)

    def _generate_p97_from_clean_session(self) -> None:
        if not P97_REPORT_ENABLED:
            self._log("Generación P97 deshabilitada por configuración")
            return

        self._log("Generando reporte P97 desde sesión limpia")
        date_from, date_to = self._p97_date_range()
        self.run_summary["p97"]["date_from"] = date_from.strftime("%Y-%m-%d")
        self.run_summary["p97"]["date_to"] = date_to.strftime("%Y-%m-%d")

        try:
            self._close_siesa_if_open()
            self._open_siesa()
            self._login()
            generated_file = self._generate_p97_report(date_from, date_to)
            self._close_siesa_if_open()
            self._upload_p97_file(generated_file, date_from, date_to)
        except Exception as exc:  # noqa: BLE001
            self.run_summary["p97"]["error"] = str(exc)
            self._log(f"No se pudo generar o subir el P97: {exc}")
            try:
                self._close_siesa_if_open()
            except Exception as close_exc:  # noqa: BLE001
                self._log(f"No se pudo cerrar Siesa tras falla P97: {close_exc}")

    def _p97_date_range(self) -> tuple[datetime, datetime]:
        date_to = datetime.now()
        date_from = date_to - timedelta(days=P97_LOOKBACK_DAYS)

        return date_from, date_to

    def _generate_p97_report(self, date_from: datetime, date_to: datetime) -> Path:
        self._log(
            "Navegando al reporte P97 "
            f"desde {date_from.strftime('%Y-%m-%d')} hasta {date_to.strftime('%Y-%m-%d')}"
        )
        self._activate_siesa_window()
        p97_file = SIESA_P97_PATH / P97_REPORT_FILE_NAME
        before_meta = self._file_meta(p97_file)

        for token in P97_MENU_SEQUENCE:
            self._execute_sequence_token(token, date_from, date_to)

        for token in P97_REPORT_SEQUENCE:
            self._execute_sequence_token(token, date_from, date_to)

        after_meta = self._file_meta(p97_file)
        if after_meta is None:
            raise RuntimeError(f"Siesa no generó el archivo P97 esperado: {p97_file}")

        if before_meta is not None and after_meta == before_meta:
            raise RuntimeError(f"El archivo P97 no cambió después de generar el reporte: {p97_file}")

        self.run_summary["p97"]["generated"] = True
        self.run_summary["p97"]["local_file"] = str(p97_file)

        return p97_file

    def _execute_sequence_token(self, token: Any, date_from: datetime, date_to: datetime) -> None:
        if isinstance(token, dict):
            wait_seconds = token.get("wait")
            if wait_seconds is not None:
                time.sleep(float(wait_seconds))
            return

        value = str(token)
        if value == "{date_from}":
            self._write_text(date_from.strftime("%y%m%d"))
        elif value == "{date_to}":
            self._write_text(date_to.strftime("%y%m%d"))
        elif len(value) > 1 and value.lower() not in {"enter", "tab", "f2", "f10"}:
            self._write_text(value)
        else:
            self._press_key(value)

        time.sleep(MENU_STEP_WAIT_SECONDS)

    def _upload_p97_file(self, p97_file: Path, date_from: datetime, date_to: datetime) -> None:
        s3_name = (
            f"UCVE1064_{date_from.strftime('%Y%m%d')}_"
            f"{date_to.strftime('%Y%m%d')}_{self.run_id}.P97"
        )
        s3_key = f"{S3_CONFIRMACIONES_PREFIX}{s3_name}"

        self._log(f"Subiendo P97 {p97_file} -> s3://{AWS_BUCKET}/{s3_key}")
        self.s3_client.upload_file(str(p97_file), AWS_BUCKET, s3_key)
        self.run_summary["p97"]["s3_key"] = s3_key

    def _file_meta(self, path: Path) -> tuple[int, int] | None:
        if not path.exists():
            return None

        stat = path.stat()

        return (stat.st_mtime_ns, stat.st_size)

    def _handle_p99_files(self, order: PendingOrder, p99_files: list[Path]) -> ErrorResult:
        self._log(f"Se detectaron {len(p99_files)} archivo(s) P99 para {order.file_name}")

        all_errors: list[str] = []
        all_warnings: list[str] = []
        uploaded_key: str | None = None

        for p99_file in p99_files:
            parsed = self._parse_p99_file(p99_file)
            all_errors.extend(parsed.errors)
            all_warnings.extend(parsed.warnings)

            uploaded_key = (
                f"{S3_ERRORES_PREFIX}{self.run_id}_{Path(order.file_name).stem}_{p99_file.name}"
            )
            self._log(f"Subiendo {p99_file} -> s3://{AWS_BUCKET}/{uploaded_key}")
            self.s3_client.upload_file(str(p99_file), AWS_BUCKET, uploaded_key)
            self._safe_unlink(p99_file, "P99")

        unique_errors = list(dict.fromkeys(all_errors))
        unique_warnings = list(dict.fromkeys(all_warnings))

        return ErrorResult(
            file_name=order.file_name,
            s3_key=order.s3_key,
            p99_key=uploaded_key,
            errors=unique_errors,
            warnings=unique_warnings,
        )

    def _parse_p99_file(self, p99_file: Path) -> P99ParseResult:
        lines = p99_file.read_text(encoding="latin-1", errors="ignore").splitlines()
        errors: list[str] = []
        warnings: list[str] = []

        for line in lines:
            cleaned_line = line.strip()
            if not cleaned_line:
                continue

            # Ignora encabezados, cajas del reporte y pie.
            if (
                "GENERACION DE PEDIDOS" in cleaned_line
                or "PEDIDO" in cleaned_line and "CAMPO_INCONSISTENTE" in cleaned_line
                or "FIN REPORTE" in cleaned_line
                or cleaned_line[0] in {"U", "A", "+", "-", "=", "_", "³", "À", "Ä", "Ã", "Ú", "Ù", "¿", "´", "⁄", "ƒ", "≥", "√", "¥", "Ÿ"}
            ):
                continue

            match = re.match(r"^\s*(\*?)(\d{8,10})\s+(\S+)\s+(.+?)\s*$", line)
            if not match:
                continue

            warning_marker, _, field_code, message = match.groups()
            clean_message = self._format_p99_message(field_code, message)

            if warning_marker == "*":
                warnings.append(clean_message)
                continue

            errors.append(clean_message)

        return P99ParseResult(errors=errors, warnings=warnings)

    def _format_p99_message(self, field_code: str, message: str) -> str:
        clean_message = re.sub(r"^\d{2}\s+", "", message.strip())
        return f"[{field_code}] {clean_message}"

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
        self._focus_window(window)

    def _wait_for_screen(self, screenshot_path: Path, screen_name: str) -> None:
        deadline = time.time() + SCREEN_CHECK_TIMEOUT_SECONDS

        while time.time() < deadline:
            self._try_activate_siesa_window()
            if self._is_screen_visible(screenshot_path):
                return
            time.sleep(SCREEN_CHECK_INTERVAL_SECONDS)

        debug_path = self._save_debug_screenshot(screen_name)
        raise RuntimeError(
            f"No se pudo validar la {screen_name} dentro de {SCREEN_CHECK_TIMEOUT_SECONDS} segundos. "
            f"Se guardó captura de depuración en {debug_path}."
        )

    def _ensure_screen_visible(self, screenshot_path: Path, screen_name: str) -> None:
        self._activate_siesa_window()
        if not self._is_screen_visible(screenshot_path):
            raise RuntimeError(
                f"Siesa no está en la {screen_name}. La corrida se detendrá para evitar marcar pedidos como exitosos."
            )

    def _ensure_screen_not_visible(self, screenshot_path: Path, screen_name: str) -> None:
        self._activate_siesa_window()
        if self._is_screen_template_visible(screenshot_path):
            raise RuntimeError(
                f"Siesa sigue en la {screen_name} después del login. "
                "La corrida se detendrá para evitar marcar pedidos como exitosos."
            )

    def _warn_if_screen_still_visible(self, screenshot_path: Path, screen_name: str) -> None:
        self._activate_siesa_window()
        if self._is_screen_template_visible(screenshot_path):
            self._log(
                f"Advertencia: la validación visual todavía detecta la {screen_name} después del login. "
                "Se continuará y la navegación siguiente validará si Siesa realmente quedó listo."
            )

    def _try_activate_siesa_window(self) -> bool:
        windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
        if not windows:
            return False

        window = windows[0]
        self._focus_window(window)
        return True

    def _focus_window(self, window, reset_layout: bool = False) -> None:
        try:
            if window.isMinimized:
                window.restore()
        except Exception:
            pass

        try:
            window.activate()
            time.sleep(1)
        except Exception:
            pass

        if reset_layout and SIESA_RESET_WINDOW_LAYOUT:
            self._reset_window_layout(window)

        if not SIESA_FORCE_MAXIMIZE:
            return

        try:
            if not window.isMaximized:
                window.maximize()
                time.sleep(1)
        except Exception:
            pass

        # Siesa 8.5 puede ignorar maximize() y responder mejor al menú del sistema.
        try:
            if not window.isMaximized:
                pyautogui.hotkey("alt", "space")
                time.sleep(0.5)
                self._press_key("x")
                time.sleep(1)
        except Exception:
            pass

    def _reset_window_layout(self, window) -> None:
        # Siesa 8.5 a veces queda visualmente corrupto hasta que la ventana cambia
        # de estado y vuelve a maximizarse. Esta secuencia replica el workaround manual.
        try:
            if window.isMaximized:
                window.restore()
                time.sleep(1)
                window.maximize()
                time.sleep(1)
                return
        except Exception:
            pass

        try:
            window.maximize()
            time.sleep(1)
        except Exception:
            pass

        try:
            window.restore()
            time.sleep(1)
        except Exception:
            pass

        try:
            window.maximize()
            time.sleep(1)
        except Exception:
            pass

    def _is_screen_visible(self, screenshot_path: Path) -> bool:
        if not screenshot_path.exists():
            raise RuntimeError(f"No existe la captura de referencia requerida: {screenshot_path}")

        search_region = self._get_siesa_window_region()
        if search_region:
            similarity = self._window_similarity(screenshot_path, search_region)
            if similarity >= SCREEN_VISUAL_SIMILARITY_THRESHOLD:
                return True

        return self._is_screen_template_visible(screenshot_path, search_region)

    def _is_screen_template_visible(
        self,
        screenshot_path: Path,
        search_region: tuple[int, int, int, int] | None = None,
    ) -> bool:
        if not screenshot_path.exists():
            raise RuntimeError(f"No existe la captura de referencia requerida: {screenshot_path}")

        if search_region is None:
            search_region = self._get_siesa_window_region()

        locate_kwargs = {"grayscale": True}
        if SCREEN_MATCH_CONFIDENCE:
            locate_kwargs["confidence"] = SCREEN_MATCH_CONFIDENCE
        reference_size = self._get_image_size(screenshot_path)
        if search_region and reference_size:
            _, _, region_width, region_height = search_region
            reference_width, reference_height = reference_size
            if reference_width <= region_width and reference_height <= region_height:
                locate_kwargs["region"] = search_region

        try:
            match = pyautogui.locateOnScreen(str(screenshot_path), **locate_kwargs)
        except pyautogui.ImageNotFoundException:
            return False
        except PyAutoGuiFailSafeException as exc:
            raise RuntimeError(
                "PyAutoGUI activó el fail-safe porque el cursor llegó a una esquina de la pantalla "
                "durante la validación visual."
            ) from exc
        except TypeError:
            fallback_kwargs = {"grayscale": True}
            if "region" in locate_kwargs:
                fallback_kwargs["region"] = locate_kwargs["region"]
            match = pyautogui.locateOnScreen(str(screenshot_path), **fallback_kwargs)
        except NotImplementedError:
            fallback_kwargs = {"grayscale": True}
            if "region" in locate_kwargs:
                fallback_kwargs["region"] = locate_kwargs["region"]
            match = pyautogui.locateOnScreen(str(screenshot_path), **fallback_kwargs)
        except ValueError as exc:
            if "region" in locate_kwargs and "needle dimension" in str(exc).lower():
                fallback_kwargs = {"grayscale": True}
                if SCREEN_MATCH_CONFIDENCE:
                    fallback_kwargs["confidence"] = SCREEN_MATCH_CONFIDENCE
                try:
                    match = pyautogui.locateOnScreen(str(screenshot_path), **fallback_kwargs)
                except pyautogui.ImageNotFoundException:
                    return False
            else:
                raise RuntimeError(f"No fue posible validar la pantalla usando {screenshot_path.name}: {exc}") from exc
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"No fue posible validar la pantalla usando {screenshot_path.name}: {exc}") from exc

        return match is not None

    def _window_similarity(self, screenshot_path: Path, region: tuple[int, int, int, int]) -> float:
        left, top, width, height = region
        current = self._take_screenshot(region=(left, top, width, height))

        with Image.open(screenshot_path) as reference:
            reference_gray = ImageOps.grayscale(reference)
            current_gray = ImageOps.grayscale(current)

            if current_gray.size != reference_gray.size:
                current_gray = current_gray.resize(reference_gray.size)

            reference_crop = self._crop_for_similarity(reference_gray)
            current_crop = self._crop_for_similarity(current_gray)

            if current_crop.size != reference_crop.size:
                current_crop = current_crop.resize(reference_crop.size)

            diff = ImageChops.difference(reference_crop, current_crop)
            mean_diff = ImageStat.Stat(diff).mean[0]

        return max(0.0, 1.0 - (mean_diff / 255.0))

    def _crop_for_similarity(self, image: Image.Image) -> Image.Image:
        width, height = image.size

        left = int(width * 0.08)
        top = int(height * 0.10)
        right = int(width * 0.92)
        bottom = int(height * 0.90)

        return image.crop((left, top, right, bottom))

    def _close_siesa_if_open(self) -> None:
        windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
        if not windows:
            return

        self._log("Cerrando ventana de Siesa")
        window = windows[0]
        if window.isMinimized:
            window.restore()
        window.activate()
        time.sleep(1)
        self._hotkey("alt", "f4")
        time.sleep(2)

    def _snapshot_p99_files(self) -> dict[Path, tuple[int, int]]:
        snapshot: dict[Path, tuple[int, int]] = {}

        for path in SIESA_P99_PATH.glob("*.P99"):
            stat = path.stat()
            snapshot[path] = (stat.st_mtime_ns, stat.st_size)

        return snapshot

    def _detect_changed_p99_files(
        self,
        before_snapshot: dict[Path, tuple[int, int]],
        after_snapshot: dict[Path, tuple[int, int]],
    ) -> list[Path]:
        changed: list[Path] = []

        for path, after_meta in after_snapshot.items():
            before_meta = before_snapshot.get(path)
            if before_meta is None or before_meta != after_meta:
                changed.append(path)

        return sorted(changed)

    def _archive_local_file(self, local_path: Path) -> None:
        archive_name = f"{self.run_id}_{local_path.name}"
        archive_path = ARCHIVE_DIR / archive_name
        shutil.move(str(local_path), archive_path)

    def _safe_unlink(self, path: Path, label: str, attempts: int = 5, wait_seconds: float = 1.0) -> bool:
        for attempt in range(1, attempts + 1):
            try:
                path.unlink()
                return True
            except FileNotFoundError:
                return True
            except PermissionError as exc:
                if attempt == attempts:
                    self._log(
                        f"No se pudo eliminar {label} local {path} porque sigue bloqueado: {exc}. "
                        "La corrida continuará."
                    )
                    return False
                time.sleep(wait_seconds)

        return False

    def _delete_source_object(self, s3_key: str) -> None:
        try:
            self.s3_client.delete_object(Bucket=AWS_BUCKET, Key=s3_key)
            self._log(f"Objeto fuente eliminado de S3: {s3_key}")
        except Exception as exc:  # noqa: BLE001
            self._log(f"No se pudo eliminar {s3_key} de S3: {exc}")

    def _mark_processed(self, s3_key: str, file_name: str, object_version: dict[str, Any]) -> None:
        self.state.setdefault("processed_keys", {})[s3_key] = {
            "file_name": file_name,
            "object_version": object_version,
            "run_id": self.run_id,
            "processed_at": self._iso_now(),
        }

    def _build_object_version(self, item: dict[str, Any]) -> dict[str, Any]:
        last_modified = item.get("LastModified")

        return {
            "etag": str(item.get("ETag", "")).strip('"'),
            "size": int(item.get("Size", 0)),
            "last_modified": last_modified.isoformat() if last_modified else None,
        }

    def _should_process_object(
        self,
        s3_key: str,
        current_version: dict[str, Any],
        processed_objects: dict[str, Any],
    ) -> bool:
        previous_entry = processed_objects.get(s3_key)
        if not previous_entry:
            return True

        previous_version = previous_entry.get("object_version", {})

        return previous_version != current_version

    def _ensure_directories(self) -> None:
        for path in [DOWNLOADS_DIR, ARCHIVE_DIR, LOGS_DIR]:
            path.mkdir(parents=True, exist_ok=True)

        if not SIESA_SHORTCUT_PATH.exists():
            raise RuntimeError(f"No se encontró el acceso directo de Siesa: {SIESA_SHORTCUT_PATH}")

        for screenshot_path in [self.login_screenshot, self.import_screenshot]:
            if not screenshot_path.exists():
                raise RuntimeError(f"No existe la captura de referencia requerida: {screenshot_path}")

        for path in [SIESA_WORKING_DIR, SIESA_PEDIDOS_PATH, SIESA_P99_PATH, SIESA_P97_PATH]:
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
            self._clear_stale_lock_if_needed()
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

    def _clear_stale_lock_if_needed(self) -> None:
        if not LOCK_FILE.exists():
            return

        try:
            pid = int(LOCK_FILE.read_text(encoding="utf-8").strip())
        except (OSError, ValueError):
            self._log("run.lock inválido, se eliminará para continuar")
            LOCK_FILE.unlink(missing_ok=True)
            return

        if pid == os.getpid():
            return

        if self._is_process_running(pid):
            raise RuntimeError(
                f"Ya existe otra instancia del bot ejecutándose con PID {pid}. "
                "No se iniciará una nueva corrida."
            )

        self._log(f"Se encontró un run.lock huérfano del PID {pid}, se eliminará")
        LOCK_FILE.unlink(missing_ok=True)

    def _is_process_running(self, pid: int) -> bool:
        try:
            os.kill(pid, 0)
        except OSError:
            return False
        return True

    def _resolve_screenshot(self, pattern: str) -> Path:
        matches = sorted(SCREENSHOTS_DIR.glob(pattern))
        if not matches:
            raise RuntimeError(f"No se encontró ninguna captura en {SCREENSHOTS_DIR} con el patrón {pattern}")
        return matches[0]

    def _save_debug_screenshot(self, screen_name: str) -> Path:
        safe_name = screen_name.lower().replace(" ", "_")
        debug_path = LOGS_DIR / f"debug_{safe_name}_{self.run_id}.png"
        screenshot = self._take_screenshot()
        screenshot.save(debug_path)
        self._log(f"Captura de depuración guardada en {debug_path}")
        return debug_path

    def _get_siesa_window_region(self) -> tuple[int, int, int, int] | None:
        windows = gw.getWindowsWithTitle(SIESA_WINDOW_TITLE)
        if not windows:
            return None

        window = windows[0]
        width = max(int(window.width), 1)
        height = max(int(window.height), 1)
        return (int(window.left), int(window.top), width, height)

    def _get_image_size(self, image_path: Path) -> tuple[int, int] | None:
        try:
            import PIL.Image

            with PIL.Image.open(image_path) as image:
                return image.size
        except Exception:
            return None

    def _press_key(self, key: str) -> None:
        try:
            pyautogui.press(key)
        except PyAutoGuiFailSafeException as exc:
            raise RuntimeError(
                "PyAutoGUI activó el fail-safe porque el cursor llegó a una esquina de la pantalla. "
                "La corrida se detuvo para evitar enviar teclas fuera de contexto."
            ) from exc

    def _write_text(self, text: str) -> None:
        try:
            pyautogui.write(text)
        except PyAutoGuiFailSafeException as exc:
            raise RuntimeError(
                "PyAutoGUI activó el fail-safe porque el cursor llegó a una esquina de la pantalla. "
                "La corrida se detuvo para evitar escribir fuera de contexto."
            ) from exc

    def _write_login_text(self, text: str) -> None:
        previous_caps_lock = self._is_caps_lock_on()

        try:
            self._set_caps_lock(True)
            self._write_text(text.lower())
        finally:
            self._set_caps_lock(previous_caps_lock)

    def _is_caps_lock_on(self) -> bool:
        try:
            return bool(ctypes.windll.user32.GetKeyState(0x14) & 1)
        except Exception:
            return False

    def _set_caps_lock(self, enabled: bool) -> None:
        if self._is_caps_lock_on() != enabled:
            self._press_key("capslock")
            time.sleep(0.2)

    def _hotkey(self, *keys: str) -> None:
        try:
            pyautogui.hotkey(*keys)
        except PyAutoGuiFailSafeException as exc:
            raise RuntimeError(
                "PyAutoGUI activó el fail-safe porque el cursor llegó a una esquina de la pantalla."
            ) from exc

    def _take_screenshot(self, region: tuple[int, int, int, int] | None = None):
        try:
            if region is None:
                return pyautogui.screenshot()
            return pyautogui.screenshot(region=region)
        except PyAutoGuiFailSafeException as exc:
            raise RuntimeError(
                "PyAutoGUI activó el fail-safe porque el cursor llegó a una esquina de la pantalla."
            ) from exc



def main() -> int:
    bot = RpaBot()
    return bot.run()


if __name__ == "__main__":
    sys.exit(main())
