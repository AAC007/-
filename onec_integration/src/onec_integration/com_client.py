from __future__ import annotations

import platform
import socket
import subprocess
from datetime import datetime
from typing import Any

from .config import ConnectionSettings
from .errors import ErrorCode, IntegrationError


class OneCComClient:
    def __init__(self, settings: ConnectionSettings) -> None:
        self.settings = settings
        self._connection: Any | None = None

    def connect(self) -> Any:
        if self._connection is not None:
            return self._connection
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise IntegrationError(
                ErrorCode.DATABASE_UNAVAILABLE,
                "Python-модуль pywin32 не установлен или недоступен. Установите pywin32 той же разрядности, что и COM-коннектор 1С.",
            ) from exc

        try:
            connector = win32com.client.Dispatch("V83.COMConnector")
            self._connection = connector.Connect(self.settings.connection_string)
            return self._connection
        except Exception as exc:  # COM exception text can contain 1C details.
            raise IntegrationError(ErrorCode.DATABASE_UNAVAILABLE, f"Не удалось подключиться к 1С: {exc}") from exc

    def metadata_report(self) -> dict[str, Any]:
        connection = self.connect()
        metadata = connection.Метаданные
        return {
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "database": self.settings.database,
            "configuration": _safe_get(metadata, "Имя"),
            "configuration_version": _safe_get(metadata, "Версия"),
            "documents": _metadata_objects(metadata.Документы, ["Комплект", "Перемещ"]),
            "accumulation_registers": _metadata_objects(metadata.РегистрыНакопления, ["Товары", "Склад", "Остат", "Резерв", "Партии"]),
            "catalogs": _metadata_objects(metadata.Справочники, ["Номенклатура", "Склады", "Организации"]),
            "roles": _metadata_objects(metadata.Роли, []),
        }

    def query(self, text: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        connection = self.connect()
        query = connection.NewObject("Запрос")
        query.Текст = text
        for name, value in (params or {}).items():
            query.УстановитьПараметр(name, value)
        result = query.Выполнить().Выбрать()
        rows: list[dict[str, Any]] = []
        while result.Следующий():
            row: dict[str, Any] = {}
            for index in range(result.Колонки.Количество()):
                column = result.Колонки.Получить(index).Имя
                row[column] = getattr(result, column)
            rows.append(row)
        return rows


def diagnose_environment(host: str, port: int = 1541) -> dict[str, Any]:
    resolved = ""
    tcp_ok = False
    error = ""
    try:
        resolved = socket.gethostbyname(host)
        with socket.create_connection((host, port), timeout=5):
            tcp_ok = True
    except Exception as exc:
        error = str(exc)

    return {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "python": platform.python_version(),
        "python_bits": platform.architecture()[0],
        "host": host,
        "port": port,
        "resolved_address": resolved,
        "tcp_ok": tcp_ok,
        "error": error,
        "onec_executables": find_1c_executables(),
    }


def find_1c_executables() -> list[str]:
    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        "Get-ChildItem -Path 'C:\\Program Files','C:\\Program Files (x86)' -Filter '1cv8*.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName",
    ]
    try:
        completed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=20)
    except Exception:
        return []
    return [line.strip() for line in completed.stdout.splitlines() if line.strip()]


def _metadata_objects(collection: Any, keywords: list[str]) -> list[dict[str, Any]]:
    objects: list[dict[str, Any]] = []
    for index in range(collection.Количество()):
        item = collection.Получить(index)
        name = _safe_get(item, "Имя")
        synonym = _safe_get(item, "Синоним")
        if keywords and not any(word.lower() in f"{name} {synonym}".lower() for word in keywords):
            continue
        objects.append(
            {
                "name": name,
                "synonym": synonym,
                "attributes": _child_names(_safe_get(item, "Реквизиты")),
                "tabular_sections": _child_names(_safe_get(item, "ТабличныеЧасти")),
                "dimensions": _child_names(_safe_get(item, "Измерения")),
                "resources": _child_names(_safe_get(item, "Ресурсы")),
            }
        )
    return objects


def _child_names(collection: Any) -> list[str]:
    if not collection:
        return []
    try:
        return [_safe_get(collection.Получить(index), "Имя") for index in range(collection.Количество())]
    except Exception:
        return []


def _safe_get(obj: Any, attr: str) -> Any:
    try:
        return getattr(obj, attr)
    except Exception:
        return ""
