from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import ErrorCode, IntegrationError


@dataclass(frozen=True, slots=True)
class ConnectionSettings:
    server: str
    database: str
    username: str
    password: str

    @property
    def connection_string(self) -> str:
        if not self.password:
            raise IntegrationError(ErrorCode.AUTHENTICATION_ERROR, "Не задан ONEC_PASSWORD.")
        return f'Srvr="{self.server}";Ref="{self.database}";Usr="{self.username}";Pwd="{self.password}";'


@dataclass(frozen=True, slots=True)
class SafetySettings:
    dry_run: bool
    allow_document_write: bool
    allow_document_posting: bool
    integration_comment: str


@dataclass(frozen=True, slots=True)
class AppConfig:
    raw: dict[str, Any]
    connection: ConnectionSettings
    safety: SafetySettings
    log_directory: Path


def load_config(path: str | Path = "onec_integration/config.example.json") -> AppConfig:
    config_path = Path(path)
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    connection = raw["connection"]
    safety = raw["safety"]
    logging_cfg = raw.get("logging", {})

    username = os.environ.get(connection.get("username_env", "ONEC_USERNAME")) or "Аракелян Армен Седракович"
    password = os.environ.get(connection.get("password_env", "ONEC_PASSWORD"), "")

    return AppConfig(
        raw=raw,
        connection=ConnectionSettings(
            server=os.environ.get("ONEC_SERVER", connection["server"]),
            database=os.environ.get("ONEC_DATABASE", connection["database"]),
            username=username,
            password=password,
        ),
        safety=SafetySettings(
            dry_run=_env_bool(safety.get("dry_run_env", "DRY_RUN"), default=True),
            allow_document_write=_env_bool(safety.get("allow_document_write_env", "ALLOW_DOCUMENT_WRITE"), default=False),
            allow_document_posting=_env_bool(safety.get("allow_document_posting_env", "ALLOW_DOCUMENT_POSTING"), default=False),
            integration_comment=safety["integration_comment"],
        ),
        log_directory=Path(logging_cfg.get("directory", "onec_integration/logs")),
    )


def _env_bool(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "да"}


def redacted_connection(settings: ConnectionSettings) -> dict[str, str]:
    return {
        "server": settings.server,
        "database": settings.database,
        "username": settings.username,
        "password": "***" if settings.password else "",
    }
