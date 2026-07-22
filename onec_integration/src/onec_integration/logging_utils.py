from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any


SECRET_KEYS = {"password", "pwd", "ONEC_PASSWORD"}


def configure_logging(directory: Path) -> logging.Logger:
    directory.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("onec_integration")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.FileHandler(directory / "onec_integration.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(handler)
    return logger


def audit(logger: logging.Logger, operation: str, payload: dict[str, Any]) -> None:
    event = {
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "operation": operation,
        **_sanitize(payload),
    }
    logger.info(json.dumps(event, ensure_ascii=False, default=str))


def _sanitize(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: ("***" if key.lower() in SECRET_KEYS else _sanitize(item)) for key, item in value.items()}
    if isinstance(value, list):
        return [_sanitize(item) for item in value]
    return value
