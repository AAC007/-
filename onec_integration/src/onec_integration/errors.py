from __future__ import annotations

from dataclasses import dataclass
from typing import Any


class ErrorCode:
    AUTHENTICATION_ERROR = "AUTHENTICATION_ERROR"
    ACCESS_DENIED = "ACCESS_DENIED"
    DATABASE_UNAVAILABLE = "DATABASE_UNAVAILABLE"
    OBJECT_NOT_FOUND = "OBJECT_NOT_FOUND"
    WAREHOUSE_NOT_FOUND = "WAREHOUSE_NOT_FOUND"
    NOMENCLATURE_NOT_FOUND = "NOMENCLATURE_NOT_FOUND"
    CHARACTERISTIC_NOT_FOUND = "CHARACTERISTIC_NOT_FOUND"
    INSUFFICIENT_STOCK = "INSUFFICIENT_STOCK"
    INVALID_QUANTITY = "INVALID_QUANTITY"
    VALIDATION_ERROR = "VALIDATION_ERROR"
    DOCUMENT_WRITE_ERROR = "DOCUMENT_WRITE_ERROR"
    DOCUMENT_POSTING_ERROR = "DOCUMENT_POSTING_ERROR"
    DUPLICATE_EXTERNAL_ID = "DUPLICATE_EXTERNAL_ID"
    INTERNAL_ERROR = "INTERNAL_ERROR"


@dataclass(slots=True)
class IntegrationError(Exception):
    code: str
    message: str
    details: dict[str, Any] | None = None

    def to_response(self) -> dict[str, Any]:
        return {
            "success": False,
            "error": {
                "code": self.code,
                "message": self.message,
                "details": self.details or {},
            },
        }


def ok(payload: dict[str, Any]) -> dict[str, Any]:
    return {"success": True, **payload}
