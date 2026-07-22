from __future__ import annotations

from datetime import datetime
from typing import Any

from .com_client import OneCComClient
from .config import AppConfig
from .errors import ErrorCode, IntegrationError, ok


class OneCIntegrationService:
    def __init__(self, config: AppConfig, client: OneCComClient) -> None:
        self.config = config
        self.client = client

    def warehouses(self) -> dict[str, Any]:
        rows = self.client.query(
            """
            ВЫБРАТЬ
                Склады.Ссылка КАК warehouse_id,
                Склады.Код КАК code,
                Склады.Наименование КАК name
            ИЗ
                Справочник.Склады КАК Склады
            ГДЕ
                НЕ Склады.ПометкаУдаления
            """
        )
        return ok({"generated_at": _now(), "database": self.config.connection.database, "items": rows})

    def nomenclature(self, code: str = "", article: str = "", name: str = "") -> dict[str, Any]:
        filters = []
        params: dict[str, Any] = {}
        if code:
            filters.append("Номенклатура.Код = &Код")
            params["Код"] = code
        if article:
            filters.append("Номенклатура.Артикул = &Артикул")
            params["Артикул"] = article
        if name:
            filters.append("Номенклатура.Наименование ПОДОБНО &Наименование")
            params["Наименование"] = f"%{name}%"

        where = " И НЕ Номенклатура.ПометкаУдаления"
        if filters:
            where += " И " + " И ".join(filters)

        rows = self.client.query(
            f"""
            ВЫБРАТЬ ПЕРВЫЕ 500
                Номенклатура.Ссылка КАК nomenclature_id,
                Номенклатура.Код КАК code,
                Номенклатура.Артикул КАК article,
                Номенклатура.Наименование КАК name,
                Номенклатура.БазоваяЕдиницаИзмерения.Наименование КАК unit
            ИЗ
                Справочник.Номенклатура КАК Номенклатура
            ГДЕ
                ИСТИНА {where}
            УПОРЯДОЧИТЬ ПО
                Номенклатура.Наименование
            """,
            params,
        )
        return ok({"generated_at": _now(), "database": self.config.connection.database, "items": rows})

    def stocks(self, positive_only: bool = False, at_date: str | None = None) -> dict[str, Any]:
        metadata = self.config.raw.get("metadata", {})
        register = metadata.get("stock_register", "")
        dimensions = metadata.get("stock_dimensions", {})
        resources = metadata.get("stock_resources", {})
        if not register or not dimensions.get("nomenclature") or not dimensions.get("warehouse") or not resources.get("quantity"):
            raise IntegrationError(
                ErrorCode.VALIDATION_ERROR,
                "Не заполнены реальные имена регистра остатков, измерений и ресурсов в config.json. Сначала выполните inspect-metadata.",
            )

        date_expr = "&ДатаОстатков" if at_date else ""
        where = f"ГДЕ Остатки.{resources['quantity']}Остаток > 0" if positive_only else ""
        params = {"ДатаОстатков": at_date} if at_date else {}
        reserved_expr = _resource_expr(resources.get("reserved"), "reserved")
        to_provide_expr = _resource_expr(resources.get("to_provide"), "to_provide")
        query = f"""
        ВЫБРАТЬ
            Остатки.{dimensions['nomenclature']} КАК nomenclature_id,
            Остатки.{dimensions['nomenclature']}.Код КАК code,
            Остатки.{dimensions['nomenclature']}.Артикул КАК article,
            Остатки.{dimensions['nomenclature']}.Наименование КАК name,
            Остатки.{dimensions['nomenclature']}.БазоваяЕдиницаИзмерения.Наименование КАК unit,
            Остатки.{dimensions['warehouse']} КАК warehouse_id,
            Остатки.{dimensions['warehouse']}.Наименование КАК warehouse,
            Остатки.{resources['quantity']}Остаток КАК quantity
            {reserved_expr}
            {to_provide_expr}
        ИЗ
            РегистрНакопления.{register}.Остатки({date_expr}) КАК Остатки
        {where}
        """
        rows = self.client.query(query, params)
        items = []
        for row in rows:
            quantity = _number(row.get("quantity"))
            reserved = _number(row.get("reserved"))
            row["available"] = quantity - reserved
            items.append(row)
        return ok({"generated_at": _now(), "database": self.config.connection.database, "items": items})

    def create_assembly(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._validate_external_id(payload)
        self._ensure_write_allowed("Комплектование")
        return ok({"dry_run": self.config.safety.dry_run, "operation": "assembly", "external_id": payload["external_id"], "document_ref": None})

    def create_transfer(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._validate_external_id(payload)
        if payload.get("source_warehouse_id") == payload.get("target_warehouse_id"):
            raise IntegrationError(ErrorCode.VALIDATION_ERROR, "Склад-отправитель и склад-получатель должны различаться.")
        self._ensure_write_allowed("Перемещение")
        return ok({"dry_run": self.config.safety.dry_run, "operation": "transfer", "external_id": payload["external_id"], "document_ref": None})

    def _validate_external_id(self, payload: dict[str, Any]) -> None:
        if not payload.get("external_id"):
            raise IntegrationError(ErrorCode.VALIDATION_ERROR, "Не задан external_id.")

    def _ensure_write_allowed(self, document_name: str) -> None:
        if self.config.safety.dry_run:
            return
        if not self.config.safety.allow_document_write:
            raise IntegrationError(ErrorCode.DOCUMENT_WRITE_ERROR, f"Запись документа {document_name} заблокирована: ALLOW_DOCUMENT_WRITE не равен true.")


def _resource_expr(resource: str | None, alias: str) -> str:
    if not resource:
        return f", 0 КАК {alias}"
    return f", Остатки.{resource}Остаток КАК {alias}"


def _number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    return float(value)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
