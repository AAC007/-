from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode
from urllib.error import HTTPError
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
DEFAULT_ENV = Path(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")
DEFAULT_ODATA_URL = "http://pz-sql1/ut_dev/odata/standard.odata/"
STOCK_REGISTER = "AccumulationRegister_ТоварыНаСкладах"
WAREHOUSE_ENTITY = "Catalog_Склады"
NOMENCLATURE_ENTITY = "Catalog_Номенклатура"
UNIT_ENTITY = "Catalog_ЕдиницыИзмерения"


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Read-only OData синхронизация остатков 1С по секциям ЦМО/ПЗМЦ")
    parser.add_argument("--env", default=str(DEFAULT_ENV), help="Путь к .env с ONEC_*")
    parser.add_argument("--odata-url", default="", help="OData URL 1С, например http://pz-sql1/ut_dev/odata/standard.odata/")
    parser.add_argument("--wip-warehouse", nargs="+", default=["44"], help="Секции НЗП/ЦМО")
    parser.add_argument("--production-warehouse", nargs="+", default=["1"], help="Склады производственных остатков")
    parser.add_argument("--out", default=str(REPORTS / "stock_sections_sync_odata.json"))
    parser.add_argument("--timeout", type=float, default=12.0)
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    load_env(Path(args.env))
    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    result: dict[str, Any] = {
        "ok": False,
        "generated_at": generated_at,
        "provider": "odata",
        "odata_url": safe_base_url(args.odata_url),
        "main_register": "ТоварыНаСкладах",
        "warehouses": {
            "wip": args.wip_warehouse,
            "production": args.production_warehouse,
        },
        "rows": [],
        "errors": [],
        "timings_ms": {},
    }

    started = time.perf_counter()
    try:
        client = ODataClient(args.odata_url or os.environ.get("ONEC_ODATA_URL") or DEFAULT_ODATA_URL, args.timeout)
        roles = (("wip", args.wip_warehouse), ("production", args.production_warehouse))
        rows: list[dict[str, str]] = []
        for role, warehouse_filters in roles:
            for warehouse_filter in warehouse_filters:
                try:
                    rows.extend(sync_warehouse(client, warehouse_filter, role))
                except Exception as exc:
                    result["errors"].append({"warehouse": warehouse_filter, "role": role, "error": str(exc)})

        result["rows"] = rows
        result["timings_ms"]["total"] = round((time.perf_counter() - started) * 1000)
        result["ok"] = len(result["errors"]) == 0 and len(rows) > 0
        if len(rows) == 0 and len(result["errors"]) == 0:
            result["errors"].append({"warehouse": "", "role": "", "error": "OData вернул 0 строк остатков; результат признан сомнительным."})
    except Exception as exc:
        result["errors"].append({"warehouse": "", "role": "", "error": str(exc)})

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"1C OData stock section sync at {generated_at}")
    print(f"Rows: {len(result['rows'])}")
    if result["errors"]:
        print(f"Errors: {len(result['errors'])}")
    print(f"JSON: {out}")
    return 0 if result["ok"] else 2


class ODataClient:
    def __init__(self, base_url: str, timeout: float) -> None:
        self.base_url = base_url.rstrip("/") + "/"
        self.timeout = timeout
        self._warehouses: list[dict[str, Any]] | None = None
        self.username = os.environ.get("ONEC_USERNAME") or os.environ.get("ONEC_USER") or ""
        self.password = os.environ.get("ONEC_PASSWORD") or ""
        if not self.username or not self.password:
            raise RuntimeError("Не заданы ONEC_USERNAME/ONEC_USER и ONEC_PASSWORD для OData 1С.")

    def get_rows(self, path: str, params: dict[str, str] | None = None) -> list[dict[str, Any]]:
        payload = self.get_json(path, params)
        return extract_rows(payload)

    def get_json(self, path: str, params: dict[str, str] | None = None) -> Any:
        url = self.base_url + quote(path.lstrip("/"), safe="/(),'=$")
        query = {"$format": "json", **(params or {})}
        if query:
            url += "?" + urlencode(query, safe="$'(),")
        token = base64.b64encode(f"{self.username}:{self.password}".encode("utf-8")).decode("ascii")
        request = Request(url, headers={"Authorization": f"Basic {token}", "Accept": "application/json"})
        try:
            with urlopen(request, timeout=self.timeout) as response:
                return json.loads(response.read().decode("utf-8-sig"))
        except HTTPError as exc:
            body = exc.read().decode("utf-8-sig", errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {body[:700]}") from exc

    def warehouses(self) -> list[dict[str, Any]]:
        if self._warehouses is None:
            self._warehouses = self.get_rows(WAREHOUSE_ENTITY, {
                "$top": "1000",
                "$select": "Ref_Key,Code,Description",
            })
        return self._warehouses


def sync_warehouse(client: ODataClient, warehouse_filter: str, role: str) -> list[dict[str, str]]:
    warehouse = resolve_warehouse(client, warehouse_filter)
    balances = query_balance(client, warehouse["ref_key"])
    positive_balances = [row for row in balances if decimal_value(first_field(row, "КоличествоBalance", "КоличествоОстаток", "Quantity", "Количество")) > 0]
    item_keys = distinct_keys(first_field(row, "Номенклатура_Key", "Номенклатура") for row in positive_balances)
    warehouse_keys = distinct_keys(first_field(row, "Склад_Key", "Склад") for row in positive_balances)
    nomenclature = load_nomenclature(client, item_keys)
    warehouses = load_warehouses(client, warehouse_keys)
    units = load_units(client, distinct_keys(item.get("unit_key") for item in nomenclature.values()))

    rows: list[dict[str, str]] = []
    for balance in positive_balances:
        item_key = normalize_ref(first_field(balance, "Номенклатура_Key", "Номенклатура"))
        warehouse_key = normalize_ref(first_field(balance, "Склад_Key", "Склад"))
        item = nomenclature.get(item_key, {})
        warehouse_row = warehouses.get(warehouse_key, warehouse)
        quantity = decimal_value(first_field(balance, "КоличествоBalance", "КоличествоОстаток", "Quantity", "Количество"))
        if quantity <= 0:
            continue
        rows.append({
            "role": role,
            "code": str(item.get("code") or ""),
            "article": str(item.get("article") or ""),
            "name": str(item.get("name") or item_key),
            "unit": str(units.get(str(item.get("unit_key") or ""), item.get("unit") or "")),
            "warehouse_code": str(warehouse_row.get("code") or ""),
            "warehouse": str(warehouse_row.get("name") or warehouse_filter),
            "quantity": format(quantity.normalize(), "f"),
        })
    return rows


def resolve_warehouse(client: ODataClient, query: str) -> dict[str, str]:
    rows = [
        row for row in client.warehouses()
        if warehouse_matches(row, query)
    ]
    if not rows:
        raise RuntimeError(f"Склад/секция не найдены через OData: {query}")
    selected = sorted(rows, key=lambda row: warehouse_score(row, query), reverse=True)[0]
    return {
        "ref_key": normalize_ref(selected.get("Ref_Key")),
        "code": str(selected.get("Code") or ""),
        "name": str(selected.get("Description") or selected.get("Наименование") or query),
    }


def warehouse_matches(row: dict[str, Any], query: str) -> bool:
    text = query.strip().lower()
    code = str(row.get("Code") or "").strip().lower()
    name = str(row.get("Description") or row.get("Наименование") or "").strip().lower()
    if not text:
        return False
    if code == text or name == text or text in name:
        return True
    return text.isdigit() and name.startswith(f"{text} секция")


def warehouse_score(row: dict[str, Any], query: str) -> tuple[int, int]:
    text = query.strip().lower()
    code = str(row.get("Code") or "").strip().lower()
    name = str(row.get("Description") or row.get("Наименование") or "").strip().lower()
    return (
        3 if name == text else 2 if code == text else 1 if name.startswith(text) else 0,
        len(name),
    )


def query_balance(client: ODataClient, warehouse_ref: str) -> list[dict[str, Any]]:
    conditions = [
        f"Склад_Key eq guid'{warehouse_ref}'",
        f"Склад eq guid'{warehouse_ref}'",
        f"Склад/Ref_Key eq guid'{warehouse_ref}'",
    ]
    errors: list[str] = []
    for condition in conditions:
        for path, params in balance_call_variants(condition):
            try:
                rows = client.get_rows(path, params)
                if rows:
                    return rows
            except Exception as exc:
                errors.append(str(exc))
    raise RuntimeError("Не удалось получить Balance через OData: " + "; ".join(errors[:3]))


def balance_call_variants(condition: str) -> list[tuple[str, dict[str, str]]]:
    escaped = condition.replace("'", "''")
    dimensions = "Номенклатура,Склад"
    return [
        (f"{STOCK_REGISTER}/Balance", {"Condition": condition, "Dimensions": dimensions}),
        (f"{STOCK_REGISTER}/Balance(Condition='{escaped}',Dimensions='{dimensions}')", {}),
    ]


def load_nomenclature(client: ODataClient, refs: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for chunk in chunks(refs, 20):
        rows = client.get_rows(NOMENCLATURE_ENTITY, {
            "$select": "Ref_Key,Code,Description,Артикул,БазоваяЕдиницаИзмерения_Key",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        })
        for row in rows:
            ref = normalize_ref(row.get("Ref_Key"))
            result[ref] = {
                "code": str(row.get("Code") or ""),
                "article": str(row.get("Артикул") or row.get("Article") or ""),
                "name": str(row.get("Description") or row.get("Наименование") or ""),
                "unit_key": normalize_ref(row.get("БазоваяЕдиницаИзмерения_Key")),
            }
    return result


def load_warehouses(client: ODataClient, refs: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for chunk in chunks(refs, 30):
        rows = client.get_rows(WAREHOUSE_ENTITY, {
            "$select": "Ref_Key,Code,Description",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        })
        for row in rows:
            ref = normalize_ref(row.get("Ref_Key"))
            result[ref] = {
                "code": str(row.get("Code") or ""),
                "name": str(row.get("Description") or row.get("Наименование") or ""),
            }
    return result


def load_units(client: ODataClient, refs: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for chunk in chunks(refs, 30):
        rows = client.get_rows(UNIT_ENTITY, {
            "$select": "Ref_Key,Description",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        })
        for row in rows:
            result[normalize_ref(row.get("Ref_Key"))] = str(row.get("Description") or row.get("Наименование") or "")
    return result


def extract_rows(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, dict):
        if isinstance(payload.get("value"), list):
            return payload["value"]
        data = payload.get("d")
        if isinstance(data, dict) and isinstance(data.get("results"), list):
            return data["results"]
        if isinstance(data, list):
            return data
    if isinstance(payload, list):
        return payload
    return []


def first_field(row: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in row:
            return row[name]
    return None


def normalize_ref(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        value = value.get("Ref_Key") or value.get("Ref") or value.get("Value")
    return str(value).strip().strip("{}")


def distinct_keys(values: Any) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = normalize_ref(value)
        if key and key not in seen:
            seen.add(key)
            result.append(key)
    return result


def decimal_value(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value).replace(",", "."))


def chunks(values: list[str], size: int) -> list[list[str]]:
    return [values[index:index + size] for index in range(0, len(values), size)]


def odata_string(value: str) -> str:
    return value.replace("'", "''")


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        text = line.strip()
        if not text or text.startswith("#") or "=" not in text:
            continue
        key, value = text.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def safe_base_url(value: str) -> str:
    return (value or os.environ.get("ONEC_ODATA_URL") or DEFAULT_ODATA_URL).rstrip("/") + "/"


if __name__ == "__main__":
    raise SystemExit(main())
