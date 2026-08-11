from __future__ import annotations

import argparse
import base64
import json
import os
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
DEFAULT_ENV = Path(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")
DEFAULT_ODATA_URL = "http://pz-sql1/ut_dev/odata/standard.odata/"

NOMENCLATURE_ENTITY = "Catalog_Номенклатура"
ASSEMBLY_DOCUMENT = "Document_КомплектацияНоменклатуры"

ORGANIZATION_KEY = "163b8f79-346a-11e6-80b6-00155d0c9b04"  # СТП ПЗМЦ АО
WAREHOUSE_KEY = "ba59fa3b-e5f1-11eb-860c-0cc47adb8d79"  # 44 секция НЗП
RESPONSIBLE_KEY = "ce72f423-5baf-11ef-864f-0cc47adb8d79"  # Аракелян А.С.
DEPARTMENT_KEY = "14c61f15-1510-11f0-8660-0cc47adb8d79"  # Цех механической обработки
EMPTY_REF = "00000000-0000-0000-0000-000000000000"


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Create 1C production assembly document from Планирование ЦМО payload")
    parser.add_argument("--payload", required=True)
    parser.add_argument("--out", default=str(REPORTS / "production_launch_result.json"))
    parser.add_argument("--env", default=str(DEFAULT_ENV))
    parser.add_argument("--odata-url", default="")
    parser.add_argument("--allow-write", action="store_true")
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    load_env(Path(args.env))
    payload = json.loads(Path(args.payload).read_text(encoding="utf-8-sig"))
    result: dict[str, Any] = {
        "ok": False,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "number": "",
        "ref_key": "",
        "posted": False,
        "comment": payload.get("Comment") or payload.get("comment") or "",
        "created": None,
        "errors": [],
    }

    try:
        if not args.allow_write:
            raise RuntimeError("Запись документа 1С заблокирована: не передан флаг --allow-write.")

        client = ODataClient(args.odata_url or os.environ.get("ONEC_ODATA_URL") or DEFAULT_ODATA_URL)
        created = create_assembly(client, payload)
        result["ok"] = True
        result["created"] = created
        result["number"] = str(created.get("Number") or "")
        result["ref_key"] = str(created.get("Ref_Key") or "")
        result["posted"] = bool(created.get("Posted"))
        result["comment"] = str(created.get("Комментарий") or result["comment"])
    except Exception as exc:
        result["errors"].append(str(exc))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON: {out}")
    return 0 if result["ok"] else 1


class ODataClient:
    def __init__(self, base_url: str) -> None:
        self.base_url = base_url.rstrip("/") + "/"
        self.username = os.environ.get("ONEC_USERNAME") or os.environ.get("ONEC_USER") or ""
        self.password = os.environ.get("ONEC_PASSWORD") or ""
        if not self.username or not self.password:
            raise RuntimeError("Не заданы ONEC_USERNAME/ONEC_USER и ONEC_PASSWORD для OData 1С.")

    def get_rows(self, path: str, params: dict[str, str] | None = None) -> list[dict[str, Any]]:
        payload = self.request_json("GET", path, params=params)
        if isinstance(payload, dict):
            if isinstance(payload.get("value"), list):
                return payload["value"]
            data = payload.get("d")
            if isinstance(data, dict) and isinstance(data.get("results"), list):
                return data["results"]
            if isinstance(data, list):
                return data
        return []

    def post_json(self, path: str, body: dict[str, Any]) -> dict[str, Any]:
        payload = self.request_json("POST", path, body=body)
        if isinstance(payload, dict):
            data = payload.get("d")
            if isinstance(data, dict):
                return data
            return payload
        raise RuntimeError("OData вернул неожиданный ответ при создании документа.")

    def request_json(self, method: str, path: str, params: dict[str, str] | None = None, body: dict[str, Any] | None = None) -> Any:
        url = self.base_url + quote(path.lstrip("/"), safe="/(),'=$")
        query = {"$format": "json", **(params or {})}
        if query:
            url += "?" + encode_query(query)
        token = base64.b64encode(f"{self.username}:{self.password}".encode("utf-8")).decode("ascii")
        headers = {
            "Authorization": f"Basic {token}",
            "Accept": "application/json",
        }
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json; charset=utf-8"
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        request = Request(url, data=data, method=method, headers=headers)
        try:
            with urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8-sig"))
        except HTTPError as exc:
            text = exc.read().decode("utf-8-sig", errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {text[:1000]}") from exc


def create_assembly(client: ODataClient, payload: dict[str, Any]) -> dict[str, Any]:
    ips = str(payload.get("Ips") or payload.get("ips") or "").strip()
    part_code = str(payload.get("OneCPartCode") or payload.get("oneCPartCode") or normalize_numeric_code(ips)).strip()
    quantity = decimal_value(payload.get("Quantity") or payload.get("quantity"))
    if quantity <= 0:
        raise RuntimeError("Количество запуска должно быть больше 0.")

    components = payload.get("Components") or payload.get("components") or []
    if not components:
        raise RuntimeError("В заявке запуска нет комплектующих.")

    part = resolve_nomenclature(client, part_code)
    component_rows: list[dict[str, Any]] = []
    for index, component in enumerate(components, start=1):
        code = str(component.get("OneCCode") or component.get("oneCCode") or "").strip()
        component_item = resolve_nomenclature(client, code)
        component_rows.append({
            "LineNumber": str(index),
            "Номенклатура_Key": component_item["ref_key"],
            "ЕдиницаИзмерения_Key": component_item["unit_key"],
            "Количество": float(decimal_value(component.get("Quantity") or component.get("quantity"))),
            "Коэффициент": 1,
            "СерияНоменклатуры_Key": EMPTY_REF,
            "ХарактеристикаНоменклатуры_Key": EMPTY_REF,
            "СпособСписанияОстаткаТоваров": "СоСклада",
            "ДоляСтоимости": 0,
            "УдалитьСумма": 0,
            "ЦенаЗакупа": 0,
        })

    document = {
        "Date": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "DeletionMark": False,
        "Posted": False,
        "ВидОперации": "Комплектация",
        "Организация_Key": ORGANIZATION_KEY,
        "Склад_Key": WAREHOUSE_KEY,
        "Ответственный_Key": RESPONSIBLE_KEY,
        "Подразделение_Key": DEPARTMENT_KEY,
        "Номенклатура_Key": part["ref_key"],
        "ЕдиницаИзмерения_Key": part["unit_key"],
        "Количество": float(quantity),
        "Коэффициент": 1,
        "Комментарий": str(payload.get("Comment") or payload.get("comment") or "").strip(),
        "СерияНоменклатуры_Key": EMPTY_REF,
        "ХарактеристикаНоменклатуры_Key": EMPTY_REF,
        "Проект_Key": EMPTY_REF,
        "Заказ": "",
        "Заказ_Type": "StandardODATA.Undefined",
        "КодНоменклатуры": "",
        "СпособСписанияОстаткаТоваров": "СоСклада",
        "ОтражатьВУправленческомУчете": True,
        "ОтражатьВНалоговомУчете": True,
        "ОтражатьВБухгалтерскомУчете": True,
        "НДСвСтоимостиТоваров": "",
        "УдалитьСуммаДокумента": 0,
        "Комплектующие": component_rows,
    }
    return client.post_json(ASSEMBLY_DOCUMENT, document)


def resolve_nomenclature(client: ODataClient, query: str) -> dict[str, str]:
    candidates = [query, normalize_ut_code(query), normalize_numeric_code(query)]
    rows: list[dict[str, Any]] = []
    errors: list[str] = []
    select = "Ref_Key,Code,Description,Артикул,ЕдиницаХраненияОстатков_Key,ЕдиницаДляОтчетов_Key,БазоваяЕдиницаИзмерения_Key"
    for candidate in unique(candidates):
        escaped = odata_string(candidate)
        for field in ("Code", "Артикул"):
            try:
                rows.extend(client.get_rows(NOMENCLATURE_ENTITY, {
                    "$top": "5",
                    "$select": select,
                    "$filter": f"{field} eq '{escaped}'",
                }))
            except Exception as exc:
                errors.append(f"{field}={candidate}: {exc}")
        if rows:
            break
    if not rows:
        details = "; ".join(errors[:3])
        suffix = f" ({details})" if details else ""
        raise RuntimeError(f"Номенклатура 1С не найдена по коду: {query}{suffix}")
    row = sorted(rows, key=lambda item: nomenclature_score(item, candidates), reverse=True)[0]
    return {
        "ref_key": normalize_ref(row.get("Ref_Key")),
        "code": str(row.get("Code") or ""),
        "article": str(row.get("Артикул") or ""),
        "name": str(row.get("Description") or ""),
        "unit_key": normalize_ref(row.get("ЕдиницаХраненияОстатков_Key") or row.get("ЕдиницаДляОтчетов_Key") or row.get("БазоваяЕдиницаИзмерения_Key")),
    }


def nomenclature_score(row: dict[str, Any], candidates: list[str]) -> tuple[int, str]:
    code = str(row.get("Code") or "")
    article = str(row.get("Артикул") or "")
    return (3 if code in candidates else 2 if article in candidates else 0, str(row.get("Description") or ""))


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = value.strip()
        if text and text not in seen:
            seen.add(text)
            result.append(text)
    return result


def normalize_ref(value: Any) -> str:
    if value is None:
        return EMPTY_REF
    if isinstance(value, dict):
        value = value.get("Ref_Key") or value.get("Ref") or value.get("Value")
    text = str(value).strip().strip("{}")
    return text or EMPTY_REF


def decimal_value(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value).replace(",", "."))


def normalize_numeric_code(value: str) -> str:
    digits = "".join(ch for ch in str(value).strip() if ch.isdigit())
    return digits.zfill(11) if digits else str(value).strip()


def normalize_ut_code(value: str) -> str:
    text = str(value).strip().upper().replace(" ", "")
    if text.startswith("УТ"):
        text = text[2:]
    digits = "".join(ch for ch in text if ch.isdigit())
    return "УТ" + digits.zfill(9) if digits else str(value).strip()


def odata_string(value: str) -> str:
    return value.replace("'", "''")


def encode_query(values: dict[str, str]) -> str:
    value_safe_chars = "$'(),=:"
    return "&".join(
        f"{quote(str(key), safe='$')}={quote(str(value), safe=value_safe_chars)}"
        for key, value in values.items()
    )


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        text = line.strip()
        if not text or text.startswith("#") or "=" not in text:
            continue
        key, value = text.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


if __name__ == "__main__":
    raise SystemExit(main())
