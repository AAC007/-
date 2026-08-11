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
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
DEFAULT_ENV = Path(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")
DEFAULT_ODATA_URL = "http://pz-sql1/ut_dev/odata/standard.odata/"

NOMENCLATURE_ENTITY = "Catalog_Номенклатура"
WAREHOUSE_ENTITY = "Catalog_Склады"
ORGANIZATION_ENTITY = "Catalog_Организации"
TRANSFER_DOCUMENT = "Document_ПеремещениеТоваров"

DEFAULT_ORGANIZATION = "СТП ПЗМЦ АО"
DEFAULT_SOURCE_WAREHOUSE = "44 секция НЗП (незавершенное производство)"
DEFAULT_TARGET_WAREHOUSE = "Детали МУ в обработке на стороне"
RESPONSIBLE_KEY = "ce72f423-5baf-11ef-864f-0cc47adb8d79"  # Аракелян А.С.
DEPARTMENT_KEY = "14c61f15-1510-11f0-8660-0cc47adb8d79"  # Цех механической обработки
EMPTY_REF = "00000000-0000-0000-0000-000000000000"
QUALITY_NEW_KEY = "d05404a0-6bce-449b-a798-41ebe5e5b977"


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Create unposted 1C goods transfer document through OData")
    parser.add_argument("--item", action="append", nargs=2, metavar=("CODE", "QTY"), default=[], help="Nomenclature code/article and quantity. Repeat for multiple rows.")
    parser.add_argument("--source", default=DEFAULT_SOURCE_WAREHOUSE)
    parser.add_argument("--target", default=DEFAULT_TARGET_WAREHOUSE)
    parser.add_argument("--organization", default=DEFAULT_ORGANIZATION)
    parser.add_argument("--comment", default="Создано ПО Планирование ЦМО")
    parser.add_argument("--empty-trial", action="store_true", help="Allow creating a trial document without goods rows.")
    parser.add_argument("--out", default=str(REPORTS / "goods_transfer_result.json"))
    parser.add_argument("--env", default=str(DEFAULT_ENV))
    parser.add_argument("--odata-url", default="")
    parser.add_argument("--allow-write", action="store_true")
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    load_env(Path(args.env))
    result: dict[str, Any] = {
        "ok": False,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "number": "",
        "ref_key": "",
        "posted": False,
        "created": None,
        "verified": None,
        "errors": [],
    }

    try:
        if not args.allow_write:
            raise RuntimeError("Запись документа 1С заблокирована: не передан флаг --allow-write.")
        if not args.item and not args.empty_trial:
            raise RuntimeError("Для перемещения нужны строки --item CODE QTY. Пустой пробный документ разрешен только с --empty-trial.")

        client = ODataClient(args.odata_url or os.environ.get("ONEC_ODATA_URL") or DEFAULT_ODATA_URL)
        created = create_transfer(client, args)
        ref_key = normalize_ref(created.get("Ref_Key"))
        verified = client.request_json("GET", f"{TRANSFER_DOCUMENT}(guid'{ref_key}')")
        result.update({
            "ok": True,
            "number": str(created.get("Number") or ""),
            "ref_key": ref_key,
            "posted": bool(created.get("Posted")),
            "created": created,
            "verified": verified,
        })
    except Exception as exc:
        result["errors"].append(str(exc))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON: {out}")
    if result["ok"]:
        print(f"Document: {result['number']} Ref_Key={result['ref_key']} Posted={result['posted']}")
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
            return data if isinstance(data, dict) else payload
        raise RuntimeError("OData вернул неожиданный ответ при создании документа.")

    def request_json(self, method: str, path: str, params: dict[str, str] | None = None, body: dict[str, Any] | None = None) -> Any:
        url = self.base_url + quote(path.lstrip("/"), safe="/(),'=$")
        query = {"$format": "json", **(params or {})}
        if query:
            url += "?" + urlencode(query, quote_via=quote)
        token = base64.b64encode(f"{self.username}:{self.password}".encode("utf-8")).decode("ascii")
        headers = {"Authorization": f"Basic {token}", "Accept": "application/json"}
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


def create_transfer(client: ODataClient, args: argparse.Namespace) -> dict[str, Any]:
    source = resolve_by_description(client, WAREHOUSE_ENTITY, args.source, "склад-отправитель")
    target = resolve_by_description(client, WAREHOUSE_ENTITY, args.target, "склад-получатель")
    organization = resolve_by_description(client, ORGANIZATION_ENTITY, args.organization, "организация")
    if source["ref_key"] == target["ref_key"]:
        raise RuntimeError("Склад-отправитель и склад-получатель должны отличаться.")

    goods_rows = []
    for index, (code, quantity_text) in enumerate(args.item, start=1):
        item = resolve_nomenclature(client, code)
        quantity = decimal_value(quantity_text)
        if quantity <= 0:
            raise RuntimeError(f"Количество должно быть больше 0: {code} {quantity_text}")
        goods_rows.append({
            "LineNumber": str(index),
            "Номенклатура_Key": item["ref_key"],
            "КоличествоМест": 0,
            "ЕдиницаИзмерения_Key": item["unit_key"],
            "ЕдиницаИзмеренияМест_Key": EMPTY_REF,
            "Коэффициент": 1,
            "Количество": float(quantity),
            "Цена": 0,
            "ХарактеристикаНоменклатуры_Key": EMPTY_REF,
            "УдалитьСпособСписанияОстаткаТоваров": "",
            "СерияНоменклатуры_Key": EMPTY_REF,
            "ДокументРезерва": "",
            "ДокументРезерва_Type": "StandardODATA.Undefined",
            "Качество_Key": QUALITY_NEW_KEY,
            "ВнутреннийЗаказ_Key": EMPTY_REF,
            "НоменклатураIPS_Key": EMPTY_REF,
            "МестоХранения_Key": EMPTY_REF,
            "Ост": 0,
        })

    document = {
        "Date": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "DeletionMark": False,
        "Posted": False,
        "ВидОперации": "",
        "Организация_Key": organization["ref_key"],
        "Ответственный_Key": RESPONSIBLE_KEY,
        "Подразделение_Key": DEPARTMENT_KEY,
        "СкладОтправитель_Key": source["ref_key"],
        "СкладПолучатель_Key": target["ref_key"],
        "ОтражатьВУправленческомУчете": True,
        "ОтражатьВБухгалтерскомУчете": True,
        "ОтражатьВНалоговомУчете": True,
        "Комментарий": args.comment.strip(),
        "ВнутреннийЗаказ_Key": EMPTY_REF,
        "Проект_Key": EMPTY_REF,
        "ДокументОснование_Key": EMPTY_REF,
        "ОтветственныйСБ_Key": EMPTY_REF,
        "ТТНИсходящая_Key": EMPTY_REF,
        "ТТНВходящая_Key": EMPTY_REF,
        "Товары": goods_rows,
        "ВозвратнаяТара": [],
    }
    return client.post_json(TRANSFER_DOCUMENT, document)


def resolve_by_description(client: ODataClient, entity: str, query: str, label: str) -> dict[str, str]:
    rows = client.get_rows(entity, {"$select": "Ref_Key,Code,Description", "$top": "1000"})
    normalized_query = query.casefold()
    matches = [row for row in rows if normalized_query in str(row.get("Description") or "").casefold()]
    if not matches:
        raise RuntimeError(f"Не найдено через OData: {label} '{query}'")
    exact = [row for row in matches if str(row.get("Description") or "").casefold() == normalized_query]
    row = (exact or matches)[0]
    return {
        "ref_key": normalize_ref(row.get("Ref_Key")),
        "code": str(row.get("Code") or ""),
        "name": str(row.get("Description") or ""),
    }


def resolve_nomenclature(client: ODataClient, query: str) -> dict[str, str]:
    candidates = unique([query, normalize_ut_code(query), normalize_numeric_code(query)])
    select = "Ref_Key,Code,Description,Артикул,ЕдиницаХраненияОстатков_Key,ЕдиницаДляОтчетов_Key,БазоваяЕдиницаИзмерения_Key"
    rows: list[dict[str, Any]] = []
    errors: list[str] = []
    for candidate in candidates:
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
        suffix = f" ({'; '.join(errors[:3])})" if errors else ""
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


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        os.environ.setdefault(key, value)


def decimal_value(value: Any) -> Decimal:
    return Decimal(str(value).replace(",", "."))


def normalize_ref(value: Any) -> str:
    text = str(value or "").strip()
    return text or EMPTY_REF


def normalize_ut_code(value: str) -> str:
    text = value.strip().upper()
    if text.startswith("УТ") and len(text) < 10:
        digits = "".join(ch for ch in text[2:] if ch.isdigit())
        return "УТ" + digits.zfill(9)
    return text


def normalize_numeric_code(value: str) -> str:
    text = value.strip()
    return text.zfill(11) if text.isdigit() and len(text) <= 11 else text


def odata_string(value: str) -> str:
    return value.replace("'", "''")


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = value.strip()
        if text and text not in seen:
            seen.add(text)
            result.append(text)
    return result


if __name__ == "__main__":
    raise SystemExit(main())
