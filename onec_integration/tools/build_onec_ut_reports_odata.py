from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import sys
from collections import defaultdict
from datetime import date, datetime, time, timedelta
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

TRANSFER_DOCUMENT = "Document_ПеремещениеТоваров"
ASSEMBLY_DOCUMENT = "Document_КомплектацияНоменклатуры"
NOMENCLATURE_ENTITY = "Catalog_Номенклатура"
WAREHOUSE_ENTITY = "Catalog_Склады"
UNIT_ENTITY = "Catalog_ЕдиницыИзмерения"

WAREHOUSE_44_WIP = "ba59fa3b-e5f1-11eb-860c-0cc47adb8d79"
WAREHOUSE_SECTION_1_PRODUCTION = "04287e48-1b7d-11e8-8277-00155d586401"
WAREHOUSE_PAINT = "e1228c71-3243-11f1-866c-0cc47adb8d79"


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Read-only reports from 1C UT OData for Planning CMO")
    parser.add_argument("--start", default="", help="Start date, YYYY-MM-DD. Default: today.")
    parser.add_argument("--end", default="", help="End date inclusive, YYYY-MM-DD. Default: start date.")
    parser.add_argument("--include-unposted", action="store_true", help="Include unposted documents. Default: only Posted=true.")
    parser.add_argument("--env", default=str(DEFAULT_ENV))
    parser.add_argument("--odata-url", default="")
    parser.add_argument("--out", default=str(REPORTS / "onec_ut_period_reports.json"))
    parser.add_argument("--csv-dir", default=str(REPORTS))
    parser.add_argument("--max-docs", type=int, default=5000)
    parser.add_argument("--timeout", type=float, default=60.0)
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    load_env(Path(args.env))
    start_date, end_date = parse_period(args.start, args.end)

    result: dict[str, Any] = {
        "ok": False,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "provider": "odata",
        "period": {
            "start": start_date.isoformat(),
            "end": end_date.isoformat(),
            "include_unposted": args.include_unposted,
        },
        "reports": {
            "transfers_from_44_to_production_or_paint": {"rows": [], "summary": []},
            "assembly_materials_44_raw_to_details": {"rows": [], "summary": []},
        },
        "files": {},
        "errors": [],
    }

    try:
        client = ODataClient(args.odata_url or os.environ.get("ONEC_ODATA_URL") or DEFAULT_ODATA_URL, args.timeout)
        transfers = load_transfer_report(client, start_date, end_date, args.include_unposted, args.max_docs)
        assemblies = load_assembly_report(client, start_date, end_date, args.include_unposted, args.max_docs)
        result["reports"]["transfers_from_44_to_production_or_paint"] = transfers
        result["reports"]["assembly_materials_44_raw_to_details"] = assemblies
        result["ok"] = True
    except Exception as exc:
        result["errors"].append(str(exc))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    result["files"]["json"] = str(out)

    csv_dir = Path(args.csv_dir)
    csv_dir.mkdir(parents=True, exist_ok=True)
    period_slug = f"{start_date:%Y%m%d}_{end_date:%Y%m%d}"
    transfer_csv = csv_dir / f"onec_ut_transfers_from_44_{period_slug}.csv"
    assembly_csv = csv_dir / f"onec_ut_assembly_materials_44_{period_slug}.csv"
    write_csv(transfer_csv, result["reports"]["transfers_from_44_to_production_or_paint"]["rows"], TRANSFER_COLUMNS)
    write_csv(assembly_csv, result["reports"]["assembly_materials_44_raw_to_details"]["rows"], ASSEMBLY_COLUMNS)
    result["files"]["transfers_csv"] = str(transfer_csv)
    result["files"]["assembly_csv"] = str(assembly_csv)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    transfer_count = len(result["reports"]["transfers_from_44_to_production_or_paint"]["rows"])
    assembly_count = len(result["reports"]["assembly_materials_44_raw_to_details"]["rows"])
    print(f"Period: {start_date.isoformat()}..{end_date.isoformat()}, include_unposted={args.include_unposted}")
    print(f"Transfer rows: {transfer_count}")
    print(f"Assembly material rows: {assembly_count}")
    print(f"JSON: {out}")
    return 0 if result["ok"] else 1


class ODataClient:
    def __init__(self, base_url: str, timeout: float) -> None:
        self.base_url = base_url.rstrip("/") + "/"
        self.timeout = timeout
        self.username = os.environ.get("ONEC_USERNAME") or os.environ.get("ONEC_USER") or ""
        self.password = os.environ.get("ONEC_PASSWORD") or ""
        if not self.username or not self.password:
            raise RuntimeError("Не заданы ONEC_USERNAME/ONEC_USER и ONEC_PASSWORD для OData 1С.")

    def rows(self, path: str, params: dict[str, str] | None = None, max_rows: int = 5000) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        skip = 0
        page_size = min(500, max_rows)
        while len(result) < max_rows:
            page_params = dict(params or {})
            page_params.setdefault("$top", str(page_size))
            if skip:
                page_params["$skip"] = str(skip)
            payload = self.request_json("GET", path, page_params)
            rows = extract_rows(payload)
            if not rows:
                break
            result.extend(rows)
            if len(rows) < page_size:
                break
            skip += len(rows)
        return result[:max_rows]

    def request_json(self, method: str, path: str, params: dict[str, str] | None = None) -> Any:
        url = self.base_url + quote(path.lstrip("/"), safe="/(),'=$")
        query = {"$format": "json", **(params or {})}
        if query:
            url += "?" + urlencode(query, quote_via=quote)
        token = base64.b64encode(f"{self.username}:{self.password}".encode("utf-8")).decode("ascii")
        request = Request(url, method=method, headers={"Authorization": f"Basic {token}", "Accept": "application/json"})
        try:
            with urlopen(request, timeout=self.timeout) as response:
                return json.loads(response.read().decode("utf-8-sig"))
        except HTTPError as exc:
            text = exc.read().decode("utf-8-sig", errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {text[:1000]}") from exc


def load_transfer_report(client: ODataClient, start: date, end: date, include_unposted: bool, max_docs: int) -> dict[str, Any]:
    target_keys = [WAREHOUSE_SECTION_1_PRODUCTION, WAREHOUSE_PAINT]
    posted_filter = "" if include_unposted else " and Posted eq true"
    target_filter = " or ".join(f"СкладПолучатель_Key eq guid'{key}'" for key in target_keys)
    filter_text = (
        f"Date ge datetime'{start_datetime(start)}' and Date lt datetime'{end_datetime(end)}' "
        f"and СкладОтправитель_Key eq guid'{WAREHOUSE_44_WIP}' and ({target_filter}){posted_filter}"
    )
    docs = client.rows(TRANSFER_DOCUMENT, {
        "$filter": filter_text,
        "$orderby": "Date desc",
    }, max_rows=max_docs)
    nomenclature, units, warehouses = load_references_for_transfers(client, docs)

    rows: list[dict[str, Any]] = []
    for doc in docs:
        for line in doc.get("Товары") or []:
            item_key = normalize_ref(line.get("Номенклатура_Key"))
            unit_key = normalize_ref(line.get("ЕдиницаИзмерения_Key"))
            target_key = normalize_ref(doc.get("СкладПолучатель_Key"))
            item = nomenclature.get(item_key, {})
            rows.append({
                "document_number": str(doc.get("Number") or ""),
                "document_date": str(doc.get("Date") or ""),
                "posted": bool(doc.get("Posted")),
                "target_warehouse": warehouses.get(target_key, {}).get("name", target_key),
                "target_warehouse_code": warehouses.get(target_key, {}).get("code", ""),
                "line_number": str(line.get("LineNumber") or ""),
                "code": item.get("code", ""),
                "article": item.get("article", ""),
                "nomenclature": item.get("name", item_key),
                "unit": units.get(unit_key, ""),
                "quantity": decimal_text(line.get("Количество")),
                "comment": str(doc.get("Комментарий") or ""),
                "ref_key": normalize_ref(doc.get("Ref_Key")),
                "item_ref_key": item_key,
            })
    return {"rows": rows, "summary": summarize_transfer_rows(rows)}


def load_assembly_report(client: ODataClient, start: date, end: date, include_unposted: bool, max_docs: int) -> dict[str, Any]:
    posted_filter = "" if include_unposted else " and Posted eq true"
    filter_text = (
        f"Date ge datetime'{start_datetime(start)}' and Date lt datetime'{end_datetime(end)}' "
        f"and Склад_Key eq guid'{WAREHOUSE_44_WIP}'{posted_filter}"
    )
    docs = client.rows(ASSEMBLY_DOCUMENT, {
        "$filter": filter_text,
        "$orderby": "Date desc",
    }, max_rows=max_docs)
    nomenclature, units = load_references_for_assemblies(client, docs)

    rows: list[dict[str, Any]] = []
    for doc in docs:
        product_key = normalize_ref(doc.get("Номенклатура_Key"))
        product_unit_key = normalize_ref(doc.get("ЕдиницаИзмерения_Key"))
        product = nomenclature.get(product_key, {})
        for line in doc.get("Комплектующие") or []:
            material_key = normalize_ref(line.get("Номенклатура_Key"))
            material_unit_key = normalize_ref(line.get("ЕдиницаИзмерения_Key"))
            material = nomenclature.get(material_key, {})
            rows.append({
                "document_number": str(doc.get("Number") or ""),
                "document_date": str(doc.get("Date") or ""),
                "posted": bool(doc.get("Posted")),
                "product_code": product.get("code", ""),
                "product_article": product.get("article", ""),
                "product": product.get("name", product_key),
                "product_unit": units.get(product_unit_key, ""),
                "product_quantity": decimal_text(doc.get("Количество")),
                "material_code": material.get("code", ""),
                "material_article": material.get("article", ""),
                "material": material.get("name", material_key),
                "material_unit": units.get(material_unit_key, ""),
                "material_quantity": decimal_text(line.get("Количество")),
                "component_line_number": str(line.get("LineNumber") or ""),
                "comment": str(doc.get("Комментарий") or ""),
                "ref_key": normalize_ref(doc.get("Ref_Key")),
                "product_ref_key": product_key,
                "material_ref_key": material_key,
            })
    return {"rows": rows, "summary": summarize_assembly_rows(rows)}


def load_references_for_transfers(client: ODataClient, docs: list[dict[str, Any]]) -> tuple[dict[str, dict[str, str]], dict[str, str], dict[str, dict[str, str]]]:
    item_refs = sorted({normalize_ref(line.get("Номенклатура_Key")) for doc in docs for line in (doc.get("Товары") or []) if normalize_ref(line.get("Номенклатура_Key"))})
    unit_refs = sorted({normalize_ref(line.get("ЕдиницаИзмерения_Key")) for doc in docs for line in (doc.get("Товары") or []) if normalize_ref(line.get("ЕдиницаИзмерения_Key"))})
    warehouse_refs = sorted({normalize_ref(doc.get("СкладПолучатель_Key")) for doc in docs if normalize_ref(doc.get("СкладПолучатель_Key"))})
    return load_nomenclature(client, item_refs), load_units(client, unit_refs), load_warehouses(client, warehouse_refs)


def load_references_for_assemblies(client: ODataClient, docs: list[dict[str, Any]]) -> tuple[dict[str, dict[str, str]], dict[str, str]]:
    item_refs: set[str] = set()
    unit_refs: set[str] = set()
    for doc in docs:
        item_refs.add(normalize_ref(doc.get("Номенклатура_Key")))
        unit_refs.add(normalize_ref(doc.get("ЕдиницаИзмерения_Key")))
        for line in doc.get("Комплектующие") or []:
            item_refs.add(normalize_ref(line.get("Номенклатура_Key")))
            unit_refs.add(normalize_ref(line.get("ЕдиницаИзмерения_Key")))
    item_refs.discard("")
    unit_refs.discard("")
    return load_nomenclature(client, sorted(item_refs)), load_units(client, sorted(unit_refs))


def load_nomenclature(client: ODataClient, refs: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for chunk in chunks(refs, 20):
        rows = client.rows(NOMENCLATURE_ENTITY, {
            "$select": "Ref_Key,Code,Description,Артикул",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        }, max_rows=1000)
        for row in rows:
            ref = normalize_ref(row.get("Ref_Key"))
            result[ref] = {
                "code": str(row.get("Code") or ""),
                "article": str(row.get("Артикул") or ""),
                "name": str(row.get("Description") or ""),
            }
    return result


def load_units(client: ODataClient, refs: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for chunk in chunks(refs, 30):
        rows = client.rows(UNIT_ENTITY, {
            "$select": "Ref_Key,Description",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        }, max_rows=1000)
        for row in rows:
            result[normalize_ref(row.get("Ref_Key"))] = str(row.get("Description") or "")
    return result


def load_warehouses(client: ODataClient, refs: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for chunk in chunks(refs, 30):
        rows = client.rows(WAREHOUSE_ENTITY, {
            "$select": "Ref_Key,Code,Description",
            "$filter": " or ".join(f"Ref_Key eq guid'{ref}'" for ref in chunk),
        }, max_rows=1000)
        for row in rows:
            ref = normalize_ref(row.get("Ref_Key"))
            result[ref] = {
                "code": str(row.get("Code") or ""),
                "name": str(row.get("Description") or ""),
            }
    return result


def summarize_transfer_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str, str], Decimal] = defaultdict(Decimal)
    docs: dict[tuple[str, str, str, str], set[str]] = defaultdict(set)
    for row in rows:
        key = (row["target_warehouse"], row["code"], row["nomenclature"], row["unit"])
        groups[key] += decimal_value(row["quantity"])
        docs[key].add(row["document_number"])
    return [
        {
            "target_warehouse": key[0],
            "code": key[1],
            "nomenclature": key[2],
            "unit": key[3],
            "quantity": format_decimal(quantity),
            "document_count": len(docs[key]),
        }
        for key, quantity in sorted(groups.items(), key=lambda item: (item[0][0], item[0][2], item[0][1]))
    ]


def summarize_assembly_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str, str, str, str], Decimal] = defaultdict(Decimal)
    product_qty: dict[tuple[str, str, str, str, str, str], Decimal] = defaultdict(Decimal)
    docs: dict[tuple[str, str, str, str, str, str], set[str]] = defaultdict(set)
    for row in rows:
        key = (row["material_code"], row["material"], row["material_unit"], row["product_code"], row["product"], row["product_unit"])
        groups[key] += decimal_value(row["material_quantity"])
        product_qty[key] += decimal_value(row["product_quantity"])
        docs[key].add(row["document_number"])
    return [
        {
            "material_code": key[0],
            "material": key[1],
            "material_unit": key[2],
            "material_quantity": format_decimal(quantity),
            "product_code": key[3],
            "product": key[4],
            "product_unit": key[5],
            "product_quantity": format_decimal(product_qty[key]),
            "document_count": len(docs[key]),
        }
        for key, quantity in sorted(groups.items(), key=lambda item: (item[0][1], item[0][4]))
    ]


TRANSFER_COLUMNS = [
    "document_number", "document_date", "posted", "target_warehouse", "target_warehouse_code",
    "line_number", "code", "article", "nomenclature", "unit", "quantity", "comment", "ref_key", "item_ref_key",
]
ASSEMBLY_COLUMNS = [
    "document_number", "document_date", "posted", "product_code", "product_article", "product",
    "product_unit", "product_quantity", "material_code", "material_article", "material", "material_unit",
    "material_quantity", "component_line_number", "comment", "ref_key", "product_ref_key", "material_ref_key",
]


def write_csv(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="ignore", delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def parse_period(start_text: str, end_text: str) -> tuple[date, date]:
    start = date.fromisoformat(start_text) if start_text else date.today()
    end = date.fromisoformat(end_text) if end_text else start
    if end < start:
        raise RuntimeError("Дата окончания периода меньше даты начала.")
    return start, end


def start_datetime(value: date) -> str:
    return datetime.combine(value, time.min).strftime("%Y-%m-%dT%H:%M:%S")


def end_datetime(value: date) -> str:
    return datetime.combine(value + timedelta(days=1), time.min).strftime("%Y-%m-%dT%H:%M:%S")


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


def chunks(values: list[str], size: int) -> list[list[str]]:
    return [values[index:index + size] for index in range(0, len(values), size)]


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


def normalize_ref(value: Any) -> str:
    text = str(value or "").strip()
    return "" if text == "00000000-0000-0000-0000-000000000000" else text


def decimal_value(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    return Decimal(str(value).replace(",", "."))


def decimal_text(value: Any) -> str:
    return format_decimal(decimal_value(value))


def format_decimal(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(value.quantize(Decimal("1")))
    return format(value.normalize(), "f")


if __name__ == "__main__":
    raise SystemExit(main())
