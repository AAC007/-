from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from quick_stock_lookup import DEFAULT_ENV, MAIN_REGISTER, connect, pattern


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
ALIASES = ("code", "article", "name", "unit", "warehouse_code", "warehouse", "quantity")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Read-only синхронизация остатков 1С по секциям ЦМО/ПЗМЦ")
    parser.add_argument("--env", default=str(DEFAULT_ENV), help="Путь к .env с ONEC_*")
    parser.add_argument("--wip-warehouse", nargs="+", default=["44"], help="Секции НЗП/ЦМО")
    parser.add_argument("--production-warehouse", nargs="+", default=["1"], help="Склады производственных остатков")
    parser.add_argument("--out", default=str(REPORTS / "stock_sections_sync.json"))
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    base, connected_server = connect(Path(args.env))
    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    result = {
        "ok": True,
        "generated_at": generated_at,
        "connected_server": connected_server,
        "main_register": MAIN_REGISTER,
        "warehouses": {
            "wip": args.wip_warehouse,
            "production": args.production_warehouse,
        },
        "rows": [],
        "errors": [],
    }

    for role, warehouses in (("wip", args.wip_warehouse), ("production", args.production_warehouse)):
        try:
            result["rows"].extend(query_positive_stock(base, warehouses, role))
        except Exception as exc:
            result["errors"].append({"warehouse": "; ".join(warehouses), "role": role, "error": str(exc)})

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"1C stock section sync at {generated_at}")
    print(f"Server: {connected_server}")
    print(f"Rows: {len(result['rows'])}")
    if result["errors"]:
        print(f"Errors: {len(result['errors'])}")
    print(f"JSON: {out}")
    return 0 if not result["errors"] else 2


def query_positive_stock(base: Any, warehouses: list[str], role: str) -> list[dict[str, str]]:
    conditions: list[str] = []
    prepared: list[tuple[str, str, str]] = []
    for index, warehouse in enumerate(x.strip() for x in warehouses if x.strip()):
        exact_name = f"WarehouseExact{index}"
        pattern_name = f"WarehouseSectionPattern{index}"
        conditions.append(
            f"(Склад.Код = &{exact_name} "
            f"ИЛИ Склад.Наименование = &{exact_name} "
            f"ИЛИ Склад.Наименование ПОДОБНО &{pattern_name})"
        )
        prepared.append((exact_name, pattern_name, warehouse))
    if not prepared:
        return []

    warehouse_condition = "\n             ИЛИ ".join(conditions)
    query_text = f"""
    ВЫБРАТЬ
        Остатки.Номенклатура.Код КАК code,
        Остатки.Номенклатура.Артикул КАК article,
        Остатки.Номенклатура.Наименование КАК name,
        Остатки.Номенклатура.БазоваяЕдиницаИзмерения.Наименование КАК unit,
        Остатки.Склад.Код КАК warehouse_code,
        Остатки.Склад.Наименование КАК warehouse,
        Остатки.КоличествоОстаток КАК quantity
    ИЗ
        РегистрНакопления.{MAIN_REGISTER}.Остатки(,
            ({warehouse_condition})) КАК Остатки
    ГДЕ
        Остатки.КоличествоОстаток > 0
    """
    query = base.NewObject("Запрос")
    query.Text = query_text
    for exact_name, pattern_name, warehouse in prepared:
        query.SetParameter(exact_name, warehouse)
        query.SetParameter(pattern_name, f"{warehouse} секция%")
    selection = query.Execute().Choose()

    rows: list[dict[str, str]] = []
    while selection.Next():
        row = {alias: str(getattr(selection, alias)) for alias in ALIASES}
        row["role"] = role
        row["quantity"] = format(Decimal(row["quantity"].replace(",", ".")).normalize(), "f")
        rows.append(row)
    return rows


if __name__ == "__main__":
    raise SystemExit(main())
