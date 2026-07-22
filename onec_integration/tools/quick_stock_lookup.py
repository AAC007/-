from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
DEFAULT_ENV = Path(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")
MAIN_REGISTER = "ТоварыНаСкладах"
CONTROL_REGISTERS = ("ПартииТоваровНаСкладах", "ТоварыОрганизаций")
ALIASES = ("code", "article", "name", "warehouse_code", "warehouse", "quantity")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="Быстрый read-only запрос остатка в 1С:УТ 10.3")
    parser.add_argument("item", nargs="?", help="Код, артикул или часть наименования номенклатуры")
    parser.add_argument("warehouse_arg", nargs="?", help="Код, номер или часть наименования склада/секции")
    parser.add_argument("--warehouse", default="", help="Склад/секция для пакетного режима --items-file")
    parser.add_argument("--env", default=str(DEFAULT_ENV), help="Путь к .env с ONEC_*")
    parser.add_argument("--out", default=str(REPORTS / "quick_stock_lookup.json"))
    parser.add_argument("--control", action="store_true", help="Дополнительно сверить контрольные регистры")
    parser.add_argument("--no-control", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--items-file", default="", help="UTF-8 файл со списком кодов/наименований, до 30 строк")
    args = parser.parse_args()
    items = read_items(args.item, args.items_file)
    warehouse = args.warehouse or args.warehouse_arg
    if not warehouse:
        raise ValueError("Укажите склад/секцию вторым аргументом или через --warehouse.")

    REPORTS.mkdir(parents=True, exist_ok=True)
    base, connected_server = connect(Path(args.env))
    registers = [MAIN_REGISTER, *CONTROL_REGISTERS] if args.control else [MAIN_REGISTER]
    result = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "database": os.environ.get("ONEC_DATABASE", "UT_dev"),
        "connected_server": connected_server,
        "item": args.item or "",
        "items": items,
        "warehouse": warehouse,
        "main_register": MAIN_REGISTER,
        "control_registers": list(CONTROL_REGISTERS) if args.control else [],
        "matches": [],
        "summary": [],
        "errors": [],
    }

    for item in items:
        for register in registers:
            try:
                rows = query_register(base, register, item, warehouse)
                if rows:
                    result["matches"].append({"item": item, "register": register, "rows": rows})
            except Exception as exc:
                result["errors"].append({"item": item, "register": register, "error": str(exc)})

    result["summary"] = summarize(result["matches"])
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print_human(result, out)
    return 0


def connect(env_path: Path):
    if env_path.exists():
        load_dotenv(env_path)
    servers = [os.environ.get("ONEC_SERVER", "pz-sql1.pzmc.org:1541")]
    if servers[0].lower().startswith("pz-sql1.pzmc.org"):
        servers.append("pz-sql1:1541")
    if "pz-sql1:1541" not in servers:
        servers.append("pz-sql1:1541")

    connector = win32com.client.Dispatch("V83.COMConnector")
    last_error: Exception | None = None
    for server in servers:
        try:
            connection_string = (
                f'Srvr="{server}";'
                f'Ref="{os.environ.get("ONEC_DATABASE", "UT_dev")}";'
                f'Usr="{os.environ.get("ONEC_USERNAME", "")}";'
                f'Pwd="{os.environ.get("ONEC_PASSWORD", "")}";'
            )
            return connector.Connect(connection_string), server
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Не удалось подключиться к 1С ни по одному адресу: {last_error}")


def query_register(base: Any, register: str, item: str, warehouse: str) -> list[dict[str, str]]:
    query_text = f"""
    ВЫБРАТЬ
        Остатки.Номенклатура.Код КАК code,
        Остатки.Номенклатура.Артикул КАК article,
        Остатки.Номенклатура.Наименование КАК name,
        Остатки.Склад.Код КАК warehouse_code,
        Остатки.Склад.Наименование КАК warehouse,
        Остатки.КоличествоОстаток КАК quantity
    ИЗ
        РегистрНакопления.{register}.Остатки(,
            (Номенклатура.Код = &ItemExact
             ИЛИ Номенклатура.Артикул = &ItemExact
             ИЛИ Номенклатура.Код = &ItemUtCode
             ИЛИ Номенклатура.Артикул = &ItemUtCode
             ИЛИ Номенклатура.Код ПОДОБНО &ItemPattern
             ИЛИ Номенклатура.Артикул ПОДОБНО &ItemPattern
             ИЛИ Номенклатура.Наименование ПОДОБНО &ItemPattern)
            И (Склад.Код ПОДОБНО &WarehousePattern
                ИЛИ Склад.Наименование ПОДОБНО &WarehousePattern)) КАК Остатки
    """
    query = base.NewObject("Запрос")
    query.Text = query_text
    query.SetParameter("ItemExact", item.strip())
    query.SetParameter("ItemUtCode", normalize_ut_code(item))
    query.SetParameter("ItemPattern", pattern(item))
    query.SetParameter("WarehousePattern", pattern(warehouse))
    selection = query.Execute().Choose()

    rows: list[dict[str, str]] = []
    while selection.Next():
        rows.append({alias: str(getattr(selection, alias)) for alias in ALIASES})
    return rows


def pattern(value: str) -> str:
    value = value.strip()
    if "%" in value:
        return value
    return f"%{value}%"


def normalize_ut_code(value: str) -> str:
    cleaned = value.strip().upper().replace(" ", "")
    if cleaned.startswith("УТ"):
        cleaned = cleaned[2:]
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    if not digits:
        return value.strip()
    return "УТ" + digits.zfill(9)


def read_items(item: str | None, items_file: str) -> list[str]:
    if items_file:
        path = Path(items_file)
        lines = [line.strip() for line in path.read_text(encoding="utf-8-sig").splitlines()]
        items = [line for line in lines if line and not line.startswith("#")]
        if len(items) > 30:
            raise ValueError("Пакетный быстрый запрос ограничен 30 строками.")
        return items
    if item:
        return [item]
    raise ValueError("Укажите item или --items-file.")


def summarize(matches: list[dict[str, Any]]) -> list[dict[str, str]]:
    totals: dict[tuple[str, str, str], Decimal] = {}
    for match in matches:
        if match.get("register") != MAIN_REGISTER:
            continue
        for row in match["rows"]:
            key = (match.get("item", ""), row["code"], row["name"], row["warehouse"])
            totals[key] = totals.get(key, Decimal("0")) + Decimal(row["quantity"].replace(",", "."))
    return [
        {
            "item": item,
            "code": code,
            "name": name,
            "warehouse": warehouse,
            "quantity": format(quantity.normalize(), "f"),
        }
        for (item, code, name, warehouse), quantity in totals.items()
    ]


def print_human(result: dict[str, Any], out: Path) -> None:
    print(f"1C stock lookup at {result['generated_at']}")
    print(f"Server: {result['connected_server']}")
    if not result["summary"]:
        print("Остаток не найден по заданным фильтрам.")
    for row in result["summary"]:
        print(f"{row['code']} | {row['name']} | {row['warehouse']} | Остаток: {row['quantity']}")
    if result["errors"]:
        print(f"Есть ошибки контрольных регистров: {len(result['errors'])}")
    print(f"JSON: {out}")


if __name__ == "__main__":
    raise SystemExit(main())

