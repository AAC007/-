from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
REPORTS.mkdir(parents=True, exist_ok=True)


def connect():
    load_dotenv(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")
    server = os.environ.get("ONEC_SERVER", "pz-sql1.pzmc.org:1541")
    if server.lower().startswith("pz-sql1.pzmc.org"):
        server = "pz-sql1:1541"
    connection_string = (
        f'Srvr="{server}";'
        f'Ref="{os.environ.get("ONEC_DATABASE", "UT_dev")}";'
        f'Usr="{os.environ.get("ONEC_USERNAME", "")}";'
        f'Pwd="{os.environ.get("ONEC_PASSWORD", "")}";'
    )
    connector = win32com.client.Dispatch("V83.COMConnector")
    return connector.Connect(connection_string)


def dump_stock_metadata(base):
    rows = []
    for register in base.Metadata.AccumulationRegisters:
        name = str(register.Name)
        synonym = str(register.Synonym)
        text = (name + " " + synonym).lower()
        if any(word in text for word in ("товар", "склад", "остат", "резерв", "парт")):
            rows.append(
                {
                    "name": name,
                    "synonym": synonym,
                    "dimensions": [{"name": str(x.Name), "synonym": str(x.Synonym)} for x in register.Dimensions],
                    "resources": [{"name": str(x.Name), "synonym": str(x.Synonym)} for x in register.Resources],
                }
            )
    path = REPORTS / "live_stock_metadata.json"
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return rows


def query_rows(base, query_text, params, aliases):
    query = base.NewObject("Запрос")
    query.Text = query_text
    for key, value in params.items():
        query.SetParameter(key, value)
    selection = query.Execute().Choose()
    rows = []
    while selection.Next():
        row = {}
        for name in aliases:
            row[name] = str(getattr(selection, name))
        rows.append(row)
    return rows


def try_stock_query(base, register_name, item_pattern, warehouse_pattern):
    query_text = f"""
    ВЫБРАТЬ
        Остатки.Номенклатура.Код КАК code,
        Остатки.Номенклатура.Артикул КАК article,
        Остатки.Номенклатура.Наименование КАК name,
        Остатки.Склад.Код КАК warehouse_code,
        Остатки.Склад.Наименование КАК warehouse,
        Остатки.КоличествоОстаток КАК quantity
    ИЗ
        РегистрНакопления.{register_name}.Остатки(,
            (Номенклатура.Код ПОДОБНО &ItemPattern
             ИЛИ Номенклатура.Артикул ПОДОБНО &ItemPattern
             ИЛИ Номенклатура.Наименование ПОДОБНО &ItemPattern)
            И Склад.Наименование ПОДОБНО &WarehousePattern) КАК Остатки
    """
    return query_rows(
        base,
        query_text,
        {
            "ItemPattern": item_pattern,
            "WarehousePattern": warehouse_pattern,
        },
        ["code", "article", "name", "warehouse_code", "warehouse", "quantity"],
    )


def main() -> int:
    item_pattern = sys.argv[1] if len(sys.argv) > 1 else "%01-P002-03.002-T63.150.SS%"
    warehouse_pattern = sys.argv[2] if len(sys.argv) > 2 else "%44%"
    base = connect()
    metadata = dump_stock_metadata(base)
    candidates = [x["name"] for x in metadata if any(d["name"] == "Склад" for d in x["dimensions"]) and any(d["name"] == "Номенклатура" for d in x["dimensions"]) and any(r["name"] == "Количество" for r in x["resources"])]
    result = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "item_pattern": item_pattern,
        "warehouse_pattern": warehouse_pattern,
        "candidate_registers": candidates,
        "matches": [],
        "errors": [],
    }
    for register_name in candidates:
        try:
            rows = try_stock_query(base, register_name, item_pattern, warehouse_pattern)
            if rows:
                result["matches"].append({"register": register_name, "rows": rows})
        except Exception as exc:
            result["errors"].append({"register": register_name, "error": str(exc)})
    out = REPORTS / "live_stock_lookup.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

