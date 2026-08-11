from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
DEFAULT_ENV = Path(r"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Read-only поиск номенклатуры в 1С:УТ")
    parser.add_argument("--query", required=True, help="УТ-код, IPS, артикул или часть наименования")
    parser.add_argument("--limit", type=int, default=50)
    parser.add_argument("--env", default=str(DEFAULT_ENV), help="Путь к .env с ONEC_*")
    parser.add_argument("--out", default=str(REPORTS / "nomenclature_lookup.json"))
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    result = {
        "ok": False,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "database": os.environ.get("ONEC_DATABASE", "UT_dev"),
        "query": args.query,
        "items": [],
        "errors": [],
    }

    try:
        base, connected_server = connect(Path(args.env))
        result["connected_server"] = connected_server
        result["items"] = search(base, args.query, max(1, min(args.limit, 500)))
        result["ok"] = True
    except Exception as exc:
        result["errors"].append(str(exc))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON: {out}")
    return 0 if result["ok"] else 1


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


def search(base: Any, query_text: str, limit: int) -> list[dict[str, str]]:
    query = query_text.strip()
    if not query:
        return []

    onec_query = base.NewObject("Запрос")
    onec_query.Text = f"""
    ВЫБРАТЬ ПЕРВЫЕ {limit}
        Номенклатура.Код КАК code,
        Номенклатура.Артикул КАК article,
        Номенклатура.Наименование КАК name,
        Номенклатура.БазоваяЕдиницаИзмерения.Наименование КАК unit
    ИЗ
        Справочник.Номенклатура КАК Номенклатура
    ГДЕ
        НЕ Номенклатура.ПометкаУдаления
        И (
            Номенклатура.Код = &Exact
            ИЛИ Номенклатура.Артикул = &Exact
            ИЛИ Номенклатура.Код = &UtCode
            ИЛИ Номенклатура.Артикул = &UtCode
            ИЛИ Номенклатура.Код = &Numeric11
            ИЛИ Номенклатура.Артикул = &Numeric11
            ИЛИ Номенклатура.Код ПОДОБНО &Pattern
            ИЛИ Номенклатура.Артикул ПОДОБНО &Pattern
            ИЛИ Номенклатура.Наименование ПОДОБНО &Pattern
        )
    УПОРЯДОЧИТЬ ПО
        Номенклатура.Наименование
    """
    onec_query.SetParameter("Exact", query)
    onec_query.SetParameter("UtCode", normalize_ut_code(query))
    onec_query.SetParameter("Numeric11", normalize_numeric_code(query))
    onec_query.SetParameter("Pattern", pattern(query))
    selection = onec_query.Execute().Choose()

    rows: list[dict[str, str]] = []
    while selection.Next():
        rows.append({
            "code": str(selection.code),
            "article": str(selection.article),
            "name": str(selection.name),
            "unit": str(selection.unit),
        })
    return rows


def pattern(value: str) -> str:
    value = value.strip()
    return value if "%" in value else f"%{value}%"


def normalize_ut_code(value: str) -> str:
    cleaned = value.strip().upper().replace(" ", "")
    if cleaned.startswith("УТ"):
        cleaned = cleaned[2:]
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    return "УТ" + digits.zfill(9) if digits else value.strip()


def normalize_numeric_code(value: str) -> str:
    digits = "".join(ch for ch in value.strip() if ch.isdigit())
    return digits.zfill(11) if digits else value.strip()


if __name__ == "__main__":
    raise SystemExit(main())
