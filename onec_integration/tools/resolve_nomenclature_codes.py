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

    parser = argparse.ArgumentParser(description="Read-only пакетное получение наименований номенклатуры 1С:УТ")
    parser.add_argument("--code", nargs="+", default=[], help="УТ-коды, IPS или коды 1С")
    parser.add_argument("--env", default=str(DEFAULT_ENV), help="Путь к .env с ONEC_*")
    parser.add_argument("--out", default=str(REPORTS / "nomenclature_codes_lookup.json"))
    args = parser.parse_args()

    REPORTS.mkdir(parents=True, exist_ok=True)
    codes = unique([code.strip() for code in args.code if code.strip()])[:500]
    result = {
        "ok": False,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "database": os.environ.get("ONEC_DATABASE", "UT_dev"),
        "requested": codes,
        "items": [],
        "errors": [],
    }

    try:
        base, connected_server = connect(Path(args.env))
        result["connected_server"] = connected_server
        result["items"] = resolve_codes(base, codes)
        result["ok"] = True
    except Exception as exc:
        result["errors"].append(str(exc))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON: {out}")
    print(f"Items: {len(result['items'])}")
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


def resolve_codes(base: Any, codes: list[str]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for chunk in chunks(codes, 40):
        query = base.NewObject("Запрос")
        conditions: list[str] = []
        for index, code in enumerate(chunk):
            conditions.append(
                f"(Номенклатура.Код = &Exact{index} "
                f"ИЛИ Номенклатура.Артикул = &Exact{index} "
                f"ИЛИ Номенклатура.Код = &Ut{index} "
                f"ИЛИ Номенклатура.Артикул = &Ut{index} "
                f"ИЛИ Номенклатура.Код = &Numeric{index} "
                f"ИЛИ Номенклатура.Артикул = &Numeric{index})"
            )
        query.Text = f"""
        ВЫБРАТЬ
            Номенклатура.Код КАК code,
            Номенклатура.Артикул КАК article,
            Номенклатура.Наименование КАК name,
            Номенклатура.БазоваяЕдиницаИзмерения.Наименование КАК unit
        ИЗ
            Справочник.Номенклатура КАК Номенклатура
        ГДЕ
            НЕ Номенклатура.ПометкаУдаления
            И ({" ИЛИ ".join(conditions)})
        """
        for index, code in enumerate(chunk):
            query.SetParameter(f"Exact{index}", code)
            query.SetParameter(f"Ut{index}", normalize_ut_code(code))
            query.SetParameter(f"Numeric{index}", normalize_numeric_code(code))

        selection = query.Execute().Choose()
        while selection.Next():
            row = {
                "code": str(selection.code),
                "article": str(selection.article),
                "name": str(selection.name),
                "unit": str(selection.unit),
            }
            key = (row["code"], row["article"])
            if key not in seen:
                seen.add(key)
                rows.append(row)
    return rows


def normalize_ut_code(value: str) -> str:
    cleaned = value.strip().upper().replace(" ", "")
    if cleaned.startswith("УТ"):
        cleaned = cleaned[2:]
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    return "УТ" + digits.zfill(9) if digits else value.strip()


def normalize_numeric_code(value: str) -> str:
    digits = "".join(ch for ch in value.strip() if ch.isdigit())
    return digits.zfill(11) if digits else value.strip()


def chunks(values: list[str], size: int) -> list[list[str]]:
    return [values[index:index + size] for index in range(0, len(values), size)]


def unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.upper()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


if __name__ == "__main__":
    raise SystemExit(main())
