from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from .com_client import OneCComClient, diagnose_environment
from .config import load_config, redacted_connection
from .errors import IntegrationError
from .logging_utils import audit, configure_logging
from .service import OneCIntegrationService


def main() -> int:
    parser = argparse.ArgumentParser(description="Безопасная интеграция с 1С:УТ 10.3")
    parser.add_argument("--config", default="onec_integration/config.example.json")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("diagnose-environment")
    _out(sub.add_parser("inspect-metadata"))
    _out(sub.add_parser("warehouses"))
    nomenclature = _out(sub.add_parser("nomenclature"))
    nomenclature.add_argument("--code", default="")
    nomenclature.add_argument("--article", default="")
    nomenclature.add_argument("--name", default="")
    stocks = _out(sub.add_parser("stocks"))
    stocks.add_argument("--positive-only", action="store_true")
    stocks.add_argument("--at-date", default="")
    assembly = _out(sub.add_parser("assembly-dry-run"))
    assembly.add_argument("--payload", required=True)
    transfer = _out(sub.add_parser("transfer-dry-run"))
    transfer.add_argument("--payload", required=True)

    args = parser.parse_args()
    config = load_config(args.config)
    logger = configure_logging(config.log_directory)

    try:
        if args.command == "diagnose-environment":
            host, _, port = config.connection.server.partition(":")
            result = diagnose_environment(host, int(port or "1541"))
            result["connection"] = redacted_connection(config.connection)
        else:
            client = OneCComClient(config.connection)
            service = OneCIntegrationService(config, client)
            if args.command == "inspect-metadata":
                result = client.metadata_report()
            elif args.command == "warehouses":
                result = service.warehouses()
            elif args.command == "nomenclature":
                result = service.nomenclature(code=args.code, article=args.article, name=args.name)
            elif args.command == "stocks":
                result = service.stocks(positive_only=args.positive_only, at_date=args.at_date or None)
            elif args.command == "assembly-dry-run":
                result = service.create_assembly(_read_json(args.payload))
            elif args.command == "transfer-dry-run":
                result = service.create_transfer(_read_json(args.payload))
            else:
                raise AssertionError(args.command)
        audit(logger, args.command, result)
        _emit(result, getattr(args, "out", ""))
        return 0
    except IntegrationError as exc:
        result = exc.to_response()
        audit(logger, args.command, result)
        _emit(result, getattr(args, "out", ""))
        return 2


def _out(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    parser.add_argument("--out", default="")
    return parser


def _read_json(value: str) -> dict[str, Any]:
    path = Path(value)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return json.loads(value)


def _emit(payload: dict[str, Any], out: str = "") -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if out:
        path = Path(out)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    raise SystemExit(main())
