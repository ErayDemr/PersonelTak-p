"""Command line interface for the PersonelTak toolkit."""
from __future__ import annotations

import argparse
from datetime import datetime
from typing import Any, Dict

from .config import load_config
from .logging_utils import setup_logging
from .record import record_evaluation
from .report import ScoreResult, export_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PersonelTak report toolkit")
    parser.add_argument(
        "excel",
        nargs="?",
        help="Path to the source Excel workbook (defaults to config excel_path)",
    )
    parser.add_argument(
        "--config",
        dest="config_path",
        help="Path to config file (YAML or JSON)",
    )
    parser.add_argument(
        "--asof",
        help="Reference date in YYYY-MM-DD format (defaults to now)",
    )

    subparsers = parser.add_subparsers(dest="command")

    summarize = subparsers.add_parser("summarize", help="Generate weekly report")
    summarize.add_argument(
        "--output",
        dest="output_path",
        default=None,
        help="Directory to write the Excel report. Defaults to config report_path.",
    )

    record = subparsers.add_parser("record", help="Append an evaluation row")
    record.add_argument("--sicil", required=True)
    record.add_argument("--rol", required=True)
    record.add_argument("--po", required=True, type=int)
    record.add_argument("--puan", required=True, type=float)
    record.add_argument("--not", dest="note")
    record.add_argument(
        "--tarih",
        help="Date of the evaluation (YYYY-MM-DD). Defaults to today.",
    )

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    config = load_config(args.config_path) if args.config_path else load_config()
    logger = setup_logging(config)

    workbook_path = args.excel or config.excel_path

    asof_dt = None
    if args.asof:
        asof_dt = datetime.fromisoformat(args.asof)

    if args.command == "record":
        record_data: Dict[str, Any] = {
            "Sicil": args.sicil,
            "Rol": args.rol,
            "Po": args.po,
            "Puan": args.puan,
            "Not": args.note,
        }
        if args.tarih:
            record_data["Tarih"] = datetime.fromisoformat(args.tarih)
        else:
            record_data["Tarih"] = datetime.now(tz=config.timezone)
        record_evaluation(workbook_path, record_data, timeout=config.lock_timeout)
        logger.info("Evaluation recorded for sicil=%s, po=%s", args.sicil, args.po)
        return 0

    # default or summarize
    result: ScoreResult
    output_path = args.output_path or config.report_path
    result = export_report(workbook_path, output_path, config, asof=asof_dt)

    logger.info("Skor tablosu %d kayıt içeriyor", len(result.scores))
    if not result.missing.empty:
        logger.warning("Eksik puanlamalar bulundu: %d kayıt", len(result.missing))
    for msg in result.warnings:
        logger.warning(msg)

    print(result.scores)
    if not result.missing.empty:
        print("Eksik puanlamalar:")
        print(result.missing)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
