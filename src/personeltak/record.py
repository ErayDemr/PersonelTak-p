"""Utilities for appending new evaluations to the workbook."""
from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from filelock import FileLock, Timeout

from .loader import _iso_week

logger = logging.getLogger("personeltak.record")


def record_evaluation(workbook_path: str | Path, record: Dict[str, Any], timeout: float = 30.0) -> None:
    """Append an evaluation to the Degerlendirmeler sheet.

    The function preserves other sheets and writes back the workbook.
    The new record must include at least the following keys:
    ``Tarih`` (datetime/date), ``Sicil``, ``Rol``, ``Po``, ``Puan``.
    ``HaftaYili`` will be generated when missing.
    """

    path = Path(workbook_path)
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    lock = FileLock(str(path) + ".lock", timeout=timeout)
    try:
        with lock:
            excel = pd.ExcelFile(path)
            sheets = {name: excel.parse(name) for name in excel.sheet_names}
            evaluations = sheets.get("Degerlendirmeler")
            if evaluations is None:
                raise ValueError("Workbook missing 'Degerlendirmeler' sheet")

            new_row = dict(record)
            if "HaftaYili" not in new_row or not new_row.get("HaftaYili"):
                new_row["HaftaYili"] = _iso_week(new_row.get("Tarih"))
            if "Tarih" in new_row and not isinstance(new_row["Tarih"], datetime):
                new_row["Tarih"] = pd.to_datetime(new_row["Tarih"]).to_pydatetime()

            updated = pd.concat([evaluations, pd.DataFrame([new_row])], ignore_index=True)
            sheets["Degerlendirmeler"] = updated

            with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
                for sheet_name, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Timeout as exc:
        raise TimeoutError(
            f"Unable to acquire lock for workbook {path} within {timeout} seconds"
        ) from exc

    logger.info("Yeni değerlendirme eklendi: %s", new_row)


__all__ = ["record_evaluation"]
