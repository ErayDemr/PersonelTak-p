"""Excel loading and validation helpers."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd


logger = logging.getLogger("personeltak.loader")


class ValidationError(Exception):
    """Raised when input data fails validation rules."""


@dataclass(frozen=True)
class WorkbookData:
    """Container for source workbook data."""

    criteria: pd.DataFrame
    employees: pd.DataFrame
    evaluations: pd.DataFrame


def load_data(path: str | Path, employees_path: Optional[str | Path] = None) -> WorkbookData:
    """Load workbook sheets and apply base validation."""

    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    xl = pd.ExcelFile(workbook_path)
    try:
        criteria = xl.parse("Kriterler")
        evaluations = xl.parse("Degerlendirmeler")
        if employees_path:
            override_path = Path(employees_path)
            if not override_path.exists():
                raise FileNotFoundError(f"Employees workbook not found: {override_path}")
            try:
                employees = pd.read_excel(override_path, sheet_name="Calisanlar")
            except ValueError as exc:
                raise ValidationError("Employees workbook must contain Calisanlar sheet") from exc
            logger.debug("Employees loaded from override path %s", override_path)
        else:
            employees = xl.parse("Calisanlar")
            logger.debug("Employees loaded from primary workbook %s", workbook_path)
    except ValueError as exc:
        raise ValidationError("Workbook must contain Kriterler, Calisanlar, Degerlendirmeler sheets") from exc

    criteria = criteria.copy()
    employees = employees.copy()
    evaluations = evaluations.copy()

    if "Po" not in criteria:
        raise ValidationError("Kriterler sheet must contain 'Po' column")
    if criteria["Po"].duplicated().any():
        duplicates = criteria["Po"][criteria["Po"].duplicated()].unique()
        raise ValidationError(f"Duplicate Po values in Kriterler: {duplicates}")

    if (criteria.get("PuanMax", 5).fillna(5) <= 0).any():
        raise ValidationError("PuanMax must be > 0 for all criteria")

    # Normalize column names for evaluations
    if "HaftaYili" not in evaluations:
        evaluations["HaftaYili"] = evaluations["Tarih"].apply(_iso_week)
    else:
        evaluations["HaftaYili"] = evaluations["HaftaYili"].fillna("")
        missing_week_mask = evaluations["HaftaYili"] == ""
        evaluations.loc[missing_week_mask, "HaftaYili"] = evaluations.loc[missing_week_mask, "Tarih"].apply(_iso_week)

    if (evaluations["Puan"] < 0).any():
        raise ValidationError("Negative scores are not allowed")

    return WorkbookData(criteria=criteria, employees=employees, evaluations=evaluations)


def _iso_week(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, datetime):
        iso_year, iso_week, _ = value.isocalendar()
        return f"{iso_year}-W{iso_week:02d}"
    try:
        value = pd.to_datetime(value)
    except Exception as exc:  # pragma: no cover - fallback path
        raise ValidationError(f"Invalid Tarih value: {value}") from exc
    iso_year, iso_week, _ = value.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


__all__ = ["ValidationError", "WorkbookData", "load_data"]
