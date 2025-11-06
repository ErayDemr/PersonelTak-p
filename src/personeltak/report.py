"""Score calculation and report generation."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Sequence

import pandas as pd
from filelock import FileLock, Timeout

from .config import AppConfig
from .loader import ValidationError, WorkbookData, load_data

ROLES: Sequence[str] = ("Personel", "Şef", "Yönetici")

logger = logging.getLogger("personeltak.report")


@dataclass(frozen=True)
class ScoreResult:
    scores: pd.DataFrame
    missing: pd.DataFrame
    warnings: List[str]


def summarize_scores(
    workbook: WorkbookData,
    config: AppConfig,
    asof: datetime | None = None,
) -> ScoreResult:
    """Calculate scores and missing evaluations according to the business rules."""

    asof = _normalize_asof(asof, config)
    criteria = workbook.criteria
    employees = workbook.employees
    evaluations = _prepare_evaluations(workbook.evaluations, config)

    role_weights = config.role_weights
    category_weights = config.category_weights

    iso_week = asof.isocalendar()
    current_week = f"{iso_week[0]}-W{iso_week[1]:02d}"
    tespit_since = asof - timedelta(days=config.tespit_days)

    rows: List[Dict[str, object]] = []
    missing_rows: List[Dict[str, object]] = []
    warnings: List[str] = []

    criteria_lookup = criteria.set_index("Po")
    allowed_role_map: Dict[int, List[str]] = {
        int(po): _allowed_roles(criterion) for po, criterion in criteria_lookup.iterrows()
    }

    valid_po = set(criteria_lookup.index.astype(int))
    invalid_po_mask = ~evaluations["Po"].isin(valid_po)
    if invalid_po_mask.any():
        invalid_pos = evaluations.loc[invalid_po_mask, "Po"].unique()
        warnings.append(f"Evaluations with unknown Po ignored: {sorted(map(int, invalid_pos))}")
        evaluations = evaluations.loc[~invalid_po_mask]

    def _role_allowed(row: pd.Series) -> bool:
        allowed = allowed_role_map.get(int(row["Po"]), [])
        return row["Rol"] in allowed

    invalid_role_mask = ~evaluations.apply(_role_allowed, axis=1)
    if invalid_role_mask.any():
        invalid_roles = evaluations.loc[invalid_role_mask, ["Po", "Rol"]]
        for _, rec in invalid_roles.iterrows():
            warnings.append(f"Role {rec['Rol']} not allowed for Po={rec['Po']}; record ignored")
        evaluations = evaluations.loc[~invalid_role_mask]

    for _, employee in employees.iterrows():
        sicil = employee.get("Sicil")
        if pd.isna(sicil):
            continue
        employee_scores: List[float] = []
        employee_weights: List[float] = []
        employee_missing: List[Dict[str, object]] = []

        for po, criterion in criteria_lookup.iterrows():
            try:
                category = criterion.get("Kategori", "İş")
                period = str(criterion.get("Period", "Haftalık"))
                puan_max = criterion.get("PuanMax", 5) or 5
                puan_max = float(puan_max)
            except Exception as exc:
                raise ValidationError(f"Invalid criterion definition for Po={po}") from exc

            allowed_roles = allowed_role_map.get(int(po), [])
            if not allowed_roles:
                continue

            role_scores: Dict[str, float] = {}
            for role in allowed_roles:
                role_filter = (
                    (evaluations["Sicil"] == sicil)
                    & (evaluations["Po"] == po)
                    & (evaluations["Rol"] == role)
                )
                if period.lower() == "haftalık":
                    role_filter &= evaluations["HaftaYili"] == current_week
                elif period.lower() == "tespit":
                    role_filter &= evaluations["Tarih"] >= tespit_since
                    role_filter &= evaluations["Tarih"] <= asof
                else:
                    warnings.append(f"Unknown period '{period}' for Po={po}")
                    continue

                candidate = evaluations.loc[role_filter]
                if candidate.empty:
                    continue
                last_record = candidate.sort_values("Tarih").iloc[-1]
                score_value = float(last_record.get("Puan", 0))
                score_value = max(0.0, min(score_value, puan_max))
                role_scores[role] = score_value / puan_max if puan_max else 0

            if not role_scores:
                employee_missing.append(
                    {
                        "Sicil": sicil,
                        "AdSoyad": employee.get("AdSoyad"),
                        "Po": po,
                        "Değerlendirme": criterion.get("Değerlendirme"),
                        "Period": period,
                        "Eksik_Roller": ", ".join(sorted(allowed_roles)),
                    }
                )
                continue

            numerator = 0.0
            denominator = 0.0
            missing_roles = []
            for role in allowed_roles:
                weight = float(role_weights.get(role, 0))
                if role in role_scores:
                    numerator += weight * role_scores[role]
                    denominator += weight
                else:
                    missing_roles.append(role)

            if denominator == 0:
                employee_missing.append(
                    {
                        "Sicil": sicil,
                        "AdSoyad": employee.get("AdSoyad"),
                        "Po": po,
                        "Değerlendirme": criterion.get("Değerlendirme"),
                        "Period": period,
                        "Eksik_Roller": ", ".join(sorted(allowed_roles)),
                    }
                )
                continue

            category_weight = float(category_weights.get(category, 1.0))
            employee_scores.append((numerator / denominator) * category_weight)
            employee_weights.append(category_weight)

            if missing_roles:
                employee_missing.append(
                    {
                        "Sicil": sicil,
                        "AdSoyad": employee.get("AdSoyad"),
                        "Po": po,
                        "Değerlendirme": criterion.get("Değerlendirme"),
                        "Period": period,
                        "Eksik_Roller": ", ".join(sorted(missing_roles)),
                    }
                )

        if employee_scores and employee_weights:
            total_score = 100 * sum(employee_scores) / sum(employee_weights)
        else:
            total_score = 0.0

        rows.append(
            {
                "Sicil": sicil,
                "AdSoyad": employee.get("AdSoyad"),
                "Departman": employee.get("Departman"),
                "Unvan": employee.get("Unvan"),
                "ToplamSkor": round(total_score, 2),
                "Hafta": current_week,
            }
        )

        missing_rows.extend(employee_missing)

    scores_df = pd.DataFrame(rows)
    missing_df = pd.DataFrame(missing_rows)

    if config.missing_threshold is not None and config.missing_threshold > 0 and not missing_df.empty:
        counts = missing_df.groupby("Sicil").size().rename("EksikAdet")
        missing_df = missing_df.merge(counts, on="Sicil")
        missing_df = missing_df[missing_df["EksikAdet"] >= config.missing_threshold]

    return ScoreResult(scores=scores_df, missing=missing_df, warnings=warnings)


def export_report(
    workbook_path: str | Path,
    output_path: str | Path,
    config: AppConfig,
    asof: datetime | None = None,
) -> ScoreResult:
    """High-level helper that loads data and exports the report to Excel."""

    workbook_path = Path(workbook_path)
    lock = FileLock(str(workbook_path) + ".lock", timeout=config.lock_timeout)
    try:
        with lock:
            workbook = load_data(workbook_path, employees_path=config.employees_path)
    except Timeout as exc:
        raise TimeoutError(
            f"Unable to acquire lock for workbook {workbook_path} within {config.lock_timeout} seconds"
        ) from exc

    logger.info("Workbook %s loaded for reporting", workbook_path)
    result = summarize_scores(workbook, config, asof=asof)

    report_dir = Path(output_path)
    report_dir.mkdir(parents=True, exist_ok=True)

    iso_week = result.scores["Hafta"].iloc[0] if not result.scores.empty else _normalize_asof(asof, config).strftime("%G-W%V")
    report_file = report_dir / f"rapor_{iso_week}.xlsx"

    with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
        result.scores.to_excel(writer, sheet_name="Skorlar", index=False)
        result.missing.to_excel(writer, sheet_name="EksikPuanlamalar", index=False)

    logger.info("Excel report written to %s", report_file)

    if config.csv_export:
        _export_csv(result, report_dir, iso_week)
    if config.powerbi_export:
        _export_powerbi(result, config, iso_week)

    return result


def _export_csv(result: ScoreResult, report_dir: Path, iso_week: str) -> None:
    scores_csv = report_dir / f"rapor_{iso_week}_Skorlar.csv"
    missing_csv = report_dir / f"rapor_{iso_week}_EksikPuanlamalar.csv"
    result.scores.to_csv(scores_csv, index=False, encoding="utf-8-sig")
    result.missing.to_csv(missing_csv, index=False, encoding="utf-8-sig")
    logger.info("CSV raporları oluşturuldu: %s, %s", scores_csv, missing_csv)


def _export_powerbi(result: ScoreResult, config: AppConfig, iso_week: str) -> None:
    output_dir = config.powerbi_output or config.report_path
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    dataset_file = output_dir / f"personeltak_powerbi_{iso_week}.csv"
    dataset = result.scores.copy()
    if not result.missing.empty:
        counts = result.missing.groupby("Sicil").size()
        dataset["EksikSayisi"] = dataset["Sicil"].map(counts).fillna(0).astype(int)
    else:
        dataset["EksikSayisi"] = 0
    dataset.to_csv(dataset_file, index=False, encoding="utf-8-sig")
    logger.info("Power BI dataseti güncellendi: %s", dataset_file)


def _normalize_asof(asof: datetime | None, config: AppConfig) -> datetime:
    if asof is None:
        asof = datetime.now(tz=config.timezone)
    elif asof.tzinfo is None:
        asof = asof.replace(tzinfo=config.timezone)
    else:
        asof = asof.astimezone(config.timezone)
    return asof


def _prepare_evaluations(evaluations: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    df = evaluations.copy()
    df["Tarih"] = pd.to_datetime(df["Tarih"], errors="coerce")
    df = df.dropna(subset=["Tarih", "Sicil", "Po", "Rol"])
    if df["Tarih"].dt.tz is None:
        df["Tarih"] = df["Tarih"].dt.tz_localize(
            config.timezone, nonexistent="shift_forward", ambiguous="NaT"
        )
    else:
        df["Tarih"] = df["Tarih"].dt.tz_convert(config.timezone)
    df = df.dropna(subset=["Tarih"])
    df["Po"] = df["Po"].astype(int)
    df["Rol"] = df["Rol"].astype(str)
    df["HaftaYili"] = df["HaftaYili"].astype(str)
    df["Puan"] = pd.to_numeric(df["Puan"], errors="coerce")
    df = df.dropna(subset=["Puan"])
    return df


def _allowed_roles(criterion: pd.Series) -> List[str]:
    allowed: List[str] = []
    for role in ROLES:
        value = criterion.get(role)
        if isinstance(value, str) and value.strip().lower() == "x":
            allowed.append(role)
    return allowed


__all__ = ["ScoreResult", "summarize_scores", "export_report"]
