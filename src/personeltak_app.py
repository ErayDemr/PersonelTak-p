"""All-in-one PersonelTak toolkit packed into a single file."""
from __future__ import annotations

import argparse
import json
import logging
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta, tzinfo
from pathlib import Path
from typing import Any, Dict, List, Mapping, MutableMapping, Optional, Sequence

import pandas as pd
import yaml
from filelock import FileLock, Timeout
from flask import Flask, redirect, render_template_string, request, send_file, url_for

DEFAULT_CONFIG: Mapping[str, Any] = {
    "role_weights": {"Personel": 0.20, "Şef": 0.40, "Yönetici": 0.40},
    "category_weights": {"İş": 1.0, "Kanaat": 0.7},
    "tespit_days": 30,
    "timezone": "Europe/Istanbul",
    "excel_path": "data/input.xlsx",
    "employees_path": "C:/ProgramData/PersonelTak/Calisanlar.xlsx",
    "report_path": "reports",
    "log_path": "logs/personeltak.log",
    "log_level": "INFO",
    "missing_threshold": None,
    "csv_export": False,
    "powerbi_export": False,
    "powerbi_output": "reports",
    "lock_timeout": 30,
}


@dataclass(frozen=True)
class AppConfig:
    """Typed configuration container."""

    role_weights: Mapping[str, float]
    category_weights: Mapping[str, float]
    tespit_days: int
    timezone: tzinfo
    excel_path: Path
    report_path: Path
    employees_path: Optional[Path]
    log_path: Path
    log_level: str
    missing_threshold: Optional[int] = None
    csv_export: bool = False
    powerbi_export: bool = False
    powerbi_output: Optional[Path] = None
    lock_timeout: float = 30.0

    extra: Mapping[str, object] = field(default_factory=dict)


def _parse_timezone(name: str) -> tzinfo:
    try:
        from zoneinfo import ZoneInfo

        return ZoneInfo(name)
    except Exception as exc:  # pragma: no cover - ZoneInfo availability
        raise ValueError(f"Unknown timezone '{name}'") from exc


def _load_config_file(path: Path) -> MutableMapping[str, object]:
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() in {".yaml", ".yml"}:
        data = yaml.safe_load(text) or {}
    elif path.suffix.lower() == ".json":
        data = json.loads(text or "{}")
    else:
        raise ValueError(f"Unsupported config format: {path.suffix}")
    if not isinstance(data, MutableMapping):
        raise ValueError("Config root must be a mapping")
    return data


def load_config(path: Optional[str | Path] = None, overrides: Optional[Mapping[str, object]] = None) -> AppConfig:
    merged: MutableMapping[str, object] = dict(DEFAULT_CONFIG)
    if path:
        config_path = Path(path)
        merged.update(_load_config_file(config_path))
    if overrides:
        merged.update(overrides)

    tz_name = str(merged.get("timezone", DEFAULT_CONFIG["timezone"]))
    tz = _parse_timezone(tz_name)

    excel_path = Path(str(merged.get("excel_path", DEFAULT_CONFIG["excel_path"]))).expanduser()
    report_path = Path(str(merged.get("report_path", DEFAULT_CONFIG["report_path"]))).expanduser()

    employees_value = merged.get("employees_path")
    employees_path = Path(str(employees_value)).expanduser() if employees_value else None

    log_path = Path(str(merged.get("log_path", DEFAULT_CONFIG["log_path"]))).expanduser()
    log_level = str(merged.get("log_level", DEFAULT_CONFIG["log_level"])).upper()

    missing_threshold_value = merged.get("missing_threshold")
    if missing_threshold_value is not None:
        try:
            missing_threshold_value = int(missing_threshold_value)
        except (TypeError, ValueError) as exc:
            raise ValueError("missing_threshold must be int or null") from exc

    csv_export = bool(merged.get("csv_export", DEFAULT_CONFIG["csv_export"]))
    powerbi_export = bool(merged.get("powerbi_export", DEFAULT_CONFIG["powerbi_export"]))
    powerbi_output_value = merged.get("powerbi_output", DEFAULT_CONFIG["powerbi_output"])
    powerbi_output = Path(str(powerbi_output_value)).expanduser() if powerbi_output_value else None

    lock_timeout_value = merged.get("lock_timeout", DEFAULT_CONFIG["lock_timeout"])
    try:
        lock_timeout = float(lock_timeout_value)
    except (TypeError, ValueError) as exc:
        raise ValueError("lock_timeout must be numeric") from exc

    return AppConfig(
        role_weights=dict(merged.get("role_weights", DEFAULT_CONFIG["role_weights"])),
        category_weights=dict(merged.get("category_weights", DEFAULT_CONFIG["category_weights"])),
        tespit_days=int(merged.get("tespit_days", DEFAULT_CONFIG["tespit_days"])),
        timezone=tz,
        excel_path=excel_path,
        report_path=report_path,
        employees_path=employees_path,
        log_path=log_path,
        log_level=log_level,
        missing_threshold=missing_threshold_value,
        csv_export=csv_export,
        powerbi_export=powerbi_export,
        powerbi_output=powerbi_output,
        lock_timeout=lock_timeout,
        extra={k: v for k, v in merged.items() if k not in DEFAULT_CONFIG},
    )


def setup_logging(config: AppConfig) -> logging.Logger:
    logger = logging.getLogger("personeltak")
    logger.setLevel(getattr(logging, config.log_level.upper(), logging.INFO))
    logger.propagate = False

    log_path = Path(config.log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    if not any(isinstance(handler, RotatingFileHandler) for handler in logger.handlers):
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=2 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    if not any(isinstance(handler, logging.StreamHandler) for handler in logger.handlers):
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logger.addHandler(stream_handler)

    logger.debug("Logging initialized at %s level", config.log_level.upper())
    return logger


class ValidationError(Exception):
    """Raised when input data fails validation rules."""


@dataclass(frozen=True)
class WorkbookData:
    criteria: pd.DataFrame
    employees: pd.DataFrame
    evaluations: pd.DataFrame


def load_data(path: str | Path, employees_path: Optional[str | Path] = None) -> WorkbookData:
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
            logging.getLogger("personeltak.loader").debug("Employees loaded from %s", override_path)
        else:
            employees = xl.parse("Calisanlar")
            logging.getLogger("personeltak.loader").debug(
                "Employees loaded from primary workbook %s", workbook_path
            )
    except ValueError as exc:
        raise ValidationError(
            "Workbook must contain Kriterler, Calisanlar, Degerlendirmeler sheets"
        ) from exc

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

    if "HaftaYili" not in evaluations:
        evaluations["HaftaYili"] = evaluations["Tarih"].apply(_iso_week)
    else:
        evaluations["HaftaYili"] = evaluations["HaftaYili"].fillna("")
        missing_week_mask = evaluations["HaftaYili"] == ""
        evaluations.loc[missing_week_mask, "HaftaYili"] = evaluations.loc[missing_week_mask, "Tarih"].apply(
            _iso_week
        )

    if (evaluations["Puan"] < 0).any():
        raise ValidationError("Negative scores are not allowed")

    return WorkbookData(criteria=criteria, employees=employees, evaluations=evaluations)


def _iso_week(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, datetime):
        iso_year, iso_week, _ = value.isocalendar()
        return f"{iso_year}-W{iso_week:02d}"
    try:
        parsed = pd.to_datetime(value)
    except Exception as exc:  # pragma: no cover - fallback path
        raise ValidationError(f"Invalid Tarih value: {value}") from exc
    if pd.isna(parsed):
        return ""
    iso_year, iso_week, _ = parsed.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


ROLES: Sequence[str] = ("Personel", "Şef", "Yönetici")


@dataclass(frozen=True)
class ScoreResult:
    scores: pd.DataFrame
    missing: pd.DataFrame
    warnings: List[str]
    report_file: Optional[Path] = None


def summarize_scores(workbook: WorkbookData, config: AppConfig, asof: datetime | None = None) -> ScoreResult:
    asof_dt = _normalize_asof(asof, config)
    criteria = workbook.criteria
    employees = workbook.employees
    evaluations = _prepare_evaluations(workbook.evaluations, config)

    role_weights = config.role_weights
    category_weights = config.category_weights

    iso_week = asof_dt.isocalendar()
    current_week = f"{iso_week[0]}-W{iso_week[1]:02d}"
    tespit_since = asof_dt - timedelta(days=config.tespit_days)

    rows: List[Dict[str, Any]] = []
    missing_rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    criteria_lookup = criteria.set_index("Po")
    allowed_role_map: Dict[int, List[str]] = {
        int(po): _allowed_roles(criterion) for po, criterion in criteria_lookup.iterrows()
    }

    valid_po = set(criteria_lookup.index.astype(int))
    invalid_po_mask = ~evaluations["Po"].isin(valid_po)
    if invalid_po_mask.any():
        invalid_pos = evaluations.loc[invalid_po_mask, "Po"].unique()
        warnings.append(
            f"Evaluations with unknown Po ignored: {sorted(map(int, invalid_pos))}"
        )
        evaluations = evaluations.loc[~invalid_po_mask]

    def _role_allowed(row: pd.Series) -> bool:
        allowed = allowed_role_map.get(int(row["Po"]), [])
        return row["Rol"] in allowed

    invalid_role_mask = ~evaluations.apply(_role_allowed, axis=1)
    if invalid_role_mask.any():
        invalid_roles = evaluations.loc[invalid_role_mask, ["Po", "Rol"]]
        for _, rec in invalid_roles.iterrows():
            warnings.append(
                f"Role {rec['Rol']} not allowed for Po={rec['Po']}; record ignored"
            )
        evaluations = evaluations.loc[~invalid_role_mask]

    for _, employee in employees.iterrows():
        sicil = employee.get("Sicil")
        if pd.isna(sicil):
            continue
        employee_scores: List[float] = []
        employee_weights: List[float] = []
        employee_missing: List[Dict[str, Any]] = []

        for po, criterion in criteria_lookup.iterrows():
            try:
                category = criterion.get("Kategori", "İş")
                period = str(criterion.get("Period", "Haftalık"))
                puan_max = float(criterion.get("PuanMax", 5) or 5)
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
                    role_filter &= evaluations["Tarih"] <= asof_dt
                else:
                    warnings.append(f"Unknown period '{period}' for Po={po}")
                    continue

                candidate = evaluations.loc[role_filter]
                if candidate.empty:
                    continue
                last_record = candidate.sort_values("Tarih").iloc[-1]
                score_value = float(last_record.get("Puan", 0))
                score_value = max(0.0, min(score_value, puan_max))
                role_scores[role] = score_value / puan_max if puan_max else 0.0

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
            missing_roles: List[str] = []
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

    if (
        config.missing_threshold is not None
        and config.missing_threshold > 0
        and not missing_df.empty
    ):
        counts = missing_df.groupby("Sicil").size().rename("EksikAdet")
        missing_df = missing_df.merge(counts, on="Sicil")
        missing_df = missing_df[missing_df["EksikAdet"] >= config.missing_threshold]

    return ScoreResult(
        scores=scores_df,
        missing=missing_df,
        warnings=warnings,
        report_file=None,
    )


def export_report(
    workbook_path: str | Path,
    output_path: str | Path,
    config: AppConfig,
    asof: datetime | None = None,
) -> ScoreResult:
    workbook_path = Path(workbook_path)
    lock = FileLock(str(workbook_path) + ".lock", timeout=config.lock_timeout)
    try:
        with lock:
            workbook = load_data(workbook_path, employees_path=config.employees_path)
    except Timeout as exc:
        raise TimeoutError(
            f"Unable to acquire lock for workbook {workbook_path} within {config.lock_timeout} seconds"
        ) from exc

    logging.getLogger("personeltak.report").info(
        "Workbook %s loaded for reporting", workbook_path
    )
    result = summarize_scores(workbook, config, asof=asof)

    report_dir = Path(output_path)
    report_dir.mkdir(parents=True, exist_ok=True)

    if not result.scores.empty and "Hafta" in result.scores:
        iso_week = result.scores["Hafta"].iloc[0]
    else:
        iso_week = _normalize_asof(asof, config).strftime("%G-W%V")
    report_file = report_dir / f"rapor_{iso_week}.xlsx"

    with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
        result.scores.to_excel(writer, sheet_name="Skorlar", index=False)
        result.missing.to_excel(writer, sheet_name="EksikPuanlamalar", index=False)

    logging.getLogger("personeltak.report").info(
        "Excel report written to %s", report_file
    )

    if config.csv_export:
        _export_csv(result, report_dir, iso_week)
    if config.powerbi_export:
        _export_powerbi(result, config, iso_week)

    return replace(result, report_file=report_file)


def _export_csv(result: ScoreResult, report_dir: Path, iso_week: str) -> None:
    scores_csv = report_dir / f"rapor_{iso_week}_Skorlar.csv"
    missing_csv = report_dir / f"rapor_{iso_week}_EksikPuanlamalar.csv"
    result.scores.to_csv(scores_csv, index=False, encoding="utf-8-sig")
    result.missing.to_csv(missing_csv, index=False, encoding="utf-8-sig")
    logging.getLogger("personeltak.report").info(
        "CSV raporları oluşturuldu: %s, %s", scores_csv, missing_csv
    )


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
    logging.getLogger("personeltak.report").info(
        "Power BI dataseti güncellendi: %s", dataset_file
    )


def _normalize_asof(asof: datetime | None, config: AppConfig) -> datetime:
    if asof is None:
        return datetime.now(tz=config.timezone)
    if asof.tzinfo is None:
        return asof.replace(tzinfo=config.timezone)
    return asof.astimezone(config.timezone)


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


def record_evaluation(workbook_path: str | Path, record: Dict[str, Any], timeout: float = 30.0) -> None:
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

    logging.getLogger("personeltak.record").info("Yeni değerlendirme eklendi: %s", new_row)


HTML_TEMPLATE = """<!doctype html>
<html lang=\"tr\">
<head>
  <meta charset=\"utf-8\" />
  <title>PersonelTak HTML Arayüzü</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 24px; background-color: #f4f6f8; color: #1d1d1d; }
    header { margin-bottom: 24px; }
    h1 { margin: 0 0 8px 0; font-size: 28px; }
    section { background: #ffffff; padding: 16px 20px; margin-bottom: 20px; border-radius: 12px;
              box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08); }
    .controls { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; }
    .controls label { font-weight: 600; }
    .controls input[type=date] { padding: 6px 10px; border-radius: 6px; border: 1px solid #c7c7c7; }
    button, .button { background-color: #0052cc; color: #fff; border: none; border-radius: 6px; padding: 8px 16px;
                      cursor: pointer; text-decoration: none; display: inline-block; }
    button:hover, .button:hover { background-color: #003f9e; }
    .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; }
    .alert-error { background: #ffe8e6; color: #a4262c; }
    .alert-success { background: #e7f5e4; color: #2f7d32; }
    .alert-warning { background: #fff4e5; color: #8a5300; }
    .empty { color: #666; font-style: italic; }
    table.table { width: 100%; border-collapse: collapse; }
    table.table th, table.table td { padding: 8px 10px; border-bottom: 1px solid #e5e5e5; text-align: left; }
    table.table tr:nth-child(even) { background-color: #fafafa; }
    form.grid { display: grid; gap: 12px; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); }
    form.grid label { display: flex; flex-direction: column; font-weight: 600; font-size: 14px; }
    form.grid input, form.grid select, form.grid textarea { margin-top: 4px; padding: 8px 10px; border: 1px solid #c7c7c7;
                                                           border-radius: 6px; font-size: 14px; }
    textarea { resize: vertical; }
    footer { text-align: right; font-size: 12px; color: #666; margin-top: 12px; }
    .warning-list ul { margin: 8px 0 0 20px; }
  </style>
</head>
<body>
  <header>
    <h1>PersonelTak HTML Arayüzü</h1>
    <p>Haftalık ve tespit bazlı skorları görüntüleyip yeni değerlendirme ekleyin.</p>
  </header>

  <section>
    <form method=\"get\" class=\"controls\">
      <label>Rapor Tarihi
        <input type=\"date\" name=\"asof\" value=\"{{ asof_text }}\" required />
      </label>
      <button type=\"submit\">Raporu Yenile</button>
      <a class=\"button\" href=\"{{ download_url }}\">Excel İndir</a>
    </form>
  </section>

  {% if summary_error %}
  <div class=\"alert alert-error\">{{ summary_error }}</div>
  {% endif %}
  {% if form_error %}
  <div class=\"alert alert-error\">{{ form_error }}</div>
  {% endif %}
  {% if message %}
  <div class=\"alert alert-success\">{{ message }}</div>
  {% endif %}
  {% if warnings %}
  <div class=\"alert alert-warning warning-list\">
    <strong>Uyarılar:</strong>
    <ul>
    {% for item in warnings %}
      <li>{{ item }}</li>
    {% endfor %}
    </ul>
  </div>
  {% endif %}

  <section>
    <h2>Çalışan Skorları</h2>
    {{ scores_html|safe }}
  </section>

  <section>
    <h2>Eksik Puanlamalar</h2>
    {{ missing_html|safe }}
  </section>

  <section>
    <h2>Yeni Değerlendirme Kaydı</h2>
    <form method=\"post\" class=\"grid\">
      <input type=\"hidden\" name=\"form_id\" value=\"record\" />
      <label>Sicil
        <input name=\"sicil\" required />
      </label>
      <label>Rol
        <select name=\"rol\" required>
          <option value=\"\">Rol Seçin</option>
          {% for role in roles %}
          <option value=\"{{ role }}\">{{ role }}</option>
          {% endfor %}
        </select>
      </label>
      <label>Kriter (Po)
        <input name=\"po\" type=\"number\" min=\"1\" required />
      </label>
      <label>Puan
        <input name=\"puan\" type=\"number\" step=\"0.01\" min=\"0\" required />
      </label>
      <label>Tarih
        <input name=\"tarih\" type=\"datetime-local\" />
      </label>
      <label>Not
        <textarea name=\"note\" rows=\"2\"></textarea>
      </label>
      <div>
        <button type=\"submit\">Kaydı Ekle</button>
      </div>
    </form>
  </section>

  <footer>
    <span>Son güncelleme: {{ generated_at }}</span>
  </footer>
</body>
</html>
"""


def _dataframe_to_html(df: pd.DataFrame, empty_message: str) -> str:
    if df.empty:
        return f"<p class='empty'>{empty_message}</p>"
    return df.fillna("").to_html(classes="table", index=False, border=0, justify="left")


def _parse_input_date(value: str, config: AppConfig) -> datetime:
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as exc:  # pragma: no cover - user input handling
        raise ValueError("Geçersiz tarih formatı. YYYY-AA-GG olarak girin.") from exc
    return _normalize_asof(parsed, config)


def create_web_app(config: AppConfig, workbook_path: Path) -> Flask:
    data_path = Path(workbook_path).expanduser()
    app = Flask(__name__)

    @app.route("/", methods=["GET", "POST"])
    def index() -> str:
        message = request.args.get("message")
        form_error: Optional[str] = None
        summary_error: Optional[str] = None

        asof_text = request.values.get("asof") or datetime.now(tz=config.timezone).date().isoformat()
        asof_dt: Optional[datetime] = None
        try:
            asof_dt = _parse_input_date(asof_text, config)
        except ValueError as exc:
            summary_error = str(exc)

        if request.method == "POST" and request.form.get("form_id") == "record":
            sicil = request.form.get("sicil", "").strip()
            rol = request.form.get("rol", "").strip()
            po_text = request.form.get("po", "").strip()
            puan_text = request.form.get("puan", "").strip()
            tarih_text = request.form.get("tarih", "").strip()
            note = request.form.get("note", "").strip()

            errors: List[str] = []
            if not sicil:
                errors.append("Sicil zorunludur.")
            if not rol:
                errors.append("Rol zorunludur.")

            po_value: Optional[int] = None
            try:
                po_value = int(po_text)
                if po_value <= 0:
                    errors.append("Po 1 veya daha büyük olmalıdır.")
            except ValueError:
                errors.append("Po numerik olmalıdır.")

            puan_value: Optional[float] = None
            try:
                puan_value = float(puan_text)
                if puan_value < 0:
                    errors.append("Puan negatif olamaz.")
            except ValueError:
                errors.append("Puan numerik olmalıdır.")

            tarih_value: Optional[datetime]
            if tarih_text:
                try:
                    tarih_value = datetime.fromisoformat(tarih_text)
                except ValueError:
                    errors.append(
                        "Tarih alanı ISO formatında olmalıdır (YYYY-AA-GG veya YYYY-AA-GGTHH:MM)."
                    )
                    tarih_value = None
            else:
                tarih_value = datetime.now(tz=config.timezone)

            if errors:
                form_error = " ".join(errors)
            else:
                if po_value is None or puan_value is None:
                    form_error = "Form verileri işlenemedi."
                else:
                    if tarih_value is not None:
                        if tarih_value.tzinfo is None:
                            tarih_value = tarih_value.replace(tzinfo=config.timezone)
                        else:
                            tarih_value = tarih_value.astimezone(config.timezone)

                    record = {
                        "Sicil": sicil,
                        "Rol": rol,
                        "Po": po_value,
                        "Puan": puan_value,
                        "Not": note or None,
                        "Tarih": tarih_value,
                    }

                    try:
                        record_evaluation(data_path, record, timeout=config.lock_timeout)
                    except Exception as exc:  # pragma: no cover - runtime safeguard
                        form_error = f"Kayıt eklenemedi: {exc}"
                    else:
                        return redirect(
                            url_for(
                                "index",
                                asof=asof_text,
                                message="Puan kaydı başarıyla eklendi.",
                            )
                        )

        warnings: List[str] = []
        scores_html = _dataframe_to_html(pd.DataFrame(), "Henüz skor bulunmuyor.")
        missing_html = _dataframe_to_html(pd.DataFrame(), "Eksik puanlama yok.")

        if summary_error is None:
            try:
                lock = FileLock(str(data_path) + ".lock", timeout=config.lock_timeout)
                with lock:
                    workbook = load_data(data_path, employees_path=config.employees_path)
                summary = summarize_scores(workbook, config, asof=asof_dt)
            except Timeout as exc:
                summary_error = f"Dosya kilidi alınamadı: {exc}"
            except Exception as exc:  # pragma: no cover - runtime safeguards
                summary_error = f"Rapor oluşturulamadı: {exc}"
            else:
                warnings = summary.warnings
                scores_html = _dataframe_to_html(summary.scores, "Henüz skor bulunmuyor.")
                missing_html = _dataframe_to_html(summary.missing, "Eksik puanlama yok.")

        download_url = url_for("download", asof=asof_text)
        generated_at = datetime.now(tz=config.timezone).strftime("%d.%m.%Y %H:%M")

        return render_template_string(
            HTML_TEMPLATE,
            asof_text=asof_text,
            scores_html=scores_html,
            missing_html=missing_html,
            message=message,
            form_error=form_error,
            summary_error=summary_error,
            warnings=warnings,
            roles=ROLES,
            generated_at=generated_at,
            download_url=download_url,
        )

    @app.route("/download")
    def download():
        asof_text = request.args.get("asof")
        asof_dt: Optional[datetime] = None
        if asof_text:
            try:
                asof_dt = _parse_input_date(asof_text, config)
            except ValueError as exc:
                return str(exc), 400

        try:
            result = export_report(data_path, config.report_path, config, asof=asof_dt)
        except Exception as exc:  # pragma: no cover - runtime safeguard
            return (f"Rapor oluşturulamadı: {exc}", 500)

        report_path = result.report_file
        if report_path is None:
            return ("Rapor dosyası bulunamadı", 500)

        report_path = Path(report_path)
        if not report_path.exists():
            return ("Rapor dosyası henüz oluşturulmadı", 500)

        return send_file(
            report_path,
            as_attachment=True,
            download_name=report_path.name,
            max_age=0,
        )

    return app


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PersonelTak report toolkit (single file)")
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

    web = subparsers.add_parser("web", help="Start the HTML dashboard")
    web.add_argument("--host", default="127.0.0.1", help="Host/IP for the web server")
    web.add_argument("--port", type=int, default=5000, help="Port for the web server")
    web.add_argument("--debug", action="store_true", help="Enable Flask debug mode")

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    config = load_config(args.config_path) if args.config_path else load_config()
    logger = setup_logging(config)

    workbook_path = Path(args.excel).expanduser() if args.excel else config.excel_path

    asof_dt = None
    if args.asof:
        asof_dt = datetime.fromisoformat(args.asof)

    if args.command == "web":
        app = create_web_app(config, workbook_path)
        logger.info(
            "HTML arayüzü başlatılıyor: http://%s:%s", args.host, args.port
        )
        app.run(host=args.host, port=args.port, debug=args.debug)
        return 0

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

    output_path = args.output_path or config.report_path if hasattr(args, "output_path") else config.report_path
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
