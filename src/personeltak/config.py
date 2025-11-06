"""Configuration loading utilities for PersonelTak."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import timezone
from pathlib import Path
from typing import Mapping, MutableMapping, Optional

import yaml

DEFAULT_CONFIG = {
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
    """Typed configuration container with default fallbacks."""

    role_weights: Mapping[str, float]
    category_weights: Mapping[str, float]
    tespit_days: int
    timezone: timezone
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


def _parse_timezone(name: str) -> timezone:
    try:
        from zoneinfo import ZoneInfo

        return ZoneInfo(name)
    except Exception as exc:  # pragma: no cover - ZoneInfo may be missing
        raise ValueError(f"Unknown timezone '{name}'") from exc


def _load_file(path: Path) -> MutableMapping[str, object]:
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
    """Load application configuration merging defaults and overrides."""

    merged: MutableMapping[str, object] = dict(DEFAULT_CONFIG)
    config_path = None
    if path:
        config_path = Path(path)
        merged.update(_load_file(config_path))
    if overrides:
        merged.update(overrides)

    tz_name = str(merged.get("timezone", DEFAULT_CONFIG["timezone"]))
    tz = _parse_timezone(tz_name)

    excel_path = Path(str(merged.get("excel_path", DEFAULT_CONFIG["excel_path"]))).expanduser()
    report_path = Path(str(merged.get("report_path", DEFAULT_CONFIG["report_path"]))).expanduser()
    employees_value = merged.get("employees_path")
    employees_path = None
    if employees_value:
        employees_path = Path(str(employees_value)).expanduser()

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

    config = AppConfig(
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
    return config


__all__ = ["AppConfig", "DEFAULT_CONFIG", "load_config"]
