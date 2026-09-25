from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any


class ConfigError(ValueError):
    """Raised when sales configuration validation fails."""
    pass


@dataclass(frozen=True)
class SalesConfig:
    pms_sheet_name: str = "PMS"
    ago_sheet_name: str = "AGO"
    lpg_sheet_name: str = "LPG"
    pms_header_row: int = 2
    ago_header_row: int = 2
    lpg_header_row: int = 3
    sync_mode: str = "all"  # "all" (all sheets in file), "latest" (last sheet), "specific_date"
    source_sheet_index: int = -1
    target_date: date | None = None
    match_threshold: float = 0.55

    def validate(self) -> None:
        if not self.pms_sheet_name.strip():
            raise ConfigError("PMS sheet name cannot be empty.")
        if not self.ago_sheet_name.strip():
            raise ConfigError("AGO sheet name cannot be empty.")
        if not self.lpg_sheet_name.strip():
            raise ConfigError("LPG sheet name cannot be empty.")

        for row, name in [
            (self.pms_header_row, "PMS header row"),
            (self.ago_header_row, "AGO header row"),
            (self.lpg_header_row, "LPG header row"),
        ]:
            if row < 1:
                raise ConfigError(f"{name} must be >= 1 (got {row}).")

        if self.sync_mode not in {"all", "latest", "specific_date"}:
            raise ConfigError(
                f"Invalid sync_mode: '{self.sync_mode}'. Must be 'all', 'latest', or 'specific_date'."
            )

        if self.sync_mode == "specific_date" and self.target_date is None:
            raise ConfigError("target_date must be specified when sync_mode is 'specific_date'.")

        if not (0.0 <= self.match_threshold <= 1.0):
            raise ConfigError("match_threshold must be between 0.0 and 1.0.")
