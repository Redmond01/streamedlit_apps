from __future__ import annotations

from dataclasses import dataclass
from typing import Any


class ConfigError(ValueError):
    """Raised when configuration validation fails."""
    pass


@dataclass(frozen=True)
class GearConfig:
    gear_sheet_name: str = "Sheet1"
    header_row_for_tanks: int = 2
    sales_meter_row: int = 130
    sales_dipping_row: int = 131
    actual_gear_row: int = 132
    expected_gear_row: int = 133
    gear_percent_row: int = 134
    average_gear_row: int = 135
    previous_average_gear_row: int = 128
    protected_last_columns_count: int = 2
    source_sheet_index: int = -1

    def validate(self) -> None:
        rows = [
            self.sales_meter_row,
            self.sales_dipping_row,
            self.actual_gear_row,
            self.expected_gear_row,
            self.gear_percent_row,
            self.average_gear_row,
            self.previous_average_gear_row,
        ]
        if len(rows) != len(set(rows)):
            raise ConfigError("Destination row numbers must all be unique.")
        for row in rows:
            if row < 1:
                raise ConfigError(f"Row numbers must be >= 1 (got {row}).")
        if self.header_row_for_tanks < 1:
            raise ConfigError(f"Header row for tanks must be >= 1 (got {self.header_row_for_tanks}).")
        if self.protected_last_columns_count < 0:
            raise ConfigError("Protected columns count cannot be negative.")
