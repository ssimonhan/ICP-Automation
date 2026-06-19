from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CalibrationRange:
    """Thresholds for selecting and highlighting ICP ppb values."""

    orange_min: float = 1.0
    green_min: float = 10.0
    green_max: float = 400.0

    def __post_init__(self) -> None:
        if self.orange_min < 0:
            raise ValueError("Orange-range minimum must be 0 or greater.")
        if self.green_min <= self.orange_min:
            raise ValueError("Green-range minimum must be greater than orange-range minimum.")
        if self.green_max < self.green_min:
            raise ValueError("Green-range maximum must be greater than or equal to green-range minimum.")

    @property
    def label(self) -> str:
        return f"{format_number(self.green_min)}-{format_number(self.green_max)}"


DEFAULT_CALIBRATION_RANGE = CalibrationRange()


@dataclass(frozen=True)
class HighlightThresholds:
    """Thresholds for internal-standard percent error and blank ppb checks."""

    internal_light_orange: float = 20.0
    internal_orange: float = 40.0
    blank_light_orange: float = 0.2
    blank_orange: float = 1.0

    def __post_init__(self) -> None:
        if self.internal_light_orange < 0:
            raise ValueError("Internal-standard tolerance must be 0 or greater.")
        if self.internal_orange <= self.internal_light_orange:
            raise ValueError("Internal-standard orange threshold must be greater than light-orange threshold.")
        if self.blank_light_orange < 0:
            raise ValueError("Blank-sample tolerance must be 0 or greater.")
        if self.blank_orange <= self.blank_light_orange:
            raise ValueError("Blank-sample orange threshold must be greater than light-orange threshold.")


DEFAULT_HIGHLIGHT_THRESHOLDS = HighlightThresholds()


def format_number(value: float) -> str:
    return f"{value:g}"
