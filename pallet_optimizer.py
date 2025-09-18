"""Core logic for pallet layout optimization.

This module contains data structures and helper functions that encapsulate the
calculation rules required by the Streamlit application.  Keeping the logic
separate from the UI makes it easier to unit test the behaviour independently
from the web front-end.
"""

from __future__ import annotations

from dataclasses import dataclass
import itertools
import math
import re
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


# Type aliases for readability.
MM = int
Orientation = Tuple[MM, MM, MM]


@dataclass
class OrientationSummary:
    """The result of placing a box on a pallet using a specific orientation."""

    orientation: Orientation
    per_layer: int
    layers: int
    total: int
    grid: Tuple[int, int]

    @property
    def as_tuple(self) -> Tuple[int, int, int, int, int, int]:
        """Return a tuple representation that is convenient for caching."""

        l, w, h = self.orientation
        gx, gy = self.grid
        return l, w, h, gx, gy, self.total


@dataclass
class BoxMetrics:
    """Aggregated information about the best orientations for a box."""

    dims_mm: Orientation
    sorted_dims_mm: Orientation
    best_by_height: Dict[int, Optional[OrientationSummary]]
    error: Optional[str] = None


DIMENSION_ALIASES: Dict[str, Tuple[str, ...]] = {
    "length": ("len", "length", "длин", "глуб", "depth", "толщ"),
    "width": ("wid", "width", "шир"),
    "height": ("hei", "height", "выс"),
}


def detect_dimension_columns(columns: Iterable[str]) -> Dict[str, Optional[str]]:
    """Attempt to detect the columns that describe the three dimensions.

    Parameters
    ----------
    columns:
        Names of the columns in the uploaded spreadsheet.

    Returns
    -------
    dict
        Mapping from logical dimension name (``length``, ``width`` and
        ``height``) to the detected column name.  Columns that cannot be
        detected are returned as ``None``.
    """

    detected: Dict[str, Optional[str]] = {"length": None, "width": None, "height": None}
    lowercase_columns = {col: col.lower() for col in columns}
    taken: set[str] = set()

    for logical, aliases in DIMENSION_ALIASES.items():
        for column, lower_column in lowercase_columns.items():
            if column in taken:
                continue
            if any(alias in lower_column for alias in aliases):
                detected[logical] = column
                taken.add(column)
                break

    # If not all dimensions have been detected, pick remaining columns in the
    # order they were provided to avoid leaving values unset.
    for column in columns:
        if column in taken:
            continue
        for logical in ("length", "width", "height"):
            if detected[logical] is None:
                detected[logical] = column
                taken.add(column)
                break

    return detected


def parse_dimension_value(value: object) -> Optional[float]:
    """Parse a single dimension value expressed in centimetres.

    The helper accepts both numeric values and strings that may use commas as a
    decimal separator.  Non-numeric strings (e.g. containing units) are
    filtered out by keeping digits, signs and decimal separators.

    Parameters
    ----------
    value:
        The raw cell value extracted from the spreadsheet.

    Returns
    -------
    float or None
        The parsed value in centimetres. ``None`` is returned when the value
        cannot be interpreted as a number.
    """

    if value is None:
        return None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            return None

        # Replace commas used as decimal separators and remove any characters
        # that are not part of a floating-point number.
        cleaned = cleaned.replace(",", ".")
        cleaned = re.sub(r"[^0-9+\-.]", "", cleaned)
        if cleaned in {"", "+", "-", "."}:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    return None


def cm_to_mm(value_cm: float) -> MM:
    """Convert a measurement from centimetres to millimetres."""

    return int(round(value_cm * 10))


def unique_permutations(values: Sequence[MM]) -> Iterable[Orientation]:
    """Generate unique permutations of three values.

    ``itertools.permutations`` returns duplicate permutations when some
    dimensions are equal.  Converting to a set and sorting gives deterministic
    ordering and avoids redundant work.
    """

    return sorted({tuple(perm) for perm in itertools.permutations(values, 3)})


def find_best_orientation(
    dims_mm: Orientation,
    pallet_length: MM,
    pallet_width: MM,
    height_limit: MM,
    overhang: MM,
) -> Optional[OrientationSummary]:
    """Return the orientation that yields the largest number of boxes."""

    best: Optional[OrientationSummary] = None

    max_length = pallet_length + overhang
    max_width = pallet_width + overhang

    for orientation in unique_permutations(dims_mm):
        length, width, height = orientation

        if height > height_limit:
            continue

        fit_length = math.floor(max_length / length) if length else 0
        fit_width = math.floor(max_width / width) if width else 0

        if fit_length <= 0 or fit_width <= 0:
            continue

        layers = math.floor(height_limit / height)
        if layers <= 0:
            continue

        per_layer = fit_length * fit_width
        total = per_layer * layers

        summary = OrientationSummary(
            orientation=orientation,
            per_layer=per_layer,
            layers=layers,
            total=total,
            grid=(fit_length, fit_width),
        )

        if best is None:
            best = summary
            continue

        if summary.total > best.total:
            best = summary
            continue

        if summary.total == best.total:
            # Break ties preferring more layers, then more boxes per layer, and
            # finally the orientation with the lowest height.
            if summary.layers > best.layers:
                best = summary
                continue
            if summary.layers == best.layers and summary.per_layer > best.per_layer:
                best = summary
                continue
            if (
                summary.layers == best.layers
                and summary.per_layer == best.per_layer
                and summary.orientation[2] < best.orientation[2]
            ):
                best = summary

    return best


def compute_box_metrics(
    dims_mm: Orientation,
    pallet_length: MM,
    pallet_width: MM,
    overhang: MM,
    height_limits: Sequence[MM],
) -> BoxMetrics:
    """Compute the best orientation for each requested height limit."""

    sorted_dims = tuple(sorted(dims_mm, reverse=True))  # type: ignore[assignment]
    best_by_height: Dict[int, Optional[OrientationSummary]] = {}

    for limit in height_limits:
        best_by_height[limit] = find_best_orientation(
            dims_mm=dims_mm,
            pallet_length=pallet_length,
            pallet_width=pallet_width,
            height_limit=limit,
            overhang=overhang,
        )

    return BoxMetrics(
        dims_mm=dims_mm,
        sorted_dims_mm=sorted_dims,  # type: ignore[arg-type]
        best_by_height=best_by_height,
    )


def evaluate_combination_pair(
    dims_a: Orientation,
    dims_b: Orientation,
    pallet_length: MM,
    pallet_width: MM,
    overhang: MM,
    height_limit: MM,
) -> Tuple[bool, Optional[Dict[str, object]]]:
    """Check whether a pair of boxes can share a pallet layer.

    The heuristic tries three simple layouts:

    * Side-by-side along the pallet length.
    * Side-by-side along the pallet width.
    * Stacked on top of each other (sharing the same footprint).

    Parameters
    ----------
    dims_a, dims_b:
        The three dimensions of each box in millimetres.

    Returns
    -------
    tuple
        ``(True, details)`` if a feasible layout is found, otherwise
        ``(False, None)``.
    """

    max_length = pallet_length + overhang
    max_width = pallet_width + overhang

    for orient_a in unique_permutations(dims_a):
        for orient_b in unique_permutations(dims_b):
            length_a, width_a, height_a = orient_a
            length_b, width_b, height_b = orient_b

            if height_a > height_limit or height_b > height_limit:
                continue

            # Option 1: place along the pallet length (A next to B).
            if (
                length_a + length_b <= max_length
                and max(width_a, width_b) <= max_width
            ):
                return True, {
                    "orientation_a": orient_a,
                    "orientation_b": orient_b,
                    "arrangement": "length",
                }

            # Option 2: place along the pallet width (A behind B).
            if (
                max(length_a, length_b) <= max_length
                and width_a + width_b <= max_width
            ):
                return True, {
                    "orientation_a": orient_a,
                    "orientation_b": orient_b,
                    "arrangement": "width",
                }

            # Option 3: stack one on top of the other.
            if (
                max(length_a, length_b) <= max_length
                and max(width_a, width_b) <= max_width
                and height_a + height_b <= height_limit
            ):
                return True, {
                    "orientation_a": orient_a,
                    "orientation_b": orient_b,
                    "arrangement": "stack",
                }

    return False, None


def evaluate_combination_for_group(
    dims_list: Sequence[Orientation],
    pallet_length: MM,
    pallet_width: MM,
    overhang: MM,
    height_limits: Sequence[MM],
) -> Dict[int, Dict[str, object]]:
    """Evaluate feasibility for a group of SKUs sharing the same pallet ID."""

    results: Dict[int, Dict[str, object]] = {}

    if len(dims_list) < 2:
        return results

    # Only the first two SKUs are considered for the heuristic.  When more
    # SKUs share the same pallet we flag the case for manual verification.
    if len(dims_list) > 2:
        for limit in height_limits:
            results[limit] = {
                "ok": False,
                "detail": None,
                "note": "Более двух артикулов в группе",
            }
        return results

    dims_a, dims_b = dims_list[0], dims_list[1]

    for limit in height_limits:
        ok, detail = evaluate_combination_pair(
            dims_a=dims_a,
            dims_b=dims_b,
            pallet_length=pallet_length,
            pallet_width=pallet_width,
            overhang=overhang,
            height_limit=limit,
        )
        results[limit] = {"ok": ok, "detail": detail, "note": None}

    return results


def format_orientation(summary: Optional[OrientationSummary]) -> str:
    """Return a human-readable representation of an orientation summary."""

    if summary is None:
        return "Нет подходящей ориентации"

    length, width, height = summary.orientation
    grid_x, grid_y = summary.grid
    total = summary.total
    layers = summary.layers
    per_layer = summary.per_layer
    return (
        f"{length}×{width}×{height} мм,"
        f" слоёв: {layers}, в ряду: {grid_x}×{grid_y} = {per_layer},"
        f" всего: {total}"
    )


def mm_to_cm_string(value: MM) -> str:
    """Convert a millimetre value back to a string expressed in centimetres."""

    return f"{value / 10:.1f}"


def build_orientation_grid(
    summary: OrientationSummary,
) -> List[Tuple[float, float, float, float]]:
    """Return rectangle definitions for plotting a single-layer layout.

    Each tuple represents ``(x, y, width, height)`` in millimetres.
    """

    rectangles: List[Tuple[float, float, float, float]] = []
    length, width, _ = summary.orientation
    grid_x, grid_y = summary.grid

    for ix in range(grid_x):
        for iy in range(grid_y):
            rectangles.append((ix * length, iy * width, length, width))

    return rectangles


def build_combination_rectangles(
    detail: Dict[str, object]
) -> List[Tuple[float, float, float, float, str]]:
    """Return rectangle definitions for plotting a two-SKU layout."""

    rectangles: List[Tuple[float, float, float, float, str]] = []
    orient_a: Orientation = detail["orientation_a"]  # type: ignore[assignment]
    orient_b: Orientation = detail["orientation_b"]  # type: ignore[assignment]
    arrangement: str = detail.get("arrangement", "length")  # type: ignore[assignment]

    if arrangement == "length":
        offset_b = orient_a[0]
        rectangles.append((0, 0, orient_a[0], orient_a[1], "A"))
        rectangles.append((offset_b, 0, orient_b[0], orient_b[1], "B"))
    elif arrangement == "width":
        offset_b = orient_a[1]
        rectangles.append((0, 0, orient_a[0], orient_a[1], "A"))
        rectangles.append((0, offset_b, orient_b[0], orient_b[1], "B"))
    else:  # stack
        rectangles.append((0, 0, orient_a[0], orient_a[1], "A"))
        rectangles.append((0, 0, orient_b[0], orient_b[1], "B"))

    return rectangles

