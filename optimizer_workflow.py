"""Вспомогательные функции, общие для веб- и настольного интерфейсов."""

from __future__ import annotations

import io
from typing import Dict, List, Optional, Sequence, Tuple

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd

from pallet_optimizer import (
    BoxMetrics,
    OrientationSummary,
    build_combination_rectangles,
    build_orientation_grid,
    cm_to_mm,
    compute_box_metrics,
    evaluate_combination_for_group,
    mm_to_cm_string,
    parse_dimension_value,
)

DEFAULT_PALLET_LENGTH = 1200  # мм
DEFAULT_PALLET_WIDTH = 800  # мм
DEFAULT_PALLET_HEIGHT_LIMITS = [1800, 1700]  # мм
DEFAULT_OVERHANG = 30  # мм


def ensure_unique_selections(selection: Dict[str, Optional[str]]) -> bool:
    """Return ``True`` если каждая выбранная колонка уникальна."""

    values = [value for value in selection.values() if value]
    return len(values) == len(set(values))


def convert_row_to_mm(row: pd.Series, mapping: Dict[str, str]) -> Optional[List[int]]:
    """Convert сантиметровые размеры строки в миллиметры."""

    dims_cm: List[Optional[float]] = []
    for logical in ("length", "width", "height"):
        value = row[mapping[logical]]
        dims_cm.append(parse_dimension_value(value))

    if any(value is None for value in dims_cm):
        return None

    dims_mm = [cm_to_mm(value) for value in dims_cm if value is not None]
    return dims_mm if len(dims_mm) == 3 else None


def format_orientation(summary: Optional[OrientationSummary]) -> str:
    """Вернуть удобочитаемую подпись ориентации."""

    if summary is None:
        return "—"

    length, width, height = summary.orientation
    grid_x, grid_y = summary.grid
    return (
        f"{mm_to_cm_string(length)}×{mm_to_cm_string(width)}×{mm_to_cm_string(height)} см;"
        f" слой: {grid_x}×{grid_y}, слоёв: {summary.layers}, всего: {summary.total}"
    )


def export_to_excel(df: pd.DataFrame) -> bytes:
    """Serialize таблицу результатов в Excel."""

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.read()


def process_dataframe(
    df_raw: pd.DataFrame,
    mapping: Dict[str, str],
    pallet_length: int,
    pallet_width: int,
    overhang: int,
    height_limits: Sequence[int],
    pallet_id_column: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[object, BoxMetrics], Dict[object, Dict[object, object]]]:
    """Подготовить таблицу результатов и служебные структуры для UI."""

    metrics_map: Dict[object, BoxMetrics] = {}
    errors: Dict[object, str] = {}

    for idx, row in df_raw.iterrows():
        dims_mm = convert_row_to_mm(row, mapping)
        if dims_mm is None:
            errors[idx] = "Не удалось распознать размеры"
            continue

        metrics = compute_box_metrics(
            dims_mm=(dims_mm[0], dims_mm[1], dims_mm[2]),
            pallet_length=pallet_length,
            pallet_width=pallet_width,
            overhang=overhang,
            height_limits=height_limits,
        )
        metrics_map[idx] = metrics

    result_df = df_raw.copy()
    length_values: List[Optional[int]] = []
    width_values: List[Optional[int]] = []
    height_values: List[Optional[int]] = []
    max_values: Dict[int, List[Optional[int]]] = {limit: [] for limit in height_limits}
    scheme_values: Dict[int, List[str]] = {limit: [] for limit in height_limits}
    notes: List[str] = []

    for idx in result_df.index:
        metrics = metrics_map.get(idx)
        if not metrics:
            length_values.append(None)
            width_values.append(None)
            height_values.append(None)
            for limit in height_limits:
                max_values[limit].append(None)
                scheme_values[limit].append("—")
            notes.append(errors.get(idx, ""))
            continue

        length_values.append(metrics.sorted_dims_mm[0])
        width_values.append(metrics.sorted_dims_mm[1])
        height_values.append(metrics.sorted_dims_mm[2])

        for limit in height_limits:
            summary = metrics.best_by_height.get(limit)
            max_values[limit].append(summary.total if summary else 0)
            scheme_values[limit].append(format_orientation(summary))
        notes.append("")

    result_df["Длина, мм"] = length_values
    result_df["Ширина, мм"] = width_values
    result_df["Высота, мм"] = height_values

    for limit in height_limits:
        cm_label = f"{limit / 10:.0f}"
        result_df[f"Макс. на паллету ({cm_label} см)"] = max_values[limit]
        result_df[f"Схема укладки {cm_label} см"] = scheme_values[limit]

    result_df["Примечание"] = notes

    combination_details: Dict[object, Dict[object, object]] = {}
    combination_column: List[str] = ["—" for _ in result_df.index]

    if pallet_id_column:
        grouped = result_df.groupby(pallet_id_column)
        for group_id, group_df in grouped:
            dims_list = []
            for row_idx in group_df.index:
                metrics = metrics_map.get(row_idx)
                if metrics:
                    dims_list.append(metrics.dims_mm)

            if not dims_list:
                continue

            combination_result = evaluate_combination_for_group(
                dims_list=dims_list,
                pallet_length=pallet_length,
                pallet_width=pallet_width,
                overhang=overhang,
                height_limits=height_limits,
            )
            combination_details[group_id] = combination_result

            label = "Нет"
            detail_height: Optional[int] = None

            if combination_result:
                for limit in height_limits:
                    info = combination_result.get(limit)
                    if info and info.get("ok"):
                        label = f"Да ({limit/10:.0f} см)"
                        detail_height = limit
                        break
                else:
                    if any((info or {}).get("note") for info in combination_result.values()):
                        label = "Проверить вручную"

            for row_idx in group_df.index:
                position = result_df.index.get_loc(row_idx)
                combination_column[position] = label

            if detail_height is not None:
                combination_details[group_id]["selected_height"] = detail_height

    result_df["Комбинация допустима?"] = combination_column
    return result_df, metrics_map, combination_details


def draw_single_layout(
    summary: OrientationSummary,
    pallet_length: int,
    pallet_width: int,
    overhang: int,
    color: str = "tab:blue",
):
    """Построить схему одного слоя паллеты."""

    rectangles = build_orientation_grid(summary)
    fig, ax = plt.subplots(figsize=(6, 4))

    for rect in rectangles:
        x, y, width, height = rect
        ax.add_patch(
            patches.Rectangle(
                (x, y),
                width,
                height,
                linewidth=1,
                edgecolor="black",
                facecolor=color,
                alpha=0.6,
            )
        )

    max_length = pallet_length + overhang
    max_width = pallet_width + overhang
    ax.add_patch(
        patches.Rectangle(
            (0, 0),
            max_length,
            max_width,
            linewidth=2,
            edgecolor="gray",
            facecolor="none",
            linestyle="--",
            label="Паллет",
        )
    )

    ax.set_xlim(0, max_length)
    ax.set_ylim(0, max_width)
    ax.set_xlabel("Длина, мм")
    ax.set_ylabel("Ширина, мм")
    ax.set_title("Схема одного слоя")
    ax.set_aspect("equal", adjustable="box")
    ax.grid(True, linestyle=":", linewidth=0.5, alpha=0.5)
    return fig


def draw_combination_layout(
    detail: Dict[str, object],
    pallet_length: int,
    pallet_width: int,
    overhang: int,
):
    """Построить схему комбинированной паллеты."""

    rectangles = build_combination_rectangles(detail)
    fig, ax = plt.subplots(figsize=(6, 4))
    colors = {"A": "tab:blue", "B": "tab:red"}

    for rect in rectangles:
        x, y, width, height, label = rect
        ax.add_patch(
            patches.Rectangle(
                (x, y),
                width,
                height,
                linewidth=1,
                edgecolor="black",
                facecolor=colors[label],
                alpha=0.6,
            )
        )

    max_length = pallet_length + overhang
    max_width = pallet_width + overhang
    ax.add_patch(
        patches.Rectangle(
            (0, 0),
            max_length,
            max_width,
            linewidth=2,
            edgecolor="gray",
            facecolor="none",
            linestyle="--",
        )
    )

    arrangement = detail.get("arrangement", "length")
    ax.set_xlim(0, max_length)
    ax.set_ylim(0, max_width)
    ax.set_xlabel("Длина, мм")
    ax.set_ylabel("Ширина, мм")
    ax.set_title(f"Комбинированная раскладка ({arrangement})")
    ax.set_aspect("equal", adjustable="box")
    ax.grid(True, linestyle=":", linewidth=0.5, alpha=0.5)
    return fig
