import io
from typing import Dict, List, Optional

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
import streamlit as st

from pallet_optimizer import (
    BoxMetrics,
    OrientationSummary,
    build_combination_rectangles,
    build_orientation_grid,
    cm_to_mm,
    compute_box_metrics,
    detect_dimension_columns,
    evaluate_combination_for_group,
    mm_to_cm_string,
    parse_dimension_value,
)


st.set_page_config(page_title="Паллетный оптимизатор", layout="wide")


DEFAULT_PALLET_LENGTH = 1200  # мм
DEFAULT_PALLET_WIDTH = 800  # мм
DEFAULT_PALLET_HEIGHT_LIMITS = [1800, 1700]  # мм
DEFAULT_OVERHANG = 30  # мм


def _draw_single_layout(
    summary: OrientationSummary,
    pallet_length: int,
    pallet_width: int,
    overhang: int,
    color: str = "tab:blue",
):
    rectangles = build_orientation_grid(summary)
    fig, ax = plt.subplots(figsize=(6, 4))

    for rect in rectangles:
        x, y, width, height = rect
        ax.add_patch(
            patches.Rectangle(
                (x, y), width, height, linewidth=1, edgecolor="black", facecolor=color, alpha=0.6
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


def _draw_combination_layout(
    detail: Dict[str, object],
    pallet_length: int,
    pallet_width: int,
    overhang: int,
):
    rectangles = build_combination_rectangles(detail)
    fig, ax = plt.subplots(figsize=(6, 4))
    colors = {"A": "tab:blue", "B": "tab:red"}

    for rect in rectangles:
        x, y, width, height, label = rect
        ax.add_patch(
            patches.Rectangle(
                (x, y), width, height, linewidth=1, edgecolor="black", facecolor=colors[label], alpha=0.6
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

    ax.set_xlim(0, max_length)
    ax.set_ylim(0, max_width)
    ax.set_xlabel("Длина, мм")
    ax.set_ylabel("Ширина, мм")
    arrangement = detail.get("arrangement", "length")
    ax.set_title(f"Комбинированная раскладка ({arrangement})")
    ax.set_aspect("equal", adjustable="box")
    ax.grid(True, linestyle=":", linewidth=0.5, alpha=0.5)
    return fig


def _ensure_unique_selections(selection: Dict[str, Optional[str]]) -> bool:
    values = [value for value in selection.values() if value]
    return len(values) == len(set(values))


def _convert_row_to_mm(row, mapping: Dict[str, str]) -> Optional[List[int]]:
    dims_cm: List[Optional[float]] = []
    for logical in ("length", "width", "height"):
        value = row[mapping[logical]]
        dims_cm.append(parse_dimension_value(value))

    if any(value is None for value in dims_cm):
        return None

    dims_mm = [cm_to_mm(value) for value in dims_cm if value is not None]
    return dims_mm if len(dims_mm) == 3 else None


def _format_orientation(summary: Optional[OrientationSummary]) -> str:
    if summary is None:
        return "—"

    length, width, height = summary.orientation
    grid_x, grid_y = summary.grid
    return (
        f"{mm_to_cm_string(length)}×{mm_to_cm_string(width)}×{mm_to_cm_string(height)} см;"
        f" слой: {grid_x}×{grid_y}, слоёв: {summary.layers}, всего: {summary.total}"
    )


def _export_to_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.read()


def main() -> None:
    st.title("Оптимизация укладки товаров на европаллету")

    st.sidebar.header("Параметры паллеты")
    pallet_length = st.sidebar.number_input(
        "Длина паллеты (мм)", min_value=200, max_value=2400, value=DEFAULT_PALLET_LENGTH, step=10
    )
    pallet_width = st.sidebar.number_input(
        "Ширина паллеты (мм)", min_value=200, max_value=1600, value=DEFAULT_PALLET_WIDTH, step=10
    )
    overhang = st.sidebar.number_input(
        "Допустимый свес (мм)", min_value=0, max_value=100, value=DEFAULT_OVERHANG, step=5
    )

    st.sidebar.markdown(
        """Расчёт выполняется для двух ограничений по высоте: 1800 и 1700 мм."""
    )

    uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx", "xlsm", "xls"])  # type: ignore[arg-type]

    if not uploaded_file:
        st.info("Загрузите файл, чтобы увидеть расчёты.")
        return

    try:
        df_raw = pd.read_excel(uploaded_file)
    except ValueError as exc:
        st.error(f"Не удалось прочитать файл: {exc}")
        return

    if df_raw.empty:
        st.warning("Файл не содержит данных.")
        return

    df_raw = df_raw.ffill()

    detected = detect_dimension_columns(df_raw.columns)
    st.subheader("Назначение колонок")
    col1, col2, col3, col4 = st.columns(4)
    length_column = col1.selectbox(
        "Колонка длины/глубины (см)",
        options=list(df_raw.columns),
        index=list(df_raw.columns).index(detected["length"]) if detected["length"] in df_raw.columns else 0,
    )
    width_column = col2.selectbox(
        "Колонка ширины (см)",
        options=list(df_raw.columns),
        index=list(df_raw.columns).index(detected["width"]) if detected["width"] in df_raw.columns else 0,
    )
    height_column = col3.selectbox(
        "Колонка высоты (см)",
        options=list(df_raw.columns),
        index=list(df_raw.columns).index(detected["height"]) if detected["height"] in df_raw.columns else 0,
    )
    pallet_column = col4.selectbox(
        "Колонка ID паллеты (опционально)",
        options=["—"] + list(df_raw.columns),
        index=0,
    )

    mapping = {"length": length_column, "width": width_column, "height": height_column}

    if not _ensure_unique_selections(mapping):
        st.error("Каждая размерная колонка должна быть уникальной.")
        return

    pallet_id_column = None if pallet_column == "—" else pallet_column

    height_limits = DEFAULT_PALLET_HEIGHT_LIMITS

    processed_rows: List[Dict[str, object]] = []
    metrics_map: Dict[int, BoxMetrics] = {}

    for idx, row in df_raw.iterrows():
        dims_mm = _convert_row_to_mm(row, mapping)
        if dims_mm is None:
            processed_rows.append(
                {
                    "index": idx,
                    "dims": None,
                    "error": "Не удалось распознать размеры",
                }
            )
            continue

        metrics = compute_box_metrics(
            dims_mm=(dims_mm[0], dims_mm[1], dims_mm[2]),
            pallet_length=pallet_length,
            pallet_width=pallet_width,
            overhang=overhang,
            height_limits=height_limits,
        )
        metrics_map[idx] = metrics

        processed_rows.append(
            {
                "index": idx,
                "dims": metrics.sorted_dims_mm,
                "best": metrics.best_by_height,
                "error": None,
            }
        )

    result_df = df_raw.copy()
    length_values: List[Optional[int]] = []
    width_values: List[Optional[int]] = []
    height_values: List[Optional[int]] = []
    max_180: List[Optional[int]] = []
    max_170: List[Optional[int]] = []
    scheme_180: List[str] = []
    scheme_170: List[str] = []
    notes: List[str] = []

    for item in processed_rows:
        idx = item["index"]
        metrics = metrics_map.get(idx)

        if metrics is None:
            length_values.append(None)
            width_values.append(None)
            height_values.append(None)
            max_180.append(None)
            max_170.append(None)
            scheme_180.append("—")
            scheme_170.append("—")
            notes.append(item.get("error", ""))
            continue

        length_values.append(metrics.sorted_dims_mm[0])
        width_values.append(metrics.sorted_dims_mm[1])
        height_values.append(metrics.sorted_dims_mm[2])

        summary_180 = metrics.best_by_height.get(1800)
        summary_170 = metrics.best_by_height.get(1700)

        max_180.append(summary_180.total if summary_180 else 0)
        max_170.append(summary_170.total if summary_170 else 0)
        scheme_180.append(_format_orientation(summary_180))
        scheme_170.append(_format_orientation(summary_170))
        notes.append("")

    result_df["Длина, мм"] = length_values
    result_df["Ширина, мм"] = width_values
    result_df["Высота, мм"] = height_values
    result_df["Макс. на паллету (180 см)"] = max_180
    result_df["Макс. на паллету (170 см)"] = max_170
    result_df["Схема укладки 180 см"] = scheme_180
    result_df["Схема укладки 170 см"] = scheme_170
    result_df["Примечание"] = notes

    combination_details: Dict[object, Dict[object, object]] = {}
    combination_column: List[str] = ["—" for _ in result_df.index]

    if pallet_id_column:
        grouped = result_df.groupby(pallet_id_column)
        for group_id, group_df in grouped:
            if group_df.shape[0] < 2:
                continue

            dims_list = []
            for row_idx in group_df.index:
                metrics = metrics_map.get(row_idx)
                if not metrics:
                    continue
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
            detail_height = None

            if combination_result:
                for limit in height_limits:
                    info = combination_result.get(limit)
                    if info and info.get("ok"):
                        label = f"Да ({limit/10:.0f} см)"
                        detail_height = limit
                        break
                else:
                    if any(info.get("note") for info in combination_result.values()):
                        label = "Проверить вручную"

            for row_idx in group_df.index:
                position = result_df.index.get_loc(row_idx)
                combination_column[position] = label

            if detail_height is not None:
                combination_details[group_id]["selected_height"] = detail_height

    result_df["Комбинация допустима?"] = combination_column

    st.subheader("Результаты расчёта")
    st.dataframe(result_df, use_container_width=True)

    st.download_button(
        "Скачать результат в Excel",
        data=_export_to_excel(result_df),
        file_name="pallet_optimization.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Визуализация")
    selected_index = st.selectbox(
        "Выберите строку для схемы",
        options=list(result_df.index),
        format_func=lambda idx: str(result_df.loc[idx].get(result_df.columns[0], idx)),
    )

    selected_height_option = st.radio("Ограничение по высоте", options=height_limits, format_func=lambda v: f"{v/10:.0f} см")

    selected_metrics = metrics_map.get(selected_index)
    if selected_metrics:
        selected_summary = selected_metrics.best_by_height.get(selected_height_option)
        if selected_summary:
            fig = _draw_single_layout(selected_summary, pallet_length, pallet_width, overhang)
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.info("Для выбранной высоты коробка не размещается на паллете.")
    else:
        st.warning("Не удалось построить схему для выбранной строки.")

    if pallet_id_column and combination_details:
        st.subheader("Комбинированные паллеты")
        combo_ids = list(combination_details.keys())
        if combo_ids:
            selected_combo = st.selectbox("ID паллеты", options=combo_ids)
            detail_info = combination_details.get(selected_combo, {})
            selected_height = detail_info.get("selected_height")
            if selected_height is None:
                selected_height = height_limits[0]
            detail = detail_info.get(selected_height)
            if detail and detail.get("ok") and detail.get("detail"):
                fig = _draw_combination_layout(detail["detail"], pallet_length, pallet_width, overhang)
                st.pyplot(fig)
                plt.close(fig)
            else:
                st.info("Для выбранной паллеты нет подтверждённой схемы.")


if __name__ == "__main__":
    main()

