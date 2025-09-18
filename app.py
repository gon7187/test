import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

from optimizer_workflow import (
    DEFAULT_PALLET_HEIGHT_LIMITS,
    DEFAULT_PALLET_LENGTH,
    DEFAULT_PALLET_WIDTH,
    DEFAULT_OVERHANG,
    draw_combination_layout,
    draw_single_layout,
    ensure_unique_selections,
    export_to_excel,
    process_dataframe,
)
from pallet_optimizer import detect_dimension_columns


st.set_page_config(page_title="Паллетный оптимизатор", layout="wide")


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

    if not ensure_unique_selections(mapping):
        st.error("Каждая размерная колонка должна быть уникальной.")
        return

    pallet_id_column = None if pallet_column == "—" else pallet_column

    height_limits = DEFAULT_PALLET_HEIGHT_LIMITS

    result_df, metrics_map, combination_details = process_dataframe(
        df_raw=df_raw,
        mapping=mapping,
        pallet_length=pallet_length,
        pallet_width=pallet_width,
        overhang=overhang,
        height_limits=height_limits,
        pallet_id_column=pallet_id_column,
    )

    st.subheader("Результаты расчёта")
    st.dataframe(result_df, use_container_width=True)

    st.download_button(
        "Скачать результат в Excel",
        data=export_to_excel(result_df),
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
            fig = draw_single_layout(selected_summary, pallet_length, pallet_width, overhang)
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
                fig = draw_combination_layout(detail["detail"], pallet_length, pallet_width, overhang)
                st.pyplot(fig)
                plt.close(fig)
            else:
                st.info("Для выбранной паллеты нет подтверждённой схемы.")


if __name__ == "__main__":
    main()

