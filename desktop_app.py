"""Графический интерфейс для настольной версии паллетного оптимизатора."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Optional

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from optimizer_workflow import (
    DEFAULT_PALLET_HEIGHT_LIMITS,
    DEFAULT_PALLET_LENGTH,
    DEFAULT_PALLET_WIDTH,
    DEFAULT_OVERHANG,
    draw_combination_layout,
    draw_single_layout,
    ensure_unique_selections,
    export_to_excel,
    format_orientation,
    process_dataframe,
)
from pallet_optimizer import BoxMetrics, detect_dimension_columns

matplotlib.use("TkAgg")


class OptimizerApp(tk.Tk):
    """Основное окно настольного приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Паллетный оптимизатор")
        self.geometry("1100x720")

        self.df_raw: Optional[pd.DataFrame] = None
        self.result_df: Optional[pd.DataFrame] = None
        self.metrics_map: Dict[object, BoxMetrics] = {}
        self.combination_details: Dict[object, Dict[object, object]] = {}

        self.height_limits = DEFAULT_PALLET_HEIGHT_LIMITS
        self.pallet_length = DEFAULT_PALLET_LENGTH
        self.pallet_width = DEFAULT_PALLET_WIDTH
        self.overhang = DEFAULT_OVERHANG
        self.selected_row_index: Optional[object] = None

        self._tree_iid_to_index: Dict[str, object] = {}
        self._combo_label_to_id: Dict[str, object] = {}

        self._build_ui()

    # ------------------------------------------------------------------ UI ---
    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.file_var = tk.StringVar()
        file_frame = ttk.LabelFrame(main_frame, text="Исходный файл")
        file_frame.pack(fill=tk.X)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_var)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 6), pady=8)
        ttk.Button(file_frame, text="Выбрать…", command=self._open_file_dialog).pack(
            side=tk.LEFT, padx=(0, 8), pady=8
        )

        params_frame = ttk.LabelFrame(main_frame, text="Параметры паллеты")
        params_frame.pack(fill=tk.X, pady=(12, 0))
        for i in range(6):
            params_frame.columnconfigure(i, weight=1)

        self.length_var = tk.StringVar(value=str(DEFAULT_PALLET_LENGTH))
        self.width_var = tk.StringVar(value=str(DEFAULT_PALLET_WIDTH))
        self.overhang_var = tk.StringVar(value=str(DEFAULT_OVERHANG))

        ttk.Label(params_frame, text="Длина (мм)").grid(row=0, column=0, sticky=tk.W, padx=8, pady=6)
        ttk.Entry(params_frame, textvariable=self.length_var, width=10).grid(
            row=0, column=1, sticky=tk.W
        )
        ttk.Label(params_frame, text="Ширина (мм)").grid(row=0, column=2, sticky=tk.W, padx=8)
        ttk.Entry(params_frame, textvariable=self.width_var, width=10).grid(
            row=0, column=3, sticky=tk.W
        )
        ttk.Label(params_frame, text="Свес (мм)").grid(row=0, column=4, sticky=tk.W, padx=8)
        ttk.Entry(params_frame, textvariable=self.overhang_var, width=10).grid(
            row=0, column=5, sticky=tk.W
        )

        columns_frame = ttk.LabelFrame(main_frame, text="Соответствие колонок")
        columns_frame.pack(fill=tk.X, pady=(12, 0))
        for i in range(2):
            columns_frame.columnconfigure(i, weight=1)

        self.length_column_var = tk.StringVar()
        self.width_column_var = tk.StringVar()
        self.height_column_var = tk.StringVar()
        self.pallet_column_var = tk.StringVar()
        self.use_pallet_column_var = tk.BooleanVar(value=False)

        ttk.Label(columns_frame, text="Длина / глубина").grid(
            row=0, column=0, sticky=tk.W, padx=8, pady=(8, 2)
        )
        self.length_column_box = ttk.Combobox(
            columns_frame, textvariable=self.length_column_var, state="disabled"
        )
        self.length_column_box.grid(row=0, column=1, sticky=tk.EW, padx=8, pady=(8, 2))

        ttk.Label(columns_frame, text="Ширина").grid(row=1, column=0, sticky=tk.W, padx=8)
        self.width_column_box = ttk.Combobox(
            columns_frame, textvariable=self.width_column_var, state="disabled"
        )
        self.width_column_box.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=2)

        ttk.Label(columns_frame, text="Высота").grid(row=2, column=0, sticky=tk.W, padx=8)
        self.height_column_box = ttk.Combobox(
            columns_frame, textvariable=self.height_column_var, state="disabled"
        )
        self.height_column_box.grid(row=2, column=1, sticky=tk.EW, padx=8, pady=2)

        self.pallet_checkbox = ttk.Checkbutton(
            columns_frame,
            text="Есть колонка с ID паллеты",
            variable=self.use_pallet_column_var,
            command=self._toggle_pallet_column,
        )
        self.pallet_checkbox.grid(row=3, column=0, sticky=tk.W, padx=8, pady=(6, 2))

        self.pallet_column_box = ttk.Combobox(
            columns_frame, textvariable=self.pallet_column_var, state="disabled"
        )
        self.pallet_column_box.grid(row=3, column=1, sticky=tk.EW, padx=8, pady=(6, 2))

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(12, 0))

        self.calculate_button = ttk.Button(
            buttons_frame, text="Рассчитать", command=self._calculate, state="disabled"
        )
        self.calculate_button.pack(side=tk.LEFT, padx=(8, 6))

        self.save_button = ttk.Button(
            buttons_frame, text="Сохранить в Excel", command=self._save_excel, state="disabled"
        )
        self.save_button.pack(side=tk.LEFT)

        table_frame = ttk.LabelFrame(main_frame, text="Результаты расчёта")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=8)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=(8, 0))
        self.tree.configure(xscrollcommand=hsb.set)

        detail_frame = ttk.LabelFrame(main_frame, text="Подробности по строке")
        detail_frame.pack(fill=tk.BOTH, expand=False, pady=(12, 0))

        height_frame = ttk.Frame(detail_frame)
        height_frame.pack(anchor=tk.W, padx=8, pady=(6, 4))
        self.height_var = tk.IntVar(value=self.height_limits[0])
        for limit in self.height_limits:
            label = f"{limit / 10:.0f} см"
            ttk.Radiobutton(
                height_frame,
                text=label,
                value=limit,
                variable=self.height_var,
                command=self._on_height_change,
            ).pack(side=tk.LEFT, padx=(0, 12))

        self.show_layout_button = ttk.Button(
            detail_frame,
            text="Показать схему слоя",
            command=self._show_single_layout,
            state="disabled",
        )
        self.show_layout_button.pack(anchor=tk.W, padx=8)

        self.detail_text = tk.Text(detail_frame, height=6, wrap="word", state="disabled")
        self.detail_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))

        combo_frame = ttk.LabelFrame(main_frame, text="Комбинированные паллеты")
        combo_frame.pack(fill=tk.BOTH, expand=False, pady=(12, 0))

        combo_top = ttk.Frame(combo_frame)
        combo_top.pack(fill=tk.X, padx=8, pady=(6, 4))

        ttk.Label(combo_top, text="ID паллеты:").pack(side=tk.LEFT)
        self.combo_var = tk.StringVar()
        self.combo_box = ttk.Combobox(combo_top, textvariable=self.combo_var, state="disabled", width=18)
        self.combo_box.pack(side=tk.LEFT, padx=(6, 12))
        self.combo_box.bind("<<ComboboxSelected>>", self._on_combo_change)

        ttk.Label(combo_top, text="Высота:").pack(side=tk.LEFT)
        self.combo_height_var = tk.StringVar()
        self.combo_height_box = ttk.Combobox(combo_top, textvariable=self.combo_height_var, state="disabled", width=10)
        self.combo_height_box.pack(side=tk.LEFT, padx=(6, 12))
        self.combo_height_box.bind("<<ComboboxSelected>>", self._on_combo_height_change)

        self.combo_button = ttk.Button(
            combo_top,
            text="Показать комбинированную схему",
            command=self._show_combination_layout,
            state="disabled",
        )
        self.combo_button.pack(side=tk.LEFT)

        self.combo_text = tk.Text(combo_frame, height=5, wrap="word", state="disabled")
        self.combo_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    # --------------------------------------------------------------- Helpers ---
    def _open_file_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm")],
        )
        if not path:
            return
        self.file_var.set(path)
        self._load_dataframe(path)

    def _load_dataframe(self, path: str) -> None:
        try:
            df = pd.read_excel(path)
        except ValueError as exc:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {exc}")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Ошибка", f"Неожиданная ошибка: {exc}")
            return

        if df.empty:
            messagebox.showwarning("Внимание", "Файл не содержит данных.")
            return

        self.df_raw = df.ffill()
        columns = list(self.df_raw.columns)
        if not columns:
            messagebox.showwarning("Внимание", "Файл не содержит колонок.")
            return

        detected = detect_dimension_columns(columns)
        self.length_column_box.configure(values=columns, state="readonly")
        self.width_column_box.configure(values=columns, state="readonly")
        self.height_column_box.configure(values=columns, state="readonly")
        self.pallet_column_box.configure(values=columns)

        self.length_column_var.set(detected["length"] or columns[0])
        self.width_column_var.set(detected["width"] or columns[0])
        self.height_column_var.set(detected["height"] or columns[0])
        self.pallet_column_var.set(columns[0])

        self.calculate_button.configure(state="normal")
        self._reset_results()

    def _toggle_pallet_column(self) -> None:
        if self.use_pallet_column_var.get():
            self.pallet_column_box.configure(state="readonly")
        else:
            self.pallet_column_box.configure(state="disabled")

    def _reset_results(self) -> None:
        self.result_df = None
        self.metrics_map = {}
        self.combination_details = {}
        self.selected_row_index = None

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree.configure(columns=())
        self._tree_iid_to_index.clear()
        self._combo_label_to_id.clear()

        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.configure(state="disabled")
        self.show_layout_button.configure(state="disabled")

        self.combo_box.configure(state="disabled", values=())
        self.combo_box.set("")
        self.combo_height_box.configure(state="disabled", values=())
        self.combo_height_var.set("")
        self.combo_button.configure(state="disabled")
        self.combo_text.configure(state="normal")
        self.combo_text.delete("1.0", tk.END)
        self.combo_text.configure(state="disabled")

        self.save_button.configure(state="disabled")

    # -------------------------------------------------------------- Actions ---
    def _calculate(self) -> None:
        if self.df_raw is None:
            messagebox.showwarning("Внимание", "Сначала выберите Excel-файл.")
            return

        try:
            pallet_length = int(self.length_var.get())
            pallet_width = int(self.width_var.get())
            overhang = int(self.overhang_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Параметры паллеты должны быть целыми числами.")
            return

        mapping = {
            "length": self.length_column_var.get(),
            "width": self.width_column_var.get(),
            "height": self.height_column_var.get(),
        }

        if not ensure_unique_selections(mapping):
            messagebox.showerror("Ошибка", "Колонки длины, ширины и высоты должны различаться.")
            return

        pallet_column = None
        if self.use_pallet_column_var.get():
            pallet_column = self.pallet_column_var.get()
            if not pallet_column:
                messagebox.showerror("Ошибка", "Укажите колонку с ID паллеты или снимите галочку.")
                return

        self.pallet_length = pallet_length
        self.pallet_width = pallet_width
        self.overhang = overhang

        try:
            result_df, metrics_map, combination_details = process_dataframe(
                df_raw=self.df_raw,
                mapping=mapping,
                pallet_length=pallet_length,
                pallet_width=pallet_width,
                overhang=overhang,
                height_limits=self.height_limits,
                pallet_id_column=pallet_column,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Ошибка", f"Не удалось выполнить расчёт: {exc}")
            return

        self.result_df = result_df
        self.metrics_map = metrics_map
        self.combination_details = combination_details

        self._populate_table()
        self._update_combination_controls()
        self.save_button.configure(state="normal" if not result_df.empty else "disabled")

    def _populate_table(self) -> None:
        if self.result_df is None:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        columns = list(self.result_df.columns)
        self.tree.configure(columns=columns, show="headings")

        for col in columns:
            values = self.result_df[col].fillna("").astype(str).tolist()
            max_len = max(len(col), *(len(value) for value in values))
            width = min(40, max_len) * 8 + 24
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor=tk.W)

        self._tree_iid_to_index.clear()
        for idx in self.result_df.index:
            row = self.result_df.loc[idx]
            values = ["" if pd.isna(value) else value for value in row]
            display_values = ["" if value is None else str(value) for value in values]
            iid = str(idx)
            self._tree_iid_to_index[iid] = idx
            self.tree.insert("", tk.END, iid=iid, values=display_values)

        if len(self.result_df.index) > 0:
            first_index = self.result_df.index[0]
            iid = str(first_index)
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self._update_row_details(first_index)

    def _update_row_details(self, index: object) -> None:
        self.selected_row_index = index
        metrics = self.metrics_map.get(index)

        lines = [f"Строка: {index}"]
        if metrics:
            for limit in self.height_limits:
                summary = metrics.best_by_height.get(limit)
                lines.append(f"{limit / 10:.0f} см: {format_orientation(summary)}")
        else:
            lines.append("Для строки не удалось выполнить расчёт.")

        if self.result_df is not None and index in self.result_df.index:
            note = self.result_df.loc[index].get("Примечание", "")
            combo_label = self.result_df.loc[index].get("Комбинация допустима?", "")
            if isinstance(note, str) and note:
                lines.append(f"Примечание: {note}")
            if isinstance(combo_label, str) and combo_label:
                lines.append(f"Комбинация: {combo_label}")

        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, "\n".join(lines))
        self.detail_text.configure(state="disabled")

        self._update_layout_button_state()

    def _update_layout_button_state(self) -> None:
        if self.selected_row_index is None:
            self.show_layout_button.configure(state="disabled")
            return

        metrics = self.metrics_map.get(self.selected_row_index)
        selected_height = self.height_var.get()
        if metrics and metrics.best_by_height.get(selected_height):
            self.show_layout_button.configure(state="normal")
        else:
            self.show_layout_button.configure(state="disabled")

    def _update_combination_controls(self) -> None:
        if not self.combination_details:
            self.combo_box.configure(state="disabled", values=())
            self.combo_button.configure(state="disabled")
            self.combo_height_box.configure(state="disabled", values=())
            self.combo_height_var.set("")
            self.combo_text.configure(state="normal")
            self.combo_text.delete("1.0", tk.END)
            self.combo_text.configure(state="disabled")
            return

        ids = list(self.combination_details.keys())
        self._combo_label_to_id = {self._format_combo_label(value): value for value in ids}
        labels = list(self._combo_label_to_id.keys())
        self.combo_box.configure(state="readonly", values=labels)
        self.combo_box.set(labels[0])
        self.combo_var.set(labels[0])

        height_labels = [self._format_height_label(limit) for limit in self.height_limits]
        self.combo_height_box.configure(state="readonly", values=height_labels)
        self.combo_height_box.set(height_labels[0])
        self.combo_height_var.set(height_labels[0])

        self._update_combination_text()

    def _update_combination_text(self) -> None:
        label = self.combo_var.get()
        if not label:
            return

        group_id = self._combo_label_to_id.get(label)
        if group_id is None:
            return

        detail_info = self.combination_details.get(group_id, {})
        lines = [f"Паллета {label}"]

        selected_label = self.combo_height_var.get()
        selected_limit = self._parse_height_label(selected_label)

        for limit in self.height_limits:
            info = detail_info.get(limit)
            if not info:
                lines.append(f"{limit / 10:.0f} см: нет данных")
                continue
            note = info.get("note")
            if info.get("ok") and info.get("detail"):
                lines.append(f"{limit / 10:.0f} см: можно объединять")
            elif note:
                lines.append(f"{limit / 10:.0f} см: {note}")
            else:
                lines.append(f"{limit / 10:.0f} см: схема не найдена")

        recommended = detail_info.get("selected_height")
        if recommended:
            lines.append(f"Рекомендуемая высота: {recommended / 10:.0f} см")

        self.combo_text.configure(state="normal")
        self.combo_text.delete("1.0", tk.END)
        self.combo_text.insert(tk.END, "\n".join(lines))
        self.combo_text.configure(state="disabled")

        if selected_limit is None:
            self.combo_button.configure(state="disabled")
            return

        info = detail_info.get(selected_limit)
        if info and info.get("ok") and info.get("detail"):
            self.combo_button.configure(state="normal")
        else:
            self.combo_button.configure(state="disabled")

    # -------------------------------------------------------------- Events ---
    def _on_row_select(self, event: tk.Event) -> None:  # type: ignore[override]
        selection = self.tree.selection()
        if not selection:
            return
        iid = selection[0]
        index = self._tree_iid_to_index.get(iid)
        if index is None:
            return
        self._update_row_details(index)

    def _on_height_change(self) -> None:
        self._update_layout_button_state()

    def _on_combo_change(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._update_combination_text()

    def _on_combo_height_change(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._update_combination_text()

    # ----------------------------------------------------------- UI actions ---
    def _show_single_layout(self) -> None:
        if self.selected_row_index is None:
            return

        metrics = self.metrics_map.get(self.selected_row_index)
        if not metrics:
            messagebox.showinfo("Схема", "Для выбранной строки нет данных.")
            return

        selected_height = self.height_var.get()
        summary = metrics.best_by_height.get(selected_height)
        if not summary:
            messagebox.showinfo(
                "Схема",
                "Для выбранной высоты коробка не размещается на паллете.",
            )
            return

        fig = draw_single_layout(summary, self.pallet_length, self.pallet_width, self.overhang)
        self._open_figure_window("Схема слоя", fig)

    def _show_combination_layout(self) -> None:
        label = self.combo_var.get()
        if not label:
            return

        group_id = self._combo_label_to_id.get(label)
        if group_id is None:
            return

        detail_info = self.combination_details.get(group_id, {})
        selected_limit = self._parse_height_label(self.combo_height_var.get())
        if selected_limit is None:
            messagebox.showinfo("Схема", "Выберите ограничение по высоте.")
            return

        info = detail_info.get(selected_limit)
        if not info or not info.get("ok") or not info.get("detail"):
            messagebox.showinfo("Схема", "Для выбранного ограничения нет схемы.")
            return

        fig = draw_combination_layout(info["detail"], self.pallet_length, self.pallet_width, self.overhang)
        self._open_figure_window(f"Комбинированная паллета {label}", fig)

    def _open_figure_window(self, title: str, figure) -> None:
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("720x520")

        canvas = FigureCanvasTkAgg(figure, master=window)
        canvas.draw()
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)

        def _on_close() -> None:
            plt.close(figure)
            window.destroy()

        ttk.Button(window, text="Закрыть", command=_on_close).pack(pady=8)
        window.protocol("WM_DELETE_WINDOW", _on_close)

    def _save_excel(self) -> None:
        if self.result_df is None or self.result_df.empty:
            messagebox.showinfo("Сохранение", "Сначала выполните расчёт.")
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return

        data = export_to_excel(self.result_df)
        try:
            with open(path, "wb") as fh:
                fh.write(data)
        except OSError as exc:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {exc}")
            return

        messagebox.showinfo("Сохранение", "Результат успешно сохранён.")

    # ----------------------------------------------------------- Misc utils ---
    def _format_height_label(self, value: int) -> str:
        return f"{value / 10:.0f} см"

    def _parse_height_label(self, label: str) -> Optional[int]:
        for limit in self.height_limits:
            if label == self._format_height_label(limit):
                return limit
        return None

    def _format_combo_label(self, value: object) -> str:
        return str(value)


def main() -> None:
    app = OptimizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
