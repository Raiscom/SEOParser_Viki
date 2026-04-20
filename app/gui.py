"""Графический интерфейс приложения."""

from __future__ import annotations

import asyncio
import json
import threading
import tkinter as tk
from concurrent.futures import Future
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import ttkbootstrap as ttk
from loguru import logger

from app.config import AppSettings, get_settings, save_settings
from app.models import SerpRiverResult, WordstatResult, XmlRiverResult
from app.services.serpriver import SerpRiverClient
from app.services.wordstat import WordstatClient
from app.services.xmlriver import XmlRiverClient
from app.utils.autocomplete import ComboboxAutocomplete
from app.utils.export import build_export_path, export_to_csv, export_to_xlsx
from app.utils.io import read_csv_lines, read_xlsx_lines
from app.utils.reference_data import load_reference_catalogs

TREE_ERROR_TAG = "error"
YANDEX_DOMAIN_VALUES = ("ru", "com", "ua", "com.tr", "by", "kz")
DEVICE_VALUES = ("desktop", "tablet", "mobile")
WORDSTAT_DEVICE_VALUES = (
    "",
    "desktop",
    "tablet",
    "phone",
    "desktop,tablet",
    "desktop,phone",
    "tablet,phone",
    "desktop,tablet,phone",
)
WORDSTAT_PERIOD_VALUES = ("", "month", "week", "day")
WORDSTAT_PAGETYPE_VALUES = ("words", "history")
SERPRIVER_OUTPUT_FORMAT_VALUES = ("json", "xml")
SORT_DIRECTION_VALUES = ("A -> Я / 1 -> 9", "Я -> A / 9 -> 1")
COMBOBOX_MAX_VALUES = 200
FILTER_DEBOUNCE_MS = 300
AUTOCOMPLETE_DEBOUNCE_MS = 120
APP_FONT = ("Segoe UI", 10)
APP_FONT_BOLD = ("Segoe UI", 10, "bold")
APP_FONT_SMALL = ("Segoe UI", 9)
ACCENT_COLOR = "#FF8C00"
ACCENT_DARK = "#CC7000"
TREE_ROW_HEIGHT = 26
PANED_HANDLE_SIZE = 14
GOOGLE_FIXED_COUNTRY_LABEL = "Россия"
GOOGLE_FIXED_COUNTRY_VALUE = "RU"
GOOGLE_FIXED_LANGUAGE_LABEL = "Русский"
GOOGLE_FIXED_LANGUAGE_VALUE = "ru"
GOOGLE_FIXED_DOMAIN_VALUE = "ru"
QUERY_TEXT_MIN_HEIGHT = 170
RESULTS_MIN_HEIGHT = 300
WORDSTAT_DATE_FORMAT = "%Y-%m-%d"
SERPRIVER_OPTION_LABELS = {
    "result_cnt": "Количество результатов",
    "domain": "Домен поиска",
    "device": "Устройство",
    "lr": "Регион Яндекса (lr)",
    "location": "Локация Google (location)",
    "hl": "Язык интерфейса (hl)",
    "gl": "Страна поиска (gl)",
    "output_format": "Формат ответа",
}


class SeoParserApp:
    """Создает главное окно и вкладки приложения."""

    def __init__(self) -> None:
        self.settings = get_settings()
        self.reference_catalogs, self.reference_errors = load_reference_catalogs()
        self.reference_maps = {
            catalog_name: {option.label: option.value for option in options}
            for catalog_name, options in self.reference_catalogs.items()
        }

        self.root = ttk.Window(
            themename="darkly",
            title="SEO API Парсер VIKI",
            size=(1420, 860),
            minsize=(1180, 720),
        )
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._configure_styles()
        self._init_edit_menu()

        self.background_loop = asyncio.new_event_loop()
        self.loop_thread = threading.Thread(
            target=self._run_background_loop,
            name="seo-parser-loop",
            daemon=True,
        )
        self.loop_thread.start()

        self._filter_debounce_ids: dict[str, str] = {}
        self.autocomplete_helpers: list[ComboboxAutocomplete] = []

        self.is_google_xml_running = False
        self.is_yandex_xml_running = False
        self.is_wordstat_running = False
        self.is_serpriver_running = False

        self.xmlriver_user_var = tk.StringVar(value=self.settings.xmlriver_user_id)
        self.xmlriver_key_var = tk.StringVar(value=self.settings.xmlriver_api_key)
        self.is_xmlriver_key_visible = False
        self.xml_shared_controls: list[tuple[tk.Widget, str]] = []

        self.google_loc_var = tk.StringVar(
            value=self._get_saved_reference_label("google_geo", self.settings.google_last_location_label),
        )
        self.google_country_var = tk.StringVar(value=GOOGLE_FIXED_COUNTRY_LABEL)
        self.google_lr_var = tk.StringVar(value=GOOGLE_FIXED_LANGUAGE_LABEL)
        self.google_domain_var = tk.StringVar(value=GOOGLE_FIXED_DOMAIN_VALUE)
        self.google_device_var = tk.StringVar(value="desktop")
        self.google_controls: list[tuple[tk.Widget, str]] = []
        self.google_required_catalogs = ("google_geo",)
        self.is_google_blocked = False

        self.yandex_lr_var = tk.StringVar(
            value=self._get_saved_reference_label("yandex_geo", self.settings.yandex_last_region_label),
        )
        self.yandex_domain_var = tk.StringVar(value="ru")
        self.yandex_lang_var = tk.StringVar(value="ru")
        self.yandex_device_var = tk.StringVar(value="desktop")
        self.yandex_controls: list[tuple[tk.Widget, str]] = []
        self.yandex_required_catalogs = ("yandex_geo",)
        self.is_yandex_blocked = False

        self.wordstat_region_var = tk.StringVar(
            value=self._get_saved_reference_label("yandex_geo", self.settings.wordstat_last_region_label),
        )
        self.wordstat_regions_extra_var = tk.StringVar(value="")
        self.wordstat_device_var = tk.StringVar(value="")
        self.wordstat_period_var = tk.StringVar(value="")
        self.wordstat_start_var = tk.StringVar(value=self._get_today_date_text())
        self.wordstat_end_var = tk.StringVar(value=self._get_today_date_text())
        self.wordstat_pagetype_var = tk.StringVar(value="words")
        self.wordstat_controls: list[tuple[tk.Widget, str]] = []
        self.wordstat_required_catalogs = ("yandex_geo",)
        self.is_wordstat_blocked = False

        self.google_results: list[XmlRiverResult] = []
        self.yandex_results: list[XmlRiverResult] = []
        self.wordstat_results: list[WordstatResult] = []
        self.serpriver_results: list[SerpRiverResult] = []
        self.serpriver_raw_response = ""

        self.google_filter_var = tk.StringVar(value="")
        self.google_sort_var = tk.StringVar(value="query")
        self.google_sort_direction_var = tk.StringVar(value=SORT_DIRECTION_VALUES[0])
        self.yandex_filter_var = tk.StringVar(value="")
        self.yandex_sort_var = tk.StringVar(value="query")
        self.yandex_sort_direction_var = tk.StringVar(value=SORT_DIRECTION_VALUES[0])
        self.wordstat_filter_var = tk.StringVar(value="")
        self.wordstat_sort_var = tk.StringVar(value="query")
        self.wordstat_sort_direction_var = tk.StringVar(value=SORT_DIRECTION_VALUES[0])
        self.serpriver_key_var = tk.StringVar(value=self.settings.serpriver_api_key)
        self.serpriver_engine_var = tk.StringVar(value="yandex")
        self.serpriver_domain_var = tk.StringVar(value="yandex.ru")
        self.serpriver_limit_var = tk.StringVar(value=str(self.settings.serpriver_max_concurrency))
        self.serpriver_result_cnt_var = tk.StringVar(value="10")
        self.serpriver_search_domain_var = tk.StringVar(value="ru")
        self.serpriver_lr_var = tk.StringVar(value="213")
        self.serpriver_location_var = tk.StringVar(value="New York,United States")
        self.serpriver_hl_var = tk.StringVar(value="en")
        self.serpriver_gl_var = tk.StringVar(value="US")
        self.serpriver_device_var = tk.StringVar(value="desktop")
        self.serpriver_output_format_var = tk.StringVar(value="json")
        self.is_serpriver_key_visible = False
        self.serpriver_controls: list[tuple[tk.Widget, str]] = []
        self.serpriver_option_entries: dict[str, ttk.Entry] = {}

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(12, 0))

        self.xmlriver_frame = ttk.Frame(self.notebook)
        self.serpriver_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.xmlriver_frame, text="  XMLRiver  ")
        self.notebook.add(self.serpriver_frame, text="  SERPRiver  ")

        self._build_xmlriver_tab()
        self._build_serpriver_tab()

        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Frame(self.root)
        status_bar.pack(fill="x", side="bottom")
        tk.Frame(status_bar, height=2, bg=ACCENT_COLOR).pack(fill="x")
        ttk.Label(
            status_bar,
            textvariable=self.status_var,
            font=APP_FONT_SMALL,
            padding=(12, 5),
        ).pack(fill="x")

    # ── Фоновый цикл ──

    def _run_background_loop(self) -> None:
        """Запускает фоновый event loop для асинхронных запросов."""
        asyncio.set_event_loop(self.background_loop)
        self.background_loop.run_forever()

    def _on_close(self) -> None:
        """Корректно завершает фоновый event loop."""
        if self.background_loop.is_running():
            self.background_loop.call_soon_threadsafe(self.background_loop.stop)
        self.root.destroy()

    # ── Стили ──

    def _configure_styles(self) -> None:
        """Настраивает стили с тёмной темой и оранжевыми акцентами."""
        style = ttk.Style()
        style.configure("Treeview", rowheight=TREE_ROW_HEIGHT, font=APP_FONT)
        style.configure("Treeview.Heading", font=APP_FONT_BOLD)
        style.map(
            "Treeview",
            background=[("selected", ACCENT_COLOR)],
            foreground=[("selected", "white")],
        )
        style.configure("TLabel", font=APP_FONT)
        style.configure("TLabelframe.Label", font=APP_FONT_BOLD, foreground=ACCENT_COLOR)
        style.configure("TButton", font=APP_FONT)
        style.configure("TNotebook.Tab", font=APP_FONT, padding=(18, 8))
        style.configure("TRadiobutton", font=APP_FONT)
        style.configure("TCheckbutton", font=APP_FONT)

    # ── Контекстное меню полей ──

    def _init_edit_menu(self) -> None:
        """Создает контекстное меню для редактируемых полей."""
        self.edit_menu = tk.Menu(self.root, tearoff=False)
        self.edit_menu.add_command(label="Вырезать", command=lambda: self._generate_edit_event("<<Cut>>"))
        self.edit_menu.add_command(label="Копировать", command=lambda: self._generate_edit_event("<<Copy>>"))
        self.edit_menu.add_command(label="Вставить", command=self._paste_into_widget)
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="Выделить всё", command=self._select_all_in_widget)
        for widget_class in ("Entry", "TEntry", "Text", "TCombobox"):
            self.root.bind_class(widget_class, "<Button-3>", self._show_edit_menu, add="+")
            self.root.bind_class(widget_class, "<Control-a>", self._select_all_shortcut, add="+")
            self.root.bind_class(widget_class, "<Control-A>", self._select_all_shortcut, add="+")
            self.root.bind_class(widget_class, "<Control-v>", self._paste_shortcut, add="+")
            self.root.bind_class(widget_class, "<Control-V>", self._paste_shortcut, add="+")

    def _select_all_shortcut(self, event: tk.Event) -> str:
        """Обрабатывает Ctrl+A для редактируемых полей."""
        self.edit_menu_widget = event.widget
        self._select_all_in_widget()
        return "break"

    def _show_edit_menu(self, event: tk.Event) -> str:
        """Показывает контекстное меню для текущего виджета."""
        self.edit_menu_widget = event.widget
        try:
            event.widget.focus_set()
            self.edit_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.edit_menu.grab_release()
        return "break"

    def _generate_edit_event(self, event_name: str) -> None:
        """Отправляет виртуальное событие редактирования в активный виджет."""
        widget = getattr(self, "edit_menu_widget", None)
        if widget is not None:
            widget.event_generate(event_name)

    def _paste_shortcut(self, event: tk.Event) -> str:
        """Обрабатывает Ctrl+V для полей ввода и combobox."""
        self.edit_menu_widget = event.widget
        self._paste_into_widget()
        return "break"

    def _paste_into_widget(self) -> None:
        """Вставляет текст из буфера обмена в текущий виджет."""
        widget = getattr(self, "edit_menu_widget", None)
        if widget is None:
            return
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            return
        if isinstance(widget, tk.Text):
            self._replace_text_selection(widget)
            widget.insert("insert", clipboard_text)
            widget.see("insert")
            return
        try:
            self._replace_entry_selection(widget)
            widget.insert("insert", clipboard_text)
            widget.icursor("end")
        except tk.TclError:
            return

    def _replace_text_selection(self, widget: tk.Text) -> None:
        """Удаляет текущее выделение в текстовом поле перед вставкой."""
        try:
            if widget.tag_ranges("sel"):
                widget.delete("sel.first", "sel.last")
        except tk.TclError:
            return

    def _replace_entry_selection(self, widget: tk.Widget) -> None:
        """Удаляет текущее выделение в однострочном поле перед вставкой."""
        try:
            if bool(widget.selection_present()):
                widget.delete("sel.first", "sel.last")
        except (AttributeError, tk.TclError):
            return

    def _select_all_in_widget(self) -> None:
        """Выделяет всё содержимое активного текстового виджета."""
        widget = getattr(self, "edit_menu_widget", None)
        if widget is None:
            return
        if isinstance(widget, tk.Text):
            widget.tag_add("sel", "1.0", "end-1c")
            widget.mark_set("insert", "1.0")
            widget.see("insert")
            return
        try:
            widget.selection_range(0, "end")
            widget.icursor("end")
        except tk.TclError:
            return

    def _bind_query_text_shortcuts(self, text_widget: tk.Text) -> None:
        """Добавляет прямые горячие клавиши для текстовых полей запросов."""
        text_widget.bind("<Control-a>", self._handle_query_select_all_shortcut, add="+")
        text_widget.bind("<Control-A>", self._handle_query_select_all_shortcut, add="+")
        text_widget.bind("<Control-KeyPress>", self._handle_query_select_all_shortcut, add="+")

    def _handle_query_select_all_shortcut(self, event: tk.Event) -> str | None:
        """Обрабатывает Ctrl+A для английской и русской раскладки."""
        keysym = (event.keysym or "").lower()
        char = (event.char or "").lower()
        if keysym not in {"a", "cyrillic_ef"} and char not in {"a", "ф"}:
            return None
        if not isinstance(event.widget, tk.Text):
            return None
        self.edit_menu_widget = event.widget
        event.widget.tag_add("sel", "1.0", "end-1c")
        event.widget.mark_set("insert", "1.0")
        event.widget.see("insert")
        return "break"

    def _build_resizable_sections(self, parent: ttk.Frame) -> tuple[ttk.LabelFrame, ttk.LabelFrame]:
        """Создает стабильные секции запросов и результатов без ломающего splitter."""
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True, pady=(10, 0))
        input_frame = ttk.LabelFrame(container, text="Запросы")
        input_frame.pack(fill="x")
        results_frame = ttk.LabelFrame(container, text="Результаты")
        results_frame.pack(fill="both", expand=True, pady=(10, 0))
        return input_frame, results_frame

    def _set_default_google_location(self) -> None:
        """Выбирает базовую локацию Russia, если она есть в справочнике."""
        saved_value = self._get_saved_reference_label("google_geo", self.google_loc_var.get())
        if saved_value:
            self.google_loc_var.set(saved_value)
            return
        labels = [option.label for option in self.reference_catalogs.get("google_geo", [])]
        for label in labels:
            if label.lower() == "russia":
                self.google_loc_var.set(label)
                return
        if labels:
            self.google_loc_var.set(labels[0])

    def _get_saved_reference_label(self, catalog_name: str, saved_label: str) -> str:
        """Возвращает сохранённый label, если он всё ещё существует в справочнике."""
        cleaned_label = saved_label.strip()
        if not cleaned_label:
            return ""
        for option_label in self.reference_maps.get(catalog_name, {}):
            if option_label.casefold() == cleaned_label.casefold():
                return option_label
        return ""

    def _get_today_date_text(self) -> str:
        """Возвращает сегодняшнюю дату в формате Wordstat."""
        return datetime.now().strftime(WORDSTAT_DATE_FORMAT)

    def _parse_date_text(self, raw_value: str) -> datetime | None:
        """Преобразует текст даты в datetime или возвращает None."""
        cleaned_value = raw_value.strip()
        if not cleaned_value:
            return None
        try:
            return datetime.strptime(cleaned_value, WORDSTAT_DATE_FORMAT)
        except ValueError:
            return None

    # ── Копирование из Treeview ──

    def _init_tree_copy_bindings(self, tree: ttk.Treeview) -> None:
        """Добавляет копирование строк для Treeview."""
        tree.bind("<Button-3>", lambda e: self._show_tree_context_menu(e, tree))
        tree.bind("<Control-c>", lambda e: self._copy_tree_selection(tree))
        tree.bind("<Control-C>", lambda e: self._copy_tree_selection(tree))

    def _show_tree_context_menu(self, event: tk.Event, tree: ttk.Treeview) -> None:
        """Показывает контекстное меню правой кнопкой для Treeview."""
        row_id = tree.identify_row(event.y)
        if row_id:
            tree.selection_set(row_id)
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Копировать строку (Ctrl+C)", command=lambda: self._copy_tree_selection(tree))
        menu.add_command(label="Копировать все строки", command=lambda: self._copy_tree_all(tree))
        menu.add_separator()
        menu.add_command(label="Выделить всё", command=lambda: tree.selection_set(tree.get_children()))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy_tree_selection(self, tree: ttk.Treeview) -> None:
        """Копирует выделенные строки Treeview в буфер обмена."""
        selected = tree.selection()
        if not selected:
            return
        columns = tree["columns"]
        header = "\t".join(columns)
        rows: list[str] = []
        for item_id in selected:
            values = tree.item(item_id, "values")
            rows.append("\t".join(str(v) for v in values))
        text = header + "\n" + "\n".join(rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status(f"Скопировано строк: {len(rows)}")

    def _copy_tree_all(self, tree: ttk.Treeview) -> None:
        """Копирует все строки Treeview в буфер обмена."""
        children = tree.get_children()
        if not children:
            return
        columns = tree["columns"]
        header = "\t".join(columns)
        rows: list[str] = []
        for item_id in children:
            values = tree.item(item_id, "values")
            rows.append("\t".join(str(v) for v in values))
        text = header + "\n" + "\n".join(rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status(f"Скопировано все строки: {len(rows)}")

    # ── Вкладка XMLRiver ──

    def _build_xmlriver_tab(self) -> None:
        """Создает вкладку XMLRiver с внутренними подвкладками."""
        access_frame = ttk.LabelFrame(self.xmlriver_frame, text="Доступ XMLRiver")
        access_frame.pack(fill="x")

        ttk.Label(access_frame, text="User ID").grid(row=0, column=0, sticky="w")
        xml_user_entry = ttk.Entry(access_frame, textvariable=self.xmlriver_user_var, width=20)
        xml_user_entry.grid(row=0, column=1, padx=(8, 12), sticky="we")
        ttk.Label(access_frame, text="API Key").grid(row=0, column=2, sticky="w")
        self.xmlriver_key_entry = ttk.Entry(access_frame, textvariable=self.xmlriver_key_var, width=30, show="*")
        self.xmlriver_key_entry.grid(row=0, column=3, padx=(8, 12), sticky="we")
        xml_toggle_button = ttk.Button(access_frame, text="👁", width=3, command=self._toggle_xmlriver_key, bootstyle="dark")
        xml_toggle_button.grid(row=0, column=4, sticky="w")
        xml_save_button = ttk.Button(access_frame, text="Сохранить", command=self._save_xmlriver_credentials, bootstyle="warning-outline")
        xml_save_button.grid(row=0, column=5, sticky="e", padx=(12, 0))

        self.xml_shared_controls = [
            (xml_user_entry, "normal"),
            (self.xmlriver_key_entry, "normal"),
            (xml_toggle_button, "normal"),
            (xml_save_button, "normal"),
        ]

        tk.Frame(self.xmlriver_frame, height=2, bg=ACCENT_COLOR).pack(fill="x", pady=(10, 0))

        self.xmlriver_notebook = ttk.Notebook(self.xmlriver_frame)
        self.xmlriver_notebook.pack(fill="both", expand=True, pady=(10, 0))

        self.google_frame = ttk.Frame(self.xmlriver_notebook)
        self.yandex_frame = ttk.Frame(self.xmlriver_notebook)
        self.wordstat_frame = ttk.Frame(self.xmlriver_notebook)

        self.xmlriver_notebook.add(self.google_frame, text="  Google  ")
        self.xmlriver_notebook.add(self.yandex_frame, text="  Yandex  ")
        self.xmlriver_notebook.add(self.wordstat_frame, text="  Wordstat  ")

        self._build_google_tab()
        self._build_yandex_tab()
        self._build_wordstat_tab()

    # ── Подвкладка Google ──

    def _build_google_tab(self) -> None:
        """Создает подвкладку Google."""
        self.google_warning_label = ttk.Label(self.google_frame, text="", bootstyle="danger")
        self.google_warning_label.pack(anchor="w", pady=(0, 6))

        params_frame = ttk.LabelFrame(self.google_frame, text="Параметры Google")
        params_frame.pack(fill="x")
        self.google_loc_combo = self._add_reference_combobox(
            params_frame,
            "Локация",
            self.google_loc_var,
            "google_geo",
            0,
            0,
            28,
            self.google_controls,
            searchable=True,
            on_commit=lambda _value: self._save_ui_settings(),
        )
        self._set_default_google_location()
        self.google_device_combo = self._add_static_combobox(
            params_frame, "Устройство", self.google_device_var, DEVICE_VALUES, 0, 2, 16, self.google_controls, searchable=True,
        )
        self.google_country_entry = self._add_fixed_entry(
            params_frame, "Страна", self.google_country_var, 1, 0, 24, self.google_controls,
        )
        self.google_lr_entry = self._add_fixed_entry(
            params_frame, "Язык", self.google_lr_var, 1, 2, 20, self.google_controls,
        )
        self.google_domain_entry = self._add_fixed_entry(
            params_frame, "Домен Google", self.google_domain_var, 2, 0, 16, self.google_controls,
        )

        input_frame, results_frame = self._build_resizable_sections(self.google_frame)
        self.google_queries_text = tk.Text(input_frame, height=6, wrap="word", bg="#2b3e50", fg="white",
                                           insertbackground="white", selectbackground=ACCENT_COLOR, font=APP_FONT)
        self.google_queries_text.pack(fill="both", expand=True)
        self._bind_query_text_shortcuts(self.google_queries_text)

        actions_frame = ttk.Frame(input_frame)
        actions_frame.pack(fill="x", pady=(8, 0))
        self.google_import_csv_button = ttk.Button(actions_frame, text="Импорт CSV", command=self._import_google_csv, bootstyle="secondary-outline")
        self.google_import_csv_button.pack(side="left")
        self.google_import_xlsx_button = ttk.Button(actions_frame, text="Импорт XLSX", command=self._import_google_xlsx, bootstyle="secondary-outline")
        self.google_import_xlsx_button.pack(side="left", padx=(8, 0))
        self.google_start_button = ttk.Button(
            actions_frame, text="▶  Запустить", command=self._start_google_xmlriver, bootstyle="warning",
        )
        self.google_start_button.pack(side="left", padx=(16, 0))
        self.google_export_csv_button = ttk.Button(actions_frame, text="Экспорт CSV", command=self._export_google_csv, bootstyle="info-outline")
        self.google_export_csv_button.pack(side="right")
        self.google_export_xlsx_button = ttk.Button(actions_frame, text="Экспорт XLSX", command=self._export_google_xlsx, bootstyle="info-outline")
        self.google_export_xlsx_button.pack(side="right", padx=(0, 8))

        progress_frame = ttk.Frame(self.google_frame)
        progress_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(progress_frame, text="Прогресс обработки").pack(anchor="w")
        self.google_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=1, value=0, bootstyle="warning-striped")
        self.google_progress.pack(fill="x", pady=(4, 0))

        self._build_result_tools(
            results_frame, self.google_filter_var, self.google_sort_var, self.google_sort_direction_var,
            ("query", "position", "url", "domain", "title", "snippet"), self._apply_google_view, "google",
        )
        self.google_tree = self._build_scrolled_tree(
            results_frame,
            ("query", "position", "url", "domain", "title", "snippet"),
            (("query", "Запрос"), ("position", "Позиция"), ("url", "URL"), ("domain", "Домен"), ("title", "Заголовок"), ("snippet", "Сниппет")),
        )

        self._register_control(self.google_controls, self.google_queries_text, "normal")
        for widget in (
            self.google_import_csv_button, self.google_import_xlsx_button, self.google_start_button,
            self.google_export_csv_button, self.google_export_xlsx_button,
        ):
            self._register_control(self.google_controls, widget)
        self._apply_catalog_requirements(
            self.google_warning_label, self.google_controls, self.google_required_catalogs, "google",
        )

    # ── Подвкладка Yandex ──

    def _build_yandex_tab(self) -> None:
        """Создает подвкладку Yandex."""
        self.yandex_warning_label = ttk.Label(self.yandex_frame, text="", bootstyle="danger")
        self.yandex_warning_label.pack(anchor="w", pady=(0, 6))

        params_frame = ttk.LabelFrame(self.yandex_frame, text="Параметры Yandex")
        params_frame.pack(fill="x")
        self.yandex_lr_combo = self._add_reference_combobox(
            params_frame,
            "Регион",
            self.yandex_lr_var,
            "yandex_geo",
            0,
            0,
            28,
            self.yandex_controls,
            searchable=True,
            on_commit=lambda _value: self._save_ui_settings(),
        )
        self.yandex_domain_combo = self._add_static_combobox(
            params_frame, "Домен Яндекса", self.yandex_domain_var, YANDEX_DOMAIN_VALUES, 0, 2, 16, self.yandex_controls, searchable=True,
        )
        ttk.Label(params_frame, text="Язык").grid(row=1, column=0, sticky="w")
        self.yandex_lang_entry = ttk.Entry(params_frame, textvariable=self.yandex_lang_var, width=16)
        self.yandex_lang_entry.grid(row=1, column=1, padx=(8, 14), sticky="we")
        self._register_control(self.yandex_controls, self.yandex_lang_entry)
        self.yandex_device_combo = self._add_static_combobox(
            params_frame, "Устройство", self.yandex_device_var, DEVICE_VALUES, 1, 2, 16, self.yandex_controls, searchable=True,
        )

        input_frame, results_frame = self._build_resizable_sections(self.yandex_frame)
        self.yandex_queries_text = tk.Text(input_frame, height=6, wrap="word", bg="#2b3e50", fg="white",
                                           insertbackground="white", selectbackground=ACCENT_COLOR, font=APP_FONT)
        self.yandex_queries_text.pack(fill="both", expand=True)
        self._bind_query_text_shortcuts(self.yandex_queries_text)

        actions_frame = ttk.Frame(input_frame)
        actions_frame.pack(fill="x", pady=(8, 0))
        self.yandex_import_csv_button = ttk.Button(actions_frame, text="Импорт CSV", command=self._import_yandex_csv, bootstyle="secondary-outline")
        self.yandex_import_csv_button.pack(side="left")
        self.yandex_import_xlsx_button = ttk.Button(actions_frame, text="Импорт XLSX", command=self._import_yandex_xlsx, bootstyle="secondary-outline")
        self.yandex_import_xlsx_button.pack(side="left", padx=(8, 0))
        self.yandex_start_button = ttk.Button(
            actions_frame, text="▶  Запустить", command=self._start_yandex_xmlriver, bootstyle="warning",
        )
        self.yandex_start_button.pack(side="left", padx=(16, 0))
        self.yandex_export_csv_button = ttk.Button(actions_frame, text="Экспорт CSV", command=self._export_yandex_csv, bootstyle="info-outline")
        self.yandex_export_csv_button.pack(side="right")
        self.yandex_export_xlsx_button = ttk.Button(actions_frame, text="Экспорт XLSX", command=self._export_yandex_xlsx, bootstyle="info-outline")
        self.yandex_export_xlsx_button.pack(side="right", padx=(0, 8))

        progress_frame = ttk.Frame(self.yandex_frame)
        progress_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(progress_frame, text="Прогресс обработки").pack(anchor="w")
        self.yandex_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=1, value=0, bootstyle="warning-striped")
        self.yandex_progress.pack(fill="x", pady=(4, 0))

        self._build_result_tools(
            results_frame, self.yandex_filter_var, self.yandex_sort_var, self.yandex_sort_direction_var,
            ("query", "position", "url", "domain", "title", "snippet"), self._apply_yandex_view, "yandex",
        )
        self.yandex_tree = self._build_scrolled_tree(
            results_frame,
            ("query", "position", "url", "domain", "title", "snippet"),
            (("query", "Запрос"), ("position", "Позиция"), ("url", "URL"), ("domain", "Домен"), ("title", "Заголовок"), ("snippet", "Сниппет")),
        )

        self._register_control(self.yandex_controls, self.yandex_queries_text, "normal")
        for widget in (
            self.yandex_import_csv_button, self.yandex_import_xlsx_button, self.yandex_start_button,
            self.yandex_export_csv_button, self.yandex_export_xlsx_button,
        ):
            self._register_control(self.yandex_controls, widget)
        self._apply_catalog_requirements(
            self.yandex_warning_label, self.yandex_controls, self.yandex_required_catalogs, "yandex",
        )

    # ── Подвкладка Wordstat ──

    def _build_wordstat_tab(self) -> None:
        """Создает подвкладку Wordstat."""
        self.wordstat_warning_label = ttk.Label(self.wordstat_frame, text="", bootstyle="danger")
        self.wordstat_warning_label.pack(anchor="w", pady=(0, 6))

        params_frame = ttk.LabelFrame(self.wordstat_frame, text="Параметры Wordstat")
        params_frame.pack(fill="x")
        self.wordstat_region_combo = self._add_reference_combobox(
            params_frame,
            "Основной регион",
            self.wordstat_region_var,
            "yandex_geo",
            0,
            0,
            28,
            self.wordstat_controls,
            searchable=True,
            on_commit=lambda _value: self._save_ui_settings(),
        )
        ttk.Label(params_frame, text="Доп. регионы (id)").grid(row=0, column=2, sticky="w")
        self.wordstat_regions_extra_entry = ttk.Entry(params_frame, textvariable=self.wordstat_regions_extra_var, width=24)
        self.wordstat_regions_extra_entry.grid(row=0, column=3, padx=(8, 14), sticky="we")
        self._register_control(self.wordstat_controls, self.wordstat_regions_extra_entry)
        self.wordstat_device_combo = self._add_static_combobox(
            params_frame, "Устройство", self.wordstat_device_var, WORDSTAT_DEVICE_VALUES, 1, 0, 24, self.wordstat_controls, searchable=True,
        )
        self.wordstat_period_combo = self._add_static_combobox(
            params_frame, "Группировка", self.wordstat_period_var, WORDSTAT_PERIOD_VALUES, 1, 2, 18, self.wordstat_controls, searchable=True,
        )
        self.wordstat_start_entry = self._add_date_entry(
            params_frame, "Дата начала", self.wordstat_start_var, 2, 0, 18, self.wordstat_controls,
        )
        self.wordstat_end_entry = self._add_date_entry(
            params_frame, "Дата окончания", self.wordstat_end_var, 2, 2, 18, self.wordstat_controls,
        )
        self.wordstat_pagetype_combo = self._add_static_combobox(
            params_frame, "Тип данных", self.wordstat_pagetype_var, WORDSTAT_PAGETYPE_VALUES, 3, 0, 18, self.wordstat_controls, searchable=True,
        )

        input_frame, results_frame = self._build_resizable_sections(self.wordstat_frame)
        self.wordstat_queries_text = tk.Text(input_frame, height=6, wrap="word", bg="#2b3e50", fg="white",
                                             insertbackground="white", selectbackground=ACCENT_COLOR, font=APP_FONT)
        self.wordstat_queries_text.pack(fill="both", expand=True)
        self._bind_query_text_shortcuts(self.wordstat_queries_text)

        actions_frame = ttk.Frame(input_frame)
        actions_frame.pack(fill="x", pady=(8, 0))
        self.wordstat_import_csv_button = ttk.Button(actions_frame, text="Импорт CSV", command=self._import_wordstat_csv, bootstyle="secondary-outline")
        self.wordstat_import_csv_button.pack(side="left")
        self.wordstat_import_xlsx_button = ttk.Button(actions_frame, text="Импорт XLSX", command=self._import_wordstat_xlsx, bootstyle="secondary-outline")
        self.wordstat_import_xlsx_button.pack(side="left", padx=(8, 0))
        self.wordstat_start_button = ttk.Button(
            actions_frame, text="▶  Запустить", command=self._start_wordstat, bootstyle="warning",
        )
        self.wordstat_start_button.pack(side="left", padx=(16, 0))
        self.wordstat_export_csv_button = ttk.Button(actions_frame, text="Экспорт CSV", command=self._export_wordstat_csv, bootstyle="info-outline")
        self.wordstat_export_csv_button.pack(side="right")
        self.wordstat_export_xlsx_button = ttk.Button(actions_frame, text="Экспорт XLSX", command=self._export_wordstat_xlsx, bootstyle="info-outline")
        self.wordstat_export_xlsx_button.pack(side="right", padx=(0, 8))

        progress_frame = ttk.Frame(self.wordstat_frame)
        progress_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(progress_frame, text="Прогресс обработки").pack(anchor="w")
        self.wordstat_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=1, value=0, bootstyle="warning-striped")
        self.wordstat_progress.pack(fill="x", pady=(4, 0))

        self._build_result_tools(
            results_frame, self.wordstat_filter_var, self.wordstat_sort_var, self.wordstat_sort_direction_var,
            ("query", "result_type", "phrase", "value"), self._apply_wordstat_view, "wordstat",
        )
        self.wordstat_tree = self._build_scrolled_tree(
            results_frame,
            ("query", "result_type", "phrase", "value"),
            (("query", "Запрос"), ("result_type", "Тип"), ("phrase", "Фраза"), ("value", "Значение")),
        )

        self._register_control(self.wordstat_controls, self.wordstat_queries_text, "normal")
        for widget in (
            self.wordstat_import_csv_button, self.wordstat_import_xlsx_button, self.wordstat_start_button,
            self.wordstat_export_csv_button, self.wordstat_export_xlsx_button,
        ):
            self._register_control(self.wordstat_controls, widget)
        self._apply_catalog_requirements(
            self.wordstat_warning_label, self.wordstat_controls, self.wordstat_required_catalogs, "wordstat",
        )

    # ── Вкладка SERPRiver ──

    def _build_serpriver_tab(self) -> None:
        """Создает вкладку SERPRiver."""
        access_frame = ttk.LabelFrame(self.serpriver_frame, text="Настройки доступа")
        access_frame.pack(fill="x")
        ttk.Label(access_frame, text="API Key").grid(row=0, column=0, sticky="w")
        self.serpriver_key_entry = ttk.Entry(access_frame, textvariable=self.serpriver_key_var, width=28, show="*")
        self.serpriver_key_entry.grid(row=0, column=1, padx=(8, 12), sticky="we")
        serp_toggle_button = ttk.Button(access_frame, text="👁", width=3, command=self._toggle_serpriver_key, bootstyle="dark")
        serp_toggle_button.grid(row=0, column=2, sticky="w")
        ttk.Label(access_frame, text="Домен").grid(row=0, column=3, sticky="w", padx=(12, 0))
        serp_domain_entry = ttk.Entry(access_frame, textvariable=self.serpriver_domain_var, width=24)
        serp_domain_entry.grid(row=0, column=4, padx=(8, 12), sticky="we")
        ttk.Label(access_frame, text="Потоки").grid(row=0, column=5, sticky="w")
        serp_limit_entry = ttk.Entry(access_frame, textvariable=self.serpriver_limit_var, width=8)
        serp_limit_entry.grid(row=0, column=6, padx=(8, 12), sticky="w")
        serp_save_button = ttk.Button(access_frame, text="Сохранить", command=self._save_serpriver_credentials, bootstyle="warning-outline")
        serp_save_button.grid(row=0, column=7, sticky="e")

        tk.Frame(self.serpriver_frame, height=2, bg=ACCENT_COLOR).pack(fill="x", pady=(10, 0))

        engine_frame = ttk.LabelFrame(self.serpriver_frame, text="Поисковик")
        engine_frame.pack(fill="x", pady=(10, 0))
        self.serpriver_yandex_radio = ttk.Radiobutton(
            engine_frame, text="Яндекс", value="yandex", variable=self.serpriver_engine_var,
            command=self._update_serpriver_engine_fields, bootstyle="warning",
        )
        self.serpriver_yandex_radio.pack(side="left")
        self.serpriver_google_radio = ttk.Radiobutton(
            engine_frame, text="Google", value="google", variable=self.serpriver_engine_var,
            command=self._update_serpriver_engine_fields, bootstyle="warning",
        )
        self.serpriver_google_radio.pack(side="left", padx=(12, 0))

        options_frame = ttk.LabelFrame(self.serpriver_frame, text="Параметры поиска")
        options_frame.pack(fill="x", pady=(10, 0))
        self._add_serpriver_option(options_frame, "result_cnt", self.serpriver_result_cnt_var, 0, 0, 8)
        self._add_serpriver_option(options_frame, "domain", self.serpriver_search_domain_var, 0, 2, 12)
        self._add_serpriver_option(options_frame, "device", self.serpriver_device_var, 0, 4, 12)
        self._add_serpriver_option(options_frame, "lr", self.serpriver_lr_var, 1, 0, 10)
        self._add_serpriver_option(options_frame, "location", self.serpriver_location_var, 1, 2, 26)
        self._add_serpriver_option(options_frame, "hl", self.serpriver_hl_var, 2, 0, 8)
        self._add_serpriver_option(options_frame, "gl", self.serpriver_gl_var, 2, 2, 8)
        self.serpriver_output_format_combo = self._add_static_combobox(
            options_frame,
            SERPRIVER_OPTION_LABELS["output_format"],
            self.serpriver_output_format_var,
            SERPRIVER_OUTPUT_FORMAT_VALUES,
            3,
            0,
            12,
            [],
            searchable=False,
        )

        input_frame, results_frame = self._build_resizable_sections(self.serpriver_frame)
        self.serpriver_queries_text = tk.Text(input_frame, height=6, wrap="word", bg="#2b3e50", fg="white",
                                              insertbackground="white", selectbackground=ACCENT_COLOR, font=APP_FONT)
        self.serpriver_queries_text.pack(fill="both", expand=True)
        self._bind_query_text_shortcuts(self.serpriver_queries_text)

        actions_frame = ttk.Frame(input_frame)
        actions_frame.pack(fill="x", pady=(8, 0))
        self.serpriver_import_csv_button = ttk.Button(actions_frame, text="Импорт CSV", command=self._import_serpriver_csv, bootstyle="secondary-outline")
        self.serpriver_import_csv_button.pack(side="left")
        self.serpriver_import_xlsx_button = ttk.Button(actions_frame, text="Импорт XLSX", command=self._import_serpriver_xlsx, bootstyle="secondary-outline")
        self.serpriver_import_xlsx_button.pack(side="left", padx=(8, 0))
        self.serpriver_start_button = ttk.Button(
            actions_frame, text="▶  Запустить", command=self._start_serpriver, bootstyle="warning",
        )
        self.serpriver_start_button.pack(side="left", padx=(16, 0))
        self.serpriver_export_button = ttk.Button(
            actions_frame,
            text="Экспорт",
            command=self._export_serpriver_raw_response,
            bootstyle="info-outline",
        )
        self.serpriver_export_button.pack(side="right")

        progress_frame = ttk.Frame(self.serpriver_frame)
        progress_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(progress_frame, text="Прогресс обработки").pack(anchor="w")
        self.serpriver_progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=1, value=0, bootstyle="warning-striped")
        self.serpriver_progress.pack(fill="x", pady=(4, 0))

        code_frame = ttk.Frame(results_frame)
        code_frame.pack(fill="both", expand=True)
        response_container = ttk.Frame(code_frame)
        response_container.pack(fill="both", expand=True)
        response_scrollbar_y = ttk.Scrollbar(
            response_container,
            orient="vertical",
            command=lambda *args: self.serpriver_response_text.yview(*args),
            bootstyle="warning-round",
        )
        response_scrollbar_x = ttk.Scrollbar(
            response_container,
            orient="horizontal",
            command=lambda *args: self.serpriver_response_text.xview(*args),
            bootstyle="warning-round",
        )
        self.serpriver_response_text = tk.Text(
            response_container,
            wrap="none",
            bg="#1f1f1f",
            fg="#f3f3f3",
            insertbackground="#f3f3f3",
            selectbackground=ACCENT_COLOR,
            font=("Consolas", 10),
            yscrollcommand=response_scrollbar_y.set,
            xscrollcommand=response_scrollbar_x.set,
        )
        response_scrollbar_y.pack(side="right", fill="y")
        response_scrollbar_x.pack(side="bottom", fill="x")
        self.serpriver_response_text.pack(side="left", fill="both", expand=True)
        self.serpriver_response_text.configure(state="disabled")

        self.serpriver_controls = [
            (self.serpriver_key_entry, "normal"),
            (serp_toggle_button, "normal"),
            (serp_domain_entry, "normal"),
            (serp_limit_entry, "normal"),
            (serp_save_button, "normal"),
            (self.serpriver_queries_text, "normal"),
            (self.serpriver_yandex_radio, "normal"),
            (self.serpriver_google_radio, "normal"),
            (self.serpriver_output_format_combo, "readonly"),
            (self.serpriver_import_csv_button, "normal"),
            (self.serpriver_import_xlsx_button, "normal"),
            (self.serpriver_start_button, "normal"),
            (self.serpriver_export_button, "normal"),
        ] + [(entry, "normal") for entry in self.serpriver_option_entries.values()]
        self._update_serpriver_engine_fields()

    # ── Построение таблиц ──

    def _build_scrolled_tree(
        self,
        parent: ttk.Frame,
        columns: tuple[str, ...],
        headings: tuple[tuple[str, str], ...],
        height: int = 12,
    ) -> ttk.Treeview:
        """Создает Treeview со скроллбаром и контекстным меню копирования."""
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        tree = ttk.Treeview(container, columns=columns, show="headings", height=height, selectmode="extended")
        scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview, bootstyle="warning-round")
        scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=tree.xview, bootstyle="warning-round")
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        for col_name, title in headings:
            tree.heading(col_name, text=title)
            tree.column(col_name, width=180, anchor="w")
        tree.tag_configure(TREE_ERROR_TAG, foreground="#ff6b6b")

        self._init_tree_copy_bindings(tree)
        return tree

    # ── Панель фильтров ──

    def _build_result_tools(
        self,
        parent: ttk.LabelFrame,
        filter_var: tk.StringVar,
        sort_var: tk.StringVar,
        direction_var: tk.StringVar,
        sort_values: tuple[str, ...],
        apply_command: Any,
        debounce_key: str = "default",
    ) -> None:
        """Создает панель фильтрации и сортировки с реактивным обновлением."""
        tools_frame = ttk.Frame(parent)
        tools_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(tools_frame, text="🔍 Фильтр").pack(side="left")
        filter_entry = ttk.Entry(tools_frame, textvariable=filter_var, width=24)
        filter_entry.pack(side="left", padx=(6, 10))
        filter_entry.bind("<KeyRelease>", lambda e: self._debounced_apply(debounce_key, apply_command))

        ttk.Label(tools_frame, text="Поле").pack(side="left")
        sort_combo = ttk.Combobox(tools_frame, textvariable=sort_var, values=sort_values, width=14, state="readonly")
        sort_combo.pack(side="left", padx=(6, 10))
        sort_combo.bind("<<ComboboxSelected>>", lambda e: apply_command())

        ttk.Label(tools_frame, text="Порядок").pack(side="left")
        direction_combo = ttk.Combobox(tools_frame, textvariable=direction_var, values=SORT_DIRECTION_VALUES, width=16, state="readonly")
        direction_combo.pack(side="left", padx=(6, 10))
        direction_combo.bind("<<ComboboxSelected>>", lambda e: apply_command())

        ttk.Button(tools_frame, text="Применить", command=apply_command, bootstyle="warning-outline").pack(side="left")
        ttk.Button(
            tools_frame, text="Сбросить",
            command=lambda: self._reset_result_tools(filter_var, sort_var, direction_var, apply_command),
            bootstyle="secondary",
        ).pack(side="left", padx=(6, 0))

    def _debounced_apply(self, key: str, apply_func: Any) -> None:
        """Задержка фильтрации для плавной работы при вводе."""
        if key in self._filter_debounce_ids:
            self.root.after_cancel(self._filter_debounce_ids[key])
        self._filter_debounce_ids[key] = self.root.after(FILTER_DEBOUNCE_MS, apply_func)

    def _reset_result_tools(
        self,
        filter_var: tk.StringVar,
        sort_var: tk.StringVar,
        direction_var: tk.StringVar,
        apply_command: Any,
    ) -> None:
        """Сбрасывает фильтр и сортировку таблицы."""
        filter_var.set("")
        sort_var.set("query")
        direction_var.set(SORT_DIRECTION_VALUES[0])
        apply_command()

    # ── Combobox-хелперы ──

    def _add_reference_combobox(
        self,
        parent: ttk.LabelFrame,
        label_text: str,
        value_var: tk.StringVar,
        catalog_name: str,
        row_index: int,
        column_index: int,
        width: int,
        controls: list[tuple[tk.Widget, str]],
        searchable: bool = False,
        on_commit: Any = None,
    ) -> ttk.Combobox:
        """Добавляет combobox на базе справочника."""
        ttk.Label(parent, text=label_text).grid(row=row_index, column=column_index, sticky="w")
        labels = [option.label for option in self.reference_catalogs.get(catalog_name, [])]
        if labels and not searchable:
            value_var.set(labels[0])
        combobox_state = "normal" if searchable else "readonly"
        initial_values = labels[:COMBOBOX_MAX_VALUES] if searchable else labels
        combobox = ttk.Combobox(parent, textvariable=value_var, values=initial_values, width=width, state=combobox_state)
        combobox.grid(row=row_index, column=column_index + 1, padx=(8, 14), sticky="we")
        if searchable:
            self._attach_combobox_autocomplete(combobox, value_var, labels, on_commit=on_commit)
        self._register_control(controls, combobox, combobox_state)
        return combobox

    def _add_fixed_entry(
        self,
        parent: ttk.LabelFrame,
        label_text: str,
        value_var: tk.StringVar,
        row_index: int,
        column_index: int,
        width: int,
        controls: list[tuple[tk.Widget, str]],
    ) -> ttk.Entry:
        """Добавляет read-only поле с фиксированным значением."""
        ttk.Label(parent, text=label_text).grid(row=row_index, column=column_index, sticky="w")
        entry = ttk.Entry(parent, textvariable=value_var, width=width, state="readonly")
        entry.grid(row=row_index, column=column_index + 1, padx=(8, 14), sticky="we")
        self._register_control(controls, entry, "readonly")
        return entry

    def _add_date_entry(
        self,
        parent: ttk.LabelFrame,
        label_text: str,
        value_var: tk.StringVar,
        row_index: int,
        column_index: int,
        width: int,
        controls: list[tuple[tk.Widget, str]],
    ) -> ttk.DateEntry:
        """Добавляет поле даты с мини-календарём."""
        ttk.Label(parent, text=label_text).grid(row=row_index, column=column_index, sticky="w")
        start_date = self._parse_date_text(value_var.get()) or datetime.now()
        date_entry = ttk.DateEntry(
            parent,
            dateformat=WORDSTAT_DATE_FORMAT,
            firstweekday=0,
            startdate=start_date,
            width=width,
            bootstyle="warning",
        )
        date_entry.grid(row=row_index, column=column_index + 1, padx=(8, 14), sticky="we")
        self._sync_date_entry(date_entry, value_var, start_date)
        date_entry.bind(
            "<<DateEntrySelected>>",
            lambda _event, widget=date_entry, var=value_var: self._sync_date_entry(widget, var),
            add="+",
        )
        date_entry.entry.bind(
            "<FocusOut>",
            lambda _event, widget=date_entry, var=value_var: self._sync_date_entry(widget, var),
            add="+",
        )
        self._register_control(controls, date_entry)
        return date_entry

    def _add_static_combobox(
        self,
        parent: ttk.LabelFrame,
        label_text: str,
        value_var: tk.StringVar,
        values: tuple[str, ...],
        row_index: int,
        column_index: int,
        width: int,
        controls: list[tuple[tk.Widget, str]],
        searchable: bool = False,
        on_commit: Any = None,
    ) -> ttk.Combobox:
        """Добавляет статический combobox."""
        ttk.Label(parent, text=label_text).grid(row=row_index, column=column_index, sticky="w")
        if values and not value_var.get():
            value_var.set(values[0])
        items = list(values)
        combobox_state = "normal" if searchable else "readonly"
        initial_values = items[:COMBOBOX_MAX_VALUES] if searchable else items
        combobox = ttk.Combobox(parent, textvariable=value_var, values=initial_values, width=width, state=combobox_state)
        combobox.grid(row=row_index, column=column_index + 1, padx=(8, 14), sticky="we")
        if searchable:
            self._attach_combobox_autocomplete(combobox, value_var, items, on_commit=on_commit)
        self._register_control(controls, combobox, combobox_state)
        return combobox

    def _attach_combobox_autocomplete(
        self,
        combobox: ttk.Combobox,
        value_var: tk.StringVar,
        items: list[str],
        on_commit: Any = None,
    ) -> None:
        """Подключает модульное автодополнение к combobox."""
        autocomplete = ComboboxAutocomplete(
            self.root,
            combobox,
            value_var,
            items,
            max_values=COMBOBOX_MAX_VALUES,
            debounce_ms=AUTOCOMPLETE_DEBOUNCE_MS,
            on_commit=on_commit,
        )
        self.autocomplete_helpers.append(autocomplete)

    def _sync_date_entry(
        self,
        date_entry: ttk.DateEntry,
        value_var: tk.StringVar,
        parsed_date: datetime | None = None,
    ) -> None:
        """Синхронизирует DateEntry и StringVar в каноническом формате."""
        normalized_date = parsed_date or self._parse_date_text(date_entry.entry.get()) or datetime.now()
        normalized_value = normalized_date.strftime(WORDSTAT_DATE_FORMAT)
        value_var.set(normalized_value)
        date_entry.set_date(normalized_date)

    def _add_serpriver_option(
        self,
        parent: ttk.LabelFrame,
        option_name: str,
        value_var: tk.StringVar,
        row_index: int,
        column_index: int,
        entry_width: int,
    ) -> None:
        """Добавляет поле параметра SERPRiver."""
        label_text = SERPRIVER_OPTION_LABELS.get(option_name, option_name)
        ttk.Label(parent, text=label_text).grid(row=row_index, column=column_index, sticky="w")
        entry = ttk.Entry(parent, textvariable=value_var, width=entry_width)
        entry.grid(row=row_index, column=column_index + 1, padx=(8, 14), sticky="we")
        self.serpriver_option_entries[option_name] = entry

    def _register_control(self, controls: list[tuple[tk.Widget, str]], widget: tk.Widget, normal_state: str = "normal") -> None:
        """Регистрирует виджет для временной блокировки."""
        controls.append((widget, normal_state))

    def _apply_catalog_requirements(
        self,
        warning_label: ttk.Label,
        controls: list[tuple[tk.Widget, str]],
        required_catalogs: tuple[str, ...],
        tab_name: str,
    ) -> None:
        """Показывает ошибки загрузки справочников и блокирует подвкладку."""
        missing = [catalog_name for catalog_name in required_catalogs if catalog_name in self.reference_errors]
        if not missing:
            warning_label.configure(text="")
            return
        warning_text = "⚠ Не загружены справочники: " + "; ".join(self.reference_errors[name] for name in missing)
        warning_label.configure(text=warning_text)
        self._set_controls_state(controls, disabled=True)
        if tab_name == "google":
            self.is_google_blocked = True
        elif tab_name == "yandex":
            self.is_yandex_blocked = True
        else:
            self.is_wordstat_blocked = True

    # ── Переключатели видимости ──

    def _toggle_xmlriver_key(self) -> None:
        """Переключает видимость ключа XMLRiver."""
        self.is_xmlriver_key_visible = not self.is_xmlriver_key_visible
        self.xmlriver_key_entry.configure(show="" if self.is_xmlriver_key_visible else "*")

    def _toggle_serpriver_key(self) -> None:
        """Переключает видимость ключа SERPRiver."""
        self.is_serpriver_key_visible = not self.is_serpriver_key_visible
        self.serpriver_key_entry.configure(show="" if self.is_serpriver_key_visible else "*")

    # ── Сохранение настроек ──

    def _build_app_settings(self) -> AppSettings:
        """Собирает полный снимок настроек приложения для сохранения."""
        return AppSettings(
            XMLRIVER_USER_ID=self.xmlriver_user_var.get().strip(),
            XMLRIVER_API_KEY=self.xmlriver_key_var.get().strip(),
            SERPRIVER_API_KEY=self.serpriver_key_var.get().strip(),
            GOOGLE_LAST_LOCATION_LABEL=self.google_loc_var.get().strip(),
            YANDEX_LAST_REGION_LABEL=self.yandex_lr_var.get().strip(),
            WORDSTAT_LAST_REGION_LABEL=self.wordstat_region_var.get().strip(),
            request_connect_timeout=self.settings.request_connect_timeout,
            request_read_timeout=self.settings.request_read_timeout,
            xmlriver_max_concurrency=self.settings.xmlriver_max_concurrency,
            serpriver_max_concurrency=self.settings.serpriver_max_concurrency,
            import_row_limit=self.settings.import_row_limit,
        )

    def _save_ui_settings(self) -> None:
        """Тихо сохраняет выбранные локации, если они изменились."""
        has_changes = any(
            [
                self.settings.google_last_location_label != self.google_loc_var.get().strip(),
                self.settings.yandex_last_region_label != self.yandex_lr_var.get().strip(),
                self.settings.wordstat_last_region_label != self.wordstat_region_var.get().strip(),
            ]
        )
        if not has_changes:
            return
        try:
            save_settings(self._build_app_settings())
        except OSError as error:
            logger.error("UI settings save failed | error={}", str(error))
            self._set_status("Не удалось сохранить выбранные локации")
            return
        self.settings = get_settings()

    def _save_xmlriver_credentials(self) -> None:
        """Сохраняет учетные данные XMLRiver в .env."""
        save_settings(self._build_app_settings())
        self.settings = get_settings()
        self._set_status("Данные XMLRiver сохранены")
        messagebox.showinfo("Сохранение", "Данные XMLRiver сохранены в .env")

    def _save_serpriver_credentials(self) -> None:
        """Сохраняет ключ SERPRiver в .env."""
        save_settings(self._build_app_settings())
        self.settings = get_settings()
        self._set_status("Данные SERPRiver сохранены")
        messagebox.showinfo("Сохранение", "Данные SERPRiver сохранены в .env")

    def _update_serpriver_engine_fields(self) -> None:
        """Переключает доступность полей SERPRiver по типу поисковика."""
        is_google = self.serpriver_engine_var.get() == "google"
        current_domain = self.serpriver_search_domain_var.get().strip().lower()
        if current_domain in {"", "ru", "com"}:
            self.serpriver_search_domain_var.set("com" if is_google else "ru")
        self.serpriver_option_entries["lr"].configure(state="disabled" if is_google else "normal")
        self.serpriver_option_entries["location"].configure(state="normal" if is_google else "disabled")
        self.serpriver_option_entries["hl"].configure(state="normal" if is_google else "disabled")
        self.serpriver_option_entries["gl"].configure(state="normal" if is_google else "disabled")

    # ── Импорт файлов ──

    def _import_google_csv(self) -> None:
        self._load_text_rows(self.google_queries_text, "csv")

    def _import_google_xlsx(self) -> None:
        self._load_text_rows(self.google_queries_text, "xlsx")

    def _import_yandex_csv(self) -> None:
        self._load_text_rows(self.yandex_queries_text, "csv")

    def _import_yandex_xlsx(self) -> None:
        self._load_text_rows(self.yandex_queries_text, "xlsx")

    def _import_wordstat_csv(self) -> None:
        self._load_text_rows(self.wordstat_queries_text, "csv")

    def _import_wordstat_xlsx(self) -> None:
        self._load_text_rows(self.wordstat_queries_text, "xlsx")

    def _import_serpriver_csv(self) -> None:
        self._load_text_rows(self.serpriver_queries_text, "csv")

    def _import_serpriver_xlsx(self) -> None:
        self._load_text_rows(self.serpriver_queries_text, "xlsx")

    def _load_text_rows(self, text_widget: tk.Text, file_type: str) -> None:
        """Загружает строки из файла в текстовое поле."""
        filetypes = [("Поддерживаемые файлы", "*.csv *.xlsx"), ("Все файлы", "*.*")]
        file_path = filedialog.askopenfilename(title="Выберите файл", filetypes=filetypes)
        if not file_path:
            return
        reader = read_csv_lines if file_type == "csv" else read_xlsx_lines
        try:
            rows = reader(Path(file_path), self.settings.import_row_limit)
        except ValueError as error:
            logger.error("Import failed | error={}", str(error))
            messagebox.showerror("Ошибка импорта", str(error))
            return
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", "\n".join(rows))
        self._set_status(f"Импортировано строк: {len(rows)}")

    # ── Экспорт результатов ──

    def _export_google_csv(self) -> None:
        self._export_tree_rows(self.google_tree, "export_xmlriver_google", "csv")

    def _export_google_xlsx(self) -> None:
        self._export_tree_rows(self.google_tree, "export_xmlriver_google", "xlsx")

    def _export_yandex_csv(self) -> None:
        self._export_tree_rows(self.yandex_tree, "export_xmlriver_yandex", "csv")

    def _export_yandex_xlsx(self) -> None:
        self._export_tree_rows(self.yandex_tree, "export_xmlriver_yandex", "xlsx")

    def _export_wordstat_csv(self) -> None:
        self._export_tree_rows(self.wordstat_tree, "export_xmlriver_wordstat", "csv")

    def _export_wordstat_xlsx(self) -> None:
        self._export_tree_rows(self.wordstat_tree, "export_xmlriver_wordstat", "xlsx")

    def _export_serpriver_raw_response(self) -> None:
        """Экспортирует сырой ответ SERPRiver в выбранном формате."""
        raw_response = self.serpriver_raw_response.strip()
        if not raw_response:
            messagebox.showwarning("Экспорт", "Сырой ответ SERPRiver пуст")
            return
        output_format = self.serpriver_output_format_var.get().strip().lower()
        suffix = output_format if output_format in SERPRIVER_OUTPUT_FORMAT_VALUES else "json"
        suggested_name = build_export_path(Path.cwd(), "export_serpriver", suffix).name
        file_path = filedialog.asksaveasfilename(
            title="Сохранить экспорт",
            defaultextension=f".{suffix}",
            initialfile=suggested_name,
            filetypes=[(suffix.upper(), f"*.{suffix}"), ("Все файлы", "*.*")],
        )
        if not file_path:
            return
        export_path = Path(file_path)
        try:
            export_path.write_text(raw_response, encoding="utf-8")
        except OSError as error:
            logger.error("SERPRiver export failed | path={} error={}", str(export_path), str(error))
            messagebox.showerror("Экспорт", f"Не удалось сохранить файл: {error}")
            return
        self._set_status(f"Экспорт: {export_path.name}")
        messagebox.showinfo("Экспорт", f"Файл сохранен: {export_path.name}")

    def _export_tree_rows(self, tree: ttk.Treeview, prefix: str, suffix: str) -> None:
        """Экспортирует строки таблицы в выбранный формат."""
        rows = self._collect_tree_rows(tree)
        if not rows:
            messagebox.showwarning("Экспорт", "Таблица пуста")
            return
        suggested_name = build_export_path(Path.cwd(), prefix, suffix).name
        file_path = filedialog.asksaveasfilename(
            title="Сохранить экспорт",
            defaultextension=f".{suffix}",
            initialfile=suggested_name,
            filetypes=[(suffix.upper(), f"*.{suffix}"), ("Все файлы", "*.*")],
        )
        if not file_path:
            return
        export_path = Path(file_path)
        if suffix == "csv":
            export_to_csv(rows, export_path)
        else:
            export_to_xlsx(rows, export_path)
        self._set_status(f"Экспорт: {export_path.name}")
        messagebox.showinfo("Экспорт", f"Файл сохранен: {export_path.name}")

    def _collect_tree_rows(self, tree: ttk.Treeview) -> list[dict[str, str]]:
        """Преобразует строки таблицы в список словарей."""
        rows: list[dict[str, str]] = []
        for item_id in tree.get_children():
            values = tree.item(item_id, "values")
            rows.append(dict(zip(tree["columns"], values, strict=False)))
        return rows

    # ── Фильтрация и сортировка ──

    def _apply_google_view(self) -> None:
        """Применяет фильтр и сортировку к результатам Google XMLRiver."""
        self._render_xml_results(
            self.google_tree, self.google_results,
            self.google_filter_var.get(), self.google_sort_var.get(), self.google_sort_direction_var.get(),
        )

    def _apply_yandex_view(self) -> None:
        """Применяет фильтр и сортировку к результатам Yandex XMLRiver."""
        self._render_xml_results(
            self.yandex_tree, self.yandex_results,
            self.yandex_filter_var.get(), self.yandex_sort_var.get(), self.yandex_sort_direction_var.get(),
        )

    def _apply_wordstat_view(self) -> None:
        """Применяет фильтр и сортировку к результатам Wordstat."""
        self._clear_tree(self.wordstat_tree)
        filtered_results = self._filter_results(self.wordstat_results, self.wordstat_filter_var.get())
        sorted_results = self._sort_results(filtered_results, self.wordstat_sort_var.get(), self.wordstat_sort_direction_var.get())
        for result in sorted_results:
            self._insert_wordstat_result(result)

    def _apply_serpriver_view(self) -> None:
        """Показывает в кодовом поле последний сырой ответ SERPRiver."""
        self.serpriver_raw_response = self._get_last_serpriver_raw_response(self.serpriver_results)
        self._show_serpriver_raw_response(self.serpriver_raw_response)

    def _render_xml_results(
        self,
        tree: ttk.Treeview,
        results: list[XmlRiverResult],
        filter_text: str,
        sort_field: str,
        direction: str,
    ) -> None:
        """Перерисовывает XML-результаты по фильтру и сортировке."""
        self._clear_tree(tree)
        filtered_results = self._filter_results(results, filter_text)
        sorted_results = self._sort_results(filtered_results, sort_field, direction)
        for result in sorted_results:
            self._insert_xml_result(tree, result)

    def _filter_results(self, results: list[Any], filter_text: str) -> list[Any]:
        """Фильтрует список результатов по любой текстовой ячейке."""
        cleaned_filter = filter_text.strip().lower()
        if not cleaned_filter:
            return list(results)
        filtered_results: list[Any] = []
        for result in results:
            row_text = " ".join(str(value).lower() for value in result.model_dump().values())
            if cleaned_filter in row_text:
                filtered_results.append(result)
        return filtered_results

    def _sort_results(self, results: list[Any], sort_field: str, direction: str) -> list[Any]:
        """Сортирует результаты по выбранному полю."""
        reverse = direction == SORT_DIRECTION_VALUES[1]
        return sorted(results, key=lambda item: self._build_sort_key(getattr(item, sort_field, "")), reverse=reverse)

    def _build_sort_key(self, value: Any) -> tuple[int, Any]:
        """Преобразует значение строки в ключ сортировки."""
        text_value = str(value).strip()
        if text_value.isdigit():
            return (0, int(text_value))
        return (1, text_value.lower())

    # ── Запуск операций ──

    def _start_google_xmlriver(self) -> None:
        """Запускает Google XMLRiver."""
        if self.is_google_blocked or self.is_google_xml_running:
            return
        try:
            queries = self._read_queries(self.google_queries_text)
            params = self._collect_google_params()
        except ValueError as error:
            messagebox.showerror("Ошибка Google XMLRiver", str(error))
            return

        self.is_google_xml_running = True
        self._set_controls_state(self.xml_shared_controls + self.google_controls, disabled=True)
        self._prepare_run(self.google_tree, self.google_progress, len(queries))
        self._set_status(f"Google XMLRiver: обработка {len(queries)} запросов...")
        self._submit_background_task(
            self._run_google_xmlriver_requests(queries, params),
            lambda future: self._handle_google_completion(future, len(queries)),
        )

    def _start_yandex_xmlriver(self) -> None:
        """Запускает Yandex XMLRiver."""
        if self.is_yandex_blocked or self.is_yandex_xml_running:
            return
        try:
            queries = self._read_queries(self.yandex_queries_text)
            params = self._collect_yandex_params()
        except ValueError as error:
            messagebox.showerror("Ошибка Yandex XMLRiver", str(error))
            return

        self.is_yandex_xml_running = True
        self._set_controls_state(self.xml_shared_controls + self.yandex_controls, disabled=True)
        self._prepare_run(self.yandex_tree, self.yandex_progress, len(queries))
        self._set_status(f"Yandex XMLRiver: обработка {len(queries)} запросов...")
        self._submit_background_task(
            self._run_yandex_xmlriver_requests(queries, params),
            lambda future: self._handle_yandex_completion(future, len(queries)),
        )

    def _start_wordstat(self) -> None:
        """Запускает Wordstat."""
        if self.is_wordstat_blocked or self.is_wordstat_running:
            return
        try:
            queries = self._read_queries(self.wordstat_queries_text)
            params = self._collect_wordstat_params()
        except ValueError as error:
            messagebox.showerror("Ошибка Wordstat", str(error))
            return

        self.is_wordstat_running = True
        self._set_controls_state(self.xml_shared_controls + self.wordstat_controls, disabled=True)
        self._prepare_run(self.wordstat_tree, self.wordstat_progress, len(queries))
        self._set_status(f"Wordstat: обработка {len(queries)} запросов...")
        self._submit_background_task(
            self._run_wordstat_requests(queries, params),
            lambda future: self._handle_wordstat_completion(future, len(queries)),
        )

    def _start_serpriver(self) -> None:
        """Запускает SERPRiver."""
        if self.is_serpriver_running:
            return
        try:
            queries = self._read_queries(self.serpriver_queries_text)
            target_domain = self._require_value(self.serpriver_domain_var.get(), "Введите домен для отслеживания")
            search_params = self._collect_serpriver_params()
            max_concurrency = self._parse_positive_int(self.serpriver_limit_var.get(), "Потоки")
        except ValueError as error:
            messagebox.showerror("Ошибка SERPRiver", str(error))
            return

        self.is_serpriver_running = True
        self._set_controls_state(self.serpriver_controls, disabled=True)
        self._prepare_run(None, self.serpriver_progress, len(queries))
        self._set_status(f"SERPRiver: обработка {len(queries)} запросов...")
        self._submit_background_task(
            self._run_serpriver_requests(queries, target_domain, search_params, max_concurrency),
            lambda future: self._handle_serpriver_completion(future, len(queries)),
        )

    def _prepare_run(self, tree: ttk.Treeview | None, progressbar: ttk.Progressbar, total: int) -> None:
        """Очищает результаты и подготавливает progress bar."""
        if tree is not None:
            self._clear_tree(tree)
        progressbar.configure(maximum=max(total, 1), value=0)
        if tree is self.google_tree:
            self.google_results = []
        elif tree is self.yandex_tree:
            self.yandex_results = []
        elif tree is self.wordstat_tree:
            self.wordstat_results = []
        else:
            self.serpriver_results = []
            self.serpriver_raw_response = ""
            self._show_serpriver_raw_response("")

    # ── Асинхронные запросы ──

    async def _run_google_xmlriver_requests(self, queries: list[str], params: dict[str, str]) -> list[XmlRiverResult]:
        """Выполняет запросы Google XMLRiver."""
        client = XmlRiverClient(
            user_id=self.xmlriver_user_var.get().strip(),
            api_key=self.xmlriver_key_var.get().strip(),
            connect_timeout=self.settings.request_connect_timeout,
            read_timeout=self.settings.request_read_timeout,
            max_concurrency=self.settings.xmlriver_max_concurrency,
        )
        return await client.fetch_queries(
            queries=queries, engine="google", params=params,
            progress_callback=lambda completed, total, query: self._queue_progress_update(self.google_progress, completed, total),
        )

    async def _run_yandex_xmlriver_requests(self, queries: list[str], params: dict[str, str]) -> list[XmlRiverResult]:
        """Выполняет запросы Yandex XMLRiver."""
        client = XmlRiverClient(
            user_id=self.xmlriver_user_var.get().strip(),
            api_key=self.xmlriver_key_var.get().strip(),
            connect_timeout=self.settings.request_connect_timeout,
            read_timeout=self.settings.request_read_timeout,
            max_concurrency=self.settings.xmlriver_max_concurrency,
        )
        return await client.fetch_queries(
            queries=queries, engine="yandex", params=params,
            progress_callback=lambda completed, total, query: self._queue_progress_update(self.yandex_progress, completed, total),
        )

    async def _run_wordstat_requests(self, queries: list[str], params: dict[str, str]) -> list[WordstatResult]:
        """Выполняет запросы Wordstat."""
        client = WordstatClient(
            user_id=self.xmlriver_user_var.get().strip(),
            api_key=self.xmlriver_key_var.get().strip(),
            connect_timeout=self.settings.request_connect_timeout,
            read_timeout=self.settings.request_read_timeout,
            max_concurrency=self.settings.xmlriver_max_concurrency,
        )
        return await client.fetch_queries(
            queries=queries, params=params,
            progress_callback=lambda completed, total, query: self._queue_progress_update(self.wordstat_progress, completed, total),
        )

    async def _run_serpriver_requests(
        self, queries: list[str], target_domain: str, search_params: dict[str, str], max_concurrency: int,
    ) -> list[SerpRiverResult]:
        """Выполняет запросы SERPRiver."""
        client = SerpRiverClient(
            api_key=self.serpriver_key_var.get().strip(),
            connect_timeout=self.settings.request_connect_timeout,
            read_timeout=self.settings.request_read_timeout,
            max_concurrency=max_concurrency,
        )
        return await client.fetch_queries(
            queries=queries, target_domain=target_domain, search_params=search_params,
            progress_callback=lambda completed, total, query: self._queue_progress_update(self.serpriver_progress, completed, total),
        )

    def _submit_background_task(self, coroutine: Any, on_done: Any) -> Future:
        """Отправляет coroutine в фоновый event loop."""
        future = asyncio.run_coroutine_threadsafe(coroutine, self.background_loop)
        future.add_done_callback(lambda done_future: self.root.after(0, lambda: on_done(done_future)))
        return future

    # ── Обработчики завершения ──

    def _handle_google_completion(self, future: Future, total: int) -> None:
        """Завершает цикл Google XMLRiver."""
        self.is_google_xml_running = False
        self._restore_xml_control_sets(self.google_controls)
        try:
            results = future.result()
        except Exception as error:  # noqa: BLE001
            logger.error("Google XMLRiver failed | error={}", str(error))
            self._set_status("Google XMLRiver: ошибка")
            messagebox.showerror("Ошибка Google XMLRiver", str(error))
            return
        self.google_results = results
        self._apply_google_view()
        self.google_progress.configure(value=max(total, 1))
        self._set_status(f"Google XMLRiver: завершено, строк: {len(results)}")
        messagebox.showinfo("Google XMLRiver", f"Обработка завершена. Строк: {len(results)}")

    def _handle_yandex_completion(self, future: Future, total: int) -> None:
        """Завершает цикл Yandex XMLRiver."""
        self.is_yandex_xml_running = False
        self._restore_xml_control_sets(self.yandex_controls)
        try:
            results = future.result()
        except Exception as error:  # noqa: BLE001
            logger.error("Yandex XMLRiver failed | error={}", str(error))
            self._set_status("Yandex XMLRiver: ошибка")
            messagebox.showerror("Ошибка Yandex XMLRiver", str(error))
            return
        self.yandex_results = results
        self._apply_yandex_view()
        self.yandex_progress.configure(value=max(total, 1))
        self._set_status(f"Yandex XMLRiver: завершено, строк: {len(results)}")
        messagebox.showinfo("Yandex XMLRiver", f"Обработка завершена. Строк: {len(results)}")

    def _handle_wordstat_completion(self, future: Future, total: int) -> None:
        """Завершает цикл Wordstat."""
        self.is_wordstat_running = False
        self._restore_xml_control_sets(self.wordstat_controls)
        try:
            results = future.result()
        except Exception as error:  # noqa: BLE001
            logger.error("Wordstat failed | error={}", str(error))
            self._set_status("Wordstat: ошибка")
            messagebox.showerror("Ошибка Wordstat", str(error))
            return
        self.wordstat_results = results
        self._apply_wordstat_view()
        self.wordstat_progress.configure(value=max(total, 1))
        self._set_status(f"Wordstat: завершено, строк: {len(results)}")
        messagebox.showinfo("Wordstat", f"Обработка завершена. Строк: {len(results)}")

    def _handle_serpriver_completion(self, future: Future, total: int) -> None:
        """Завершает цикл SERPRiver."""
        self.is_serpriver_running = False
        self._set_controls_state(self.serpriver_controls, disabled=False)
        self._update_serpriver_engine_fields()
        try:
            results = future.result()
        except Exception as error:  # noqa: BLE001
            logger.error("SERPRiver failed | error={}", str(error))
            self._set_status("SERPRiver: ошибка")
            messagebox.showerror("Ошибка SERPRiver", str(error))
            return
        self.serpriver_results = results
        self._apply_serpriver_view()
        self.serpriver_progress.configure(value=max(total, 1))
        self._set_status(f"SERPRiver: завершено, строк: {len(results)}")
        messagebox.showinfo("SERPRiver", f"Обработка завершена. Строк: {len(results)}")

    def _restore_xml_control_sets(self, tab_controls: list[tuple[tk.Widget, str]]) -> None:
        """Восстанавливает доступность общих полей XMLRiver и подвкладки."""
        self._set_controls_state(self.xml_shared_controls, disabled=False)
        self._set_controls_state(tab_controls, disabled=False)
        if self.is_google_blocked:
            self._set_controls_state(self.google_controls, disabled=True)
        if self.is_yandex_blocked:
            self._set_controls_state(self.yandex_controls, disabled=True)
        if self.is_wordstat_blocked:
            self._set_controls_state(self.wordstat_controls, disabled=True)

    # ── Утилиты UI ──

    def _queue_progress_update(self, progressbar: ttk.Progressbar, completed: int, total: int) -> None:
        """Безопасно обновляет progress bar из фонового потока."""
        self.root.after(0, lambda: progressbar.configure(maximum=max(total, 1), value=completed))

    def _set_controls_state(self, controls: list[tuple[tk.Widget, str]], disabled: bool) -> None:
        """Изменяет состояние списка виджетов."""
        for widget, normal_state in controls:
            widget.configure(state="disabled" if disabled else normal_state)

    def _set_status(self, text: str) -> None:
        """Обновляет текст статусбара."""
        self.status_var.set(text)

    def _clear_tree(self, tree: ttk.Treeview) -> None:
        """Удаляет все строки таблицы."""
        for item_id in tree.get_children():
            tree.delete(item_id)

    def _show_serpriver_raw_response(self, raw_response: str) -> None:
        """Обновляет кодовое окно с сырым ответом SERPRiver."""
        display_response = self._format_serpriver_response_for_display(raw_response)
        self.serpriver_response_text.configure(state="normal")
        self.serpriver_response_text.delete("1.0", "end")
        if display_response.strip():
            self.serpriver_response_text.insert("1.0", display_response)
        self.serpriver_response_text.configure(state="disabled")
        self.serpriver_response_text.xview_moveto(0)
        self.serpriver_response_text.yview_moveto(0)

    def _format_serpriver_response_for_display(self, raw_response: str) -> str:
        """Форматирует сырой ответ SERPRiver для отображения в UI."""
        cleaned_response = raw_response.strip()
        if not cleaned_response:
            return ""
        if self.serpriver_output_format_var.get().strip().lower() != "json":
            return raw_response
        try:
            parsed_response = json.loads(cleaned_response)
        except json.JSONDecodeError:
            return raw_response
        return json.dumps(parsed_response, ensure_ascii=False, indent=2)

    # ── Вставка результатов ──

    def _insert_xml_result(self, tree: ttk.Treeview, result: XmlRiverResult) -> None:
        """Добавляет строку XMLRiver в таблицу."""
        values = (
            result.query,
            result.position or ("Ошибка" if result.error_code else ""),
            result.url or result.error_code,
            result.domain,
            result.title or result.error_message,
            result.snippet,
        )
        tags = (TREE_ERROR_TAG,) if result.error_code else ()
        tree.insert("", "end", values=values, tags=tags)

    def _insert_wordstat_result(self, result: WordstatResult) -> None:
        """Добавляет строку Wordstat в таблицу."""
        values = (
            result.query,
            result.result_type or ("Ошибка" if result.error_code else ""),
            result.phrase or result.error_message,
            result.value or result.error_code,
        )
        tags = (TREE_ERROR_TAG,) if result.error_code else ()
        self.wordstat_tree.insert("", "end", values=values, tags=tags)

    def _get_last_serpriver_raw_response(self, results: list[SerpRiverResult]) -> str:
        """Возвращает последний непустой сырой ответ SERPRiver."""
        for result in reversed(results):
            raw_response = result.raw_response.strip()
            if raw_response:
                return raw_response
        return ""

    # ── Чтение и валидация ──

    def _read_queries(self, text_widget: tk.Text) -> list[str]:
        """Возвращает список запросов из текстового поля."""
        rows = [row.strip() for row in text_widget.get("1.0", "end").splitlines() if row.strip()]
        if not rows:
            raise ValueError("Введите хотя бы один запрос")
        return rows

    def _collect_google_params(self) -> dict[str, str]:
        """Собирает параметры Google XMLRiver."""
        self._require_xml_credentials()
        return {
            "engine": "google",
            "groupby": "10",
            "loc": self._get_reference_value("google_geo", self.google_loc_var.get(), "Выберите локацию"),
            "country": GOOGLE_FIXED_COUNTRY_VALUE,
            "lr": GOOGLE_FIXED_LANGUAGE_VALUE,
            "domain": GOOGLE_FIXED_DOMAIN_VALUE,
            "device": self._get_static_value(self.google_device_var.get(), DEVICE_VALUES, "Выберите устройство"),
        }

    def _collect_yandex_params(self) -> dict[str, str]:
        """Собирает параметры Yandex XMLRiver."""
        self._require_xml_credentials()
        return {
            "engine": "yandex",
            "groupby": "10",
            "lr": self._get_reference_value("yandex_geo", self.yandex_lr_var.get(), "Выберите регион"),
            "domain": self._get_static_value(self.yandex_domain_var.get(), YANDEX_DOMAIN_VALUES, "Выберите домен Яндекса"),
            "lang": self._require_value(self.yandex_lang_var.get(), "Введите язык"),
            "device": self._get_static_value(self.yandex_device_var.get(), DEVICE_VALUES, "Выберите устройство"),
        }

    def _collect_wordstat_params(self) -> dict[str, str]:
        """Собирает параметры Wordstat."""
        self._require_xml_credentials()
        region_id = self._get_reference_value("yandex_geo", self.wordstat_region_var.get(), "Выберите регион")
        extra_regions = ",".join(
            chunk.strip() for chunk in self.wordstat_regions_extra_var.get().split(",") if chunk.strip()
        )
        start_value = self._normalize_wordstat_date(self.wordstat_start_var.get(), "Укажите корректную дату начала")
        end_value = self._normalize_wordstat_date(self.wordstat_end_var.get(), "Укажите корректную дату окончания")
        if extra_regions:
            regions = ",".join([region_id, extra_regions])
        else:
            regions = region_id
        if bool(start_value) != bool(end_value):
            raise ValueError("Укажите обе даты периода: начало и окончание")
        return {
            "regions": regions,
            "device": self._get_static_value(self.wordstat_device_var.get(), WORDSTAT_DEVICE_VALUES, "", allow_empty=True),
            "period": self._get_static_value(self.wordstat_period_var.get(), WORDSTAT_PERIOD_VALUES, "", allow_empty=True),
            "start": start_value,
            "end": end_value,
            "pagetype": self._get_static_value(
                self.wordstat_pagetype_var.get(), WORDSTAT_PAGETYPE_VALUES, "Выберите тип данных",
            ) or "words",
        }

    def _collect_serpriver_params(self) -> dict[str, str]:
        """Собирает параметры SERPRiver."""
        self._require_value(self.serpriver_key_var.get(), "Введите API Key SERPRiver")
        result_cnt = self._parse_positive_int(self.serpriver_result_cnt_var.get(), "result_cnt")
        domain = self._require_value(self.serpriver_search_domain_var.get(), "Введите domain поисковика")
        device = self._require_value(self.serpriver_device_var.get(), "Введите device")
        engine = self.serpriver_engine_var.get()
        search_params = {
            "system": engine,
            "domain": domain,
            "result_cnt": str(result_cnt),
            "device": device,
            "output_format": self._get_static_value(
                self.serpriver_output_format_var.get(),
                SERPRIVER_OUTPUT_FORMAT_VALUES,
                "Выберите формат ответа",
            ),
        }
        if engine == "google":
            search_params["location"] = self._require_value(self.serpriver_location_var.get(), "Введите location для Google")
            search_params["hl"] = self._require_value(self.serpriver_hl_var.get(), "Введите hl для Google")
            search_params["gl"] = self._require_value(self.serpriver_gl_var.get(), "Введите gl для Google")
        else:
            search_params["lr"] = str(self._parse_non_negative_int(self.serpriver_lr_var.get(), "lr"))
        return search_params

    def _require_xml_credentials(self) -> None:
        """Проверяет общие учетные данные XMLRiver."""
        self._require_value(self.xmlriver_user_var.get(), "Введите User ID XMLRiver")
        self._require_value(self.xmlriver_key_var.get(), "Введите API Key XMLRiver")

    def _get_reference_value(self, catalog_name: str, label: str, error_message: str) -> str:
        """Возвращает значение справочника по выбранному label."""
        selected_label = label.strip()
        if not selected_label:
            raise ValueError(error_message)
        catalog = self.reference_maps.get(catalog_name, {})
        if selected_label in catalog:
            return catalog[selected_label]
        for option_label, option_value in catalog.items():
            if option_label.lower() == selected_label.lower():
                return option_value
        matches = [option_value for option_label, option_value in catalog.items() if selected_label.lower() in option_label.lower()]
        if len(matches) == 1:
            return matches[0]
        raise ValueError(error_message)

    def _require_value(self, value: str, error_message: str) -> str:
        """Проверяет обязательное строковое значение."""
        cleaned_value = value.strip()
        if not cleaned_value:
            raise ValueError(error_message)
        return cleaned_value

    def _normalize_wordstat_date(self, value: str, error_message: str) -> str:
        """Проверяет и нормализует дату Wordstat в формате YYYY-MM-DD."""
        cleaned_value = value.strip()
        if not cleaned_value:
            return ""
        parsed_date = self._parse_date_text(cleaned_value)
        if parsed_date is None:
            raise ValueError(error_message)
        return parsed_date.strftime(WORDSTAT_DATE_FORMAT)

    def _get_static_value(
        self,
        value: str,
        allowed_values: tuple[str, ...],
        error_message: str,
        allow_empty: bool = False,
    ) -> str:
        """Возвращает допустимое значение из статического списка."""
        cleaned_value = value.strip()
        if not cleaned_value:
            if allow_empty:
                return ""
            raise ValueError(error_message)
        lowered_value = cleaned_value.lower()
        exact_matches = [item for item in allowed_values if item.lower() == lowered_value]
        if exact_matches:
            return exact_matches[0]
        partial_matches = [item for item in allowed_values if lowered_value in item.lower()]
        if len(partial_matches) == 1:
            return partial_matches[0]
        raise ValueError(error_message)

    def _parse_positive_int(self, value: str, field_name: str) -> int:
        """Преобразует строку в положительное число."""
        try:
            parsed_value = int(value.strip())
        except ValueError as error:
            raise ValueError(f"Поле {field_name} должно быть числом") from error
        if parsed_value <= 0:
            raise ValueError(f"Поле {field_name} должно быть больше нуля")
        return parsed_value

    def _parse_non_negative_int(self, value: str, field_name: str) -> int:
        """Преобразует строку в неотрицательное число."""
        try:
            parsed_value = int(value.strip())
        except ValueError as error:
            raise ValueError(f"Поле {field_name} должно быть числом") from error
        if parsed_value < 0:
            raise ValueError(f"Поле {field_name} не может быть отрицательным")
        return parsed_value

    def run(self) -> None:
        """Запускает цикл обработки событий tkinter."""
        self.root.mainloop()
