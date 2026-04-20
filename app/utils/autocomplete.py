"""Хелпер автодополнения для ttk.Combobox."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Sequence

import ttkbootstrap as ttk


class ComboboxAutocomplete:
    """Обеспечивает живой поиск и фиксацию выбора для combobox."""

    def __init__(
        self,
        root: tk.Misc,
        combobox: ttk.Combobox,
        value_var: tk.StringVar,
        items: Sequence[str],
        *,
        max_values: int,
        debounce_ms: int,
        on_commit: Callable[[str], None] | None = None,
    ) -> None:
        self.root = root
        self.combobox = combobox
        self.value_var = value_var
        self.items = tuple(items)
        self.max_values = max_values
        self.debounce_ms = debounce_ms
        self.on_commit = on_commit

        self._debounce_id: str | None = None
        self._filtered_items: tuple[str, ...] = ()
        self._configured_values: tuple[str, ...] = ()
        self._committed_value = self._find_exact_match(self.value_var.get().strip()) or ""

        self._bind_events()
        self.refresh(show_dropdown=False)

    def refresh(self, show_dropdown: bool = False) -> None:
        """Обновляет список подсказок по текущему тексту поля."""
        typed_value = self.value_var.get().strip()
        filtered_items = self._filter_items(typed_value)
        limited_items = tuple(filtered_items[: self.max_values])
        self._filtered_items = limited_items
        if limited_items != self._configured_values:
            self.combobox.configure(values=limited_items)
            self._configured_values = limited_items
        if limited_items and show_dropdown:
            self.root.after_idle(self._open_dropdown)

    def _bind_events(self) -> None:
        """Подключает обработчики ввода, выбора и потери фокуса."""
        self.combobox.bind("<KeyRelease>", self._handle_key_release, add="+")
        self.combobox.bind("<Button-1>", self._handle_click, add="+")
        self.combobox.bind("<<ComboboxSelected>>", self._handle_selection, add="+")
        self.combobox.bind("<Return>", self._handle_return, add="+")
        self.combobox.bind("<FocusOut>", self._handle_focus_out, add="+")

    def _handle_key_release(self, event: tk.Event) -> None:
        """Запускает отложенное обновление подсказок после ввода."""
        if self._is_control_key(event):
            return
        self._schedule_refresh(show_dropdown=True)

    def _handle_click(self, _event: tk.Event) -> None:
        """Показывает актуальные подсказки при открытии списка мышью."""
        self._schedule_refresh(show_dropdown=bool(self.value_var.get().strip()))

    def _handle_selection(self, _event: tk.Event) -> None:
        """Фиксирует значение после выбора элемента из списка."""
        self._commit_value(self.combobox.get())

    def _handle_return(self, _event: tk.Event) -> str:
        """Фиксирует текущее выбранное значение по Enter."""
        self._commit_current_value(prefer_filtered=True)
        return "break"

    def _handle_focus_out(self, _event: tk.Event) -> None:
        """Нормализует текст поля после потери фокуса."""
        self._commit_current_value(prefer_filtered=False)

    def _schedule_refresh(self, *, show_dropdown: bool) -> None:
        """Перезапускает debounce перед фильтрацией списка."""
        if self._debounce_id is not None:
            self.root.after_cancel(self._debounce_id)
        self._debounce_id = self.root.after(
            self.debounce_ms,
            lambda: self._flush_refresh(show_dropdown=show_dropdown),
        )

    def _flush_refresh(self, *, show_dropdown: bool) -> None:
        """Сбрасывает debounce и выполняет фильтрацию."""
        self._debounce_id = None
        self.refresh(show_dropdown=show_dropdown)

    def _commit_current_value(self, *, prefer_filtered: bool) -> None:
        """Пытается привести текущий текст к точному значению списка."""
        resolved_value = self._resolve_value(self.combobox.get(), prefer_filtered=prefer_filtered)
        if resolved_value is None:
            return
        self._commit_value(resolved_value)

    def _commit_value(self, value: str) -> None:
        """Записывает подтверждённое значение в поле и вызывает callback."""
        normalized_value = self._find_exact_match(value.strip()) or value.strip()
        if not normalized_value:
            return
        self._committed_value = normalized_value
        self.value_var.set(normalized_value)
        self.combobox.set(normalized_value)
        self._close_dropdown()
        if self.on_commit is not None:
            self.on_commit(normalized_value)

    def _resolve_value(self, typed_value: str, *, prefer_filtered: bool) -> str | None:
        """Находит каноническое значение по точному или уникальному совпадению."""
        cleaned_value = typed_value.strip()
        if not cleaned_value:
            return None
        exact_match = self._find_exact_match(cleaned_value)
        if exact_match is not None:
            return exact_match
        unique_match = self._find_unique_contains_match(cleaned_value)
        if unique_match is not None:
            return unique_match
        current_index = self.combobox.current()
        if 0 <= current_index < len(self._filtered_items):
            return self._filtered_items[current_index]
        if prefer_filtered and len(self._filtered_items) == 1:
            return self._filtered_items[0]
        return None

    def _filter_items(self, typed_value: str) -> list[str]:
        """Возвращает элементы, содержащие введённую подстроку."""
        if not typed_value:
            return list(self.items)
        lowered_value = typed_value.casefold()
        return [item for item in self.items if lowered_value in item.casefold()]

    def _find_exact_match(self, value: str) -> str | None:
        """Ищет точное совпадение без учёта регистра."""
        lowered_value = value.casefold()
        for item in self.items:
            if item.casefold() == lowered_value:
                return item
        return None

    def _find_unique_contains_match(self, value: str) -> str | None:
        """Возвращает единственное совпадение по вхождению подстроки."""
        lowered_value = value.casefold()
        matches = [item for item in self.items if lowered_value in item.casefold()]
        if len(matches) == 1:
            return matches[0]
        return None

    def _is_control_key(self, event: tk.Event) -> bool:
        """Исключает служебные клавиши из фильтрации."""
        keysym = getattr(event, "keysym", "")
        return keysym in {
            "Up",
            "Down",
            "Left",
            "Right",
            "Home",
            "End",
            "Prior",
            "Next",
            "Return",
            "Escape",
            "Tab",
            "Shift_L",
            "Shift_R",
            "Control_L",
            "Control_R",
            "Alt_L",
            "Alt_R",
            "Caps_Lock",
        }

    def _open_dropdown(self) -> None:
        """Открывает выпадающий список с подсказками."""
        try:
            self.combobox.tk.call("ttk::combobox::Post", str(self.combobox))
            self.combobox.focus_set()
            self.combobox.icursor("end")
        except tk.TclError:
            return

    def _close_dropdown(self) -> None:
        """Закрывает выпадающий список после подтверждения выбора."""
        try:
            self.combobox.tk.call("ttk::combobox::Unpost", str(self.combobox))
        except tk.TclError:
            return
