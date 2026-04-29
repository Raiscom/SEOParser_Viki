"""Autocomplete helper for ttk.Combobox."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Sequence

import ttkbootstrap as ttk


class ComboboxAutocomplete:
    """Provides live filtering and committed selection for combobox fields."""

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
        self._popup: tk.Toplevel | None = None
        self._popup_listbox: tk.Listbox | None = None

        self._bind_events()
        self.refresh(show_dropdown=False)

    def refresh(self, show_dropdown: bool = False) -> None:
        """Refreshes suggestions from current field text."""
        typed_value = self.value_var.get().strip()
        filtered_items = self._filter_items(typed_value)
        limited_items = tuple(filtered_items[: self.max_values])
        self._filtered_items = limited_items
        if limited_items != self._configured_values:
            self.combobox.configure(values=limited_items)
            self._configured_values = limited_items
        if limited_items and show_dropdown:
            self.root.after_idle(self._show_popup)
        elif not limited_items:
            self._close_dropdown()

    def _bind_events(self) -> None:
        """Connects input, selection and focus handlers."""
        self.combobox.bind("<KeyRelease>", self._handle_key_release, add="+")
        self.combobox.bind("<Button-1>", self._handle_click, add="+")
        self.combobox.bind("<<ComboboxSelected>>", self._handle_selection, add="+")
        self.combobox.bind("<Return>", self._handle_return, add="+")
        self.combobox.bind("<FocusOut>", self._handle_focus_out, add="+")
        self.combobox.bind("<Escape>", lambda _event: self._close_dropdown(), add="+")

    def _handle_key_release(self, event: tk.Event) -> None:
        """Schedules suggestions update after typing."""
        if self._is_control_key(event):
            return
        self._schedule_refresh(show_dropdown=True)

    def _handle_click(self, _event: tk.Event) -> None:
        """Shows current suggestions when the field is clicked."""
        self._schedule_refresh(show_dropdown=True)

    def _handle_selection(self, _event: tk.Event) -> None:
        """Commits a value selected through the native combobox list."""
        self._commit_value(self.combobox.get())

    def _handle_return(self, _event: tk.Event) -> str:
        """Commits the current value on Enter."""
        self._commit_current_value(prefer_filtered=True)
        return "break"

    def _handle_focus_out(self, _event: tk.Event) -> None:
        """Commits only after focus has really left the field and popup."""
        self.root.after(120, self._commit_after_focus_out)

    def _commit_after_focus_out(self) -> None:
        """Commits text after focus leaves both combobox and autocomplete popup."""
        focus_widget = self.root.focus_get()
        if focus_widget is self.combobox or self._is_popup_widget(focus_widget):
            return
        self._commit_current_value(prefer_filtered=False)
        self._close_dropdown()

    def _schedule_refresh(self, *, show_dropdown: bool) -> None:
        """Restarts debounce before filtering values."""
        if self._debounce_id is not None:
            self.root.after_cancel(self._debounce_id)
        self._debounce_id = self.root.after(
            self.debounce_ms,
            lambda: self._flush_refresh(show_dropdown=show_dropdown),
        )

    def _flush_refresh(self, *, show_dropdown: bool) -> None:
        """Runs the debounced refresh."""
        self._debounce_id = None
        self.refresh(show_dropdown=show_dropdown)

    def _commit_current_value(self, *, prefer_filtered: bool) -> None:
        """Tries to resolve current text to a catalog value."""
        resolved_value = self._resolve_value(self.combobox.get(), prefer_filtered=prefer_filtered)
        if resolved_value is None:
            return
        self._commit_value(resolved_value)

    def _commit_value(self, value: str) -> None:
        """Stores a confirmed value in the field and invokes the callback."""
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
        """Finds an exact or unique catalog value for current text."""
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
        """Returns values containing the typed text."""
        if not typed_value:
            return list(self.items)
        lowered_value = typed_value.casefold()
        return [item for item in self.items if lowered_value in item.casefold()]

    def _find_exact_match(self, value: str) -> str | None:
        """Finds a case-insensitive exact match."""
        lowered_value = value.casefold()
        for item in self.items:
            if item.casefold() == lowered_value:
                return item
        return None

    def _find_unique_contains_match(self, value: str) -> str | None:
        """Returns the only contains-match when it is unique."""
        lowered_value = value.casefold()
        matches = [item for item in self.items if lowered_value in item.casefold()]
        if len(matches) == 1:
            return matches[0]
        return None

    def _is_control_key(self, event: tk.Event) -> bool:
        """Skips service keys that should not trigger filtering."""
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
        """Compatibility wrapper for older call sites."""
        self._show_popup()

    def _close_dropdown(self) -> None:
        """Closes the autocomplete popup."""
        if self._popup is None:
            return
        try:
            self._popup.destroy()
        except tk.TclError:
            pass
        self._popup = None
        self._popup_listbox = None

    def _show_popup(self) -> None:
        """Shows filtered values without stealing text input from the combobox."""
        if not self._filtered_items or not self.combobox.winfo_ismapped():
            self._close_dropdown()
            return
        if self._popup is None or not self._popup.winfo_exists():
            self._create_popup()
        if self._popup is None or self._popup_listbox is None:
            return
        self._position_popup()
        self._refresh_popup_items()
        self._popup.deiconify()
        self._popup.lift()
        self.combobox.focus_set()
        self.combobox.icursor("end")

    def _create_popup(self) -> None:
        """Creates a small non-modal list under the combobox."""
        popup = tk.Toplevel(self.combobox)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)

        frame = ttk.Frame(popup)
        frame.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", bootstyle="warning-round")
        listbox = tk.Listbox(
            frame,
            activestyle="dotbox",
            bg="#2b3e50",
            fg="white",
            highlightthickness=1,
            highlightbackground="#FF8C00",
            relief="solid",
            selectbackground="#FF8C00",
            selectforeground="white",
            yscrollcommand=scrollbar.set,
        )
        scrollbar.configure(command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)
        listbox.bind("<ButtonRelease-1>", self._handle_popup_selection, add="+")
        listbox.bind("<Return>", self._handle_popup_return, add="+")
        listbox.bind("<Escape>", lambda _event: self._close_dropdown(), add="+")
        listbox.bind("<MouseWheel>", self._handle_popup_mousewheel, add="+")

        self._popup = popup
        self._popup_listbox = listbox

    def _position_popup(self) -> None:
        """Positions the popup directly under the combobox."""
        if self._popup is None:
            return
        self.combobox.update_idletasks()
        row_height = 24
        visible_rows = min(max(len(self._filtered_items), 1), 8)
        width = max(self.combobox.winfo_width(), 220)
        height = visible_rows * row_height
        x = self.combobox.winfo_rootx()
        y = self.combobox.winfo_rooty() + self.combobox.winfo_height()
        self._popup.geometry(f"{width}x{height}+{x}+{y}")

    def _refresh_popup_items(self) -> None:
        """Replaces popup rows with current filtered values."""
        if self._popup_listbox is None:
            return
        self._popup_listbox.delete(0, "end")
        for item in self._filtered_items:
            self._popup_listbox.insert("end", item)

    def _handle_popup_selection(self, _event: tk.Event) -> None:
        """Commits the clicked popup item."""
        if self._popup_listbox is None:
            return
        selection = self._popup_listbox.curselection()
        if not selection:
            return
        self._commit_value(self._popup_listbox.get(selection[0]))

    def _handle_popup_return(self, event: tk.Event) -> str:
        """Commits the active popup item."""
        self._handle_popup_selection(event)
        return "break"

    def _handle_popup_mousewheel(self, event: tk.Event) -> str:
        """Scrolls the popup list with the mouse wheel."""
        if self._popup_listbox is None:
            return "break"
        delta = -1 if getattr(event, "delta", 0) > 0 else 1
        self._popup_listbox.yview_scroll(delta, "units")
        return "break"

    def _is_popup_widget(self, widget: tk.Widget | None) -> bool:
        """Returns True when the widget belongs to the autocomplete popup."""
        if widget is None or self._popup is None:
            return False
        try:
            return str(widget).startswith(str(self._popup))
        except tk.TclError:
            return False
