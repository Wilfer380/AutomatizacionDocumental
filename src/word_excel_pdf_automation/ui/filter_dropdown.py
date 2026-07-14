from __future__ import annotations

import tkinter as tk
from tkinter import BooleanVar, StringVar, ttk
from typing import Callable

from ..utils.text import normalize_for_match


class MultiSelectDropdown(ttk.Frame):
    def __init__(
        self,
        parent: tk.Widget,
        *,
        title: str,
        empty_label: str = "Todas",
        on_selection_changed: Callable[[list[str]], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title = title
        self.empty_label = empty_label
        self.on_selection_changed = on_selection_changed

        self._available_options: tuple[str, ...] = ()
        self._selected: list[str] = []
        self._variables: dict[str, BooleanVar] = {}
        self._popup: tk.Toplevel | None = None
        self._search_var = StringVar(value="")
        self._summary_var = StringVar(value=empty_label)
        self._search_var.trace_add("write", self._on_search_changed)
        self._list_container: ttk.Frame | None = None
        self._canvas: tk.Canvas | None = None
        self._window_id: int | None = None
        self._popup_selected: list[str] | None = None
        self._popup_initial_selected: list[str] | None = None

        self.columnconfigure(0, weight=1)
        self.button = ttk.Button(self, text=self._build_button_text(), style="Ghost.TButton", command=self.toggle_popup)
        self.button.grid(row=0, column=0, sticky="ew")

    def set_options(self, options: tuple[str, ...]) -> None:
        self._available_options = options
        current = set(self._selected)
        self._selected = [item for item in options if item in current]
        self._variables = {item: BooleanVar(value=item in self._selected) for item in options}
        self._sync_summary()
        self._render_options()

    def set_selected(self, selected: list[str]) -> None:
        allowed = set(self._available_options)
        self._selected = [item for item in self._available_options if item in allowed and item in set(selected)]
        for option, variable in self._variables.items():
            variable.set(option in self._selected)
        self._sync_summary()
        self._render_options()

    def clear(self) -> None:
        self.set_selected([])
        self._search_var.set("")

    def get_selected(self) -> list[str]:
        return list(self._selected)

    def set_enabled(self, enabled: bool) -> None:
        self.button.configure(state="normal" if enabled else "disabled")
        if not enabled:
            self._close_popup()

    def toggle_popup(self) -> None:
        if self._popup and self._popup.winfo_exists():
            self._close_popup()
            return
        self._open_popup()

    def _open_popup(self) -> None:
        popup = tk.Toplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.winfo_toplevel())
        popup.configure(bg="#d9e3ee")
        popup.bind("<FocusOut>", lambda _event: self.after(50, self._maybe_close_popup))
        popup.bind("<Escape>", lambda _event: self._close_popup())

        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        self._popup_initial_selected = list(self._selected)
        self._popup_selected = list(self._selected)

        ttk.Label(frame, text=self.title, style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w")

        search_frame = ttk.Frame(frame)
        search_frame.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        search_frame.columnconfigure(0, weight=1)
        ttk.Entry(search_frame, textvariable=self._search_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(search_frame, text="Todas", style="Ghost.TButton", command=self._select_all).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(search_frame, text="Ninguna", style="Ghost.TButton", command=self._clear_all).grid(row=0, column=2)

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=2, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        canvas = tk.Canvas(list_frame, highlightthickness=0, bg="#ffffff")
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        content = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))

        self._popup = popup
        self._canvas = canvas
        self._list_container = content
        self._window_id = window_id
        self._render_options()

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="e", pady=(10, 0))
        ttk.Button(actions, text="Cancelar", style="Ghost.TButton", command=self._cancel_popup).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Aplicar", style="Primary.TButton", command=self._apply_popup).pack(side="left")

        self.update_idletasks()
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height()
        width = max(self.button.winfo_width(), 280)
        height = 320
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.deiconify()
        popup.lift()
        popup.focus_force()

    def _close_popup(self) -> None:
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        self._list_container = None
        self._canvas = None
        self._window_id = None
        self._popup_selected = None
        self._popup_initial_selected = None

    def _maybe_close_popup(self) -> None:
        if not self._popup or not self._popup.winfo_exists():
            return
        focused = self.winfo_toplevel().focus_get()
        if focused is None:
            self._close_popup()
            return
        current = focused
        while current is not None:
            if current == self._popup:
                return
            current = current.master
        self._close_popup()

    def _on_search_changed(self, *_args) -> None:
        self._render_options()

    def _filtered_options(self) -> list[str]:
        needle = normalize_for_match(self._search_var.get())
        if not needle:
            return list(self._available_options)
        return [value for value in self._available_options if needle in normalize_for_match(value)]

    def _render_options(self) -> None:
        if self._list_container is None:
            return
        for widget in self._list_container.winfo_children():
            widget.destroy()

        filtered = self._filtered_options()
        if not filtered:
            ttk.Label(self._list_container, text="Sin resultados", style="Muted.TLabel").pack(anchor="w", padx=4, pady=4)
            return

        for option in filtered:
            ttk.Checkbutton(
                self._list_container,
                text=option,
                variable=self._variables[option],
                command=lambda current=option: self._toggle_value(current),
                style="Body.TCheckbutton",
            ).pack(anchor="w", fill="x", padx=4, pady=2)

    def _toggle_value(self, value: str) -> None:
        selected = set(self._popup_selected if self._popup_selected is not None else self._selected)
        if self._variables[value].get():
            selected.add(value)
        else:
            selected.discard(value)
        self._popup_selected = [item for item in self._available_options if item in selected]

    def _select_all(self) -> None:
        filtered = set(self._filtered_options())
        selected = set(self._popup_selected if self._popup_selected is not None else self._selected)
        selected.update(filtered)
        self._popup_selected = [item for item in self._available_options if item in selected]
        for option, variable in self._variables.items():
            variable.set(option in self._popup_selected)
        self._render_options()

    def _clear_all(self) -> None:
        filtered = set(self._filtered_options())
        current = self._popup_selected if self._popup_selected is not None else self._selected
        self._popup_selected = [item for item in current if item not in filtered]
        for option, variable in self._variables.items():
            variable.set(option in self._popup_selected)
        self._render_options()

    def _apply_popup(self) -> None:
        selected = list(self._popup_selected if self._popup_selected is not None else self._selected)
        self._selected = selected
        for option, variable in self._variables.items():
            variable.set(option in self._selected)
        self._sync_summary()
        if self.on_selection_changed:
            self.on_selection_changed(self.get_selected())
        self._close_popup()

    def _cancel_popup(self) -> None:
        original = list(self._popup_initial_selected if self._popup_initial_selected is not None else self._selected)
        self._selected = original
        for option, variable in self._variables.items():
            variable.set(option in self._selected)
        self._sync_summary()
        self._close_popup()

    def _sync_summary(self) -> None:
        if not self._selected:
            summary = self.empty_label
        elif len(self._selected) <= 2:
            summary = ", ".join(self._selected)
        else:
            summary = f"{len(self._selected)} seleccionadas"
        self._summary_var.set(summary)
        self.button.configure(text=self._build_button_text(summary))

    def _build_button_text(self, summary: str | None = None) -> str:
        caret = "\u25BE"
        if not summary:
            return f"{self.title} {caret}"
        return f"{self.title} ({summary}) {caret}"
