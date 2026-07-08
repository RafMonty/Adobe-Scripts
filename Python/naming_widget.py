
from __future__ import annotations

import sys
import tkinter as tk
from tkinter import ttk
from typing import List, Tuple


class NamingWidget(tk.Tk):
    """Small always-on-top widget to build standardised names."""

    SEPARATOR_OPTIONS: List[Tuple[str, str]] = [
        ("None", ""),
        ("Underscore _", "_"),
        ("Hyphen -", "-"),
        ("Space  ", " "),
        ("Colon :", ":"),
    ]

    DEFAULT_SEPARATOR_LABELS: List[str] = [
        "Hyphen -",       # after Client Abbreviation
        "Underscore _",   # after Supplier Abbreviation
        "Underscore _",   # after Product
        "Underscore _",   # after Type
        "Underscore _",   # after Options
    ]

    PLACEHOLDERS: List[str] = [
        "e.g. MAG",
        "e.g. ABC, ENG",
        "e.g. 2pp A4 Brochure",
        "e.g. Deluxe",
        "e.g. Landscape",
        "e.g. 2025",
    ]

    PH_FG = "#888"     # placeholder colour
    TXT_FG = "#000"    # normal text colour

    def __init__(self) -> None:
        super().__init__()
        self.title("Naming Widget")
        self.attributes("-topmost", True)
        self.resizable(False, False)
        self.configure(padx=10, pady=8)

        self.segment_labels: List[str] = [
            "Client Abbreviation",
            "Supplier Abbreviation",
            "Product",
            "Type",
            "Options",
            "Year",
        ]

        self.segment_vars: List[tk.StringVar] = [tk.StringVar(value="") for _ in self.segment_labels]
        self.separator_vars: List[tk.StringVar] = []
        self.entry_widgets: List[tk.Entry] = []  # keep references to manage placeholder styles

        for i in range(len(self.segment_labels) - 1):
            default_label = (
                self.DEFAULT_SEPARATOR_LABELS[i]
                if i < len(self.DEFAULT_SEPARATOR_LABELS)
                else self.SEPARATOR_OPTIONS[0][0]
            )
            self.separator_vars.append(tk.StringVar(value=default_label))

        self._build_ui()
        self._init_placeholders()
        self._bind_events()
        self._update_preview()

    # -------------------------- UI --------------------------
    def _build_ui(self) -> None:
        header = ttk.Frame(self)
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        ttk.Label(header, text="Consistent Name Builder", font=("Segoe UI", 10, "bold")).pack(side="left")
        self._make_drag_handle(header)

        ttk.Label(self, text="Field").grid(row=1, column=0, sticky="w")
        ttk.Label(self, text="Value").grid(row=1, column=1, sticky="w")
        ttk.Label(self, text="Separator after").grid(row=1, column=2, sticky="w")

        for i, label in enumerate(self.segment_labels):
            row = i + 2
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=2)

            # Use tk.Entry so we can control foreground colour per widget for placeholders
            entry = tk.Entry(self, textvariable=self.segment_vars[i], width=34)
            entry.grid(row=row, column=1, sticky="ew", padx=(4, 8), pady=2)
            entry.bind("<FocusIn>", lambda e, idx=i: self._on_focus_in(idx))
            entry.bind("<FocusOut>", lambda e, idx=i: self._on_focus_out(idx))
            self.entry_widgets.append(entry)

            if i < len(self.segment_labels) - 1:
                combo = ttk.Combobox(
                    self,
                    textvariable=self.separator_vars[i],
                    state="readonly",
                    width=16,
                    values=[opt[0] for opt in self.SEPARATOR_OPTIONS],
                )
                combo.grid(row=row, column=2, sticky="w", pady=2)
                combo.set(self.separator_vars[i].get())
            else:
                ttk.Label(self, text="—").grid(row=row, column=2, sticky="w", pady=2)

        ttk.Separator(self, orient="horizontal").grid(row=8, column=0, columnspan=3, sticky="ew", pady=(6, 6))

        ttk.Label(self, text="Result (auto-updates):").grid(row=9, column=0, sticky="w")
        self.preview_var = tk.StringVar(value="")
        ttk.Entry(self, textvariable=self.preview_var, width=52, state="readonly").grid(
            row=10, column=0, columnspan=3, sticky="ew", pady=(2, 0)
        )

        btns = ttk.Frame(self)
        btns.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(btns, text="Copy", command=self.copy_to_clipboard).pack(side="left")
        ttk.Button(btns, text="Clear", command=self.clear_all).pack(side="left", padx=(6, 0))

        ttk.Label(
            self,
            text="Tip: Empty fields are skipped; separators only appear between non-empty parts.",
            foreground="#555",
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _init_placeholders(self) -> None:
        # Set initial placeholder text/colour
        for i, entry in enumerate(self.entry_widgets):
            self._apply_placeholder(i)

    def _bind_events(self) -> None:
        for var in self.segment_vars:
            var.trace_add("write", lambda *_: self._update_preview())
        for var in self.separator_vars:
            var.trace_add("write", lambda *_: self._update_preview())

        self.bind_all("<Control-c>", lambda e: self.copy_to_clipboard())
        if sys.platform == "darwin":
            self.bind_all("<Command-c>", lambda e: self.copy_to_clipboard())

    def _make_drag_handle(self, widget: tk.Widget) -> None:
        def start_move(event):
            self._drag_start_x = event.x
            self._drag_start_y = event.y
        def on_move(event):
            x = self.winfo_x() + (event.x - self._drag_start_x)
            y = self.winfo_y() + (event.y - self._drag_start_y)
            self.geometry(f"+{x}+{y}")
        widget.bind("<Button-1>", start_move)
        widget.bind("<B1-Motion>", on_move)

    # ----------------------- Placeholder logic ----------------------
    def _on_focus_in(self, idx: int) -> None:
        if self._is_placeholder(idx):
            # Clear placeholder for user typing
            self.segment_vars[idx].set("")
            self.entry_widgets[idx].config(fg=self.TXT_FG)

    def _on_focus_out(self, idx: int) -> None:
        # Reinstate placeholder if user left it empty
        if not self.segment_vars[idx].get().strip():
            self._apply_placeholder(idx)

    def _apply_placeholder(self, idx: int) -> None:
        self.segment_vars[idx].set(self.PLACEHOLDERS[idx])
        self.entry_widgets[idx].config(fg=self.PH_FG)

    def _is_placeholder(self, idx: int) -> bool:
        return self.segment_vars[idx].get() == self.PLACEHOLDERS[idx]

    # ----------------------- Behaviour ----------------------
    def _update_preview(self) -> None:
        parts: List[str] = []
        clean_segments: List[str] = []
        for i, var in enumerate(self.segment_vars):
            value = var.get().strip()
            if value == self.PLACEHOLDERS[i]:
                value = ""  # treat placeholders as empty
            clean_segments.append(value)

        sep_values = [self._map_separator(v.get()) for v in self.separator_vars]

        for i, seg in enumerate(clean_segments):
            if not seg:
                continue
            parts.append(seg)
            if i < len(clean_segments) - 1:
                if any(clean_segments[j] for j in range(i + 1, len(clean_segments))):
                    sep = sep_values[i]
                    if sep:
                        parts.append(sep)

        self.preview_var.set("".join(parts))

    @staticmethod
    def _map_separator(label: str) -> str:
        lookup = {
            "None": "",
            "Underscore _": "_",
            "Hyphen -": "-",
            "Space  ": " ",
            "Colon :": ":",
        }
        return lookup.get(label, "")

    def copy_to_clipboard(self) -> None:
        text = self.preview_var.get()
        self.clipboard_clear()
        self.clipboard_append(text)
        original = self.title()
        self.title("✅ Copied")
        self.after(700, lambda: self.title(original))

    def clear_all(self) -> None:
        for i, var in enumerate(self.segment_vars):
            var.set("")
            self._apply_placeholder(i)
        self._update_preview()


if __name__ == "__main__":
    app = NamingWidget()
    app.mainloop()
