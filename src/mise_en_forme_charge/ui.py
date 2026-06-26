"""Fonctions UI Tkinter pour la CLI."""

from __future__ import annotations

import os
from pathlib import Path
from tkinter import BOTH, X, StringVar, TclError, Text, Tk, filedialog
from tkinter import ttk

try:
    from .profile_filters import (
        default_atelier_from_filters,
        format_filters_for_display,
        load_profile_filters,
        preferred_atelier,
    )
except ImportError:
    from profile_filters import (
        default_atelier_from_filters,
        format_filters_for_display,
        load_profile_filters,
        preferred_atelier,
    )


def pick_input_file() -> Path | None:
    root = Tk()
    root.withdraw()
    try:
        selected = filedialog.askopenfilename(
            title="Selectionner le fichier source",
            filetypes=[
                ("Fichiers Excel", "*.xlsm *.xlsx"),
                ("Fichiers texte", "*.txt"),
                ("Tous les fichiers", "*.*"),
            ],
        )
    except KeyboardInterrupt:
        selected = ""
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


def ask_atelier_selection(ateliers: list[str], profile_path: Path) -> str | None:
    if not ateliers:
        return None

    filters = load_profile_filters(profile_path)
    default_atelier = default_atelier_from_filters(filters)
    initial_atelier = preferred_atelier(ateliers, default_atelier)

    root = Tk()
    root.title("Selection atelier - page GLOBAL")
    root.geometry("860x520")
    root.resizable(True, True)
    root.configure(bg="#EAF3FF")

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Blue.TFrame", background="#EAF3FF")
    style.configure("Card.TFrame", background="#FFFFFF")
    style.configure(
        "Title.TLabel",
        background="#EAF3FF",
        foreground="#0B3A6E",
        font=("Segoe UI", 13, "bold"),
    )
    style.configure(
        "Hint.TLabel",
        background="#FFFFFF",
        foreground="#2A5B94",
        font=("Segoe UI", 9),
    )
    style.configure(
        "BlueLabel.TLabel",
        background="#EAF3FF",
        foreground="#143F73",
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "Primary.TButton",
        font=("Segoe UI", 10, "bold"),
        foreground="#FFFFFF",
        background="#1D5FAF",
        borderwidth=0,
        padding=(12, 8),
    )
    style.map(
        "Primary.TButton",
        background=[("active", "#174D8E"), ("pressed", "#123D72")],
    )
    style.configure(
        "Secondary.TButton",
        font=("Segoe UI", 10),
        foreground="#143F73",
        background="#DCEBFF",
        borderwidth=0,
        padding=(12, 8),
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#CAE0FF"), ("pressed", "#BED8FF")],
    )
    style.configure(
        "Blue.TCombobox",
        fieldbackground="#F8FBFF",
        background="#F8FBFF",
        foreground="#0B2C52",
        arrowsize=14,
        padding=4,
    )

    message_var = StringVar(
        master=root,
        value=format_filters_for_display(profile_path, initial_atelier),
    )
    selection_var = StringVar(master=root, value=initial_atelier)
    chosen_atelier = {"value": initial_atelier}

    main = ttk.Frame(root, style="Blue.TFrame", padding=14)
    main.pack(fill=BOTH, expand=True)

    header = ttk.Label(
        main,
        text="Choisir Atelier :",
        style="Title.TLabel",
    )
    header.pack(fill=X, pady=(0, 8))

    atelier_frame = ttk.Frame(main, style="Blue.TFrame", padding=(6, 6, 6, 10))
    atelier_frame.pack(fill=X)

    label = ttk.Label(atelier_frame, text="Atelier", style="BlueLabel.TLabel")
    label.pack(side="left")

    combo = ttk.Combobox(
        atelier_frame,
        textvariable=selection_var,
        values=ateliers,
        state="readonly",
        width=28,
        font=("Segoe UI", 10),
        style="Blue.TCombobox",
    )
    combo.pack(side="left", padx=(8, 0))

    info_frame = ttk.Frame(main, style="Blue.TFrame", padding=(6, 4, 6, 10))
    info_frame.pack(fill=BOTH, expand=True)

    filters_title = ttk.Label(
        info_frame,
        text="Filtres indicatifs",
        style="Hint.TLabel",
    )
    filters_title.pack(fill=X, pady=(4, 4))

    text = Text(
        info_frame,
        wrap="word",
        font=("Segoe UI", 9),
        height=8,
        bg="#F4F9FF",
        fg="#3B5D84",
        relief="flat",
        padx=10,
        pady=8,
    )
    text.insert("1.0", message_var.get())
    text.configure(state="disabled")
    text.pack(fill=BOTH, expand=True, pady=(0, 12))

    button_frame = ttk.Frame(main, style="Blue.TFrame")
    button_frame.pack(fill=X)
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=1)

    def _refresh_filters(*_: object) -> None:
        current_atelier = selection_var.get().strip()
        message = format_filters_for_display(profile_path, current_atelier)
        message_var.set(message)
        text.configure(state="normal")
        text.delete("1.0", "end")
        text.insert("1.0", message_var.get())
        text.configure(state="disabled")

    def _confirm_selection() -> None:
        chosen_atelier["value"] = selection_var.get().strip() or initial_atelier
        root.destroy()

    def _cancel_selection() -> None:
        chosen_atelier["value"] = initial_atelier
        root.destroy()

    selection_var.trace_add("write", _refresh_filters)

    cancel_button = ttk.Button(
        button_frame,
        text="Annuler",
        command=_cancel_selection,
        style="Secondary.TButton",
        width=20,
    )
    cancel_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))

    continue_button = ttk.Button(
        button_frame,
        text="Continuer",
        command=_confirm_selection,
        style="Primary.TButton",
        width=20,
    )
    continue_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

    def _focus_window() -> None:
        try:
            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            root.focus_force()
            root.after(300, lambda: root.attributes("-topmost", False))
        except TclError:
            return

    root.after(50, _focus_window)
    combo.focus_set()
    root.mainloop()
    return chosen_atelier["value"]


def open_generated_workbook(path: Path) -> None:
    if not path.exists():
        return
    try:
        os.startfile(str(path.resolve()))  # type: ignore[attr-defined]
    except Exception:
        try:
            from subprocess import Popen

            Popen(["cmd", "/c", "start", "", str(path.resolve())], shell=False)
        except Exception:
            print(f"Classeur genere (ouverture auto impossible): {path}")
