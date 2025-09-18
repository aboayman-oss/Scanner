"""Settings window for configuring application preferences."""
import json
import os
from tkinter import Listbox, filedialog, messagebox

import customtkinter as ctk
from customtkinter import (
    CTkButton,
    CTkCheckBox,
    CTkComboBox,
    CTkEntry,
    CTkFrame,
    CTkLabel,
    CTkRadioButton,
    CTkTabview,
    CTkToplevel,
)
from PIL import Image, ImageTk

from utils.helpers import (
    MAPPING_FILE,
    MIN_SETTINGS_SIZE,
    SETTINGS,
    SETTINGS_BG_FILE,
    SETTINGS_FILE,
    bring_window_to_front,
    ensure_initial_size,
)

class SettingsWindow(CTkToplevel):
    mapping_placeholder = "-- Select --"

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Settings")
        self.minsize(*MIN_SETTINGS_SIZE)

        self.original_bg = Image.open(SETTINGS_BG_FILE)
        self.bg_photo = ImageTk.PhotoImage(self.original_bg)
        self.bg_label = CTkLabel(self, text="", image=self.bg_photo)
        self.bg_label.place(relwidth=1, relheight=1)
        self.bind("<Configure>", self._on_resize)
        self.after(50, lambda: bring_window_to_front(self))

        self.parent_app = parent
        self.column_map = dict(parent.column_map or {})
        self.working_mapping = dict(self.column_map)
        self.mapping_fields = [
            ("Card ID", "card_id"),
            ("Student ID", "student_id"),
            ("Name", "name"),
            ("Phone no.", "phone"),
            ("Attendance", "attendance"),
            ("Notes", "notes"),
            ("Timestamp", "timestamp"),
            ("Exam", "exam"),
            ("Homework", "homework"),
        ]
        self.mapping_controls = {}
        self.mapping_source_path = None
        self.mapping_columns = []
        for value in self.working_mapping.values():
            if value and value not in self.mapping_columns:
                self.mapping_columns.append(value)

        self.stage_options = list(SETTINGS["stage_options"])
        self.center_options = list(SETTINGS["center_options"])
        self.var_exam = ctk.BooleanVar(value=SETTINGS["restrictions"].get("exam", False))
        self.var_homework = ctk.BooleanVar(value=SETTINGS["restrictions"].get("homework", False))
        self.var_file_type = ctk.StringVar(value=SETTINGS.get("file_type", "xlsx"))

        self.template_status_var = ctk.StringVar()

        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=(24, 72))

        self.tabview = CTkTabview(container)
        self.tabview.pack(fill="both", expand=True, padx=4, pady=4)

        self.template_tab = self.tabview.add("Template Mapping")
        self.stage_tab = self.tabview.add("Stage & Center")
        self.restrictions_tab = self.tabview.add("Restrictions")
        self.filetype_tab = self.tabview.add("File Type")

        self._build_template_tab()
        self._build_stage_tab()
        self._build_restrictions_tab()
        self._build_filetype_tab()

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", padx=24, pady=20)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        self.cancel_button = CTkButton(btn_frame, text="Cancel", command=self._cancel)
        self.cancel_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.apply_button = CTkButton(btn_frame, text="Apply", command=self._apply_settings, state="disabled")
        self.apply_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        if self.working_mapping:
            self.template_status_var.set("Using saved mapping. Load a sample file to update it.")
        else:
            self.template_status_var.set("Load a sample file to map template fields.")
        self._populate_template_controls()
        self._update_apply_state()
        ensure_initial_size(self, min_size=MIN_SETTINGS_SIZE)

    def _on_resize(self, event):
        if event.widget is self:
            resized = self.original_bg.resize((event.width, event.height), Image.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(resized)
            self.bg_label.configure(image=self.bg_photo)

    def _build_template_tab(self):
        CTkLabel(
            self.template_tab,
            text="Assign each template field to a column from a sample data file.",
            justify="left"
        ).pack(anchor="w", pady=(12, 8))
        option_row = CTkFrame(self.template_tab, fg_color="transparent")
        option_row.pack(fill="x", pady=(0, 12))
        CTkButton(option_row, text="Load Source File", width=140, command=self._prompt_for_columns).pack(side="left")
        CTkLabel(option_row, textvariable=self.template_status_var, justify="left", wraplength=360).pack(side="left", padx=12)

        form = CTkFrame(self.template_tab, fg_color="transparent")
        form.pack(fill="x", padx=4, pady=(4, 0))
        form.grid_columnconfigure(1, weight=1)
        self.template_form = form

        for idx, (label_text, field_key) in enumerate(self.mapping_fields):
            CTkLabel(form, text=f"{label_text}:").grid(row=idx, column=0, sticky="w", padx=(0, 12), pady=4)
            combo = CTkComboBox(form, state="readonly", values=[self.mapping_placeholder])
            combo.grid(row=idx, column=1, sticky="ew", pady=4)
            combo.set(self.mapping_placeholder)
            combo.configure(command=lambda value, key=field_key: self._on_mapping_change(key, value))
            self.mapping_controls[field_key] = combo

    def _populate_template_controls(self):
        available = []
        for col in self.mapping_columns:
            col = str(col).strip()
            if col and col not in available:
                available.append(col)
        values = [self.mapping_placeholder] + available if available else [self.mapping_placeholder]

        for field_key, combo in self.mapping_controls.items():
            combo.configure(values=values)
            current = self.working_mapping.get(field_key, "")
            if current and current in available:
                combo.set(current)
            else:
                combo.set(self.mapping_placeholder)
                self.working_mapping[field_key] = ""

    def _on_mapping_change(self, field_key, value):
        cleaned = "" if value in ("", self.mapping_placeholder) else value.strip()
        self.working_mapping[field_key] = cleaned
        self._update_apply_state()

    def _prompt_for_columns(self):
        file_type = self.var_file_type.get().lower()
        ext = "*.xlsx" if file_type == "xlsx" else "*.csv"
        path = filedialog.askopenfilename(
            parent=self,
            title=f"Select {file_type.upper()}",
            filetypes=[(f"{file_type.upper()} files", ext)]
        )
        if not path:
            return
        try:
            df = read_data(path, nrows=0)
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc), parent=self)
            return
        columns = [str(col).strip() for col in df.columns]
        self.mapping_columns = [col for col in columns if col]
        self.mapping_source_path = path
        self.template_status_var.set(f"Columns loaded from {os.path.basename(path)}.")
        self._populate_template_controls()
        self._update_apply_state()

    def _collect_mapping(self):
        mapping = {}
        for _, field_key in self.mapping_fields:
            value = self.mapping_controls[field_key].get().strip()
            if value == self.mapping_placeholder:
                value = ""
            mapping[field_key] = value
        return mapping

    def _is_mapping_valid(self):
        mapping = self._collect_mapping()
        values = list(mapping.values())
        if not values or "" in values:
            return False
        return len(values) == len(set(values))

    def _update_apply_state(self):
        state = "normal" if self._is_mapping_valid() else "disabled"
        self.apply_button.configure(state=state)

    def _build_stage_tab(self):
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            list_bg, list_fg = "#1f1f1f", "#f2f2f2"
        else:
            list_bg, list_fg = "#ffffff", "#1a1a1a"
        select_bg, select_fg = "#1f6aa5", "#ffffff"

        CTkLabel(
            self.stage_tab,
            text="Manage the stage and center choices available when starting a session."
        ).pack(anchor="w", pady=(12, 8))

        lists_frame = CTkFrame(self.stage_tab, fg_color="transparent")
        lists_frame.pack(fill="both", expand=True)
        lists_frame.grid_columnconfigure(0, weight=1)
        lists_frame.grid_columnconfigure(1, weight=1)
        lists_frame.grid_rowconfigure(1, weight=1)

        stage_frame = CTkFrame(lists_frame, fg_color="transparent")
        stage_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        CTkLabel(stage_frame, text="Stage Options").pack(anchor="w")
        self.stage_listbox = Listbox(
            stage_frame,
            height=8,
            bg=list_bg,
            fg=list_fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
            highlightthickness=0,
            relief="flat"
        )
        self.stage_listbox.pack(fill="both", expand=True, pady=(6, 8))
        for item in self.stage_options:
            self.stage_listbox.insert("end", item)
        stage_entry_row = CTkFrame(stage_frame, fg_color="transparent")
        stage_entry_row.pack(fill="x", pady=(0, 6))
        stage_entry_row.grid_columnconfigure(0, weight=1)
        self.stage_entry = CTkEntry(stage_entry_row, placeholder_text="Add stage")
        self.stage_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        CTkButton(stage_entry_row, text="Add", width=70, command=self._add_stage).grid(row=0, column=1, sticky="e")
        CTkButton(stage_frame, text="Remove Selected", command=self._remove_stage).pack(anchor="e")

        center_frame = CTkFrame(lists_frame, fg_color="transparent")
        center_frame.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        CTkLabel(center_frame, text="Center Options").pack(anchor="w")
        self.center_listbox = Listbox(
            center_frame,
            height=8,
            bg=list_bg,
            fg=list_fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
            highlightthickness=0,
            relief="flat"
        )
        self.center_listbox.pack(fill="both", expand=True, pady=(6, 8))
        for item in self.center_options:
            self.center_listbox.insert("end", item)
        center_entry_row = CTkFrame(center_frame, fg_color="transparent")
        center_entry_row.pack(fill="x", pady=(0, 6))
        center_entry_row.grid_columnconfigure(0, weight=1)
        self.center_entry = CTkEntry(center_entry_row, placeholder_text="Add center")
        self.center_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        CTkButton(center_entry_row, text="Add", width=70, command=self._add_center).grid(row=0, column=1, sticky="e")
        CTkButton(center_frame, text="Remove Selected", command=self._remove_center).pack(anchor="e")

    def _add_stage(self):
        value = self.stage_entry.get().strip()
        if not value:
            return
        existing = set(self.stage_listbox.get(0, "end"))
        if value in existing:
            self.stage_entry.delete(0, "end")
            return
        self.stage_listbox.insert("end", value)
        self.stage_entry.delete(0, "end")

    def _remove_stage(self):
        selection = self.stage_listbox.curselection()
        for index in reversed(selection):
            self.stage_listbox.delete(index)

    def _add_center(self):
        value = self.center_entry.get().strip()
        if not value:
            return
        existing = set(self.center_listbox.get(0, "end"))
        if value in existing:
            self.center_entry.delete(0, "end")
            return
        self.center_listbox.insert("end", value)
        self.center_entry.delete(0, "end")

    def _remove_center(self):
        selection = self.center_listbox.curselection()
        for index in reversed(selection):
            self.center_listbox.delete(index)

    def _build_restrictions_tab(self):
        CTkLabel(
            self.restrictions_tab,
            text="Toggle optional columns that should be collected during scans."
        ).pack(anchor="w", pady=(12, 8))
        CTkCheckBox(
            self.restrictions_tab,
            text="Enable Exam Column",
            variable=self.var_exam,
            command=self._update_apply_state
        ).pack(anchor="w", pady=6)
        CTkCheckBox(
            self.restrictions_tab,
            text="Enable Homework Column",
            variable=self.var_homework,
            command=self._update_apply_state
        ).pack(anchor="w", pady=6)

    def _build_filetype_tab(self):
        CTkLabel(
            self.filetype_tab,
            text="Choose the preferred format when importing or exporting data."
        ).pack(anchor="w", pady=(12, 8))
        CTkRadioButton(
            self.filetype_tab,
            text="CSV",
            variable=self.var_file_type,
            value="csv",
            command=self._update_apply_state
        ).pack(anchor="w", pady=6)
        CTkRadioButton(
            self.filetype_tab,
            text="XLSX",
            variable=self.var_file_type,
            value="xlsx",
            command=self._update_apply_state
        ).pack(anchor="w", pady=6)

    def _apply_settings(self):
        if not self._is_mapping_valid():
            messagebox.showerror("Invalid Mapping", "Each template field must map to a unique column.", parent=self)
            return

        mapping = self._collect_mapping()
        stage_options = list(self.stage_listbox.get(0, "end"))
        center_options = list(self.center_listbox.get(0, "end"))
        restrictions = {
            "exam": bool(self.var_exam.get()),
            "homework": bool(self.var_homework.get()),
        }
        file_type = self.var_file_type.get()

        try:
            with open(MAPPING_FILE, "w", encoding="utf-8") as file:
                json.dump(mapping, file, indent=2)
            SETTINGS["stage_options"] = stage_options
            SETTINGS["center_options"] = center_options
            SETTINGS["restrictions"].update(restrictions)
            SETTINGS["file_type"] = file_type
            with open(SETTINGS_FILE, "w", encoding="utf-8") as file:
                json.dump(SETTINGS, file, indent=2)
        except OSError as exc:
            messagebox.showerror("Save Failed", str(exc), parent=self)
            return

        self.parent_app.column_map = mapping
        self.column_map = dict(mapping)
        self.working_mapping = dict(mapping)
        self.mapping_columns = [value for value in mapping.values() if value]

        if hasattr(self.parent_app, "set_status"):
            self.parent_app.set_status("Settings saved.")
        messagebox.showinfo("Settings", "Settings saved successfully.", parent=self)
        self.on_close()

    def _cancel(self):
        if hasattr(self.parent_app, "set_status"):
            self.parent_app.set_status("Settings closed without saving.")
        self.on_close()

    def on_close(self):
        if getattr(self.parent_app, "settings_window", None) is self:
            self.parent_app.settings_window = None
        self.destroy()
