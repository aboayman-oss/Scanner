# CustomTkinter migration patch: UI widgets swapped from tkinter while logic stays identical. UX redesign comes later.
import os
import sys
import json
import subprocess
import customtkinter as ctk
from customtkinter import (
    CTk,
    CTkToplevel,
    CTkFrame,
    CTkButton,
    CTkLabel,
    CTkEntry,
    CTkCheckBox,
    CTkRadioButton,
    CTkProgressBar,
    CTkComboBox,
    CTkTabview,
)
from tkinter import ttk, filedialog, messagebox, Listbox
from PIL import Image, ImageTk
import pandas as pd
from datetime import datetime

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def get_runtime_base():
    """Return the folder containing the script or executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_assets_dir():
    """Locate bundled assets for development and PyInstaller builds.

    PyInstaller one-file executables extract bundled data into a temporary
    folder exposed via ``sys._MEIPASS``. We read assets from there so the
    binary stays portable.
    """
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', get_runtime_base())
        return os.path.join(base, 'assets')
    script_dir = get_runtime_base()
    local_assets = os.path.join(script_dir, 'assets')
    if os.path.exists(os.path.join(local_assets, 'logo.png')):
        return local_assets
    parent_dir = os.path.dirname(script_dir)
    parent_assets = os.path.join(parent_dir, 'assets')
    if os.path.exists(os.path.join(parent_assets, 'logo.png')):
        return parent_assets
    return local_assets


RUNTIME_BASE = get_runtime_base()
# Assets live beside the script during development and inside the temporary
# PyInstaller bundle when frozen, so we centralize their path resolution above.
ASSETS_DIR = get_assets_dir()

if getattr(sys, 'frozen', False):
    # Keep user-generated data next to the executable for portability.
    BASE_FOLDER = RUNTIME_BASE
else:
    BASE_FOLDER = os.path.dirname(ASSETS_DIR)

LOGO_FILE        = os.path.join(ASSETS_DIR, 'logo.png')
HOME_BG_FILE     = os.path.join(ASSETS_DIR, 'background.jpg')
SETTINGS_BG_FILE = os.path.join(ASSETS_DIR, 'backgroundnew.jpg')

DATA_FOLDER      = os.path.join(BASE_FOLDER, 'Data')
SESSIONS_FOLDER  = os.path.join(BASE_FOLDER, 'Sessions')
ARCHIVE_FOLDER   = os.path.join(BASE_FOLDER, 'Data archive')
MAPPING_FILE     = os.path.join(ARCHIVE_FOLDER, 'column_map.json')
SETTINGS_FILE    = os.path.join(ARCHIVE_FOLDER, 'app_settings.json')
LAST_DATA_FILE   = os.path.join(ARCHIVE_FOLDER, 'last_data.json')

MIN_DASHBOARD_SIZE     = (980, 640)
MIN_SCAN_SIZE          = (900, 560)
MIN_SETTINGS_SIZE      = (640, 480)
MIN_SESSION_SETUP_SIZE = (360, 240)
MIN_SUMMARY_SIZE       = (380, 320)
MIN_PAST_SESSIONS_SIZE = (720, 480)
for folder in (DATA_FOLDER, SESSIONS_FOLDER, ARCHIVE_FOLDER):
    os.makedirs(folder, exist_ok=True)

SETTINGS = {
    "stage_options":  ["2nd", "3rd"],
    "center_options": [
        "October", "Ferdous", "Helwan", "Hadayek Helwan",
        "Zayed", "Haram", "Dokki", "Maadi", "15 May"
    ],
    "restrictions": {"exam": True, "homework": True},
    "file_type": "xlsx"
}



def bring_window_to_front(window):
    """Raise a toplevel window above its siblings and give it focus."""
    if window is None:
        return
    try:
        window.deiconify()
    except Exception:
        pass
    try:
        window.lift()
    except Exception:
        pass
    try:
        window.focus_force()
    except Exception:
        pass
    try:
        window.attributes('-topmost', True)
        window.after_idle(lambda: window.attributes('-topmost', False))
    except Exception:
        pass


def ensure_initial_size(window, *, min_size=None, padding=(0, 0)):
    """Size a toplevel so its default geometry fits the current layout."""
    if window is None:
        return 0, 0
    window.update_idletasks()
    req_w = max(window.winfo_reqwidth(), window.winfo_width())
    req_h = max(window.winfo_reqheight(), window.winfo_height())
    pad_x, pad_y = padding if isinstance(padding, tuple) else (padding, padding)
    width = max(int(req_w + pad_x), 1)
    height = max(int(req_h + pad_y), 1)
    if min_size:
        min_w, min_h = min_size
        width = max(width, int(min_w))
        height = max(height, int(min_h))
    window.minsize(width, height)
    window.geometry(f"{width}x{height}")
    return width, height

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


class SessionSetupDialog(CTkToplevel):
    def __init__(self, parent, stages, centers, has_data, callback):
        super().__init__(parent)
        self.parent = parent
        self.stages = stages
        self.centers = centers
        self.callback = callback
        self.has_data = has_data
        self.title("Start New Session")
        self.resizable(False, False)
        self.minsize(*MIN_SESSION_SETUP_SIZE)
        self.transient(parent)
        self.grid_columnconfigure(1, weight=1)

        notice_text = (
            "Using the imported dataset for this session."
            if has_data
            else "No dataset imported yet. A blank roster will be created."
        )
        self.notice_var = ctk.StringVar(value=notice_text)
        self.error_var = ctk.StringVar(value="")

        self._build_form()
        ensure_initial_size(self, min_size=MIN_SESSION_SETUP_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Return>", lambda _e: self._on_submit())
        self.bind("<Escape>", lambda _e: self._on_cancel())
        self.after(10, self._center_on_parent)
        self.after(50, lambda: self.stage_cb.focus_set())

    def _build_form(self):
        CTkLabel(
            self,
            textvariable=self.notice_var,
            justify="left",
            anchor="w",
            wraplength=320
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 8))

        CTkLabel(self, text="Stage:").grid(row=1, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.stage_cb = CTkComboBox(self, values=self.stages, state="readonly")
        self.stage_cb.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(self, text="Center:").grid(row=2, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.center_cb = CTkComboBox(self, values=self.centers, state="readonly")
        self.center_cb.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(self, text="Session No.:").grid(row=3, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.session_ent = CTkEntry(self)
        self.session_ent.grid(row=3, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(
            self,
            textvariable=self.error_var,
            text_color=("#ff6b6b", "#b00020"),
            justify="left",
            anchor="w",
            wraplength=320
        ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=16, pady=(4, 0))

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(12, 16))
        CTkButton(btn_frame, text="Start Session", command=self._on_submit).pack(side="left", padx=(0, 8))
        CTkButton(btn_frame, text="Cancel", command=self._on_cancel).pack(side="left")

        if self.stages:
            self.stage_cb.set(self.stages[0])
        if self.centers:
            self.center_cb.set(self.centers[0])

    def _center_on_parent(self):
        self.update_idletasks()
        width = self.winfo_width() or self.winfo_reqwidth()
        height = self.winfo_height() or self.winfo_reqheight()
        px = self.parent.winfo_rootx()
        py = self.parent.winfo_rooty()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        x = px + max((pw - width) // 2, 0) if pw else px
        y = py + max((ph - height) // 2, 0) if ph else py
        self.geometry(f"{width}x{height}+{x}+{y}")
        bring_window_to_front(self)

    def _on_submit(self):
        stage = self.stage_cb.get().strip()
        center = self.center_cb.get().strip()
        session_no = self.session_ent.get().strip()
        if not stage or not center or not session_no.isdigit():
            self.error_var.set("Select stage, center, and enter a numeric session number.")
            return
        self.error_var.set("")
        payload = {
            "stage": stage,
            "center": center,
            "no": int(session_no),
            "name": f"{stage} {center} session {int(session_no)}"
        }
        if self.callback:
            self.callback(payload)
            self.callback = None
        self.destroy()

    def _on_cancel(self):
        if self.callback:
            self.callback(None)
            self.callback = None
        self.destroy()

class SessionManager:
    def __init__(self, name, params, column_map, data_df):
        self.params       = params
        self.name         = name
        self.mapping      = column_map
        self.data_df      = data_df
        self.records      = []
        self.restrictions = SETTINGS["restrictions"]
        # Use the correct extension based on SETTINGS
        file_type = SETTINGS.get("file_type", "csv")
        ext = "xlsx" if file_type == "xlsx" else "csv"
        self.session_path = os.path.join(SESSIONS_FOLDER, f"{name}.{ext}")
        if os.path.exists(self.session_path):
            df = read_data(self.session_path)
            # Only keep mapped columns
            mapped_keys = ["card_id", "student_id", "name", "phone", "attendance", "notes"]
            if self.restrictions.get("exam"):
                mapped_keys.append("exam")
            if self.restrictions.get("homework"):
                mapped_keys.append("homework")
            self.records = []
            for _, row in df.iterrows():
                rec = {}
                for k in mapped_keys:
                    col = self.mapping.get(k, k)
                    rec[k] = row.get(col, "")
                self.records.append(rec)

    def add_record(self, rec):
        df = read_data(self.session_path)
        card_col     = self.mapping.get("card_id", "card_id")
        att_col      = self.mapping.get("attendance", "attendance")
        notes_col    = self.mapping.get("notes", "notes")
        timestamp_col= self.mapping.get("timestamp", "timestamp")
        mask = df[card_col].astype(str) == str(rec["card_id"])

        if mask.any():
            # --- Preserve timestamp if already present ---
            existing_timestamp = ""
            if timestamp_col in df.columns:
                existing_timestamp = df.loc[mask, timestamp_col].values[0]
            # Only overwrite if rec["timestamp"] is not empty
            if rec.get("timestamp"):
                df.loc[mask, timestamp_col] = rec["timestamp"]
            else:
                df.loc[mask, timestamp_col] = existing_timestamp
            df.loc[mask, att_col]   = rec["attendance"]
            df.loc[mask, notes_col] = rec["notes"]
        else:
            row = {col: "" for col in df.columns}
            for k in ("card_id", "student_id", "name", "phone"):
                col_name = self.mapping.get(k, k)
                if col_name in df.columns:
                    row[col_name] = rec.get(k, "")
            row[att_col]      = rec["attendance"]
            row[notes_col]    = rec["notes"]
            row[timestamp_col]= rec.get("timestamp", "")
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        write_data(df, self.session_path)

def read_data(path, **kwargs):
    if path.lower().endswith(".xlsx"):
        return pd.read_excel(path, dtype=str, **kwargs)
    else:
        return pd.read_csv(path, dtype=str, **kwargs)

def write_data(df, path, **kwargs):
    if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False, **kwargs)
    else:
        df.to_csv(path, index=False, **kwargs)

class PastSessionsWindow(CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Past Sessions")
        self.minsize(*MIN_PAST_SESSIONS_SIZE)
        self._paths = {}
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(50, lambda: bring_window_to_front(self))
        header = CTkLabel(
            self,
            text="Past Sessions",
            font=("Arial", 20, "bold")
        )
        header.pack(anchor="w", padx=24, pady=(24, 12))

        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        columns = ("name", "modified", "size")
        self.tree = ttk.Treeview(container, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="Session")
        self.tree.column("name", anchor="w", width=280)
        self.tree.heading("modified", text="Last Modified")
        self.tree.column("modified", anchor="center", width=160)
        self.tree.heading("size", text="Size")
        self.tree.column("size", anchor="center", width=100)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.empty_label = CTkLabel(
            container,
            text="No session files found.",
            font=("Arial", 14)
        )
        self.empty_label.place_forget()

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _e: self._open_selected())

        button_bar = CTkFrame(self, fg_color="transparent")
        button_bar.pack(fill="x", padx=24, pady=(0, 24))
        button_bar.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="past_actions")

        self.open_btn = CTkButton(
            button_bar,
            text="Open in Scanner",
            state="disabled",
            command=self._open_selected
        )
        self.open_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.reveal_btn = CTkButton(
            button_bar,
            text="Reveal in Folder",
            state="disabled",
            command=self._reveal_selected
        )
        self.reveal_btn.grid(row=0, column=1, padx=6, sticky="ew")

        self.refresh_btn = CTkButton(button_bar, text="Refresh", command=self.refresh)
        self.refresh_btn.grid(row=0, column=2, padx=6, sticky="ew")

        self.clear_btn = CTkButton(
            button_bar,
            text="Clear All",
            state="disabled",
            command=self._clear_all_sessions
        )
        self.clear_btn.grid(row=0, column=3, padx=6, sticky="ew")

        self.close_btn = CTkButton(button_bar, text="Close", command=self._on_close)
        self.close_btn.grid(row=0, column=4, padx=(6, 0), sticky="ew")

        self.refresh()
        ensure_initial_size(self, min_size=MIN_PAST_SESSIONS_SIZE)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._paths.clear()
        if not os.path.isdir(SESSIONS_FOLDER):
            self._toggle_empty_state(True)
            self._on_select()
            self._update_clear_state()
            return
        files = []
        for entry in os.listdir(SESSIONS_FOLDER):
            path_entry = os.path.join(SESSIONS_FOLDER, entry)
            if os.path.isfile(path_entry) and entry.lower().endswith((".csv", ".xlsx")):
                stats = os.stat(path_entry)
                files.append((path_entry, stats.st_mtime, stats.st_size))
        files.sort(key=lambda item: item[1], reverse=True)
        for index, (path_entry, modified, size) in enumerate(files):
            name = os.path.splitext(os.path.basename(path_entry))[0]
            stamp = datetime.fromtimestamp(modified).strftime("%d %b %Y %H:%M")
            size_text = self._format_size(size)
            iid = f"past_{index}"
            self.tree.insert("", "end", iid=iid, values=(name, stamp, size_text))
            self._paths[iid] = path_entry
        self._toggle_empty_state(len(self._paths) == 0)
        self._on_select()
        self._update_clear_state()

    def _update_clear_state(self):
        state = "normal" if self._paths else "disabled"
        self.clear_btn.configure(state=state)

    def _toggle_empty_state(self, show):
        if show:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.empty_label.place_forget()

    def _format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"

    def _on_select(self, _event=None):
        selection = self.tree.selection()
        state = "normal" if selection else "disabled"
        self.open_btn.configure(state=state)
        self.reveal_btn.configure(state=state)

    def _get_selected_path(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return self._paths.get(selection[0])

    def _open_selected(self):
        path_entry = self._get_selected_path()
        if not path_entry:
            return
        opened = self.parent._open_session_path(path_entry, read_only=True)
        if opened:
            self.parent._refresh_recent_sessions()

    def _reveal_selected(self):
        path_entry = self._get_selected_path()
        if not path_entry:
            return
        self.parent._reveal_session_path(path_entry)

    def _clear_all_sessions(self):
        if not self._paths:
            return
        confirm = messagebox.askyesno(
            "Clear All Sessions",
            "Delete all saved session files?",
            parent=self
        )
        if not confirm:
            return
        failures = []
        for path_entry in list(self._paths.values()):
            try:
                os.remove(path_entry)
            except Exception as exc:
                failures.append(f"{os.path.basename(path_entry)}: {exc}")
        self.refresh()
        if hasattr(self.parent, "_refresh_recent_sessions"):
            try:
                self.parent._refresh_recent_sessions()
            except Exception:
                pass
        if failures:
            messagebox.showerror(
                "Delete Failed",
                "Some session files could not be deleted:\n" + "\n".join(failures),
                parent=self
            )
            if hasattr(self.parent, "set_status"):
                self.parent.set_status("Some past sessions could not be removed.")
            return
        messagebox.showinfo(
            "Sessions Cleared",
            "All past session files have been deleted.",
            parent=self
        )
        if hasattr(self.parent, "set_status"):
            self.parent.set_status("All past sessions cleared.")

    def _on_close(self):
        if hasattr(self.parent, "past_sessions_window"):
            self.parent.past_sessions_window = None
        self.destroy()

class App(CTk):
    def __init__(self):
        super().__init__()
        self.title("RFID Attendance Manager")
        self.column_map = {}
        self.data_df    = None
        self.settings_window = None  # <-- Track settings window
        self.data_panel = None
        self.data_rows_var = ctk.StringVar(value="")
        self.data_path_var = ctk.StringVar(value="")
        self.current_data_path = None
        self._session_setup = None
        self.past_sessions_window = None
        self.summary_window = None

        if os.path.exists(MAPPING_FILE):
            with open(MAPPING_FILE) as f:
                self.column_map = json.load(f)
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE) as f:
                SETTINGS.update(json.load(f))

        self._build_ui()
        width, height = ensure_initial_size(self, min_size=MIN_DASHBOARD_SIZE)
        self.minsize(width, height)
        self._load_last_data()

    def _build_ui(self):
        self.main_frame = CTkFrame(self, corner_radius=0)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=0)

        content = CTkFrame(self.main_frame, fg_color="transparent")
        content.grid(row=0, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=0)
        content.grid_rowconfigure(1, weight=0)
        content.grid_rowconfigure(2, weight=1)

        header = CTkFrame(content, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(1, weight=1)

        logo = Image.open(LOGO_FILE)
        ratio = 56 / logo.width if logo.width else 1
        logo = logo.resize((56, int(logo.height * ratio)), Image.LANCZOS)
        self.logo_img = ImageTk.PhotoImage(logo)
        CTkLabel(header, image=self.logo_img, text="").grid(row=0, column=0, sticky="w", padx=(0, 12))

        title_holder = CTkFrame(header, fg_color="transparent")
        title_holder.grid(row=0, column=1, sticky="w")
        self.title_label = CTkLabel(
            title_holder,
            text="RFID Attendance Manager",
            font=("Arial", 24, "bold")
        )
        self.title_label.pack(anchor="w")
        CTkLabel(
            title_holder,
            text="Start scans, review sessions, and adjust preferences from one place.",
            font=("Arial", 14)
        ).pack(anchor="w", pady=(4, 0))

        self._build_data_status_panel(content)

        actions_frame = CTkFrame(content, fg_color="transparent")
        actions_frame.grid(row=2, column=0, sticky="nsew")
        actions_frame.grid_columnconfigure((0, 1), weight=1, uniform="actions")
        actions_frame.grid_rowconfigure((0, 1), weight=1)

        button_specs = [
            ("Start New Session", self.open_scan_window),
            ("View Past Sessions", self.view_past_sessions),
            ("Settings", self.open_settings),
        ]
        self.dashboard_buttons = []
        for index, (label, handler) in enumerate(button_specs):
            row, col = divmod(index, 2)
            btn = CTkButton(
                actions_frame,
                text=label,
                command=handler,
                height=120,
                font=("Arial", 18, "bold")
            )
            btn.grid(row=row, column=col, padx=12, pady=12, sticky="nsew")
            self.dashboard_buttons.append(btn)

        self.status_var = ctk.StringVar(value="Ready.")
        self.status_label = CTkLabel(
            self.main_frame,
            textvariable=self.status_var,
            anchor="w",
            font=("Arial", 12)
        )
        self.status_label.grid(row=1, column=0, sticky="ew", pady=(16, 0))

        self._recent_session_paths = {}
        self._refresh_recent_sessions()

    def _build_data_status_panel(self, parent):
        panel = CTkFrame(parent, fg_color=("#f2f3f5", "#1f2933"), corner_radius=12)
        panel.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        panel.grid_columnconfigure(0, weight=1)
        CTkLabel(
            panel,
            text="Data Loaded",
            font=("Arial", 18, "bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        CTkLabel(
            panel,
            textvariable=self.data_rows_var,
            font=("Arial", 14),
            anchor="w"
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 2))
        CTkLabel(
            panel,
            textvariable=self.data_path_var,
            font=("Arial", 12),
            anchor="w",
            justify="left",
            wraplength=520
        ).grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 12))
        panel.grid_remove()
        self.data_panel = panel

    def _update_data_status_panel(self, path, rows):
        if not self.data_panel:
            return
        self.data_rows_var.set(f"Rows: {rows:,}")
        self.data_path_var.set(f"File: {path}")
        self.data_panel.grid()
        self.current_data_path = path

    def _hide_data_status_panel(self):
        if self.data_panel:
            self.data_panel.grid_remove()
        self.data_rows_var.set("")
        self.data_path_var.set("")
        self.current_data_path = None


    def show_session_summary(self, *, session_name, summary, session_path, read_only=False):
        if self.summary_window is not None and self.summary_window.winfo_exists():
            try:
                self.summary_window.destroy()
            except Exception:
                pass
        self.summary_window = SessionSummaryDialog(
            self,
            session_name=session_name,
            summary=summary,
            session_path=session_path,
            read_only=read_only,
        )

    def set_status(self, message):
        if hasattr(self, "status_var"):
            self.status_var.set(message)

    def _refresh_recent_sessions(self):
        if not hasattr(self, "recent_tree"):
            return
        for item in self.recent_tree.get_children():
            self.recent_tree.delete(item)
        self._recent_session_paths = {}
        if not os.path.isdir(SESSIONS_FOLDER):
            self._on_recent_select()
            return
        files = []
        for entry in os.listdir(SESSIONS_FOLDER):
            path_entry = os.path.join(SESSIONS_FOLDER, entry)
            if os.path.isfile(path_entry) and entry.lower().endswith((".csv", ".xlsx")):
                files.append((path_entry, os.path.getmtime(path_entry)))
        files.sort(key=lambda item: item[1], reverse=True)
        for index, (path_entry, modified) in enumerate(files[:10]):
            name = os.path.splitext(os.path.basename(path_entry))[0]
            stamp = datetime.fromtimestamp(modified).strftime("%d %b %Y %H:%M")
            iid = f"recent_{index}"
            self.recent_tree.insert("", "end", iid=iid, values=(name, stamp))
            self._recent_session_paths[iid] = path_entry
        self._on_recent_select()

    def _on_recent_select(self, _event=None):
        selection = self.recent_tree.selection() if hasattr(self, "recent_tree") else ()
        state = "normal" if selection else "disabled"
        if hasattr(self, "recent_open_button"):
            self.recent_open_button.configure(state=state)
        if hasattr(self, "recent_reveal_button"):
            self.recent_reveal_button.configure(state=state)

    def _get_selected_session_path(self):
        if not hasattr(self, "recent_tree"):
            return None
        selection = self.recent_tree.selection()
        if not selection:
            return None
        return self._recent_session_paths.get(selection[0])

    def _open_session_path(self, path_entry, *, read_only=False):
        try:
            name = os.path.splitext(os.path.basename(path_entry))[0]
            df = read_data(path_entry)
            sm = SessionManager(name, {}, self.column_map, df)
            ScanWindow(self, sm, read_only=read_only)
            if read_only:
                self.set_status(f"Session '{name}' opened in view-only mode.")
            else:
                self.set_status(f"Session '{name}' opened.")
            return True
        except Exception as exc:
            messagebox.showerror("Open Failed", str(exc))
            self.set_status("Failed to open session.")
            return False

    def _reveal_session_path(self, path_entry):
        try:
            if sys.platform.startswith("win"):
                target = os.path.normpath(path_entry)
                explorer_cmd = f'/select,"{target}"'
                subprocess.Popen(["explorer", explorer_cmd])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path_entry])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(path_entry)])
            self.set_status(f"Revealed '{os.path.basename(path_entry)}'.")
            return True
        except Exception as exc:
            messagebox.showerror("Reveal Failed", str(exc))
            self.set_status("Failed to reveal session.")
            return False

    def _open_selected_session(self):
        path_entry = self._get_selected_session_path()
        if not path_entry:
            return
        self._open_session_path(path_entry, read_only=True)

    def _reveal_selected_session(self):
        path_entry = self._get_selected_session_path()
        if not path_entry:
            return
        self._reveal_session_path(path_entry)

    def view_past_sessions(self):
        if self.past_sessions_window is not None and self.past_sessions_window.winfo_exists():
            bring_window_to_front(self.past_sessions_window)
            return
        self.past_sessions_window = PastSessionsWindow(self)
        bring_window_to_front(self.past_sessions_window)
        self.set_status("Browsing past sessions.")


    def _load_last_data(self):
        self.data_df = None  # Always reset on startup
        self._hide_data_status_panel()
        if not os.path.exists(LAST_DATA_FILE):
            self.set_status("Ready.")
            return
        try:
            with open(LAST_DATA_FILE) as f:
                info = json.load(f)
        except Exception:
            self.set_status("Ready.")
            return
        path = info.get("path")
        if not path or not os.path.exists(path):
            self.set_status("Ready.")
            return
        try:
            df = read_data(path)
        except Exception:
            self.set_status("Ready.")
            return
        self.data_df = df
        self._update_data_status_panel(path, len(df))
        self.set_status(f"Restored {len(df)} records from {os.path.basename(path)}.")

    def open_settings(self):
        # Only open one settings window at a time
        if self.settings_window is not None and self.settings_window.winfo_exists():
            bring_window_to_front(self.settings_window)
            return
        self.settings_window = SettingsWindow(self)
        bring_window_to_front(self.settings_window)
        self.settings_window.protocol("WM_DELETE_WINDOW", self._on_settings_close)

    def _on_settings_close(self):
        if self.settings_window is not None:
            if hasattr(self.settings_window, 'on_close'):
                self.settings_window.on_close()
            else:
                self.settings_window.destroy()
                self.settings_window = None

    def import_csv(self):
        # Check if template is configured
        if not self.column_map:
            messagebox.showwarning("No Template", "Please configure a template first.")
            self.set_status("Import canceled - configure column template first.")
            return False

        file_type = SETTINGS.get("file_type", "csv")
        ext = "*.xlsx" if file_type == "xlsx" else "*.csv"
        path = filedialog.askopenfilename(
            title=f"Select {file_type.upper()}",
            filetypes=[(f"{file_type.upper()} files", ext)]
        )
        if not path:
            self.set_status("Import canceled.")
            return False
        try:
            df = read_data(path)
            # Pad card_id column to 8 digits and assign 'null N' for blanks
            card_col = self.column_map.get("card_id", "card_id")
            if card_col in df.columns:
                null_counter = 1
                new_card_ids = []
                for val in df[card_col]:
                    val_str = str(val).strip()
                    if not val_str or val_str.lower() == "nan":
                        new_card_ids.append(f"null {null_counter}")
                        null_counter += 1
                    elif val_str.isdigit():
                        new_card_ids.append(val_str.zfill(8))
                    else:
                        new_card_ids.append(val_str)
                df[card_col] = new_card_ids

            # Clear attendance and timestamp columns for imported data only
            att_col = self.column_map.get("attendance", "attendance")
            ts_col  = self.column_map.get("timestamp", "timestamp")
            if att_col in df.columns:
                df[att_col] = ""
            if ts_col in df.columns:
                df[ts_col] = ""
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            self.set_status("Import failed.")
            return False
        self.data_df = df
        with open(LAST_DATA_FILE, "w") as f:
            json.dump({"path": path}, f, indent=2)
        self._update_data_status_panel(path, len(df))
        self.set_status(f"Imported {len(df)} records from {os.path.basename(path)}.")
        return True

    def open_scan_window(self):
        if self._session_setup is not None and self._session_setup.winfo_exists():
            bring_window_to_front(self._session_setup)
            return

        if not self.import_csv():
            return

        self._session_setup = SessionSetupDialog(
            self,
            SETTINGS["stage_options"],
            SETTINGS["center_options"],
            has_data=self.data_df is not None,
            callback=self._on_session_setup_finished,
        )

    def _on_session_setup_finished(self, payload):
        self._session_setup = None
        if not payload:
            self.set_status("Session setup canceled.")
            return
        if self.data_df is None:
            messagebox.showwarning("No Data", "Please import data before starting a session.")
            self.set_status("Session setup aborted - no data loaded.")
            return
        name = payload["name"]
        params = {"stage": payload["stage"], "center": payload["center"], "no": payload["no"]}
        file_type = SETTINGS.get("file_type", "csv")
        ext = "xlsx" if file_type == "xlsx" else "csv"
        session_path = os.path.join(SESSIONS_FOLDER, f"{name}.{ext}")
        created = False
        if not os.path.exists(session_path):
            write_data(self.data_df, session_path)
            created = True
        session_df = read_data(session_path)
        sm = SessionManager(name, params, self.column_map, session_df)
        self._refresh_recent_sessions()
        if self.past_sessions_window is not None and self.past_sessions_window.winfo_exists():
            self.past_sessions_window.refresh()
        ScanWindow(self, sm)
        if created:
            self.set_status(f"Session '{name}' created.")
        else:
            self.set_status(f"Session '{name}' ready.")

class AddStudentDialog(CTkToplevel):
    def __init__(self, parent, *, card_id=None, on_submit=None, duplicate_checker=None, default_notes="manual addition"):
        super().__init__(parent)
        self.parent = parent
        self.card_id = card_id.zfill(8) if card_id and card_id.isdigit() else card_id
        self._on_submit = on_submit
        self._duplicate_checker = duplicate_checker
        self._default_notes = default_notes or "manual addition"
        self._focus_guard_restored = False

        self.title("Add Student")
        self.minsize(*MIN_SUMMARY_SIZE)
        self.transient(parent)
        self.grab_set()
        self.after(40, self._activate_modal)

        container = CTkFrame(self, corner_radius=16, fg_color=("#f4f6fb", "#1a1d23"))
        container.pack(fill="both", expand=True, padx=24, pady=24)
        container.grid_columnconfigure(0, weight=1)

        header = CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        CTkLabel(header, text="Add Student", font=("Arial", 22, "bold")).grid(row=0, column=0, sticky="w")
        subtitle = "Link a scanned card to a student profile." if self.card_id else "Create a manual record for a student."
        CTkLabel(header, text=subtitle, font=("Arial", 13)).grid(row=1, column=0, sticky="w", pady=(4, 0))

        if self.card_id:
            CTkLabel(
                header,
                text=f"Card ID: {self.card_id}",
                font=("Arial", 12, "bold"),
                text_color="#1f6aa5"
            ).grid(row=2, column=0, sticky="w", pady=(12, 0))

        form = CTkFrame(container, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew", pady=(20, 0))
        form.grid_columnconfigure(0, weight=0)
        form.grid_columnconfigure(1, weight=1)

        self.inputs = {}
        field_specs = [
            ("student_id", "Student ID", "e.g. 102534"),
            ("name", "Student Name", "Full name"),
            ("phone", "Phone Number", "e.g. 01012345678"),
        ]
        for row_index, (key, label_text, placeholder) in enumerate(field_specs):
            CTkLabel(form, text=label_text, font=("Arial", 12, "bold")).grid(
                row=row_index, column=0, sticky="w", padx=(0, 12), pady=(0, 10)
            )
            entry = CTkEntry(form, placeholder_text=placeholder)
            entry.grid(row=row_index, column=1, sticky="ew", pady=(0, 10))
            self.inputs[key] = entry

        self.feedback_var = ctk.StringVar(value="")
        self.feedback_label = CTkLabel(
            container,
            textvariable=self.feedback_var,
            font=("Arial", 12),
            text_color="#d64b4b"
        )
        self.feedback_label.grid(row=2, column=0, sticky="w", pady=(4, 0))

        actions = CTkFrame(container, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(24, 0))
        actions.grid_columnconfigure((0, 1), weight=1, uniform="dialog_actions")

        self.cancel_button = CTkButton(
            actions,
            text="Cancel",
            command=self._on_cancel,
            fg_color=("#e5e7eb", "#2d2f36"),
            hover_color=("#d1d5db", "#363a45"),
            text_color=("#0f172a", "#f8fafc")
        )
        self.cancel_button.grid(row=0, column=0, padx=(0, 12), sticky="ew")

        self.confirm_button = CTkButton(actions, text="Add Student", command=self._on_confirm)
        self.confirm_button.grid(row=0, column=1, sticky="ew")

        ensure_initial_size(self, min_size=MIN_SUMMARY_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._first_entry = next(iter(self.inputs.values()), None)
        self.bind("<Return>", lambda _event: self._on_confirm())
        self.bind("<Escape>", lambda _event: self._on_cancel())

    def _set_feedback(self, message, level="error"):
        colors = {
            "error": "#d64b4b",
            "warning": "#d67b2d",
            "info": "#1f6aa5",
        }
        color = colors.get(level, "#d64b4b")
        self.feedback_label.configure(text_color=color)
        self.feedback_var.set(message)

    def _activate_modal(self):
        bring_window_to_front(self)
        try:
            self.grab_set()
        except Exception:
            pass
        try:
            self.focus_force()
        except Exception:
            pass
        if self._first_entry is not None and self._first_entry.winfo_exists():
            self._first_entry.focus_set()

    def _on_confirm(self):
        values = {key: entry.get().strip() for key, entry in self.inputs.items()}
        if not all(values.values()):
            self._set_feedback("Please fill in every field before continuing.", level="warning")
            return
        if self._duplicate_checker:
            try:
                id_exists, phone_exists = self._duplicate_checker(values["student_id"], values["phone"])
            except Exception as exc:
                self._set_feedback(str(exc))
                return
            if id_exists or phone_exists:
                if id_exists and phone_exists:
                    msg = "This student ID and phone number already exist."
                elif id_exists:
                    msg = "This student ID already exists."
                else:
                    msg = "This phone number already exists."
                self._set_feedback(msg, level="warning")
                return
        if not self._on_submit:
            self._finalize()
            return
        try:
            outcome = self._on_submit(card_id=self.card_id, values=values, default_notes=self._default_notes)
        except Exception as exc:
            self._set_feedback(str(exc))
            return
        if outcome is None:
            success, message = True, None
        elif isinstance(outcome, tuple):
            success, message = outcome
        else:
            success, message = bool(outcome), None
        if success:
            self._set_feedback("")
            self._finalize()
        else:
            self._set_feedback(message or "Unable to add student.", level="error")

    def _finalize(self):
        try:
            self.grab_release()
        except Exception:
            pass
        if not getattr(self, "_focus_guard_restored", False):
            if hasattr(self.parent, "_resume_focus_guard"):
                self.parent._resume_focus_guard()
            self._focus_guard_restored = True
        if self.winfo_exists():
            self.destroy()
        if hasattr(self.parent, "scan_entry"):
            self.parent.after(120, self.parent.scan_entry.focus_set)

    def _on_cancel(self):
        self._set_feedback("")
        self._finalize()

class SessionSummaryDialog(CTkToplevel):
    def __init__(self, parent, *, session_name, summary, session_path, read_only=False):
        super().__init__(parent)
        self.parent = parent
        self.session_name = session_name
        self.summary = summary or {}
        self.session_path = session_path
        self.read_only = read_only

        self.title("Session Summary")
        self.minsize(*MIN_SUMMARY_SIZE)
        self.transient(parent)
        self.grab_set()
        self.after(40, lambda: bring_window_to_front(self))

        container = CTkFrame(self, corner_radius=16, fg_color=("#f4f6fb", "#1a1d23"))
        container.pack(fill="both", expand=True, padx=24, pady=24)
        container.grid_columnconfigure(0, weight=1)

        header = CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        CTkLabel(header, text="Session Summary", font=("Arial", 22, "bold")).grid(row=0, column=0, sticky="w")
        subtitle = f"{session_name}" if session_name else "Session details"
        CTkLabel(header, text=subtitle, font=("Arial", 13)).grid(row=1, column=0, sticky="w", pady=(4, 0))
        if read_only:
            CTkLabel(header, text="Read-only session", font=("Arial", 12), text_color="#64748b").grid(row=2, column=0, sticky="w", pady=(6, 0))

        metrics_frame = CTkFrame(container, fg_color="transparent")
        metrics_frame.grid(row=1, column=0, sticky="ew", pady=(18, 12))
        metrics_frame.grid_columnconfigure(0, weight=1)

        metrics = []
        total = self.summary.get("total")
        if total is not None:
            metrics.append(("Total students", f"{total:,}"))
        attended = self.summary.get("attended")
        if attended is not None:
            metrics.append(("Attended", f"{attended:,}"))
        rate = self.summary.get("attendance_rate")
        if rate is not None:
            metrics.append(("Attendance rate", rate))
        manual = self.summary.get("manual_additions")
        if manual is not None:
            metrics.append(("Manual additions", f"{manual:,}"))
        cancels = self.summary.get("cancellations")
        if cancels is not None:
            metrics.append(("Cancellations", f"{cancels:,}"))
        if "missing_exam" in self.summary:
            metrics.append(("Missing exam", f"{self.summary['missing_exam']:,}"))
        if "missing_hw" in self.summary:
            metrics.append(("Missing homework", f"{self.summary['missing_hw']:,}"))

        for row_index, (label_text, value_text) in enumerate(metrics):
            row = CTkFrame(metrics_frame, fg_color="transparent")
            row.grid(row=row_index, column=0, sticky="ew", pady=(0, 6))
            CTkLabel(row, text=label_text, font=("Arial", 12, "bold"), anchor="w").grid(row=0, column=0, sticky="w")
            CTkLabel(row, text=value_text, font=("Arial", 12), anchor="w").grid(row=1, column=0, sticky="w")

        path_frame = CTkFrame(container, fg_color="transparent")
        path_frame.grid(row=2, column=0, sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)
        CTkLabel(path_frame, text="Saved file", font=("Arial", 12, "bold"), anchor="w").grid(row=0, column=0, sticky="w")
        display_path = session_path if session_path else "Unavailable"
        CTkLabel(path_frame, text=display_path, font=("Arial", 11), anchor="w", wraplength=320, justify="left").grid(row=1, column=0, sticky="ew", pady=(2, 0))

        actions = CTkFrame(container, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(24, 0))
        actions.grid_columnconfigure((0, 1), weight=1, uniform="summary_actions")

        self.view_button = CTkButton(actions, text="View File Location", command=self._open_location)
        self.view_button.grid(row=0, column=0, padx=(0, 12), sticky="ew")
        if not session_path:
            self.view_button.configure(state="disabled")

        close_button = CTkButton(actions, text="Close", command=self._on_close)
        close_button.grid(row=0, column=1, sticky="ew")

        ensure_initial_size(self, min_size=MIN_SUMMARY_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _open_location(self):
        path = self.session_path
        if not path:
            return
        abs_path = os.path.abspath(path)
        if os.path.isdir(abs_path):
            target = abs_path
        else:
            target = os.path.dirname(abs_path) or abs_path
        if not os.path.exists(target):
            messagebox.showerror("Location Not Found", "The saved file or folder could not be located.", parent=self)
            return
        try:
            if sys.platform.startswith("win"):
                if os.path.isdir(abs_path):
                    os.startfile(abs_path)
                elif os.path.isfile(abs_path):
                    norm = os.path.normpath(abs_path)
                    subprocess.Popen(["explorer", "/select,", norm])
                else:
                    os.startfile(target)
            elif sys.platform == "darwin":
                if os.path.isdir(abs_path):
                    subprocess.Popen(["open", target])
                else:
                    subprocess.Popen(["open", "-R", abs_path])
            else:
                subprocess.Popen(["xdg-open", target])
        except Exception as exc:
            messagebox.showerror("Unable to open location", str(exc), parent=self)

    def _on_close(self):
        if hasattr(self.parent, "summary_window") and self.parent.summary_window is self:
            self.parent.summary_window = None
        try:
            self.grab_release()
        except Exception:
            pass
        if self.winfo_exists():
            self.destroy()

class ScanWindow(CTkToplevel):
    def __init__(self, parent, session_mgr, read_only=False):
        super().__init__(parent)
        self.parent = parent
        self.sm = session_mgr
        self.read_only = read_only
        self.restrictions = self.sm.restrictions
        self.df = read_data(self.sm.session_path).fillna("")
        self.mapping = self.sm.mapping or {col: col for col in self.df.columns}

        self.original_bg = Image.open(HOME_BG_FILE)
        self.bg_photo = ImageTk.PhotoImage(self.original_bg)
        self.bg_label = CTkLabel(self, text="", image=self.bg_photo)
        self.bg_label.place(relwidth=1, relheight=1)
        self.bind("<Configure>", self._on_bg_resize)

        self.title("Scan Attendance")
        self.protocol("WM_DELETE_WINDOW", self._on_end_scan)
        self.after(50, lambda: bring_window_to_front(self))

        self._all_iids = []
        self._visible_iids = []
        self._search_entries = []
        self.search_var = None
        self._gate_context = None
        self._manual_additions = 0
        self._cancellations = 0
        self._focus_reset_job = None
        self._focus_guard_depth = 0

        self.stats_vars = {
            "total": ctk.StringVar(value="0"),
            "attended": ctk.StringVar(value="0"),
            "percent": ctk.StringVar(value="0%"),
            "missing_exam": ctk.StringVar(value="0"),
            "missing_hw": ctk.StringVar(value="0"),
        }
        self.gate_message_var = ctk.StringVar(value="")

        self._build_ui()
        self._apply_treeview_style()
        self._load_existing()
        self._refresh_stats()
        ensure_initial_size(self, min_size=MIN_SCAN_SIZE)

        if not self.read_only:
            self.bind_all("<FocusIn>", self._global_focus_in, add="+")
            self.scan_entry.focus_set()

    def _on_bg_resize(self, event):
        if event.widget is self:
            bg = self.original_bg.resize((event.width, event.height), Image.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(bg)
            self.bg_label.configure(image=self.bg_photo)

    def _focus_scan_entry(self):
        self._focus_reset_job = None
        if self.read_only or self._focus_guard_depth > 0:
            return
        try:
            self.scan_entry.focus_set()
        except Exception:
            pass

    def _pause_focus_guard(self):
        if self._focus_reset_job is not None:
            try:
                self.after_cancel(self._focus_reset_job)
            except Exception:
                pass
            self._focus_reset_job = None
        self._focus_guard_depth += 1

    def _resume_focus_guard(self):
        if self._focus_guard_depth > 0:
            self._focus_guard_depth -= 1

    def _build_ui(self):
        top_bar = CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=12, pady=(16, 8))
        top_bar.grid_columnconfigure(1, weight=1)

        scan_block = CTkFrame(top_bar, fg_color="transparent")
        scan_block.grid(row=0, column=0, sticky="w")
        CTkLabel(scan_block, text="Scan Entry", font=("Arial", 12, "bold")).pack(anchor="w")
        self.scan_entry = CTkEntry(scan_block, width=240, placeholder_text="Scan card ID")
        self.scan_entry.pack(anchor="w", pady=(4, 0))
        self.scan_entry.bind("<Return>", lambda _e: self._on_scan())

        self.pb = CTkProgressBar(scan_block, mode="indeterminate", width=240)
        self.pb.pack(anchor="w", pady=(6, 0))
        self.pb.stop()

        if self.read_only:
            self.scan_entry.configure(state="disabled")
            self.scan_entry.unbind("<Return>")
            scan_block.grid_remove()

        session_text = f"Session: {self.sm.name}"
        if self.read_only:
            session_text += " (read-only)"
        session_label = CTkLabel(top_bar, text=session_text, font=("Arial", 12))
        session_label.grid(row=0, column=1, sticky="w", padx=18)

        self.end_button = CTkButton(top_bar, text="End Session", command=self._on_end_scan)
        self.end_button.grid(row=0, column=2, sticky="e")
        if self.read_only:
            self.end_button.configure(text="Close")

        self._build_stats_strip()

        search_frame = CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=12, pady=(0, 10))
        self.search_var = ctk.StringVar()
        search_entry = CTkEntry(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="Search by card, ID, name, phone",
        )
        search_entry.grid(row=0, column=0, sticky="ew", padx=5)
        search_frame.grid_columnconfigure(0, weight=1)
        self.search_var.trace_add("write", self._on_search_change)
        self._search_entries.append(search_entry)
        search_entry.bind("<FocusIn>", lambda _e: self._pause_focus_guard())
        search_entry.bind("<FocusOut>", lambda _e: self._resume_focus_guard())
        self.smart_search_entry = search_entry

        self.gate_panel = CTkFrame(self, fg_color=("#e2e8f0", "#1f2933"), corner_radius=8)
        self.gate_label = CTkLabel(
            self.gate_panel,
            textvariable=self.gate_message_var,
            wraplength=760,
            font=("Arial", 12)
        )
        self.gate_label.pack(pady=(10, 4), padx=12, anchor="w")
        gate_buttons = CTkFrame(self.gate_panel, fg_color="transparent")
        gate_buttons.pack(pady=(0, 10))
        self.gate_cancel_btn = CTkButton(gate_buttons, text="Cancel", command=self._gate_cancel)
        self.gate_cancel_btn.pack(side="left", padx=6)
        self.gate_attend_btn = CTkButton(gate_buttons, text="Attend", command=self._gate_attend)
        self.gate_attend_btn.pack(side="left", padx=6)
        self.gate_panel.pack(fill="x", padx=12, pady=(0, 10))
        self.gate_panel.pack_forget()
        if self.read_only:
            self.gate_cancel_btn.configure(state="disabled")
            self.gate_attend_btn.configure(state="disabled")

        cols = ["card_id", "student_id", "name", "phone"]
        if self.restrictions.get("exam"):
            cols.append("exam")
        if self.restrictions.get("homework"):
            cols.append("homework")
        cols += ["attendance", "notes", "timestamp"]

        tree_container = CTkFrame(self, fg_color="transparent")
        tree_container.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings", selectmode="browse")
        for col in cols:
            self.tree.heading(col, text=col.replace("_", " ").title())
            self.tree.column(col, anchor="center", width=110)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self._on_row_double_click)
        if self.read_only:
            self.tree.unbind("<Double-1>")

        self.control_frame = CTkFrame(self, fg_color="transparent")
        self.control_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.add_student_button = CTkButton(self.control_frame, text="Add Student", command=self._on_add_student_flow)
        self.add_student_button.pack(side="left")
        if self.read_only:
            self.control_frame.pack_forget()

    def _build_stats_strip(self):
        self.stats_frame = CTkFrame(self, fg_color=("#f1f5f9", "#12263a"), corner_radius=10)
        self.stats_frame.pack(fill="x", padx=12, pady=(0, 10))
        stats = [
            ("Total Rows", self.stats_vars["total"]),
            ("Attended", self.stats_vars["attended"]),
            ("Attendance %", self.stats_vars["percent"]),
        ]
        if self.restrictions.get("exam"):
            stats.append(("Missing Exam", self.stats_vars["missing_exam"]))
        if self.restrictions.get("homework"):
            stats.append(("Missing H.W.", self.stats_vars["missing_hw"]))
        for idx, (label, var) in enumerate(stats):
            block = CTkFrame(self.stats_frame, fg_color="transparent")
            block.grid(row=0, column=idx, sticky="w", padx=(12 if idx == 0 else 8, 8), pady=10)
            CTkLabel(block, text=label, font=("Arial", 11, "bold")).pack(anchor="w")
            CTkLabel(block, textvariable=var, font=("Arial", 16)).pack(anchor="w")
        for idx in range(len(stats)):
            self.stats_frame.grid_columnconfigure(idx, weight=1)

    def _apply_treeview_style(self):
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            bg = "#1e1e1e"
            fg = "#f2f2f2"
            heading_bg = "#1f6aa5"
            heading_fg = "#ffffff"
        else:
            bg = "#ffffff"
            fg = "#1a1a1a"
            heading_bg = "#e1efff"
            heading_fg = "#1a1a1a"
        select_bg = "#1f6aa5"
        select_fg = "#ffffff"

        style = ttk.Style(self)
        style.configure(
            "CTk.Treeview",
            background=bg,
            foreground=fg,
            fieldbackground=bg,
            rowheight=32,
            font=("Arial", 11)
        )
        style.map(
            "CTk.Treeview",
            background=[("selected", select_bg)],
            foreground=[("selected", select_fg)]
        )
        style.configure(
            "CTk.Treeview.Heading",
            background=heading_bg,
            foreground=heading_fg,
            font=("Arial", 11, "bold")
        )
        style.map(
            "CTk.Treeview.Heading",
            background=[("active", heading_bg)]
        )
        self.tree.configure(style="CTk.Treeview")

    def _load_existing(self):
        def pad_card_id(val):
            val_str = str(val).strip()
            if val_str.isdigit():
                return val_str.zfill(8)
            return val_str

        cols = self.tree["columns"]
        session_records = {pad_card_id(rec.get("card_id", "")): rec for rec in self.sm.records}
        self._all_iids = []

        for _, row in self.df.iterrows():
            card_id_col = self.mapping.get("card_id", "card_id")
            cid = pad_card_id(row.get(card_id_col, ""))
            rec = session_records.pop(cid, None)
            values = []
            for col in cols:
                mapped_col = self.mapping.get(col, col)
                if rec is not None and col in rec:
                    val = rec.get(col, "")
                else:
                    val = row.get(mapped_col, "")
                values.append(self._clean_value(val))
            self.tree.insert("", "end", iid=cid, values=tuple(values))
            self._all_iids.append(cid)

        for cid, rec in session_records.items():
            values = [self._clean_value(rec.get(col, "")) for col in cols]
            self.tree.insert("", "end", iid=cid, values=tuple(values))
            self._all_iids.append(cid)

        for iid in self._all_iids:
            attendance = self._clean_value(self.tree.set(iid, "attendance"))
            notes = self._clean_value(self.tree.set(iid, "notes"))
            self._update_row(iid, attendance, notes)

    def _clean_value(self, value):
        if value is None:
            return ""
        if isinstance(value, float) and pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() == "nan":
            return ""
        return text

    def _compute_summary_metrics(self):
        total = len(self._all_iids)
        attended = 0
        missing_exam = 0
        missing_hw = 0
        for iid in self._all_iids:
            if not self.tree.exists(iid):
                continue
            attendance = self._clean_value(self.tree.set(iid, "attendance")).lower()
            if attendance == "attend":
                attended += 1
            if self.restrictions.get("exam"):
                if not self._clean_value(self.tree.set(iid, "exam")):
                    missing_exam += 1
            if self.restrictions.get("homework"):
                if not self._clean_value(self.tree.set(iid, "homework")):
                    missing_hw += 1
        percent_text = "0%"
        if total:
            percent_text = f"{(attended / total) * 100:.1f}%"
        metrics = {
            "total": total,
            "attended": attended,
            "attendance_rate": percent_text,
        }
        if self.restrictions.get("exam"):
            metrics["missing_exam"] = missing_exam
        if self.restrictions.get("homework"):
            metrics["missing_hw"] = missing_hw
        return metrics

    def _build_summary_payload(self):
        metrics = self._compute_summary_metrics()
        summary = dict(metrics)
        summary["manual_additions"] = self._manual_additions
        summary["cancellations"] = self._cancellations
        return summary

    def _refresh_stats(self):
        metrics = self._compute_summary_metrics()
        self.stats_vars["total"].set(f"{metrics['total']}")
        self.stats_vars["attended"].set(f"{metrics['attended']}")
        self.stats_vars["percent"].set(metrics["attendance_rate"])
        if self.restrictions.get("exam"):
            self.stats_vars["missing_exam"].set(f"{metrics.get('missing_exam', 0)}")
        if self.restrictions.get("homework"):
            self.stats_vars["missing_hw"].set(f"{metrics.get('missing_hw', 0)}")

    def _finalize_and_close(self, status_message=None):
        if status_message is None:
            status_message = f"Session '{self.sm.name}' saved and closed."
        summary = self._build_summary_payload()
        session_name = self.sm.name
        session_path = getattr(self.sm, "session_path", None)
        parent = self.parent
        read_only = getattr(self, "read_only", False)
        if self.winfo_exists():
            self.destroy()
        if hasattr(parent, "_refresh_recent_sessions"):
            parent._refresh_recent_sessions()
        if getattr(parent, "past_sessions_window", None) and parent.past_sessions_window.winfo_exists():
            parent.past_sessions_window.refresh()
        if hasattr(parent, "set_status"):
            parent.set_status(status_message)
        if hasattr(parent, "show_session_summary"):
            try:
                parent.after(160, lambda: parent.show_session_summary(
                    session_name=session_name,
                    summary=summary,
                    session_path=session_path,
                    read_only=read_only,
                ))
            except Exception:
                pass

    def _on_search_change(self, *_):
        self._filter_all()

    def _filter_all(self):
        query = ''
        if self.search_var is not None:
            query = self._clean_value(self.search_var.get()).lower()
        terms = [term for term in query.split() if term]
        if not terms:
            for iid in self._all_iids:
                if self.tree.exists(iid):
                    self.tree.reattach(iid, '', 'end')
            return
        columns = self.tree['columns']
        for iid in self._all_iids:
            if not self.tree.exists(iid):
                continue
            values = [
                self._clean_value(self.tree.set(iid, col)).lower()
                for col in columns
            ]
            values.append(str(iid).lower())
            haystack = ' '.join(values)
            if all(term in haystack for term in terms):
                self.tree.reattach(iid, '', 'end')
            else:
                self.tree.detach(iid)

    def _on_scan(self):
        if self.read_only:
            return
        code = self._clean_value(self.scan_entry.get())
        self.scan_entry.delete(0, "end")
        if not code:
            return
        if code.isdigit():
            code = code.zfill(8)
        self._hide_gate_panel()
        if not self.tree.exists(code):
            resp = messagebox.askquestion(
                "Not found",
                f"Card ID '{code}' not found.\nDo you want to add this student?",
                icon="warning", type="yesno", default="no", parent=self
            )
            if resp == "yes":
                self._launch_add_student_dialog(card_id=code, default_notes="Diff group")
            return
        self.tree.selection_set(code)
        self.tree.focus(code)
        self._visible_iids = [iid for iid in self._all_iids if self.tree.exists(iid)]
        for iid in self._all_iids:
            if iid != code and self.tree.exists(iid):
                self.tree.detach(iid)
        self.pb.start()
        self.after(500, lambda: self._finish_scan(code))

    def _restore_visible_rows(self):
        if not self._visible_iids:
            return
        for iid in self._visible_iids:
            if self.tree.exists(iid):
                self.tree.reattach(iid, '', 'end')
        self._visible_iids = []

    def _get_missing_fields(self, code):
        missing = []
        if self.restrictions.get("exam"):
            if not self._clean_value(self.tree.set(code, "exam")):
                missing.append("exam")
        if self.restrictions.get("homework"):
            if not self._clean_value(self.tree.set(code, "homework")):
                missing.append("homework")
        return missing

    def _finish_scan(self, code):
        self.pb.stop()
        self._restore_visible_rows()
        if not self.tree.exists(code):
            return
        missing = self._get_missing_fields(code)
        if missing:
            self._show_gate_panel(code, missing)
            return
        current_att = self._clean_value(self.tree.set(code, "attendance")).lower()
        if current_att == "attend":
            self.scan_entry.focus_set()
            return
        notes = self._clean_value(self.tree.set(code, "notes"))
        self._set_attendance(code, "attend", notes, warn_on_duplicate=False)
        self._hide_gate_panel()
        self.scan_entry.focus_set()

    def _format_missing_text(self, missing):
        labels = []
        if "exam" in missing:
            labels.append("exam")
        if "homework" in missing:
            labels.append("homework")
        if len(labels) == 2:
            return "exam and homework"
        return labels[0]

    def _show_gate_panel(self, cid, missing):
        if self.read_only:
            return
        name = self._clean_value(self.tree.set(cid, "name")) or cid
        detail = self._format_missing_text(missing)
        self.gate_message_var.set(f"Gate restriction: {name} is missing {detail}.")
        self._gate_context = {"cid": cid, "missing": missing}
        if not self.gate_panel.winfo_ismapped():
            self.gate_panel.pack(fill="x", padx=12, pady=(0, 10))
        self.tree.selection_set(cid)
        self.tree.focus(cid)
        self.scan_entry.focus_set()

    def _hide_gate_panel(self):
        self._gate_context = None
        if self.gate_panel.winfo_ismapped():
            self.gate_panel.pack_forget()

    def _gate_cancel(self):
        if self.read_only or not self._gate_context:
            return
        cid = self._gate_context["cid"]
        note = f"[{datetime.now():%H:%M:%S}] canceled"
        self._cancellations += 1
        self._set_attendance(cid, "", note, warn_on_duplicate=False, append=True)
        self._hide_gate_panel()
        self.scan_entry.focus_set()

    def _gate_attend(self):
        if self.read_only or not self._gate_context:
            return
        cid = self._gate_context["cid"]
        missing = self._gate_context["missing"]
        base_note = self._clean_value(self.tree.set(cid, "notes"))
        missing_note = self._build_missing_note(missing)
        combined = base_note
        if missing_note:
            combined = f"{base_note} {missing_note}".strip()
        self._set_attendance(cid, "attend", combined, warn_on_duplicate=False)
        self._hide_gate_panel()
        self.scan_entry.focus_set()

    def _build_missing_note(self, missing):
        if not missing:
            return ""
        if len(missing) == 2:
            return "[No exam & H.W.]"
        return "[No exam]" if missing[0] == "exam" else "[No H.W.]"

    def _set_attendance(self, code, attendance, notes, *, warn_on_duplicate=True, append=False):
        if self.read_only:
            return False
        if not self.tree.exists(code):
            return False
        current_att = self._clean_value(self.tree.set(code, "attendance")).lower()
        if warn_on_duplicate and attendance.lower() == "attend" and current_att == "attend":
            messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
            return False
        current_notes = self._clean_value(self.tree.set(code, "notes"))
        if append and notes:
            final_notes = f"{current_notes} {notes}".strip()
        elif notes != "":
            final_notes = notes.strip()
        else:
            final_notes = current_notes
        timestamp = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        rec = self._build_record_payload(code, attendance, final_notes, timestamp)
        self.sm.add_record(rec)
        self._update_record_cache(rec)
        self._update_row(code, attendance, final_notes, timestamp)
        self._refresh_stats()
        return True

    def _build_record_payload(self, code, attendance, notes, timestamp):
        rec = {
            "card_id": code,
            "student_id": self._clean_value(self.tree.set(code, "student_id")),
            "name": self._clean_value(self.tree.set(code, "name")),
            "phone": self._clean_value(self.tree.set(code, "phone")),
            "attendance": attendance,
            "notes": notes,
            "timestamp": timestamp,
        }
        if self.restrictions.get("exam") and "exam" in self.tree["columns"]:
            rec["exam"] = self._clean_value(self.tree.set(code, "exam"))
        if self.restrictions.get("homework") and "homework" in self.tree["columns"]:
            rec["homework"] = self._clean_value(self.tree.set(code, "homework"))
        return rec

    def _update_record_cache(self, rec):
        cid = str(rec.get("card_id", ""))
        for existing in self.sm.records:
            if str(existing.get("card_id", "")) == cid:
                existing.update(rec)
                return
        self.sm.records.append(dict(rec))

    def _update_row(self, code, attendance, notes, timestamp=None):
        if not self.tree.exists(code):
            return
        cols = list(self.tree["columns"])
        values = list(self.tree.item(code, "values"))
        try:
            att_idx = cols.index("attendance")
            notes_idx = cols.index("notes")
            ts_idx = cols.index("timestamp")
        except ValueError:
            return
        values[att_idx] = self._clean_value(attendance)
        values[notes_idx] = self._clean_value(notes)
        if timestamp is None:
            df = read_data(self.sm.session_path)
            card_col = self.mapping.get("card_id", "card_id")
            ts_col = self.mapping.get("timestamp", "timestamp")
            mask = df[card_col].astype(str) == str(code)
            if mask.any() and ts_col in df.columns:
                timestamp = self._clean_value(df.loc[mask, ts_col].values[0])
            else:
                timestamp = ""
        values[ts_idx] = self._clean_value(timestamp)
        self.tree.item(code, values=tuple(values))

    def _on_row_double_click(self, _event):
        if self.read_only:
            return
        sel = self.tree.selection()
        if not sel:
            return
        cid = sel[0]
        missing = self._get_missing_fields(cid)
        if missing:
            self._show_gate_panel(cid, missing)
            return
        name = self._clean_value(self.tree.set(cid, "name")) or "this student"
        if messagebox.askyesno("Confirm Attendance", f"Mark {name} as attended?", parent=self):
            notes = self._clean_value(self.tree.set(cid, "notes"))
            self._set_attendance(cid, "attend", notes, warn_on_duplicate=True)
        elif self._clean_value(self.tree.set(cid, "attendance")).lower() == "attend":
            self._cancel_attendance(cid)
        self.scan_entry.focus_set()

    def _cancel_attendance(self, code):
        if self.read_only:
            return
        note = f"[{datetime.now():%H:%M:%S}] canceled"
        self._cancellations += 1
        self._set_attendance(code, "", note, warn_on_duplicate=False, append=True)

    def _on_add_student_flow(self):
        self._launch_add_student_dialog()

    def _launch_add_student_dialog(self, card_id=None, default_notes="manual addition"):
        if self.read_only:
            return
        self._pause_focus_guard()
        normalized_card = None
        if card_id:
            raw_card = str(card_id).strip()
            normalized_card = raw_card.zfill(8) if raw_card.isdigit() else raw_card
        try:
            dialog = AddStudentDialog(
                self,
                card_id=normalized_card,
                duplicate_checker=self._student_id_or_phone_exists,
                default_notes=default_notes,
                on_submit=self._handle_add_student_submission,
            )
        except Exception:
            self._resume_focus_guard()
            raise

        def _restore_focus_guard(event):
            if event.widget is dialog:
                self._resume_focus_guard()

        dialog.bind("<Destroy>", _restore_focus_guard, add="+")

    def _handle_add_student_submission(self, *, card_id, values, default_notes):
        cid = card_id
        if cid:
            cid = str(cid).strip()
            if cid.isdigit():
                cid = cid.zfill(8)
        else:
            cid = self._next_unknown_card_id()
        timestamp = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        rec = {
            "card_id": cid,
            "student_id": values["student_id"],
            "name": values["name"],
            "phone": values["phone"],
            "attendance": "attend",
            "notes": default_notes,
            "timestamp": timestamp,
        }
        if self.restrictions.get("exam"):
            rec["exam"] = ""
        if self.restrictions.get("homework"):
            rec["homework"] = ""
        self.sm.add_record(rec)
        self._manual_additions += 1
        self._update_record_cache(rec)
        row_values = [
            rec["card_id"],
            rec["student_id"],
            rec["name"],
            rec["phone"],
        ]
        if self.restrictions.get("exam"):
            row_values.append(rec.get("exam", ""))
        if self.restrictions.get("homework"):
            row_values.append(rec.get("homework", ""))
        row_values.extend([rec["attendance"], rec["notes"], rec["timestamp"]])
        if self.tree.exists(cid):
            self.tree.item(cid, values=tuple(row_values))
        else:
            self.tree.insert("", "end", iid=cid, values=tuple(row_values))
            if cid not in self._all_iids:
                self._all_iids.append(cid)
        self._update_row(cid, rec["attendance"], rec["notes"], rec["timestamp"])
        self._refresh_stats()
        self.after(120, self.scan_entry.focus_set)
        if hasattr(self.parent, "set_status"):
            display_name = rec["name"].strip() or rec["student_id"]
            self.parent.set_status(f"Student '{display_name}' added manually.")
        return True

    def _next_unknown_card_id(self):
        if not hasattr(self, "_unknown_counter"):
            existing = []
            for record in self.sm.records:
                card_value = str(record.get("card_id", ""))
                if card_value.startswith("Unknown "):
                    suffix = card_value.split("Unknown ", 1)[1]
                    if suffix.isdigit():
                        existing.append(int(suffix))
            self._unknown_counter = max(existing, default=0)
        self._unknown_counter += 1
        return f"Unknown {self._unknown_counter}"

    def _on_end_scan(self):
        if getattr(self, "read_only", False):
            self._finalize_and_close(status_message=f"Session '{self.sm.name}' closed (view-only).")
            return
        self._finalize_and_close()

    def _global_focus_in(self, _event):
        if self._focus_reset_job is not None:
            try:
                self.after_cancel(self._focus_reset_job)
            except Exception:
                pass
            self._focus_reset_job = None
        if self.read_only or self._focus_guard_depth > 0:
            return
        widget = self.focus_get()
        if widget is None:
            return
        # Ignore focus changes coming from other toplevel dialogs (e.g. summary window).
        try:
            owning_top = widget.winfo_toplevel()
        except Exception:
            owning_top = None
        if owning_top is not None and owning_top is not self:
            return
        if widget == self.scan_entry or widget in self._search_entries:
            return
        parent = getattr(widget, "master", None)
        while parent is not None:
            if parent == self.gate_panel:
                return
            parent = getattr(parent, "master", None)
        for w in self.winfo_children():
            if isinstance(w, CTkToplevel):
                if widget == w or widget.winfo_toplevel() == w:
                    return
        self._focus_reset_job = self.after_idle(self._focus_scan_entry)

    def _student_id_or_phone_exists(self, student_id, phone):
        df = read_data(self.sm.session_path)
        sid_col = self.mapping.get("student_id", "student_id")
        phone_col = self.mapping.get("phone", "phone")
        id_exists = phone_exists = False
        if sid_col in df.columns:
            if student_id in df[sid_col].astype(str).values:
                id_exists = True
        if phone_col in df.columns:
            if phone in df[phone_col].astype(str).values:
                phone_exists = True
        return id_exists, phone_exists
if __name__ == "__main__":
    app = App()
    app.mainloop()











































