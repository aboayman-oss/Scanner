"""Primary application window for the RFID Attendance Manager."""
import json
import os
import subprocess
import sys
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk
from customtkinter import CTk, CTkButton, CTkFrame, CTkLabel
from PIL import Image, ImageTk

from core.session_manager import SessionManager
from ui.dialogs.add_student_dialog import AddStudentDialog
from ui.dialogs.session_setup_dialog import SessionSetupDialog
from ui.dialogs.session_summary_dialog import SessionSummaryDialog
from ui.past_sessions_window import PastSessionsWindow
from ui.scan_window import ScanWindow
from ui.settings_window import SettingsWindow
from utils.helpers import (
    LAST_DATA_FILE,
    LOGO_FILE,
    MAPPING_FILE,
    MIN_DASHBOARD_SIZE,
    SETTINGS,
    SETTINGS_FILE,
    SESSIONS_FOLDER,
    bring_window_to_front,
    ensure_initial_size,
    read_data,
    write_data,
)

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
