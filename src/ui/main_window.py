'''Primary application window for the RFID Attendance Manager.'''
import json
import os
import subprocess
import sys
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk
from customtkinter import CTk, CTkButton, CTkFrame, CTkLabel, CTkImage
from PIL import Image

from core.session_manager import SessionManager
from ui.dialogs.session_setup_dialog import SessionSetupDialog
from ui.dialogs.session_summary_dialog import SessionSummaryDialog

from ui.scan_window import ScanWindow
from ui.settings_window import SettingsWindow
from tkinter import messagebox # Ensure this is imported for _clear_all_sessions
from ui.components.past_session_list_item import PastSessionListItem
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
    PAST_SESSIONS_ICON_FILE,
    SETTINGS_ICON_FILE,
    IMPORT_ICON_FILE,
    NEW_SESSION_ICON_FILE,
    DASHBOARD_ICON_FILE
)

class App(CTk):
    def __init__(self):
        super().__init__()
        self.title("RFID Attendance Manager")
        self.column_map = {}
        self.data_df    = None
        self.settings_window = None
        self.current_data_path = None
        self._session_setup = None
        self.past_sessions_window = None
        self.summary_window = None

        self._load_icons()

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

    def _create_icon(self, icon_path, size=(24, 24)):
        try:
            img = Image.open(icon_path)
            return CTkImage(light_image=img, dark_image=img, size=size)
        except Exception as e:
            print(f"Warning: Could not load icon: {icon_path} - {e}")
            return None

    def _load_icons(self):
        self.past_sessions_icon = self._create_icon(PAST_SESSIONS_ICON_FILE)
        self.settings_icon = self._create_icon(SETTINGS_ICON_FILE)
        self.import_icon = self._create_icon(IMPORT_ICON_FILE)
        self.new_session_icon = self._create_icon(NEW_SESSION_ICON_FILE)
        self.dashboard_icon = self._create_icon(DASHBOARD_ICON_FILE)

    def _build_ui(self):
        # 1. Configure root window grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 2. Build the navigation rail on the left
        self._build_nav_rail()

        # 3. Build the main content frame on the right
        self.content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self._build_dashboard_view() # Initially show the dashboard

        # Status bar setup
        self.status_var = ctk.StringVar(value="Ready.")
        self.status_label = CTkLabel(
            self,
            textvariable=self.status_var,
            anchor="w",
            font=("Arial", 12)
        )
        self.status_label.grid(row=1, column=1, sticky="ew", padx=20, pady=(0, 10))

    def _build_nav_rail(self):
        nav_rail = ctk.CTkFrame(self, width=100, corner_radius=0)
        nav_rail.grid(row=0, column=0, rowspan=2, sticky="nsw")
        
        dashboard_button = ctk.CTkButton(nav_rail, text="Dashboard", image=self.dashboard_icon, compound="left", command=self._show_dashboard_view)
        dashboard_button.pack(pady=10, padx=10)
        
        past_sessions_button = ctk.CTkButton(nav_rail, text="Past Sessions", image=self.past_sessions_icon, compound="left", command=self._show_past_sessions_view)
        past_sessions_button.pack(pady=10, padx=10)
        
        settings_button = ctk.CTkButton(nav_rail, text="Settings", image=self.settings_icon, compound="left", command=self.open_settings)
        settings_button.pack(side="bottom", pady=20, padx=10)

    def _build_dashboard_view(self):
        # Clear any previous content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(2, weight=1)

        # Build the header
        self._build_header()

        # Build the two main cards
        self._build_start_session_card()
        self._build_recent_sessions_card()
        
        # Update UI based on initial data state
        self._update_ui_for_data_state()

    def _build_header(self):
        header = CTkFrame(self.content_frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(1, weight=1)

        try:
            logo = Image.open(LOGO_FILE)
            if logo.width > 0 and logo.height > 0:
                target_width = 56
                ratio = target_width / logo.width
                target_height = max(1, int(logo.height * ratio))
                logo = logo.resize((target_width, target_height), Image.Resampling.LANCZOS)
                self.logo_img = CTkImage(light_image=logo, dark_image=logo, size=(target_width, target_height))
                CTkLabel(header, image=self.logo_img, text="").grid(row=0, column=0, sticky="w", padx=(0, 12))
        except Exception as e:
            print(f"Warning: Could not load logo image: {e}")

        title_holder = CTkFrame(header, fg_color="transparent")
        title_holder.grid(row=0, column=1, sticky="w")
        CTkLabel(
            title_holder,
            text="RFID Attendance Manager",
            font=("Arial", 24, "bold")
        ).pack(anchor="w")
        CTkLabel(
            title_holder,
            text="Start scans, review sessions, and adjust preferences from one place.",
            font=("Arial", 14)
        ).pack(anchor="w", pady=(4, 0))

    def _build_start_session_card(self):
        self.start_card = ctk.CTkFrame(self.content_frame, corner_radius=12)
        self.start_card.grid(row=1, column=0, sticky="ew", pady=(20, 20))
        self.start_card.grid_columnconfigure(1, weight=1)
        
        self.start_card_title = ctk.CTkLabel(self.start_card, text="Start a New Session", font=("Arial", 18, "bold"), anchor="w")
        self.start_card_title.grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(20, 4))
        
        self.start_card_subtitle = ctk.CTkLabel(self.start_card, text="Import a student roster to begin.", anchor="w")
        self.start_card_subtitle.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))
        
        self.start_session_btn = ctk.CTkButton(
            self.start_card,
            text="Start New Session",
            image=self.new_session_icon,
            compound="right",
            command=self.open_scan_window_setup,
            font=("Arial", 14, "bold"),
            height=40
        )
        self.start_session_btn.grid(row=2, column=1, sticky="e", padx=20, pady=(0, 20))
        
        self.import_btn = ctk.CTkButton(
            self.start_card,
            text="Import Roster",
            image=self.import_icon,
            compound="left",
            command=self._handle_import,
            fg_color="transparent",
            border_width=1
        )
        self.import_btn.grid(row=2, column=0, sticky="w", padx=20, pady=(0, 20))

    def _build_recent_sessions_card(self):
        recent_card = ctk.CTkFrame(self.content_frame, corner_radius=12)
        recent_card.grid(row=2, column=0, sticky="nsew")
        recent_card.grid_columnconfigure(0, weight=1)
        recent_card.grid_rowconfigure(1, weight=1)
        
        title = ctk.CTkLabel(recent_card, text="Recent Sessions", font=("Arial", 16, "bold"), anchor="w")
        title.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        self.recent_sessions_frame = ctk.CTkScrollableFrame(recent_card, fg_color="transparent")
        self.recent_sessions_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)
        self.recent_sessions_frame.grid_columnconfigure(0, weight=1)

        self._refresh_recent_sessions()

    def _refresh_recent_sessions(self):
        if not hasattr(self, "recent_sessions_frame"):
            return

        for widget in self.recent_sessions_frame.winfo_children():
            widget.destroy()

        if not os.path.isdir(SESSIONS_FOLDER):
            return
            
        files = []
        try:
            for entry in os.listdir(SESSIONS_FOLDER):
                path_entry = os.path.join(SESSIONS_FOLDER, entry)
                if os.path.isfile(path_entry) and entry.lower().endswith((".csv", ".xlsx")):
                    files.append((path_entry, os.path.getmtime(path_entry)))
        except FileNotFoundError:
            return

        files.sort(key=lambda item: item[1], reverse=True)
        
        if not files:
            no_sessions_label = ctk.CTkLabel(self.recent_sessions_frame, text="No recent sessions found.", anchor="w", justify="left")
            no_sessions_label.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
            return

        for i, (path_entry, modified) in enumerate(files[:5]):
            name = os.path.splitext(os.path.basename(path_entry))[0]
            stamp = datetime.fromtimestamp(modified).strftime("%d %b %Y %H:%M")
            
            item_frame = ctk.CTkFrame(self.recent_sessions_frame, fg_color="transparent")
            item_frame.grid(row=i, column=0, sticky="ew", pady=5)
            item_frame.grid_columnconfigure(0, weight=1)
            
            info_label = ctk.CTkLabel(item_frame, text=f"{name}\n{stamp}", anchor="w", justify="left")
            info_label.grid(row=0, column=0, sticky="w", padx=10)
            
            btn_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
            btn_frame.grid(row=0, column=1, sticky="e")
            
            reveal_btn = ctk.CTkButton(
                btn_frame, 
                text="Show", 
                width=60,
                command=lambda p=path_entry: self._reveal_session_path(p)
            )
            reveal_btn.pack(side="right", padx=(5, 10))

            open_btn = ctk.CTkButton(
                btn_frame, 
                text="Open", 
                width=60,
                command=lambda p=path_entry: self._open_session_path(p, read_only=True)
            )
            open_btn.pack(side="right")

    def _update_ui_for_data_state(self):
        if self.data_df is not None and self.current_data_path:
            rows = len(self.data_df)
            filename = os.path.basename(self.current_data_path)
            self.start_card_subtitle.configure(text=f"Roster Loaded: {rows:,} students from '{filename}'")
            self.start_session_btn.configure(state="normal")
            self.import_btn.configure(text="Import")
        else:
            self.start_card_subtitle.configure(text="Import a student roster to begin.")
            self.start_session_btn.configure(state="disabled")
            self.import_btn.configure(text="Import Roster")

    def _handle_import(self):
        if self.import_csv():
            self._update_ui_for_data_state()

    def _show_dashboard_view(self):
        self._build_dashboard_view()

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
                subprocess.Popen(['explorer', '/select,', target])
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

    

    def _format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"

    def _clear_all_sessions(self):
        # NOTE: This method is copied almost verbatim.
        # It scans the SESSIONS_FOLDER and deletes files.
        if not os.path.isdir(SESSIONS_FOLDER): return

        paths_to_delete = [
            os.path.join(SESSIONS_FOLDER, entry)
            for entry in os.listdir(SESSIONS_FOLDER)
            if os.path.isfile(os.path.join(SESSIONS_FOLDER, entry))
        ]

        if not paths_to_delete: return

        confirm = messagebox.askyesno(
            "Clear All Sessions",
            f"This will permanently delete {len(paths_to_delete)} session file(s). Are you sure?",
            parent=self # Use 'self' as the parent window
        )
        if not confirm: return

        failures = []
        for path_entry in paths_to_delete:
            try:
                os.remove(path_entry)
            except Exception as exc:
                failures.append(f"{os.path.basename(path_entry)}: {exc}")

        self._populate_past_sessions_list() # Refresh the current view
        self._refresh_recent_sessions()     # Also refresh the dashboard view

        if failures:
            messagebox.showerror(
                "Delete Failed",
                "Some session files could not be deleted:\n" + "\n".join(failures),
                parent=self
            )
            self.set_status("Some past sessions could not be removed.")
        else:
            messagebox.showinfo(
                "Sessions Cleared",
                "All past session files have been deleted.",
                parent=self
            )
            self.set_status("All past sessions cleared.")

    def _show_past_sessions_view(self):
        """Clears the content frame and builds the past sessions view."""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        self._build_past_sessions_view()

    def _build_past_sessions_view(self):
        """Builds the UI for browsing past sessions within the content_frame."""
        self.content_frame.grid_rowconfigure(1, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Header with title and global actions
        header_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header_frame, text="Past Sessions", font=("Arial", 24, "bold")).grid(row=0, column=0, sticky="w")

        action_buttons_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        action_buttons_frame.grid(row=0, column=1, sticky="e")

        ctk.CTkButton(action_buttons_frame, text="Refresh", command=self._populate_past_sessions_list).pack(side="left", padx=(0, 10))

        self.clear_all_btn = ctk.CTkButton(action_buttons_frame, text="Clear All", command=self._clear_all_sessions)
        self.clear_all_btn.pack(side="left")

        # Scrollable frame for the list
        self.sessions_list_frame = ctk.CTkScrollableFrame(self.content_frame, label_text="Session Files")
        self.sessions_list_frame.grid(row=1, column=0, sticky="nsew")
        self.sessions_list_frame.grid_columnconfigure(0, weight=1)

        # Populate the list with session data
        self._populate_past_sessions_list()

    def _populate_past_sessions_list(self):
        """Fetches session data and populates the scrollable list with custom widgets."""
        for widget in self.sessions_list_frame.winfo_children():
            widget.destroy()

        if not os.path.isdir(SESSIONS_FOLDER):
            files = []
        else:
            files = []
            for entry in os.listdir(SESSIONS_FOLDER):
                path_entry = os.path.join(SESSIONS_FOLDER, entry)
                if os.path.isfile(path_entry) and entry.lower().endswith((".csv", ".xlsx")):
                    stats = os.stat(path_entry)
                    files.append((path_entry, stats.st_mtime, stats.st_size))
            files.sort(key=lambda item: item[1], reverse=True)

        self.clear_all_btn.configure(state="normal" if files else "disabled")

        if not files:
            ctk.CTkLabel(self.sessions_list_frame, text="No session files found.").pack(pady=20)
            return

        for path, modified, size in files:
            PastSessionListItem(
                master=self.sessions_list_frame,
                path=path,
                modified_timestamp=modified,
                size_bytes=size,
                format_size_func=self._format_size,
                open_func=self._open_session_path,
                reveal_func=self._reveal_session_path
            ).pack(fill="x", expand=True, padx=10, pady=5)

    def _load_last_data(self):
        self.data_df = None
        self.current_data_path = None
        if os.path.exists(LAST_DATA_FILE):
            try:
                with open(LAST_DATA_FILE) as f:
                    last_data = json.load(f)
                    path = last_data.get("path")
                    if path and os.path.exists(path):
                        self.data_df = read_data(path)
                        self.current_data_path = path
            except (json.JSONDecodeError, KeyError, FileNotFoundError) as e:
                print(f"Could not load last data file: {e}")
                self.data_df = None
                self.current_data_path = None
        
        if hasattr(self, "start_card_subtitle"):
            self._update_ui_for_data_state()
        self.set_status("Ready.")

    def open_settings(self):
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
        self.current_data_path = path
        with open(LAST_DATA_FILE, "w") as f:
            json.dump({"path": path}, f, indent=2)
        
        self.set_status(f"Imported {len(df)} records from {os.path.basename(path)}.")
        return True

    def open_scan_window_setup(self):
        if self._session_setup is not None and self._session_setup.winfo_exists():
            bring_window_to_front(self._session_setup)
            return

        if self.data_df is None:
             messagebox.showwarning("No Data", "Please import data before starting a session.")
             return

        try:
            self._session_setup = SessionSetupDialog(
                self,
                SETTINGS.get("stage_options", []),
                SETTINGS.get("center_options", []),
                has_data=self.data_df is not None,
                callback=self._on_session_setup_finished,
            )
        except Exception as e:
            messagebox.showerror("Dialog Error", f"Failed to open session setup dialog: {e}")
            self._session_setup = None

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
        if not os.path.exists(session_path) or messagebox.askyesno("Overwrite Session?", f"Session '{name}' already exists. Do you want to overwrite it with the currently loaded roster?"):
            write_data(self.data_df, session_path)
            created = True
        
        session_df = read_data(session_path)
        sm = SessionManager(name, params, self.column_map, session_df)
        
        self._refresh_recent_sessions()
        if self.past_sessions_window is not None and self.past_sessions_window.winfo_exists():
            self.past_sessions_window.refresh()
            
        ScanWindow(self, sm)
        
        if created:
            self.set_status(f"Session '{name}' created/overwritten.")
        else:
            self.set_status(f"Session '{name}' loaded.")