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
from ui.dialogs.password_dialog import PasswordDialog

from ui.scan_window import ScanWindow
from ui.settings_window import SettingsWindow
from tkinter import messagebox # Ensure this is imported for _clear_all_sessions
from ui.components.past_session_list_item import PastSessionListItem
from utils.helpers import (
    DEFAULT_SESSIONS_FOLDER,
    FOLDER_OPEN_ICON_FILE,
    LAST_DATA_FILE,
    LOGO_FILE,
    MAPPING_FILE,
    MIN_DASHBOARD_SIZE,
    SETTINGS,
    SETTINGS_FILE,
    bring_window_to_front,
    ensure_initial_size,
    get_sessions_folder,
    read_data,
    resolve_session_file_path,
    save_settings,
    set_dark_title_bar,
    set_sessions_folder,
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

        # --- FORCE DARK THEME ---
        ctk.set_appearance_mode("dark")
        # ------------------------

        self.title("RFID Attendance Manager")
        set_dark_title_bar(self)
        self.column_map = {}
        self.data_df    = None
        self.settings_window = None
        self.current_data_path = None
        self._session_setup = None
        self.past_sessions_window = None
        self.summary_window = None
        self.dashboard_frame = None
        self.past_sessions_frame = None
        self._dashboard_initialized = False
        self._past_sessions_initialized = False
        self._past_sessions_list_dirty = True
        self._session_files_cache = None
        self.logo_img = None
        self.session_folder = None
        self._should_prompt_session_folder = False

        self._load_icons()

        if os.path.exists(MAPPING_FILE):
            with open(MAPPING_FILE, encoding="utf-8") as f:
                self.column_map = json.load(f)

        loaded_settings = {}
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                loaded_settings = json.load(f)
                SETTINGS.update(loaded_settings)
        else:
            self._should_prompt_session_folder = True

        self._initialize_session_folder(loaded_settings)

        self._build_ui()
        width, height = ensure_initial_size(self, min_size=MIN_DASHBOARD_SIZE)
        self.minsize(width, height)
        self._load_last_data()
        self.after(150, self._maybe_prompt_for_sessions_folder)

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
        self.folder_icon = self._create_icon(FOLDER_OPEN_ICON_FILE, size=(22, 22))
        if os.path.exists(LOGO_FILE):
            try:
                logo = Image.open(LOGO_FILE)
                if logo.width > 0 and logo.height > 0:
                    target_width = 100
                    ratio = target_width / logo.width
                    target_height = max(1, int(logo.height * ratio))
                    logo = logo.resize((target_width, target_height), Image.Resampling.LANCZOS)
                    self.logo_img = CTkImage(light_image=logo, dark_image=logo, size=(target_width, target_height))
            except Exception as e:
                print(f"Warning: Could not load logo image: {e}")

    def _initialize_session_folder(self, loaded_settings):
        stored_path = loaded_settings.get("sessions_folder", SETTINGS.get("sessions_folder"))
        if not stored_path:
            stored_path = DEFAULT_SESSIONS_FOLDER
        resolved = self._apply_sessions_folder(
            stored_path,
            persist=False,
            refresh=False,
            show_status=False,
            notify_errors=False,
        )
        if resolved is None:
            resolved = self._apply_sessions_folder(
                DEFAULT_SESSIONS_FOLDER,
                persist=False,
                refresh=False,
                show_status=False,
                notify_errors=False,
            )
        if "sessions_folder" not in loaded_settings:
            self._should_prompt_session_folder = True
        if resolved is None:
            self.session_folder = get_sessions_folder()

    def _apply_sessions_folder(
        self,
        path,
        *,
        persist=True,
        refresh=True,
        show_status=False,
        notify_errors=True,
    ):
        try:
            resolved = set_sessions_folder(path)
        except OSError as exc:
            if notify_errors:
                messagebox.showerror("Session Folder", f"Could not use the selected folder:\n{exc}", parent=self)
            return None
        self.session_folder = resolved
        if persist:
            try:
                save_settings()
            except OSError as exc:
                if notify_errors:
                    messagebox.showwarning("Preferences", f"Could not save session folder preference:\n{exc}", parent=self)
        if refresh:
            self._invalidate_session_files_cache()
            if getattr(self, "recent_sessions_frame", None) is not None:
                self._refresh_recent_sessions()
            if getattr(self, "sessions_list_frame", None) is not None:
                self._populate_past_sessions_list(force_scan=True)
            if self.past_sessions_window is not None and self.past_sessions_window.winfo_exists():
                self.past_sessions_window.refresh()
        if show_status:
            self.set_status(f"Session folder set to '{resolved}'.")
        return resolved

    def _maybe_prompt_for_sessions_folder(self):
        if not self._should_prompt_session_folder:
            return
        initial_dir = self.session_folder or DEFAULT_SESSIONS_FOLDER
        selected = filedialog.askdirectory(
            parent=self,
            title="Select Sessions Folder",
            initialdir=initial_dir,
        )
        if selected:
            resolved = self._apply_sessions_folder(selected, persist=True, show_status=True)
            if resolved:
                messagebox.showinfo("Session Folder Set", f"Sessions will be saved in:\n{resolved}", parent=self)
        else:
            messagebox.showinfo(
                "Session Folder Required",
                f"A sessions folder is required. The default location will be used:\n{DEFAULT_SESSIONS_FOLDER}",
                parent=self,
            )
            self._apply_sessions_folder(DEFAULT_SESSIONS_FOLDER, persist=True, show_status=True)
        self._should_prompt_session_folder = False

    def _on_change_sessions_folder_clicked(self):
        initial_dir = self.session_folder or DEFAULT_SESSIONS_FOLDER
        selected = filedialog.askdirectory(
            parent=self,
            title="Select Sessions Folder",
            initialdir=initial_dir,
        )
        if not selected:
            self.set_status("Session folder selection canceled.")
            return
        current = os.path.abspath(self.session_folder) if self.session_folder else None
        if current and os.path.abspath(selected) == current:
            self.set_status("Session folder unchanged.")
            return
        resolved = self._apply_sessions_folder(selected, persist=True, show_status=True)
        if resolved:
            messagebox.showinfo("Session Folder Updated", f"Sessions will now be saved in:\n{resolved}", parent=self)

    def _build_ui(self):
        # 1. Configure root window grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 2. Build the navigation rail on the left
        self._build_nav_rail()

        # 3. Build the main content frame on the right
        self.content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.dashboard_frame = ctk.CTkFrame(self.content_frame, corner_radius=0, fg_color="transparent")
        self.dashboard_frame.grid(row=0, column=0, sticky="nsew")

        self.past_sessions_frame = ctk.CTkFrame(self.content_frame, corner_radius=0, fg_color="transparent")
        self.past_sessions_frame.grid(row=0, column=0, sticky="nsew")

        self._build_dashboard_view()
        self._build_past_sessions_view()
        self._show_dashboard_view()

        # Status bar setup
        self.status_var = ctk.StringVar(value="Ready.")
        self.status_label = CTkLabel(
            self,
            textvariable=self.status_var,
            anchor="w",
            font=("Roboto", 12)
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
        if self._dashboard_initialized or self.dashboard_frame is None:
            return

        frame = self.dashboard_frame
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        self._build_header(frame)
        self._build_start_session_card(frame)
        self._build_recent_sessions_card(frame)

        self._update_ui_for_data_state()
        self._dashboard_initialized = True

    def _build_header(self, parent):
        header = CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(1, weight=1)

        if self.logo_img is not None:
            CTkLabel(header, image=self.logo_img, text="").grid(row=0, column=0, sticky="w", padx=(0, 12))

        title_holder = CTkFrame(header, fg_color="transparent")
        title_holder.grid(row=0, column=1, sticky="w")
        CTkLabel(
            title_holder,
            text="Attendance Manager",
            font=("Roboto", 24, "bold")
        ).pack(anchor="w")
        CTkLabel(
            title_holder,
            text="Powered by Gawish",
            font=("Roboto", 14)
        ).pack(anchor="w", pady=(4, 0))

    def _build_start_session_card(self, parent):
        self.start_card = ctk.CTkFrame(parent, corner_radius=12)
        self.start_card.grid(row=1, column=0, sticky="ew", pady=(20, 20))
        self.start_card.grid_columnconfigure(0, weight=1)
        self.start_card.grid_columnconfigure(1, weight=0)

        self.start_card_title = ctk.CTkLabel(
            self.start_card,
            text="Start a New Session",
            font=("Roboto", 18, "bold"),
            anchor="w"
        )
        self.start_card_title.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 4))

        self.change_folder_btn = ctk.CTkButton(
            self.start_card,
            text="",
            image=self.folder_icon,
            width=32,
            height=32,
            fg_color="transparent",
            hover_color=("#2a2a2a", "#2a2a2a"),
            command=self._on_change_sessions_folder_clicked
        )
        self.change_folder_btn.grid(row=0, column=1, sticky="ne", padx=(0, 16), pady=(16, 4))

        self.start_card_subtitle = ctk.CTkLabel(self.start_card, text="Import a student list to begin.", anchor="w")
        self.start_card_subtitle.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))

        self.start_session_btn = ctk.CTkButton(
            self.start_card,
            text="Start New Session",
            image=self.new_session_icon,
            compound="right",
            command=self.open_scan_window_setup,
            font=("Roboto", 14, "bold"),
            height=40
        )
        self.start_session_btn.grid(row=2, column=1, sticky="e", padx=20, pady=(0, 20))

        self.import_btn = ctk.CTkButton(
            self.start_card,
            text="Import list",
            image=self.import_icon,
            compound="left",
            command=self._handle_import,
            fg_color="transparent",
            border_width=1
        )
        self.import_btn.grid(row=2, column=0, sticky="w", padx=20, pady=(0, 20))

    def _build_recent_sessions_card(self, parent):
        recent_card = ctk.CTkFrame(parent, corner_radius=12)
        recent_card.grid(row=2, column=0, sticky="nsew")
        recent_card.grid_columnconfigure(0, weight=1)
        recent_card.grid_rowconfigure(1, weight=1)

        title = ctk.CTkLabel(recent_card, text="Recent Sessions", font=("Roboto", 16, "bold"), anchor="w")
        title.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))

        self.recent_sessions_frame = ctk.CTkScrollableFrame(recent_card, fg_color="transparent")
        self.recent_sessions_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)
        self.recent_sessions_frame.grid_columnconfigure(0, weight=1)

        self._refresh_recent_sessions()

    def _invalidate_session_files_cache(self):
        self._session_files_cache = None
        self._past_sessions_list_dirty = True

    def _scan_session_files(self):
        files = []
        sessions_dir = get_sessions_folder()
        if not os.path.isdir(sessions_dir):
            return files
        try:
            for root, _dirs, filenames in os.walk(sessions_dir):
                for filename in filenames:
                    if not filename.lower().endswith((".csv", ".xlsx")):
                        continue
                    path_entry = os.path.join(root, filename)
                    try:
                        stats = os.stat(path_entry)
                    except OSError:
                        continue
                    files.append((path_entry, stats.st_mtime, stats.st_size))
        except OSError:
            return files
        files.sort(key=lambda item: item[1], reverse=True)
        return files

    def _get_session_files(self):
        if self._session_files_cache is None:
            self._session_files_cache = self._scan_session_files()
        return self._session_files_cache

    def _refresh_recent_sessions(self):
        if not hasattr(self, "recent_sessions_frame"):
            return

        for widget in self.recent_sessions_frame.winfo_children():
            widget.destroy()

        files = self._get_session_files()

        if not files:
            no_sessions_label = ctk.CTkLabel(self.recent_sessions_frame, text="No recent sessions found.", anchor="w", justify="left")
            no_sessions_label.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
            return

        for i, (path_entry, modified, _) in enumerate(files[:5]):
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
                command=lambda p=path_entry: self._show_session_summary_from_path(p)
            )
            reveal_btn.pack(side="right", padx=(5, 10))

            open_btn = ctk.CTkButton(
                btn_frame,
                text="Open",
                width=60,
                command=lambda p=path_entry: self._open_session_path(p, read_only=False)
            )
            open_btn.pack(side="right")

    def _update_ui_for_data_state(self):
        if self.data_df is not None and self.current_data_path:
            rows = len(self.data_df)
            filename = os.path.basename(self.current_data_path)
            self.start_card_subtitle.configure(text=f"List Loaded: {rows:,} students from '{filename}'")
            self.start_session_btn.configure(state="normal")
            self.import_btn.configure(text="Import New")
        else:
            self.start_card_subtitle.configure(text="Import a student list to begin.")
            self.start_session_btn.configure(state="disabled")
            self.import_btn.configure(text="Import List")

    def _handle_import(self):
        if self.import_csv():
            self._update_ui_for_data_state()

    def _show_dashboard_view(self):
        self._build_dashboard_view()
        if self.dashboard_frame is not None:
            self.dashboard_frame.tkraise()

    def show_session_summary(self, *, session_name, summary, session_path, params=None, read_only=False):
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
            params=params,
            read_only=read_only,
        )

    def set_status(self, message):
        if hasattr(self, "status_var"):
            self.status_var.set(message)

    def _open_session_path(self, path_entry, *, read_only=False):
        try:
            name = os.path.splitext(os.path.basename(path_entry))[0]
            df = read_data(path_entry)
            sm = SessionManager(name, {}, self.column_map, df, session_path=path_entry)
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

    def _show_session_summary_from_path(self, path_entry):
        """Reads a session file, computes a summary, and shows the summary dialog."""
        try:
            session_name = os.path.splitext(os.path.basename(path_entry))[0]
            df = read_data(path_entry).fillna("")

            # Compute summary metrics from the session DataFrame
            att_col = self.column_map.get("attendance", "attendance")
            exam_col = self.column_map.get("exam", "exam")
            card_id_col = self.column_map.get("card_id", "card_id")
            hw_col = self.column_map.get("homework", "homework")

            total = len(df)
            attended = df[att_col].astype(str).str.lower().eq('attend').sum() if att_col in df.columns else 0
            attendance_rate = f"{(attended / total) * 100:.1f}%" if total > 0 else "0%"

            missing_exam = 0
            if exam_col in df.columns and SETTINGS["restrictions"].get("exam"):
                missing_exam = df[exam_col].astype(str).str.strip().replace("", "0").eq("0").sum()

            missing_hw = 0
            if hw_col in df.columns and SETTINGS["restrictions"].get("homework"):
                missing_hw = df[hw_col].astype(str).str.strip().replace("", "0").isin(["", "0"]).sum()

            manual_additions = 0
            if card_id_col in df.columns:
                # Manually added students are identified by card IDs starting with "Unknown"
                manual_additions = df[card_id_col].astype(str).str.strip().str.startswith("Unknown ").sum()

            cancellations = 0
            notes_col = self.column_map.get("notes", "notes")
            if notes_col in df.columns:
                cancellations = df[notes_col].astype(str).str.contains("Canceled", case=False, na=False).sum()

            summary = {
                "total": total,
                "attended": attended,
                "attendance_rate": attendance_rate,
                "manual_additions": manual_additions,
                "missing_exam": missing_exam,
                "missing_hw": missing_hw,
                "cancellations": cancellations,
            }

            self.show_session_summary(session_name=session_name, summary=summary, session_path=path_entry, read_only=False)
            self.set_status(f"Showing summary for '{session_name}'.")
        except Exception as e:
            messagebox.showerror("Summary Error", f"Could not generate summary for the session:\n{e}", parent=self)

    def _format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"

    def _clear_all_sessions(self):
        # NOTE: This method is copied almost verbatim.
        # It scans the sessions directory and deletes files.
        sessions_dir = get_sessions_folder()
        if not os.path.isdir(sessions_dir):
            return

        paths_to_delete = []
        for root, _dirs, files in os.walk(sessions_dir):
            for filename in files:
                paths_to_delete.append(os.path.join(root, filename))

        if not paths_to_delete:
            return

        confirm = messagebox.askyesno(
            "Clear All Sessions",
            f"This will permanently delete {len(paths_to_delete)} session file(s). Are you sure?",
            parent=self
        )
        if not confirm:
            return

        failures = []
        for path_entry in paths_to_delete:
            try:
                os.remove(path_entry)
            except Exception as exc:
                failures.append(f"{os.path.basename(path_entry)}: {exc}")

        for root, _dirs, _files in os.walk(sessions_dir, topdown=False):
            if root == sessions_dir:
                continue
            try:
                os.rmdir(root)
            except OSError:
                pass

        self._invalidate_session_files_cache()
        self._populate_past_sessions_list()
        self._refresh_recent_sessions()

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
        """Shows the past sessions view without rebuilding the entire layout."""
        self._build_past_sessions_view()
        if self._past_sessions_list_dirty:
            self._populate_past_sessions_list()
        if self.past_sessions_frame is not None:
            self.past_sessions_frame.tkraise()

    def _build_past_sessions_view(self):
        """Builds the UI for browsing past sessions within the content_frame."""
        if self._past_sessions_initialized or self.past_sessions_frame is None:
            return

        frame = self.past_sessions_frame
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        # Header with title and global actions
        header_frame = ctk.CTkFrame(frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header_frame, text="Past Sessions", font=("Roboto", 24, "bold")).grid(row=0, column=0, sticky="w")

        action_buttons_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        action_buttons_frame.grid(row=0, column=1, sticky="e")

        ctk.CTkButton(action_buttons_frame, text="Refresh", command=lambda: self._populate_past_sessions_list(force_scan=True)).pack(side="left", padx=(0, 10))

        # Removed Clear All button

        # Scrollable frame for the list
        sessions_container = ctk.CTkFrame(frame, fg_color="transparent")
        sessions_container.grid(row=1, column=0, sticky="nsew")
        sessions_container.grid_rowconfigure(0, weight=1)
        sessions_container.grid_columnconfigure(0, weight=1)

        self.sessions_list_frame = ctk.CTkScrollableFrame(sessions_container, label_text="Session Files")
        self.sessions_list_frame.pack(fill="both", expand=True)

        # Populate the list with session data
        self._populate_past_sessions_list(force_scan=True)
        self._past_sessions_initialized = True

    def _populate_past_sessions_list(self, *, force_scan=False):
        """Fetches session data and populates the scrollable list with custom widgets."""
        if force_scan:
            self._invalidate_session_files_cache()

        if not hasattr(self, "sessions_list_frame"):
            return

        for widget in self.sessions_list_frame.winfo_children():
            widget.destroy()

        files = self._get_session_files()

        if not files:
            ctk.CTkLabel(self.sessions_list_frame, text="No session files found.").pack(pady=20)
            self._past_sessions_list_dirty = False
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
        self._past_sessions_list_dirty = False

    def _load_last_data(self):
        self.data_df = None
        self.current_data_path = None
        if os.path.exists(LAST_DATA_FILE):
            try:
                os.remove(LAST_DATA_FILE)
            except OSError as e:
                print(f"Could not clear last data file: {e}")

        if hasattr(self, "start_card_subtitle"):
            self._update_ui_for_data_state()
        self.set_status("Ready.")

    def open_settings(self):
        dialog = PasswordDialog(self)
        password = dialog.get_input()

        if password == "gawish1":
            if self.settings_window is not None and self.settings_window.winfo_exists():
                bring_window_to_front(self.settings_window)
                return
            self.settings_window = SettingsWindow(self)
            bring_window_to_front(self.settings_window)
            self.settings_window.protocol("WM_DELETE_WINDOW", self._on_settings_close)
        elif password is not None:
            messagebox.showerror("Incorrect Password", "The password you entered is incorrect.")

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

        self.set_status(f"Imported {len(df)} records from {os.path.basename(path)}.")
        return True

    def _collect_session_history(self, stages, centers):
        session_map = {}
        if not stages or not centers:
            return session_map
        sessions_dir = get_sessions_folder()
        if not os.path.isdir(sessions_dir):
            return session_map

        try:
            iterator = os.walk(sessions_dir)
        except OSError:
            return session_map

        for root, _dirs, files in iterator:
            for entry in files:
                name, ext = os.path.splitext(entry)
                if ext.lower() not in ('.csv', '.xlsx'):
                    continue
                if ' session ' not in name:
                    continue
                try:
                    prefix, number_str = name.rsplit(' session ', 1)
                except ValueError:
                    continue
                if not number_str.isdigit():
                    continue
                number = int(number_str)

                stage_candidate = None
                center_candidate = None

                rel_root = os.path.relpath(root, sessions_dir)
                if rel_root != '.':
                    parts = rel_root.split(os.sep)
                    if len(parts) >= 1:
                        maybe_stage = parts[0]
                        if maybe_stage in stages:
                            stage_candidate = maybe_stage
                    if len(parts) >= 2:
                        maybe_center = parts[1]
                        if maybe_center in centers:
                            center_candidate = maybe_center

                if not stage_candidate or not center_candidate:
                    for stage in stages:
                        stage_prefix = f"{stage} "
                        if prefix.startswith(stage_prefix):
                            potential_center = prefix[len(stage_prefix):]
                            if potential_center in centers:
                                stage_candidate = stage
                                center_candidate = potential_center
                            break

                if stage_candidate and center_candidate:
                    center_map = session_map.setdefault(stage_candidate, {})
                    center_map[center_candidate] = max(center_map.get(center_candidate, 0), number)
        return session_map

    def open_scan_window_setup(self):
        if self._session_setup is not None and self._session_setup.winfo_exists():
            bring_window_to_front(self._session_setup)
            return

        if self.data_df is None:
             messagebox.showwarning("No Data", "Please import data before starting a session.")
             return

        stage_options = SETTINGS.get("stage_options", []) or []
        center_options = SETTINGS.get("center_options", []) or []
        session_history = self._collect_session_history(stage_options, center_options)

        try:
            self._session_setup = SessionSetupDialog(
                self,
                stage_options,
                center_options,
                has_data=self.data_df is not None,
                session_data=session_history,
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
        session_path = resolve_session_file_path(
            name,
            stage=params.get("stage"),
            center=params.get("center"),
            ext=ext,
            create=True,
        )

        created = False
        if not os.path.exists(session_path) or messagebox.askyesno("Overwrite Session?", f"Session '{name}' already exists. Do you want to overwrite it with the currently loaded roster?"):
            os.makedirs(os.path.dirname(session_path), exist_ok=True)
            write_data(self.data_df, session_path)
            created = True

        session_df = read_data(session_path)
        sm = SessionManager(name, params, self.column_map, session_df, session_path=session_path)
        
        self._invalidate_session_files_cache()
        self._refresh_recent_sessions()
        self._populate_past_sessions_list()
        if self.past_sessions_window is not None and self.past_sessions_window.winfo_exists():
            self.past_sessions_window.refresh()
            
        ScanWindow(self, sm)
        
        if created:
            self.set_status(f"Session '{name}' created/overwritten.")
        else:
            self.set_status(f"Session '{name}' loaded.")
