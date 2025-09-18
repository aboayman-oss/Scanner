"""Dialog that summarizes session statistics and provides file access shortcuts."""
import os
import subprocess
import sys
from tkinter import messagebox

from customtkinter import CTkButton, CTkFrame, CTkLabel, CTkToplevel

from utils.helpers import MIN_SUMMARY_SIZE, bring_window_to_front, ensure_initial_size

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
