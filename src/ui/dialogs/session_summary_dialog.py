"""Dialog that summarizes session statistics and provides file access shortcuts."""
import os
import subprocess
import sys
from tkinter import messagebox
import tkinter as tk
from PIL import Image, ImageDraw, ImageTk, ImageFont

from customtkinter import CTkButton, CTkFrame, CTkLabel, CTkToplevel, CTkImage

from utils.helpers import MIN_SUMMARY_SIZE, bring_window_to_front, ensure_initial_size, ASSETS_DIR, set_dark_title_bar


class CircularProgressBar(CTkFrame):
    def __init__(self, parent, size=100, progress_width=10,
                 track_width=10, progress_color="#22c55e", track_color="#3e4046",
                 text_font=("Roboto", 24, "bold"),
                 text_color="#ffffff"):
        super().__init__(parent, fg_color="transparent")

        self.size = size
        self.progress_width = progress_width
        self.track_width = track_width
        self.progress_color = progress_color
        self.track_color = track_color
        self.text_font = text_font
        self.text_color = text_color
        self.value = 0
        self.scale_factor = 4  # MODIFIED: Draw at 4x the resolution

        self.canvas = tk.Canvas(
            self,
            width=self.size,
            height=self.size,
            bg="#2b2d30",
            bd=0,
            highlightthickness=0
        )
        self.canvas.pack(expand=True)

        self.photo_image = None

        try:
            # MODIFIED: Load font at scaled size for high-res drawing
            scaled_font_size = text_font[1] * self.scale_factor
            self.font = ImageFont.truetype("Roboto-Regular.ttf", scaled_font_size)
        except IOError:
            self.font = ImageFont.load_default()

    def set_value(self, value):
        self.value = max(0, min(100, value))
        self._draw_progress()

    def _draw_progress(self):
        # MODIFIED: Create a larger canvas for supersampling
        scaled_size = self.size * self.scale_factor
        image = Image.new("RGBA", (scaled_size, scaled_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        # MODIFIED: All drawing dimensions are scaled up
        scaled_progress_width = self.progress_width * self.scale_factor
        scaled_track_width = self.track_width * self.scale_factor

        bbox = [
            scaled_progress_width / 2,
            scaled_progress_width / 2,
            scaled_size - scaled_progress_width / 2,
            scaled_size - scaled_progress_width / 2,
        ]

        draw.arc(bbox, start=-90, end=270, fill=self.track_color, width=scaled_track_width)

        if self.value > 0:
            end_angle = -90 + (self.value / 100) * 360
            draw.arc(bbox, start=-90, end=end_angle, fill=self.progress_color, width=scaled_progress_width)

        text_str = f"{self.value:.1f}%"
        # Use font.getbbox() for more accurate text bounding box
        left, top, right, bottom = self.font.getbbox(text_str)
        text_width = right - left
        text_height = bottom - top
        text_pos = (
            (scaled_size - text_width) / 2,
            (scaled_size - text_height) / 2 - top, # Adjust for vertical alignment
        )
        draw.text(text_pos, text_str, font=self.font, fill=self.text_color)

        # NEW: Resize the high-resolution image down with a high-quality filter
        # This is the step that creates the smooth anti-aliasing.
        image = image.resize((self.size, self.size), Image.Resampling.LANCZOS)

        self.photo_image = ImageTk.PhotoImage(image)

        self.canvas.delete("all")
        self.canvas.create_image(self.size / 2, self.size / 2, image=self.photo_image)


class SessionSummaryDialog(CTkToplevel):
    def __init__(self, parent, *, session_name, summary, session_path, params=None, read_only=False):
        super().__init__(parent)
        set_dark_title_bar(self)
        self.parent = parent
        self.session_name = session_name
        self.summary = summary or {}
        self.session_path = session_path
        self.params = params or {}
        self.read_only = read_only

        self.title("Session Summary")
        self.minsize(*MIN_SUMMARY_SIZE)
        self.transient(parent)
        self.grab_set()
        self.after(40, lambda: bring_window_to_front(self))

        container = CTkFrame(self, corner_radius=16, fg_color="#242424")
        container.pack(fill="both", expand=True, padx=24, pady=24)
        container.grid_columnconfigure(0, weight=1)

        header = CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        CTkLabel(header, text="Session Summary", font=("Roboto", 24, "bold")).grid(row=0, column=0, sticky="w")

        # Display Session Name
        details_frame = CTkFrame(header, fg_color="transparent")
        details_frame.grid(row=1, column=0, sticky="w", pady=(4, 0))

        session_display_name = os.path.splitext(os.path.basename(self.session_path))[0] if self.session_path else self.session_name
        CTkLabel(details_frame, text=session_display_name, font=("Roboto", 15, "bold"), text_color=("#1f6aa5", "#a9c8e7")).pack(side="left")

        if read_only:
            CTkLabel(header, text="Read-only session", font=("Roboto", 14), text_color="#64748b").grid(row=2, column=0, sticky="w", pady=(8, 0))

        metrics_frame = CTkFrame(container, fg_color="transparent")
        metrics_frame.grid(row=1, column=0, sticky="ew", pady=(18, 12))
        metrics_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="metrics_cols")

        # --- Helper to create a metric row ---
        def create_metric_row(parent, label, value):
            row = CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=(0, 4))
            row.grid_columnconfigure(1, weight=1)
            CTkLabel(row, text=label, font=("Roboto", 14), text_color="#d0d0d0").grid(row=0, column=0, sticky="w")
            CTkLabel(row, text=value, font=("Roboto", 14, "bold"), text_color="#ffffff").grid(row=0, column=1, sticky="e")

        def create_issue_row(parent, label, value):
            row = CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=(0, 4))
            row.grid_columnconfigure(1, weight=1)
            CTkLabel(row, text=label, font=("Roboto", 14), text_color="#d0d0d0").grid(row=0, column=0, sticky="w")

            numeric_value = None
            if value is not None:
                cleaned_value = str(value).replace(',', '')
                try:
                    numeric_value = int(cleaned_value)
                except (TypeError, ValueError):
                    numeric_value = None

            is_warning = numeric_value is not None and numeric_value > 0
            display_value = f"{numeric_value:,}" if numeric_value is not None else ("N/A" if value is None else str(value))
            value_color = "#ef4444" if is_warning else "#f8fafc"

            CTkLabel(row, text=display_value, font=("Roboto", 16, "bold"), text_color=value_color).grid(row=0, column=1, sticky="e")


        # --- Card 1: Overview ---
        overview_card = CTkFrame(metrics_frame, fg_color="#2b2d30", corner_radius=12, border_width=1, border_color="#3e4046")
        overview_card.grid(row=0, column=0, sticky="nsew", padx=(8, 8))
        overview_card.pack_propagate(False)

        title_frame_1 = CTkFrame(overview_card, fg_color="transparent", corner_radius=0)
        title_frame_1.pack(fill="x", side="top")
        title_frame_1.grid_columnconfigure(1, weight=1) # Make the center column expandable

        # Load and place the icon
        try:
            icon_path = os.path.join(ASSETS_DIR, "location_home.png")
            icon_image = CTkImage(Image.open(icon_path), size=(20, 20))
            CTkLabel(title_frame_1, image=icon_image, text="").grid(row=0, column=0, padx=(12, 5), pady=8)
        except Exception as e:
            print(f"Warning: Could not load 'location_home.png' icon: {e}")

        CTkLabel(title_frame_1, text="Overview", font=("Roboto", 16, "bold")).grid(row=0, column=1, pady=8)
        overview_content = CTkFrame(overview_card, fg_color="transparent")
        overview_content.pack(fill="both", expand=True, padx=12, pady=(10, 14))

        if (total := self.summary.get("total")) is not None:
            create_metric_row(overview_content, "Total students:", f"{total:,}")
        if (attended := self.summary.get("attended")) is not None:
            create_metric_row(overview_content, "Attended:", f"{attended:,}")
        if (manual := self.summary.get("manual_additions")) is not None:
            create_metric_row(overview_content, "Manual additions:", f"{manual:,}")

        # --- Card 2: Attendance Rate ---
        rate_card = CTkFrame(metrics_frame, fg_color="#2b2d30", corner_radius=12, border_width=1, border_color="#3e4046")
        rate_card.grid(row=0, column=1, sticky="nsew", padx=(4, 4))
        rate_card.pack_propagate(False)
        title_frame_2 = CTkFrame(rate_card, fg_color="transparent", corner_radius=0)
        title_frame_2.pack(fill="x", side="top")
        title_frame_2.grid_columnconfigure(1, weight=1) # Make the center column expandable

        # Load and place the icon
        try:
            icon_path = os.path.join(ASSETS_DIR, "bar_chart.png")
            icon_image = CTkImage(Image.open(icon_path), size=(20, 20))
            CTkLabel(title_frame_2, image=icon_image, text="").grid(row=0, column=0, padx=(12, 5), pady=8)
        except Exception as e:
            print(f"Warning: Could not load 'bar_chart.png' icon: {e}")

        CTkLabel(title_frame_2, text="Attendance Rate", font=("Roboto", 16, "bold")).grid(row=0, column=1, pady=8)
        rate_content = CTkFrame(rate_card, fg_color="transparent")
        rate_content.pack(fill="both", expand=True, padx=12, pady=10)

        if (rate := self.summary.get("attendance_rate")) is not None:
            try:
                rate_value = float(str(rate).strip('%'))
            except (ValueError, TypeError):
                rate_value = 0.0

            if rate_value >= 80:
                progress_color = "#22c55e"
            elif rate_value < 50:
                progress_color = "#ef4444"
            else:
                progress_color = "#f59e0b"

            progress_bar = CircularProgressBar(
                rate_content,
                size=110,
                progress_width=10,
                track_width=8,
                progress_color=progress_color,
                track_color="#4b5563",
                text_font=("Roboto", 22, "bold")
            )
            progress_bar.pack(expand=True, pady=(5, 0))
            progress_bar.set_value(rate_value)

        # --- Card 3: Issues ---
        issues_card = CTkFrame(metrics_frame, fg_color="#2b2d30", corner_radius=12, border_width=1, border_color="#3e4046")
        issues_card.grid(row=0, column=2, sticky="nsew", padx=(8, 8))
        issues_card.pack_propagate(False)
        title_frame_3 = CTkFrame(issues_card, fg_color="transparent", corner_radius=0)
        title_frame_3.pack(fill="x", side="top")
        title_frame_3.grid_columnconfigure(1, weight=1) # Make the center column expandable

        # Load and place the icon
        try:
            icon_path = os.path.join(ASSETS_DIR, "shield_person.png")
            icon_image = CTkImage(Image.open(icon_path), size=(20, 20))
            CTkLabel(title_frame_3, image=icon_image, text="").grid(row=0, column=0, padx=(12, 5), pady=8)
        except Exception as e:
            print(f"Warning: Could not load 'shield_person.png' icon: {e}")

        CTkLabel(title_frame_3, text="Issues & Flags", font=("Roboto", 16, "bold")).grid(row=0, column=1, pady=8)
        issues_content = CTkFrame(issues_card, fg_color="transparent")
        issues_content.pack(fill="both", expand=True, padx=12, pady=(10, 14))

        if (cancels := self.summary.get("cancellations")) is not None:
            create_issue_row(issues_content, "Cancellations:", cancels)
        if (missing_exam := self.summary.get("missing_exam")) is not None:
            create_issue_row(issues_content, "Missing exam:", missing_exam)
        if (missing_hw := self.summary.get("missing_hw")) is not None:
            create_issue_row(issues_content, "Missing homework:", missing_hw)

        actions = CTkFrame(container, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", pady=(24, 0))

        close_button = CTkButton(actions, text="Close", command=self._on_close)
        close_button.pack(fill="x", expand=True)

        ensure_initial_size(self, min_size=MIN_SUMMARY_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if hasattr(self.parent, "summary_window") and self.parent.summary_window is self:
            self.parent.summary_window = None
        try:
            self.grab_release()
        except Exception:
            pass
        if self.winfo_exists():
            self.destroy()
