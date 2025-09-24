"""
This file contains the ScanWindow class, which manages the student attendance scanning process.

The "Focus View" feature has been redesigned into a modern, Material 3-style interface
to provide a guided, conversational user experience. All changes for this redesign are
encapsulated within this file, primarily in the `scan_focus_` prefixed methods.
"""
import os
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk
from customtkinter import CTkButton, CTkEntry, CTkFrame, CTkLabel, CTkProgressBar, CTkTextbox, CTkToplevel
from PIL import Image

from ui.dialogs.add_student_dialog import AddStudentDialog
from utils.helpers import (
    HOME_BG_FILE,
    MIN_SCAN_SIZE,
    bring_window_to_front,
    ensure_initial_size,
    read_data,
)
from .focus_view_window import FocusViewWindow

# Located at the top of scan_window.py, after the other imports

# +++ FINAL DEFINITIVE VERSION - REPLACE THE PREVIOUS BLOCK WITH THIS +++
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    def _process_arabic(text):
        """A helper that reshapes and reorders Arabic text."""
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)

except ImportError:
    print("WARNING: Arabic text support is limited. Please run: pip install arabic_reshaper python-bidi")
    # If libraries are missing, create a dummy function that does nothing.
    _process_arabic = lambda text: text

def _format_arabic_text(text):
    """
    Correctly formats Arabic text for display in the UI.
    It checks for Arabic characters before processing.
    """
    if not text:
        return text
    
    text_str = str(text)
    # Only process strings that contain Arabic characters to avoid errors.
    if not any('\u0600' <= char <= '\u06FF' for char in text_str):
        return text_str
    
    return _process_arabic(text_str)

# --- Constants for the new Focus View Design ---
ASSETS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "assets")

# Dark Mode
DARK_BG = "#1d1b20"
DARK_SURFACE = "#141218"
DARK_PRIMARY_TEXT = "#e3e2e6"
DARK_SECONDARY_TEXT = "#cac4d0"
DARK_SUCCESS = "#b5d3a7"
DARK_WARNING = "#f9d694"
DARK_ERROR = "#f2b8b5"
DARK_INFO = "#a9c8e7"

# -- Status Definitions --
STATUS_STYLES = {
    "ok": {
        "text": "All Clear",
        "icon": "check_circle.png",
        "color": DARK_SUCCESS,
    },
    "already_attended": { # New status for duplicate attendance
        "text": "Already Attended",
        "icon": "gpp_good.png", # Using a verified-style icon
        "color": DARK_INFO,
    },
    "missing_exam": {
        "text": "Tasks Missing",
        "icon": "warning.png",
        "color": DARK_WARNING,
    },
    "missing_homework": {
        "text": "Tasks Missing",
        "icon": "warning.png",
        "color": DARK_WARNING,
    },
    "not_found": {
        "text": "New Student",
        "icon": "person_add.png",
        "color": DARK_INFO,
    },
    "duplicate": {
        "text": "Duplicate Card",
        "icon": "error.png",
        "color": DARK_ERROR,
    },
}


class ScanWindow(CTkToplevel):
    def _reset_treeview_sort(self):
        """Restore Treeview rows to their original order."""
        self._tree_sort_column = None
        self._tree_sort_reverse = False
        # Detach all
        for iid in self._all_iids:
            if self.tree.exists(iid):
                self.tree.detach(iid)
        # Re-attach in original order
        for iid in self._all_iids:
            if self.tree.exists(iid):
                self.tree.reattach(iid, '', 'end')

    def __init__(self, parent, session_mgr, read_only=False):
        super().__init__(parent)
        self.parent = parent
        self.sm = session_mgr
        self.read_only = read_only
        self.state('zoomed')
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.toggle_fullscreen)
        self.bind("<s>", self._on_s_key_press)
        self.restrictions = self.sm.restrictions
        self.df = read_data(self.sm.session_path).fillna("")
        self.mapping = self.sm.mapping or {col: col for col in self.df.columns}

        # --- Icon Cache ---
        self._icon_cache = {}

        # --- Filter State ---
        self._filter_panel = None
        self._filter_vars = {
            "attendance": ctk.StringVar(value="all"),
            "missing_exam": ctk.BooleanVar(value=False),
            "missing_hw": ctk.BooleanVar(value=False),
            "has_notes": ctk.BooleanVar(value=False),
            "manual_added": ctk.BooleanVar(value=False),
        }
        self._filter_active = False

        # Remove background image; use solid surface panel for contrast
        self.bg_label = None  # No background image

        self.title("Scan Attendance")
        self.protocol("WM_DELETE_WINDOW", self._on_end_scan)
        self.after(50, lambda: bring_window_to_front(self))

        # --- Instance Variables ---
        self._all_iids = []
        self._search_entries = []
        self.search_var = None
        self._manual_additions = 0
        self._cancellations = 0
        self._focus_reset_job = None
        self._focus_guard_depth = 0
        self.scan_focus_ctx = None
        self.scan_focus_visible_cache = []
        self.scan_focus_timer = None
        self.focus_view_container = None # For integrated view
        self.stats_vars = {
            "total": ctk.StringVar(value="0"),
            "attended": ctk.StringVar(value="0"),
            "percent": ctk.StringVar(value="0%"),
            "missing_exam": ctk.StringVar(value="0"),
            "missing_hw": ctk.StringVar(value="0"),
        }

        self._build_ui()
        self._apply_treeview_style()
        self._load_existing()
        self._refresh_stats()
        ensure_initial_size(self, min_size=MIN_SCAN_SIZE)

        if not self.read_only:
            self.bind_all("<FocusIn>", self._global_focus_in, add="+")
            self.scan_entry.focus_set()

    def _on_filter_click(self):
        # Toggle filter panel visibility
        if self._filter_panel and self._filter_panel.winfo_exists():
            self._hide_filter_panel()
        else:
            self._show_filter_panel()

    def _show_filter_panel(self):
        # Create panel if not exists
        if self._filter_panel and self._filter_panel.winfo_exists():
            self._filter_panel.lift()
            return
        panel_width = 320
        panel = CTkFrame(self, fg_color="#232a36", corner_radius=12, width=panel_width)
        self._filter_panel = panel
        self.update_idletasks()
        # Center panel horizontally above filter icon
        bx = self.filter_button.winfo_rootx()
        by = self.filter_button.winfo_rooty() + self.filter_button.winfo_height()
        icon_width = self.filter_button.winfo_width()
        x = bx - self.winfo_rootx() + (icon_width // 2) - (panel_width // 2)
        y = by - self.winfo_rooty()
        panel.place(x=x, y=y)

        # Top bar with X button
        top_bar = CTkFrame(panel, fg_color="transparent")
        top_bar.pack(fill="x", padx=0, pady=(0,0))
        CTkLabel(top_bar, text="Filters", font=("Arial", 14, "bold"), anchor="w").pack(side="left", padx=(12,0), pady=(10,0))
        x_icon = self._load_icon("close.png", size=(20, 20))
        dismiss_btn = CTkButton(top_bar, text="", image=x_icon, width=32, height=32, fg_color="transparent", command=self._hide_filter_panel)
        dismiss_btn.pack(side="right", padx=(0,8), pady=(10,0))

        # Attendance Status (Radio)
        CTkLabel(panel, text="Attendance Status", font=("Arial", 12, "bold"), anchor="w").pack(anchor="w", padx=12, pady=(10,0))
        att_frame = CTkFrame(panel, fg_color="transparent")
        att_frame.pack(anchor="w", padx=12, pady=(0,4))
        for val, label in [("all", "All Students"), ("attend", "Attended"), ("absent", "Absent")]:
            ctk.CTkRadioButton(att_frame, text=label, variable=self._filter_vars["attendance"], value=val, command=self._on_filter_change).pack(side="left", padx=(0,12))

        # Task Status (Checkboxes)
        CTkLabel(panel, text="Task Status", font=("Arial", 12, "bold"), anchor="w").pack(anchor="w", padx=12, pady=(6,0))
        task_frame = CTkFrame(panel, fg_color="transparent")
        task_frame.pack(anchor="w", padx=12, pady=(0,4))
        ctk.CTkCheckBox(task_frame, text="Missing Exam", variable=self._filter_vars["missing_exam"], command=self._on_filter_change).pack(side="left", padx=(0,12))
        ctk.CTkCheckBox(task_frame, text="Missing H.W.", variable=self._filter_vars["missing_hw"], command=self._on_filter_change).pack(side="left", padx=(0,12))

        # Other Criteria (Checkboxes)
        CTkLabel(panel, text="Other Criteria", font=("Arial", 12, "bold"), anchor="w").pack(anchor="w", padx=12, pady=(6,0))
        other_frame = CTkFrame(panel, fg_color="transparent")
        other_frame.pack(anchor="w", padx=12, pady=(0,4))
        ctk.CTkCheckBox(other_frame, text="Has Notes", variable=self._filter_vars["has_notes"], command=self._on_filter_change).pack(side="left", padx=(0,12))
        ctk.CTkCheckBox(other_frame, text="Manually Added (No Card ID)", variable=self._filter_vars["manual_added"], command=self._on_filter_change).pack(side="left", padx=(0,12))

        # Clear Filters Button
        clear_btn = CTkButton(panel, text="Clear Filters", fg_color="#232a36", command=self._clear_filters)
        clear_btn.pack(fill="x", padx=12, pady=(10,10))

        self._filter_panel.lift()

    def _hide_filter_panel(self):
        if self._filter_panel and self._filter_panel.winfo_exists():
            self._filter_panel.place_forget()
            self._filter_panel.destroy()

    def _on_click_away(self, event):
        pass  # Removed click-away dismissal for filter panel

    def _on_filter_change(self):
        self._filter_active = self._is_filter_active()
        self._update_filter_icon()
        self._filter_all()

    def _clear_filters(self):
        for v in self._filter_vars.values():
            if isinstance(v, ctk.StringVar): v.set("all")
            else: v.set(False)
        self._filter_active = False
        self._update_filter_icon()
        self._filter_all()
        self._hide_filter_panel()

    def _is_filter_active(self):
        # Returns True if any filter is not default
        if self._filter_vars["attendance"].get() != "all": return True
        if self._filter_vars["missing_exam"].get(): return True
        if self._filter_vars["missing_hw"].get(): return True
        if self._filter_vars["has_notes"].get(): return True
        if self._filter_vars["manual_added"].get(): return True
        return False

    def _update_filter_icon(self):
        # Change icon to filled if filter active
        icon_name = "filter.png" if not self._filter_active else "filter_filled.png"
        self.filter_button.configure(image=self._load_icon(icon_name, size=(28, 28)))

    # --------------------------------------------------------------------------
    # Modern: Edit Notes Modal on Treeview Right-Click
    # --------------------------------------------------------------------------
    def _on_tree_right_click(self, event):
        # Get row under mouse
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        # Open modal dialog for editing notes
        self._open_edit_notes_modal(iid)

    def _open_edit_notes_modal(self, iid):
        # Get current notes
        current_notes = self.tree.set(iid, "notes")
        student_name = self.tree.set(iid, "name") or "Student"
        modal = CTkToplevel(self)
        modal.title(f"Edit Notes - {student_name}")
        modal.geometry("420x260")
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)
        modal.attributes("-topmost", True)

        # Modal styling
        frame = CTkFrame(modal, fg_color=DARK_SURFACE, corner_radius=16)
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        CTkLabel(frame, text=f"Edit Notes for {student_name}", font=("Arial", 15, "bold"), anchor="w").pack(anchor="w", pady=(0,8))
        notes_box = CTkTextbox(frame, width=360, height=90, font=("Arial", 13), corner_radius=10)
        notes_box.pack(fill="x", pady=(0,12))
        notes_box.insert("1.0", current_notes)

        # Save button
        def save_notes():
            new_notes = notes_box.get("1.0", "end-1c")
            self._update_row(iid, self.tree.set(iid, "attendance"), new_notes, self.tree.set(iid, "timestamp"))
            self._refresh_stats()
            modal.destroy()

        save_btn = CTkButton(frame, text="Save", fg_color="#a9c8e7", text_color="#232a36", font=("Arial", 13, "bold"), command=save_notes, width=120, height=38)
        save_btn.pack(side="right", pady=(8,0))

        # Focus for quick editing
        notes_box.focus_set()

        # Allow closing with Escape
        modal.bind("<Escape>", lambda e: modal.destroy())

    def _bind_tree_right_click(self):
        # Bind right-click to treeview for notes editing
        self.tree.bind("<Button-3>", self._on_tree_right_click)


    

    def toggle_fullscreen(self, event=None):
        self.attributes("-fullscreen", not self.attributes("-fullscreen"))

    def _on_s_key_press(self, event):
        """Handler for 's' key press to focus the scan entry."""
        # Check if focus is already in a text entry field to avoid interruption
        focused_widget = self.focus_get()
        if isinstance(focused_widget, (CTkEntry, CTkTextbox)):
            return  # Don't steal focus if the user is typing
        
        self.scan_entry.focus_set()

    # --------------------------------------------------------------------------
    # Redesigned Focus View (Material 3 Style)
    # --------------------------------------------------------------------------

    def _load_icon(self, name, size=(24, 24)):
        """
        Loads an icon from the assets directory and caches it.
        Icons are expected to be white for proper coloring.
        """
        if (name, size) in self._icon_cache:
            return self._icon_cache[(name, size)]
        
        try:
            img_path = os.path.join(ASSETS_PATH, name)
            img = Image.open(img_path)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=size)
            self._icon_cache[(name, size)] = ctk_img
            return ctk_img
        except FileNotFoundError:
            print(f"Warning: Icon '{name}' not found at '{ASSETS_PATH}'")
            # Return a placeholder transparent image
            return ctk.CTkImage(light_image=Image.new("RGBA", size, (0,0,0,0)),
                                dark_image=Image.new("RGBA", size, (0,0,0,0)),
                                size=size)

    def scan_focus_create_ui(self, parent):
        """
        Creates the Focus View UI using FocusViewWindow class.
        """
        self.focus_view = FocusViewWindow(
            parent,
            read_only=getattr(self, 'read_only', False),
            icon_cache=getattr(self, '_icon_cache', {}),
            on_complete=self.scan_focus_on_completed,
            on_add_student=self.scan_focus_on_add_student,
            on_override=self.scan_focus_on_override,
            on_deny=self.scan_focus_on_deny,
            on_cancel=self.scan_focus_on_cancel_attendance,
            on_dismiss=self.scan_focus_clear
        )

    def _on_notes_focus_in(self, event):
        self._pause_focus_guard()
        if self.focus_view.notes.get("1.0", "end-1c") == "Add notes here...":
            self.focus_view.notes.delete("1.0", "end")
            self.focus_view.notes.configure(text_color=DARK_PRIMARY_TEXT)

    def _on_notes_focus_out(self, event):
        self._resume_focus_guard()
        if not self.focus_view.notes.get("1.0", "end-1c"):
            self.focus_view.notes.configure(text_color="gray")
            self.focus_view.notes.insert("1.0", "Add notes here...")

    def scan_focus_show(self, scan_ctx):
        """Shows and populates the Focus View with student data."""
        self.scan_focus_cancel_timer()
        if not self.focus_view_container: return

        # Show the integrated focus view panel
        self.focus_view_container.grid()
        
        ctx = dict(scan_ctx or {})
        ctx.setdefault("original_notes", ctx.get("existing_notes", ""))
        self.scan_focus_ctx = ctx
        
        status = ctx.get("status") or self.scan_determine_status(ctx)
        ctx["status"] = status

        # Populate UI elements
        self.focus_view.name_label.configure(text=ctx.get("name") or "Unknown Student")
        student_name = ctx.get("name") or "Unknown Student"
        formatted_name = _format_arabic_text(student_name)
        self.focus_view.name_label.configure(text=formatted_name)
        card_display_val = ctx.get('card_display', '') or ''
        card_display = str(card_display_val).replace('null', '').strip() or '--'
        student_id_val = ctx.get('student_id', '') or ''
        student_id = str(student_id_val).replace('null', '').strip() or '--'
        id_text = f"Student ID: {student_id}  •  Card ID: {card_display}"
        self.focus_view.id_label.configure(text=id_text)

        # Set notes
        if not self.read_only: self.focus_view.notes.configure(state="normal")
        self.focus_view.notes.delete("1.0", "end")
        existing_notes = ctx.get("existing_notes", "")
        if existing_notes:
            self.focus_view.notes.insert("1.0", existing_notes)
            self.focus_view.notes.configure(text_color=DARK_PRIMARY_TEXT)
        else:
            self.focus_view.notes.configure(text_color="gray")
            self.focus_view.notes.insert("1.0", "Add notes here...")
        if self.read_only: self.focus_view.notes.configure(state="disabled")

        # Filter the main table view
        focus_iids = ctx.get("focus_iids") or []
        if focus_iids and not ctx.get("skip_filter") and ctx.get("iid") is not None:
            self.scan_filter_for_focus(focus_iids)
        else:
            if focus_iids:
                primary = focus_iids[0]
                if self.tree.exists(primary):
                    self.tree.selection_set(primary); self.tree.focus(primary)
            self.scan_restore_from_focus()

        # Set status and update dynamic UI parts
        self.scan_focus_set_status(status, ctx)

    def scan_focus_set_status(self, kind, context):
        """Updates the entire Focus View UI based on the student's status."""
        if self.scan_focus_ctx: self.scan_focus_ctx["status"] = kind

        # 1. Update Status Icon
        style = STATUS_STYLES.get(kind, STATUS_STYLES["ok"])
        self.focus_view.status_icon.configure(image=self._load_icon(style["icon"], size=(48, 48)))

        # 2. Update Details Cards (Homework & Exam)
        missing_tasks = context.get("missing_tasks", [])
        success_icon = self._load_icon("task_alt.png")
        problem_icon = self._load_icon("error.png")

        # Subtle container colors
        success_color = "#1b331d" # Material Green Dark
        problem_color = "#3c1b1a" # Material Red Dark

        # Homework
        hw_missing = "homework" in missing_tasks
        self.focus_view.hw_icon_label.configure(image=problem_icon if hw_missing else success_icon)
        hw_grade = context.get("homework", "")
        hw_text = ""
        if hw_grade:
            hw_text = str(hw_grade)
            if hw_grade == "0":
                hw_text += " (Fail)"
        else:
            hw_text = "Not Submitted"
        self.focus_view.hw_grade_label.configure(text=hw_text)
        self.focus_view.hw_card.configure(fg_color=problem_color if hw_missing else success_color)

        # Exam
        exam_missing = "exam" in missing_tasks
        self.focus_view.exam_icon_label.configure(image=problem_icon if exam_missing else success_icon)
        exam_grade = context.get("exam", "")
        exam_text = ""
        if exam_grade:
            exam_text = str(exam_grade)
            if exam_grade == "0":
                exam_text += " (Fail)"
        else:
            exam_text = "Not Submitted"
        self.focus_view.exam_grade_label.configure(text=exam_text)
        self.focus_view.exam_card.configure(fg_color=problem_color if exam_missing else success_color)

        # 3. Update Action Buttons
        self._update_action_buttons(kind, context)

    def _update_action_buttons(self, kind, context):
        """Shows and hides the correct action buttons using a stable grid layout."""
        # Hide all buttons first
        for btn in self.focus_view.buttons:
            btn.grid_remove()

        # Determine which buttons to show and place them in the grid
        if kind == "not_found":
            # CHANGED: Place the single button in the center column (1)
            # This leaves columns 0 and 2 as empty spacers, maintaining width.
            self.focus_view.btn_add_student.grid(row=0, column=1, sticky="ew", padx=2)

        elif kind in {"missing_exam", "missing_homework"}:
            # UNCHANGED: This layout already uses all three columns correctly.
            self.focus_view.btn_deny.grid(row=0, column=0, sticky="ew", padx=2)
            self.focus_view.btn_override.grid(row=0, column=1, sticky="ew", padx=2)
            self.focus_view.btn_complete.grid(row=0, column=2, sticky="ew", padx=2)

        elif context.get("already_attended"):
            # CHANGED: Place the single button in the center column (1)
            self.focus_view.btn_cancel.grid(row=0, column=1, sticky="ew", padx=2)

        elif kind == "ok":
            # No buttons are needed, the grid remains empty but holds its space
            pass

    def scan_focus_clear(self):
        """Hides the Focus View and resets its state."""
        self.scan_focus_cancel_timer()
        self.scan_focus_ctx = None
        
        if hasattr(self, "focus_view"):
            if not self.read_only: self.focus_view.notes.configure(state="normal")
            self.focus_view.notes.delete("1.0", "end")
            self.focus_view.notes.configure(text_color="gray")
            self.focus_view.notes.insert("1.0", "Add notes here...")

        self.scan_restore_from_focus()
        
        if self.focus_view_container:
            self.focus_view_container.grid_remove()
            
        self.after(120, self.scan_entry.focus_set)

    # --------------------------------------------------------------------------
    # Original ScanWindow methods (unchanged unless necessary for integration)
    # --------------------------------------------------------------------------

    def _on_bg_resize(self, event):
        pass  # No background image to resize

    def _focus_scan_entry(self):
        self._focus_reset_job = None
        if self.read_only or self._focus_guard_depth > 0: return
        try: self.scan_entry.focus_set() 
        except Exception: pass

    def _pause_focus_guard(self):
        if self._focus_reset_job is not None:
            try: self.after_cancel(self._focus_reset_job)
            except Exception: pass
            self._focus_reset_job = None
        self._focus_guard_depth += 1

    def _resume_focus_guard(self):
        if self._focus_guard_depth > 0: self._focus_guard_depth -= 1


    def _build_ui(self):
        # --- Header Bar ---
        top_bar = CTkFrame(self, fg_color="#1d1b20", corner_radius=16)
        top_bar.pack(fill="x", padx=24, pady=(24, 16))
        top_bar.grid_columnconfigure(0, weight=0)
        top_bar.grid_columnconfigure(1, weight=1)
        top_bar.grid_columnconfigure(2, weight=0)
        top_bar.grid_columnconfigure(3, weight=0)

        # --- Scan Entry ---
        scan_icon = self._load_icon("scan.png", size=(28, 28))
        scan_entry_frame = CTkFrame(top_bar, fg_color="transparent")
        scan_entry_frame.grid(row=0, column=0, sticky="w", padx=(0, 12))
        scan_icon_label = CTkLabel(scan_entry_frame, image=scan_icon, text="", width=32)
        scan_icon_label.pack(side="left", padx=(0, 8))
        self.scan_entry = CTkEntry(scan_entry_frame, width=260, height=44, placeholder_text="Scan card ID (place card here)")
        self.scan_entry.pack(side="left", padx=(0, 0), pady=0)
        self.scan_entry.bind("<Return>", lambda _e: self.scan_on_scan())
        self.pb = CTkProgressBar(scan_entry_frame, mode="indeterminate", width=260)
        self.pb.pack_forget()

        # --- Add Student Button ---
        add_icon = self._load_icon("person_add.png", size=(32, 32))
        self.add_student_button = CTkButton(top_bar, width=44, height=44, text="", image=add_icon, fg_color="#232a36", corner_radius=22, command=self._on_add_student_flow)
        self.add_student_button.grid(row=0, column=1, sticky="w", padx=(0, 12))
        if self.read_only:
            self.scan_entry.configure(state="disabled"); self.scan_entry.unbind("<Return>"); self.add_student_button.grid_remove()

        # --- Search & Filter ---
        search_filter_frame = CTkFrame(top_bar, fg_color="transparent")
        search_filter_frame.grid(row=0, column=2, sticky="ew", padx=(0, 12))
        self.search_var = ctk.StringVar()
        search_icon = self._load_icon("search.png", size=(24, 24))
        search_entry_frame = CTkFrame(search_filter_frame, fg_color="transparent")
        search_entry_frame.pack(side="left", padx=(0, 0))
        search_icon_label = CTkLabel(search_entry_frame, image=search_icon, text="", width=28)
        search_icon_label.pack(side="left", padx=(0, 6))
        search_entry = CTkEntry(search_entry_frame, textvariable=self.search_var, width=220, height=44, placeholder_text="Search by name, ID, card, or phone")
        search_entry.pack(side="left")
        self.search_var.trace_add("write", self._on_search_change)
        self._search_entries.append(search_entry)
        search_entry.bind("<FocusIn>", lambda _e: self._pause_focus_guard())
        search_entry.bind("<FocusOut>", lambda _e: self._resume_focus_guard())
        self.smart_search_entry = search_entry
        filter_icon = self._load_icon("filter.png", size=(28, 28))
        self.filter_button = CTkButton(search_filter_frame, width=44, height=44, text="", image=filter_icon, fg_color="#232a36", corner_radius=22, command=self._on_filter_click)
        self.filter_button.pack(side="left", padx=(8, 0))

        # --- Actions ---
        actions_frame = CTkFrame(top_bar, fg_color="transparent")
        actions_frame.grid(row=0, column=3, sticky="e", padx=(0, 0))
        self.end_button = CTkButton(actions_frame, text="End Session" if not self.read_only else "Close", command=self._on_end_scan, width=120, height=44, fg_color="#a9c8e7", text_color="#232a36", font=("Arial", 14, "bold"))
        self.end_button.pack(side="right", padx=(0, 0))

        # --- Stats strip ---
        self._build_stats_strip()

        # --- Main content area with integrated Focus View ---
        main_body = CTkFrame(self, fg_color="transparent")
        main_body.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        main_body.grid_rowconfigure(0, weight=1)
        main_body.grid_columnconfigure(0, weight=1)
        main_body.grid_columnconfigure(1, weight=0) # Focus view column, initially no weight

        # --- Treeview Container (Left/Main) ---
        tree_outer_container = CTkFrame(main_body, fg_color=DARK_SURFACE, corner_radius=18)
        tree_outer_container.grid(row=0, column=0, sticky="nsew")
        tree_outer_container.grid_rowconfigure(0, weight=1)
        tree_outer_container.grid_columnconfigure(0, weight=1)

        tree_container = CTkFrame(tree_outer_container, fg_color="transparent")
        tree_container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        cols = ["card_id", "student_id", "name", "phone"]
        if self.restrictions.get("exam"): cols.append("exam")
        if self.restrictions.get("homework"): cols.append("homework")
        cols += ["attendance", "notes", "timestamp"]

        # Manual column widths
        column_widths = {
            "card_id": 90,
            "student_id": 90,
            "name": 220,
            "phone": 130,
            "exam": 85,
            "homework": 85,
            "attendance": 100,
            "notes": 200,
            "timestamp": 100,
        }

        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings", selectmode="browse")
        for col in cols:
            width = column_widths.get(col, 110)
            self.tree.heading(col, text=col.replace("_", " ").title())
            if col == "notes":
                # Stretch notes column to fill remaining space and left-align text
                self.tree.column(col, anchor="w", width=width, stretch=True)
            else:
                self.tree.column(col, anchor="center", width=width, stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns"); self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self.scan_on_row_double_click)
        if self.read_only: self.tree.unbind("<Double-1>")
        # Bind up/down arrow keys for navigation
        self.tree.bind("<Return>", self._on_tree_enter)
        self.tree.bind("<Up>", self._on_tree_up_down)
        self.tree.bind("<Down>", self._on_tree_up_down)
        # Modern: Bind right-click for notes editing
        self._bind_tree_right_click()

        # --- Column Sorting ---
        self._tree_sort_column = None
        self._tree_sort_reverse = False
        for col in cols:
            self.tree.heading(col, command=lambda c=col: self._on_treeview_sort(c))

        # --- Focus View Container (Right, initially hidden) ---
        self.focus_view_container = CTkFrame(main_body, fg_color=DARK_SURFACE, corner_radius=18, width=400)
        self.focus_view_container.grid(row=0, column=1, sticky="ns", padx=(12, 0))
        self.focus_view_container.grid_propagate(False) # Prevent resizing
        self.scan_focus_create_ui(self.focus_view_container)
        self.focus_view_container.grid_remove() # Hide it initially

        # Configure column weights for resizing
        main_body.grid_columnconfigure(1, weight=0) # Focus view column

    def _on_treeview_sort(self, col):
        # Get all items and their values for the column
        items = [(iid, self.tree.set(iid, col)) for iid in self._all_iids if self.tree.exists(iid)]
        # Determine if numeric sort (for exam/homework)
        def is_number(val):
            try:
                float(val)
                return True
            except Exception:
                return False
        numeric_cols = {"exam", "homework"}
        # Use str(v).strip() to avoid attribute error
        def parse_score(val):
            val = str(val).strip()
            if not val:
                return float('-inf')  # Treat empty as lowest
            if '/' in val:
                try:
                    score, total = val.split('/', 1)
                    return float(score) / float(total) if float(total) != 0 else float('-inf')
                except Exception:
                    return float('-inf')
            try:
                return float(val)
            except Exception:
                return float('-inf')

        def sort_key(item):
            val = item[1]
            if col in {"exam", "homework"}:
                return parse_score(val)
            else:
                return str(val).lower()
        # Toggle sort order if same column
        if self._tree_sort_column == col:
            self._tree_sort_reverse = not self._tree_sort_reverse
        else:
            self._tree_sort_column = col
            self._tree_sort_reverse = False
        sorted_items = sorted(items, key=sort_key, reverse=self._tree_sort_reverse)
        # Detach all
        for iid in self._all_iids:
            if self.tree.exists(iid):
                self.tree.detach(iid)
        # Re-attach in sorted order
        for iid, _ in sorted_items:
            self.tree.reattach(iid, '', 'end')
    def _on_tree_enter(self, event):
        # Simulate double-click on selected row when Enter is pressed
        selected = self.tree.selection()
        if selected:
            self.scan_on_open_row(selected[0], source="manual")

    def _on_tree_up_down(self, event):
        """Move selection up or down in the Treeview respecting the current visual order."""
        selected_iid = self.tree.selection()

        # Get only the currently visible children in their visual order.
        visible_children = self.tree.get_children('')
        if not visible_children:
            return "break" # Nothing to navigate

        # If nothing is selected, select the first visible item and stop.
        if not selected_iid:
            first_item = visible_children[0]
            self.tree.selection_set(first_item)
            self.tree.focus(first_item)
            self.tree.see(first_item)
            return "break"

        current_iid = selected_iid[0]
        try:
            current_index = visible_children.index(current_iid)
        except ValueError:
            # The selected item is not visible, so select the first visible one
            first_item = visible_children[0]
            self.tree.selection_set(first_item)
            self.tree.focus(first_item)
            self.tree.see(first_item)
            return "break"

        # Determine the next index
        if event.keysym == "Up":
            next_index = current_index - 1
        else:  # Down
            next_index = current_index + 1
            
        # Select the new item if it's within bounds
        if 0 <= next_index < len(visible_children):
            next_item = visible_children[next_index]
            self.tree.selection_set(next_item)
            self.tree.focus(next_item)
            self.tree.see(next_item)
        
        # This is crucial: it prevents the default event from firing and causing a "skip".
        return "break"

    def scan_focus_cancel_timer(self):
        if self.scan_focus_timer is not None:
            try: self.after_cancel(self.scan_focus_timer)
            except Exception: pass
            self.scan_focus_timer = None

    def scan_focus_schedule_clear(self, delay=1000):
        self.scan_focus_cancel_timer()
        self.scan_focus_timer = self.after(delay, self.scan_focus_clear)

    def scan_restore_from_focus(self):
        if not self.scan_focus_visible_cache: return
        for scan_iid in self.scan_focus_visible_cache:
            if self.tree.exists(scan_iid):
                try: self.tree.reattach(scan_iid, "", "end")
                except Exception: pass
        self.scan_focus_visible_cache = []

    def scan_filter_for_focus(self, target_iids):
        self.scan_restore_from_focus()
        if not target_iids: return
        current_visible = [iid for iid in self._all_iids if self.tree.exists(iid) and not self.tree.parent(iid)]
        self.scan_focus_visible_cache = current_visible
        for scan_iid in current_visible:
            if scan_iid not in target_iids:
                try: self.tree.detach(scan_iid)
                except Exception: pass
        for scan_iid in target_iids:
            if self.tree.exists(scan_iid):
                try: self.tree.reattach(scan_iid, "", "end")
                except Exception: pass
        primary = target_iids[0]
        if self.tree.exists(primary):
            self.tree.selection_set(primary); self.tree.focus(primary)

    def scan_normalize_card(self, value):
        text = self._clean_value(value)
        return text.zfill(8) if text and text.isdigit() else text

    def scan_lookup_matches(self, card_id):
        normalized = self.scan_normalize_card(card_id)
        if not normalized: return []
        candidates = [iid for iid in self._all_iids if self.tree.exists(iid) and (self.scan_normalize_card(iid) == normalized or self.scan_normalize_card(self.scan_tree_get(iid, "card_id")) == normalized)]
        unique = sorted(list(set(candidates)), key=lambda x: (x != normalized))
        return unique

    def scan_tree_get(self, iid, column):
        if column not in self.tree["columns"]:
            return ""
        try: return self._clean_value(self.tree.set(iid, column))
        except Exception: return ""

    def scan_collect_missing_tasks(self, iid):
        missing = []
        exam_grade = self.scan_tree_get(iid, "exam")
        if self.restrictions.get("exam") and "exam" in self.tree["columns"] and (not exam_grade or exam_grade == "0"):
            missing.append("exam")
        
        hw_grade = self.scan_tree_get(iid, "homework")
        if self.restrictions.get("homework") and "homework" in self.tree["columns"] and (not hw_grade or hw_grade == "0"):
            missing.append("homework")
        return missing

    def scan_describe_tasks(self, tasks):
        if not tasks: return ""
        labels = {"exam": "Exam", "homework": "Homework"}
        mapped = [labels.get(task, str(task).title()) for task in tasks]
        if not mapped: return ""
        return mapped[0] if len(mapped) == 1 else " & ".join(mapped)

    def scan_append_notes(self, original, addition):
        original_clean, addition_clean = self._clean_value(original), self._clean_value(addition)
        if not addition_clean: return original_clean
        if not original_clean: return addition_clean
        return f"{original_clean.rstrip()}\n{addition_clean}"

    def scan_collect_new_note(self):
        if not hasattr(self, "focus_view"): return ""
        try:
            typed = self.focus_view.notes.get("1.0", "end-1c").strip()
        except Exception:
            return ""
        if typed == "Add notes here...": return ""
        return self._clean_value(typed)

    def scan_now_tag(self):
        return f"[{datetime.now():%H:%M:%S}]"

    def scan_determine_status(self, scan_ctx):
        if scan_ctx.get("status") in {"not_found", "duplicate"}: return scan_ctx["status"]
        if not scan_ctx.get("found", True): return "not_found"
        if scan_ctx.get("already_attended"): return "already_attended"
        missing = scan_ctx.get("missing_tasks", [])
        if missing: return "missing_exam" if "exam" in missing else "missing_homework"
        return "ok"

    def scan_build_context_for_iid(self, iid, *, source="manual"):
        context = {
            "iid": iid, "card_id": self.scan_normalize_card(iid),
            "card_display": self.scan_tree_get(iid, "card_id") or self.scan_normalize_card(iid),
            "name": self.scan_tree_get(iid, "name"), "student_id": self.scan_tree_get(iid, "student_id"),
            "attendance": self.scan_tree_get(iid, "attendance").lower(),
            "existing_notes": self.scan_tree_get(iid, "notes"), "timestamp": self.scan_tree_get(iid, "timestamp"),
            "source": source, "focus_iids": [iid], "found": True,
            "homework": self.scan_tree_get(iid, "homework"),
            "exam": self.scan_tree_get(iid, "exam"),
        }
        context["missing_tasks"] = self.scan_collect_missing_tasks(iid)
        context["already_attended"] = context["attendance"] == "attend"
        context["allow_cancel"] = context["already_attended"]
        context["status"] = self.scan_determine_status(context)
        context["display_name"] = context["name"] or context["student_id"] or context["card_display"] or "Student"
        return context

    def scan_build_not_found_context(self, card_id):
        return {
            "iid": None, "card_id": card_id, "card_display": card_id, "name": "Card Not Linked",
            "student_id": "", "attendance": "", "existing_notes": "", "timestamp": "",
            "source": "scan", "focus_iids": [], "found": False, "missing_tasks": [],
            "status": "not_found", "display_name": card_id or "Card",
        }

    def scan_on_scan(self):
        if self.read_only: return
        normalized = self.scan_normalize_card(self.scan_entry.get())
        self.scan_entry.delete(0, "end")
        if not normalized: return
        
        matches = self.scan_lookup_matches(normalized)
        if not matches:
            context = self.scan_build_not_found_context(normalized)
            self.scan_focus_show(context)
            return
        
        if len(matches) > 1:
            context = {
                "card_id": normalized, "card_display": normalized, "name": "Multiple Records Found",
                "student_id": "", "status": "duplicate", "focus_iids": matches, "skip_filter": True,
            }
            self.scan_focus_show(context)
            return
        
        self.scan_on_open_row(matches[0], source="scan", card_id=normalized)

    def scan_on_row_double_click(self, event):
        if self.read_only: return
        scan_iid = self.tree.identify_row(event.y) or (self.tree.selection() and self.tree.selection()[0])
        if scan_iid: self.scan_on_open_row(scan_iid, source="manual")

    def scan_on_open_row(self, iid, *, source="manual", card_id=None):
        if self.read_only or not self.tree.exists(iid): return
        
        context = self.scan_build_context_for_iid(iid, source=source)
        if card_id: context["card_id"] = context["card_display"] = card_id
        
        # --- Prevent duplicate attendance ---
        if context.get("already_attended"):
            self.scan_focus_show(context) # Show the focus window with the duplicate status
            return
        
        self.scan_focus_show(context)
        
        if context["status"] == "ok" and not context.get("already_attended"):
            self.scan_handle_auto_attend(context)

    def scan_handle_auto_attend(self, context):
        if not context: return
        tag = self.scan_now_tag()
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(context.get("existing_notes", ""), typed)
        success = self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag)
        if success:
            self.scan_focus_schedule_clear()

    def scan_commit_attendance(self, iid, attendance, notes, *, timestamp=None, warn_on_duplicate=False):
        try: return bool(self._set_attendance(iid, attendance, notes, warn_on_duplicate=warn_on_duplicate, timestamp_override=timestamp))
        except Exception as exc: messagebox.showwarning("Attendance Update Failed", str(exc), parent=self); return False # type: ignore

    def scan_focus_on_completed(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"): return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "task"
        action_note = f"{tag} Completed {desc} at center."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag):
            self.scan_focus_clear()

    def scan_focus_on_override(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"): return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "task"
        action_note = f"{tag} Attended (Didn't do {desc})."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag):
            self.scan_focus_clear()

    def scan_focus_on_deny(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"): return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "requirements"
        action_note = f"{tag} Denied Entry: No {desc}."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "", final_note, timestamp=tag):
            self.scan_focus_clear()

    def scan_focus_on_add_student(self):
        if self.read_only: return
        context = self.scan_focus_ctx or {}
        card_id = context.get("card_id") or context.get("card_display")
        typed = self.scan_collect_new_note()
        self._launch_add_student_dialog(card_id=card_id, default_notes=typed)

    def scan_focus_on_cancel_attendance(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"): return
        tag = self.scan_now_tag()
        action_note = f"{tag} Canceled."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        self._cancellations += 1
        if self.scan_commit_attendance(context["iid"], "", final_note, timestamp=tag):
            self.scan_focus_clear()

    def _build_stats_strip(self):
        # Compact horizontal stats bar
        self.stats_frame = CTkFrame(self, fg_color="#12263a", corner_radius=12, height=56)
        self.stats_frame.pack(fill="x", padx=24, pady=(0, 8))

        card_defs = [
            {"label": "Total Students", "var": self.stats_vars["total"], "icon": "group.png", "is_progress": False},
            {"label": "Attended", "var": self.stats_vars["attended"], "icon": "check_circle.png", "is_progress": False},
            {"label": "Attendance", "var": self.stats_vars["percent"], "icon": "group.png", "is_progress": True},
        ]
        if self.restrictions.get("exam"):
            card_defs.append({"label": "Missing Exam", "var": self.stats_vars["missing_exam"], "icon": "warning.png", "is_progress": False})
        if self.restrictions.get("homework"):
            card_defs.append({"label": "Missing Homework", "var": self.stats_vars["missing_hw"], "icon": "warning.png", "is_progress": False})

        # Place all cards in a single horizontal line, centered
        for idx, card in enumerate(card_defs):
            card_frame = CTkFrame(
                self.stats_frame,
                fg_color="#232a36",
                corner_radius=10,
                width=110,
                height=56
            )
            card_frame.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 10, 0), pady=4)
            self.stats_frame.grid_columnconfigure(idx, weight=1)

            # Center everything in the card
            card_inner = CTkFrame(card_frame, fg_color="transparent")
            card_inner.pack(expand=True, fill="both")

            CTkLabel(card_inner, text=card["label"], font=("Arial", 12, "bold"), text_color="#cac4d0", anchor="center", justify="center").pack(side="top", anchor="center", pady=(6, 0))

            icon_num_frame = CTkFrame(card_inner, fg_color="transparent")
            icon_num_frame.pack(side="top", anchor="center", pady=(0, 0), expand=True)
            icon_img = self._load_icon(card["icon"], size=(22, 22))
            icon_label = CTkLabel(icon_num_frame, image=icon_img, text="", width=24)
            icon_label.pack(side="left", anchor="center", padx=(0, 4))

            if card["is_progress"]:
                percent_str = self.stats_vars["percent"].get().replace("%", "")
                try:
                    percent_val = float(percent_str) / 100.0
                except Exception:
                    percent_val = 0.0
                progress = CTkProgressBar(icon_num_frame, width=40, height=6)
                progress.set(percent_val)
                progress.pack(side="left", anchor="center", padx=(0, 4))
                CTkLabel(icon_num_frame, textvariable=self.stats_vars["percent"], font=("Arial", 18, "bold"), text_color="#a9c8e7", anchor="center", justify="center").pack(side="left", anchor="center", padx=(0, 0))
            else:
                CTkLabel(icon_num_frame, textvariable=card["var"], font=("Arial", 20, "bold"), text_color="#e3e2e6", anchor="center", justify="center").pack(side="left", anchor="center", padx=(0, 0))

    def _apply_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("default")
        bg, fg, heading_bg, heading_fg = ("#1e1e1e", "#f2f2f2", "#1f6aa5", "#ffffff")
        style.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, rowheight=32, font=("Arial", 11))
        style.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background=heading_bg, foreground=heading_fg, font=("Arial", 11, "bold") )
        style.map("Treeview.Heading", background=[("active", heading_bg)])
        self.tree.configure(style="Treeview")

    def _load_existing(self):
        def pad_card_id(val):
            val_str = str(val).strip()
            return val_str.zfill(8) if val_str.isdigit() else val_str

        cols = self.tree["columns"]
        session_records = {pad_card_id(rec.get("card_id", "")): rec for rec in self.sm.records}
        self._all_iids = []

        for _, row in self.df.iterrows():
            cid = pad_card_id(row.get(self.mapping.get("card_id", "card_id"), ""))
            rec = session_records.pop(cid, None)
            # REPLACEMENT for the line above
            values = []
            for col in cols:
                val = self._clean_value(rec.get(col) if rec and col in rec else row.get(self.mapping.get(col, col), ""))
                if col == 'name':
                    val = _format_arabic_text(val)
                values.append(val)
            self.tree.insert("", "end", iid=cid, values=tuple(values))
            self._all_iids.append(cid)

        for cid, rec in session_records.items():
            values = [self._clean_value(rec.get(col, "")) for col in cols]
            self.tree.insert("", "end", iid=cid, values=tuple(values))
            self._all_iids.append(cid)

        for iid in self._all_iids:
            self._update_row(iid, self.scan_tree_get(iid, "attendance"), self.scan_tree_get(iid, "notes"), self.scan_tree_get(iid, "timestamp"))

    def _clean_value(self, value):
        import pandas as pd
        if value is None or (isinstance(value, float) and pd.isna(value)): return ""
        text = str(value).strip()
        return "" if text.lower() == "nan" else text

    def _compute_summary_metrics(self):
        total = len(self._all_iids)
        attended = sum(1 for iid in self._all_iids if self.tree.exists(iid) and self.scan_tree_get(iid, "attendance").lower() == "attend")
        metrics = {"total": total, "attended": attended, "attendance_rate": f"{(attended / total) * 100:.1f}%" if total else "0%"}
        if self.restrictions.get("exam"): metrics["missing_exam"] = sum(1 for iid in self._all_iids if self.tree.exists(iid) and not self.scan_tree_get(iid, "exam"))
        if self.restrictions.get("homework"): metrics["missing_hw"] = sum(1 for iid in self._all_iids if self.tree.exists(iid) and not self.scan_tree_get(iid, "homework") )
        return metrics

    def _build_summary_payload(self):
        summary = self._compute_summary_metrics()
        summary.update({"manual_additions": self._manual_additions, "cancellations": self._cancellations})
        return summary

    def _refresh_stats(self):
        metrics = self._compute_summary_metrics()
        self.stats_vars["total"].set(f"{metrics['total']}")
        self.stats_vars["attended"].set(f"{metrics['attended']}")
        self.stats_vars["percent"].set(metrics["attendance_rate"])
        if "missing_exam" in metrics: self.stats_vars["missing_exam"].set(f"{metrics['missing_exam']}")
        if "missing_hw" in metrics: self.stats_vars["missing_hw"].set(f"{metrics['missing_hw']}")

    def _safe_destroy(self, widget):
        """Safely destroys a widget if it exists."""
        if widget and hasattr(widget, "winfo_exists") and widget.winfo_exists():
            try:
                widget.destroy()
            except Exception:
                pass

    def _finalize_and_close(self, status_message=None):
        if status_message is None: status_message = f"Session '{self.sm.name}' saved and closed."
        summary, session_name, session_path, parent, read_only = self._build_summary_payload(), self.sm.name, getattr(self.sm, "session_path", None), self.parent, getattr(self, "read_only", False)
        
        # Safely destroy the main scan window
        self._safe_destroy(self)
        
        if hasattr(parent, "_refresh_recent_sessions"): parent._refresh_recent_sessions()
        if getattr(parent, "past_sessions_window", None) and parent.past_sessions_window.winfo_exists(): parent.past_sessions_window.refresh()
        if hasattr(parent, "set_status"): parent.set_status(status_message)
        if hasattr(parent, "show_session_summary"):
            parent.after(160, lambda: parent.show_session_summary(session_name=session_name, summary=summary, session_path=session_path, read_only=read_only))

    def _on_search_change(self, *_): self._filter_all()

    def _filter_all(self):
        query = self._clean_value(self.search_var.get()).lower() if self.search_var else ""
        terms = [term for term in query.split() if term]
        att = self._filter_vars["attendance"].get()
        missing_exam = self._filter_vars["missing_exam"].get()
        missing_hw = self._filter_vars["missing_hw"].get()
        has_notes = self._filter_vars["has_notes"].get()
        manual_added = self._filter_vars["manual_added"].get()

        for iid in self._all_iids:
            if not self.tree.exists(iid): continue
            show = True
            # Search filter
            if terms:
                values = [self._clean_value(self.tree.set(iid, col)).lower() for col in self.tree['columns']] + [str(iid).lower()]
                haystack = ' '.join(values)
                if not all(term in haystack for term in terms):
                    show = False
            # Attendance filter
            if att == "attend" and self.scan_tree_get(iid, "attendance").lower() != "attend":
                show = False
            if att == "absent" and self.scan_tree_get(iid, "attendance").lower() == "attend":
                show = False
            # Task filters
            if missing_exam and not self.scan_collect_missing_tasks(iid).count("exam"):
                show = False
            if missing_hw and not self.scan_collect_missing_tasks(iid).count("homework"):
                show = False
            # Has notes
            if has_notes and not self.scan_tree_get(iid, "notes"):
                show = False
            # Manually added (no card id is not digit)
            if manual_added and str(iid).isdigit():
                show = False
            if show:
                self.tree.reattach(iid, '', 'end')
            else:
                self.tree.detach(iid)

    def _set_attendance(self, code, attendance, notes, *, warn_on_duplicate=True, timestamp_override=None):
        if self.read_only or not self.tree.exists(code): return False
        target_attendance = self._clean_value(attendance)
        # The check for duplicate attendance is now handled in scan_on_open_row
        # if warn_on_duplicate and target_attendance.lower() == "attend" and self.scan_tree_get(code, "attendance").lower() == "attend":
        #     messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
        #     return False
        
        timestamp = timestamp_override or datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        rec = self._build_record_payload(code, target_attendance, self._clean_value(notes), timestamp)
        
        try: self.sm.add_record(rec)
        except Exception as exc: messagebox.showwarning("Attendance Update Failed", str(exc), parent=self); return False # type: ignore

        self._update_row(code, target_attendance, notes, timestamp)
        self._refresh_stats()
        return True

    def _build_record_payload(self, code, attendance, notes, timestamp):
        rec = {col: self.scan_tree_get(code, col) for col in ["student_id", "name", "phone", "exam", "homework"] if col in self.tree["columns"]}
        rec.update({"card_id": code, "attendance": attendance, "notes": notes, "timestamp": timestamp})
        return rec

    def _update_row(self, code, attendance, notes, timestamp=None):
        if not self.tree.exists(code): return
        try:
            self.tree.set(code, "attendance", self._clean_value(attendance))
            self.tree.set(code, "notes", self._clean_value(notes))
            if timestamp is not None: self.tree.set(code, "timestamp", self._clean_value(timestamp))
        except Exception: pass

    def _on_add_student_flow(self): self._launch_add_student_dialog()

    def _launch_add_student_dialog(self, card_id=None, default_notes="manual addition"):
        if self.read_only: return
        self._pause_focus_guard()
        normalized_card = None
        if card_id:
            raw_card = str(card_id).strip()
            normalized_card = raw_card.zfill(8) if raw_card.isdigit() else raw_card
        
        dialog = AddStudentDialog(self, card_id=normalized_card, duplicate_checker=self._student_id_or_phone_exists, default_notes=default_notes, on_submit=self._handle_add_student_submission)
        dialog.bind("<Destroy>", lambda e: self._resume_focus_guard(), add="+")

    def _handle_add_student_submission(self, *, card_id, values, default_notes):
        cid = str(card_id).strip() if card_id else self._next_unknown_card_id()
        if cid.isdigit(): cid = cid.zfill(8)
        
        timestamp = self.scan_now_tag()
        rec = {"card_id": cid, "attendance": "attend", "timestamp": timestamp, **values, "notes": default_notes}
        for task in ["exam", "homework"]:
            if self.restrictions.get(task): rec.setdefault(task, "")
        
        try: self.sm.add_record(rec)
        except Exception as exc: messagebox.showwarning("Unable to add student", str(exc), parent=self); return False
        
        self._manual_additions += 1
        row_values = [rec.get(col, "") for col in self.tree["columns"]]
        
        if self.tree.exists(cid): self.tree.item(cid, values=tuple(row_values))
        else: self.tree.insert("", "end", iid=cid, values=tuple(row_values)); self._all_iids.append(cid)
        
        self._refresh_stats()
        
        if self.scan_focus_ctx and self.scan_focus_ctx.get("status") == "not_found":
            self.scan_focus_clear()
        
        self.after(120, self.scan_entry.focus_set)
        return True

    def _next_unknown_card_id(self):
        if not hasattr(self, "_unknown_counter"):
            existing = [int(r.get("card_id", "").split("Unknown ")[-1]) for r in self.sm.records if str(r.get("card_id", "")).startswith("Unknown ")]
            self._unknown_counter = max(existing, default=0)
        self._unknown_counter += 1
        return f"Unknown {self._unknown_counter}"

    def _on_end_scan(self):
        msg = f"Session '{self.sm.name}' closed (view-only)." if self.read_only else None
        self._finalize_and_close(status_message=msg)

    def _global_focus_in(self, _event):
         if self._focus_reset_job is not None:
             self.after_cancel(self._focus_reset_job); self._focus_reset_job = None
         if self.read_only or self._focus_guard_depth > 0: return
         
         widget = self.focus_get()
         if widget is None or widget.winfo_toplevel() is not self: return
         
         # FIX 1: Explicitly ignore the Treeview widget itself
         if widget is self.tree: return
 
         # This check is still valid for the scan and search entries
         if widget in {self.scan_entry, *self._search_entries}: return
 
         # FIX 2: Check against the correct Focus View container
         parent = getattr(widget, "master", None)
         while parent is not None:
             # Check if the focused widget is a child of the integrated focus view
             if parent == getattr(self, "focus_view_container", None): return
             parent = getattr(parent, "master", None)
             
         self._focus_reset_job = self.after_idle(self._focus_scan_entry)

    def _student_id_or_phone_exists(self, student_id, phone):
        df = read_data(self.sm.session_path)
        sid_col, phone_col = self.mapping.get("student_id", "student_id"), self.mapping.get("phone", "phone")
        id_exists = student_id in df[sid_col].astype(str).values if sid_col in df.columns else False
        phone_exists = phone in df[phone_col].astype(str).values if phone_col in df.columns else False
        return id_exists, phone_exists