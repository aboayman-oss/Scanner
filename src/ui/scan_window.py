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

# --- Constants for the new Focus View Design ---
ASSETS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "assets")

# -- Colors --
# Light Mode
LIGHT_BG = "#f8faff"
LIGHT_SURFACE = "#fdfcff"
LIGHT_PRIMARY_TEXT = "#1b1c1e"
LIGHT_SECONDARY_TEXT = "#43474e"
LIGHT_SUCCESS = "#386a20"
LIGHT_WARNING = "#7e5700"
LIGHT_ERROR = "#b3261e"
LIGHT_INFO = "#00639c"
LIGHT_OUTLINE = "#73777f"

# Dark Mode
DARK_BG = "#1d1b20"
DARK_SURFACE = "#141218"
DARK_PRIMARY_TEXT = "#e3e2e6"
DARK_SECONDARY_TEXT = "#cac4d0"
DARK_SUCCESS = "#b5d3a7"
DARK_WARNING = "#f9d694"
DARK_ERROR = "#f2b8b5"
DARK_INFO = "#a9c8e7"
DARK_OUTLINE = "#8e9099"

# -- Status Definitions --
STATUS_STYLES = {
    "ok": {
        "text": "All Clear",
        "icon": "check_circle.png",
        "color": (LIGHT_SUCCESS, DARK_SUCCESS),
    },
    "missing_exam": {
        "text": "Tasks Missing",
        "icon": "warning.png",
        "color": (LIGHT_WARNING, DARK_WARNING),
    },
    "missing_homework": {
        "text": "Tasks Missing",
        "icon": "warning.png",
        "color": (LIGHT_WARNING, DARK_WARNING),
    },
    "not_found": {
        "text": "New Student",
        "icon": "person_add.png",
        "color": (LIGHT_INFO, DARK_INFO),
    },
    "duplicate": {
        "text": "Duplicate Card",
        "icon": "error.png",
        "color": (LIGHT_ERROR, DARK_ERROR),
    },
}


class ScanWindow(CTkToplevel):
    def __init__(self, parent, session_mgr, read_only=False):
        super().__init__(parent)
        self.parent = parent
        self.sm = session_mgr
        self.read_only = read_only
        self.state('zoomed')
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.toggle_fullscreen)
        self.restrictions = self.sm.restrictions
        self.df = read_data(self.sm.session_path).fillna("")
        self.mapping = self.sm.mapping or {col: col for col in self.df.columns}

        # --- Icon Cache ---
        self._icon_cache = {}

        # Load background image
        self.original_bg = Image.open(HOME_BG_FILE)
        self.bg_photo = ctk.CTkImage(light_image=self.original_bg, dark_image=self.original_bg, size=self.original_bg.size)
        self.bg_label = CTkLabel(self, text="", image=self.bg_photo)
        self.bg_label.place(relwidth=1, relheight=1)
        self.bind("<Configure>", self._on_bg_resize)

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
        self.scan_focus_window = None
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
            self.bind_all("<FocusIn>", self._global_focus_in, add="+ ")
            self.scan_entry.focus_set()

    def toggle_fullscreen(self, event=None):
        self.attributes("-fullscreen", not self.attributes("-fullscreen"))

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
        Creates the complete UI for the new Material 3-style Focus View.
        The UI is divided into three zones: Status, Details, and Actions.
        """
        parent.configure(fg_color=(LIGHT_BG, DARK_BG))

        # --- 1. Status Zone (Top) ---
        # Displays student name, IDs, and a prominent status badge.
        status_zone = CTkFrame(parent, fg_color="transparent")
        status_zone.pack(fill="x", padx=20, pady=(20, 12))
        status_zone.grid_columnconfigure(0, weight=1)

        self.scan_focus_name_label = CTkLabel(status_zone, text="", font=("Roboto", 32, "bold"), text_color=(LIGHT_PRIMARY_TEXT, DARK_PRIMARY_TEXT), anchor="w")
        self.scan_focus_name_label.grid(row=0, column=0, sticky="w")

        self.scan_focus_id_label = CTkLabel(status_zone, text="", font=("Roboto", 12), text_color=(LIGHT_SECONDARY_TEXT, DARK_SECONDARY_TEXT), anchor="w")
        self.scan_focus_id_label.grid(row=1, column=0, sticky="w", pady=(0, 8))

        # Status Icon
        self.scan_focus_status_icon = CTkLabel(status_zone, text="")
        self.scan_focus_status_icon.grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))

        # --- 2. Details Zone (Middle) ---
        # Contains two cards for Homework and Exam status, and a notes textbox.
        details_zone = CTkFrame(parent, fg_color="transparent")
        details_zone.pack(fill="both", expand=True, padx=20, pady=8)

        # Homework Card
        self.hw_card = CTkFrame(details_zone, corner_radius=12, fg_color=(LIGHT_SURFACE, DARK_SURFACE))
        self.hw_card.pack(fill="x", pady=(0, 8))
        self.hw_card.grid_columnconfigure(1, weight=1)
        self.hw_icon_label = CTkLabel(self.hw_card, text="")
        self.hw_icon_label.grid(row=0, column=0, padx=(12, 8), pady=12)
        CTkLabel(self.hw_card, text="Homework", font=("Roboto", 16, "bold")).grid(row=0, column=1, sticky="w")
        self.hw_grade_label = CTkLabel(self.hw_card, text="", font=("Roboto", 14))
        self.hw_grade_label.grid(row=0, column=2, sticky="e", padx=12)

        # Exam Card
        self.exam_card = CTkFrame(details_zone, corner_radius=12, fg_color=(LIGHT_SURFACE, DARK_SURFACE))
        self.exam_card.pack(fill="x")
        self.exam_card.grid_columnconfigure(1, weight=1)
        self.exam_icon_label = CTkLabel(self.exam_card, text="")
        self.exam_icon_label.grid(row=0, column=0, padx=(12, 8), pady=12)
        CTkLabel(self.exam_card, text="Exam", font=("Roboto", 16, "bold")).grid(row=0, column=1, sticky="w")
        self.exam_grade_label = CTkLabel(self.exam_card, text="", font=("Roboto", 14))
        self.exam_grade_label.grid(row=0, column=2, sticky="e", padx=12)

        # Notes Textbox
        self.scan_focus_notes = CTkTextbox(details_zone, corner_radius=12, fg_color=(LIGHT_SURFACE, DARK_SURFACE), border_width=0, font=("Roboto", 14))
        self.scan_focus_notes.pack(fill="both", expand=True, pady=(12, 0))
        self.scan_focus_notes.insert("1.0", "Add notes here...")
        self.scan_focus_notes.bind("<FocusIn>", self._on_notes_focus_in)
        self.scan_focus_notes.bind("<FocusOut>", self._on_notes_focus_out)

        # --- 3. Actions Zone (Bottom) ---
        self.actions_zone = CTkFrame(parent, fg_color="transparent")
        self.actions_zone.pack(fill="x", padx=20, pady=(12, 20))

        # -- Button Definitions --
        # Primary (Filled)
        self._btn_complete = CTkButton(self.actions_zone, text="Complete & Attend", image=self._load_icon("task_alt.png"), command=self.scan_focus_on_completed)
        self._btn_add_student = CTkButton(self.actions_zone, text="Add New Student", image=self._load_icon("person_add.png"), command=self.scan_focus_on_add_student)
        
        # Secondary (Outlined)
        self._btn_override = CTkButton(self.actions_zone, text="Override & Attend", image=self._load_icon("gpp_good.png"), command=self.scan_focus_on_override, fg_color="transparent", border_width=1, border_color=(LIGHT_OUTLINE, DARK_OUTLINE))
        
        # Negative (Text Button)
        self._btn_deny = CTkButton(self.actions_zone, text="Deny Entry", command=self.scan_focus_on_deny, fg_color="transparent", text_color=(LIGHT_ERROR, DARK_ERROR), hover=False)

        self._btn_cancel = CTkButton(self.actions_zone, text="Cancel Attendance", command=self.scan_focus_on_cancel_attendance, fg_color="transparent", border_width=1, border_color=(LIGHT_OUTLINE, DARK_OUTLINE))

        self.scan_focus_buttons = [self._btn_complete, self._btn_add_student, self._btn_override, self._btn_deny, self._btn_cancel]
        
        if self.read_only:
            self.scan_focus_notes.configure(state="disabled")
            for btn in self.scan_focus_buttons: btn.configure(state="disabled")

    def _on_notes_focus_in(self, event):
        self._pause_focus_guard()
        if self.scan_focus_notes.get("1.0", "end-1c") == "Add notes here...":
            self.scan_focus_notes.delete("1.0", "end")
            self.scan_focus_notes.configure(text_color=(LIGHT_PRIMARY_TEXT, DARK_PRIMARY_TEXT))

    def _on_notes_focus_out(self, event):
        self._resume_focus_guard()
        if not self.scan_focus_notes.get("1.0", "end-1c"):
            self.scan_focus_notes.configure(text_color="gray")
            self.scan_focus_notes.insert("1.0", "Add notes here...")

    def _ensure_scan_focus_window(self):
        """Ensures the Focus View window exists, creating it if necessary."""
        if getattr(self, "scan_focus_window", None) and self.scan_focus_window.winfo_exists():
            return self.scan_focus_window
        
        window = CTkToplevel(self)
        window.withdraw()
        window.title("Focus View")
        window.geometry("550x640")
        window.minsize(550, 600)
        window.transient(self)
        window.attributes("-topmost", True)
        window.protocol("WM_DELETE_WINDOW", self.scan_focus_clear)
        window.bind("<Destroy>", lambda e: setattr(self, "scan_focus_window", None), add="+ ")

        self.scan_focus_window = window
        self.scan_focus_create_ui(window)
        return window

    def scan_focus_show(self, scan_ctx):
        """Shows and populates the Focus View with student data."""
        self.scan_focus_cancel_timer()
        window = self._ensure_scan_focus_window()
        if not window: return

        window.deiconify()
        bring_window_to_front(window)
        
        ctx = dict(scan_ctx or {})
        ctx.setdefault("original_notes", ctx.get("existing_notes", ""))
        self.scan_focus_ctx = ctx
        
        status = ctx.get("status") or self.scan_determine_status(ctx)
        ctx["status"] = status

        # Populate UI elements
        self.scan_focus_name_label.configure(text=ctx.get("name") or "Unknown Student")
        
        card_display_val = ctx.get('card_display', '') or ''
        card_display = str(card_display_val).replace('null', '').strip() or '--'
        student_id_val = ctx.get('student_id', '') or ''
        student_id = str(student_id_val).replace('null', '').strip() or '--'
        
        id_text = f"Student ID: {student_id}  •  Card ID: {card_display}"
        self.scan_focus_id_label.configure(text=id_text)

        # Set notes
        if not self.read_only: self.scan_focus_notes.configure(state="normal")
        self.scan_focus_notes.delete("1.0", "end")
        existing_notes = ctx.get("existing_notes", "")
        if existing_notes:
            self.scan_focus_notes.insert("1.0", existing_notes)
            self.scan_focus_notes.configure(text_color=(LIGHT_PRIMARY_TEXT, DARK_PRIMARY_TEXT))
        else:
            self.scan_focus_notes.configure(text_color="gray")
            self.scan_focus_notes.insert("1.0", "Add notes here...")
        if self.read_only: self.scan_focus_notes.configure(state="disabled")

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
        self.scan_focus_status_icon.configure(image=self._load_icon(style["icon"], size=(48, 48)))

        # 2. Update Details Cards (Homework & Exam)
        missing_tasks = context.get("missing_tasks", [])
        success_icon = self._load_icon("task_alt.png")
        problem_icon = self._load_icon("error.png")

        # Subtle container colors
        success_color = ("#e8f5e9", "#1b331d") # Material Green Light/Dark
        problem_color = ("#fce8e6", "#3c1b1a") # Material Red Light/Dark

        # Homework
        hw_missing = "homework" in missing_tasks
        self.hw_icon_label.configure(image=problem_icon if hw_missing else success_icon)
        hw_grade = context.get("homework", "")
        hw_text = ""
        if hw_grade:
            hw_text = str(hw_grade)
            if hw_grade == "0":
                hw_text += " (Fail)"
        else:
            hw_text = "Not Submitted"
        self.hw_grade_label.configure(text=hw_text)
        self.hw_card.configure(fg_color=problem_color if hw_missing else success_color)

        # Exam
        exam_missing = "exam" in missing_tasks
        self.exam_icon_label.configure(image=problem_icon if exam_missing else success_icon)
        exam_grade = context.get("exam", "")
        exam_text = ""
        if exam_grade:
            exam_text = str(exam_grade)
            if exam_grade == "0":
                exam_text += " (Fail)"
        else:
            exam_text = "Not Submitted"
        self.exam_grade_label.configure(text=exam_text)
        self.exam_card.configure(fg_color=problem_color if exam_missing else success_color)

        # 3. Update Action Buttons
        self._update_action_buttons(kind, context)

    def _update_action_buttons(self, kind, context):
        """Shows and hides the correct action buttons based on the status."""
        for btn in self.scan_focus_buttons: btn.pack_forget()

        buttons_to_show = []
        if kind == "not_found":
            buttons_to_show = [self._btn_add_student]
        elif kind in {"missing_exam", "missing_homework"}:
            buttons_to_show = [self._btn_deny, self._btn_override, self._btn_complete]
        elif kind == "ok" and context.get("already_attended"):
            buttons_to_show = [self._btn_cancel]
        
        # Pack buttons with primary actions last to appear on the right
        for btn in buttons_to_show: btn.pack(side="left", fill="x", expand=True, padx=4)
        
        # Special case for a single primary button to be centered
        if len(buttons_to_show) == 1:
            buttons_to_show[0].pack(side="top", fill="x", expand=True, padx=4)

    def scan_focus_clear(self):
        """Hides the Focus View and resets its state."""
        self.scan_focus_cancel_timer()
        self.scan_focus_ctx = None
        
        if hasattr(self, "scan_focus_notes"):
            if not self.read_only: self.scan_focus_notes.configure(state="normal")
            self.scan_focus_notes.delete("1.0", "end")
            self.scan_focus_notes.configure(text_color="gray")
            self.scan_focus_notes.insert("1.0", "Add notes here...")

        self.scan_restore_from_focus()
        
        window = getattr(self, "scan_focus_window", None)
        if window and window.winfo_exists(): window.withdraw()
            
        self.after(120, self.scan_entry.focus_set)

    # --------------------------------------------------------------------------
    # Original ScanWindow methods (unchanged unless necessary for integration)
    # --------------------------------------------------------------------------

    def _on_bg_resize(self, event):
        if event.widget is self:
            bg = self.original_bg.resize((event.width, event.height), Image.Resampling.LANCZOS)
            self.bg_photo = ctk.CTkImage(light_image=bg, dark_image=bg, size=(event.width, event.height))
            self.bg_label.configure(image=self.bg_photo)

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
        top_bar = CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=12, pady=(16, 8))
        top_bar.grid_columnconfigure(1, weight=1)

        scan_block = CTkFrame(top_bar, fg_color="transparent")
        scan_block.grid(row=0, column=0, sticky="w")
        CTkLabel(scan_block, text="Scan Entry", font=("Arial", 12, "bold")).pack(anchor="w")
        self.scan_entry = CTkEntry(scan_block, width=240, placeholder_text="Scan card ID")
        self.scan_entry.pack(anchor="w", pady=(4, 0))
        self.scan_entry.bind("<Return>", lambda _e: self.scan_on_scan())

        self.pb = CTkProgressBar(scan_block, mode="indeterminate", width=240)
        self.pb.pack(anchor="w", pady=(6, 0)); self.pb.stop()

        if self.read_only:
            self.scan_entry.configure(state="disabled"); self.scan_entry.unbind("<Return>"); scan_block.grid_remove()

        session_text = f"Session: {self.sm.name}" + (" (read-only)" if self.read_only else "")
        CTkLabel(top_bar, text=session_text, font=("Arial", 12)).grid(row=0, column=1, sticky="w", padx=18)

        self.end_button = CTkButton(top_bar, text="End Session" if not self.read_only else "Close", command=self._on_end_scan)
        self.end_button.grid(row=0, column=2, sticky="e")

        self._build_stats_strip()

        search_frame = CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=12, pady=(0, 10))
        self.search_var = ctk.StringVar()
        search_entry = CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="Search by card, ID, name, phone")
        search_entry.grid(row=0, column=0, sticky="ew", padx=5); search_frame.grid_columnconfigure(0, weight=1)
        self.search_var.trace_add("write", self._on_search_change)
        self._search_entries.append(search_entry)
        search_entry.bind("<FocusIn>", lambda _e: self._pause_focus_guard())
        search_entry.bind("<FocusOut>", lambda _e: self._resume_focus_guard())
        self.smart_search_entry = search_entry

        scan_main_content = CTkFrame(self, fg_color="transparent")
        scan_main_content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        scan_main_content.grid_rowconfigure(0, weight=1); scan_main_content.grid_columnconfigure(0, weight=1)

        tree_container = CTkFrame(scan_main_content, fg_color="transparent")
        tree_container.grid(row=0, column=0, sticky="nsew"); tree_container.grid_rowconfigure(0, weight=1); tree_container.grid_columnconfigure(0, weight=1)

        cols = ["card_id", "student_id", "name", "phone"]
        if self.restrictions.get("exam"): cols.append("exam")
        if self.restrictions.get("homework"): cols.append("homework")
        cols += ["attendance", "notes", "timestamp"]

        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings", selectmode="browse")
        for col in cols:
            self.tree.heading(col, text=col.replace("_", " ").title()); self.tree.column(col, anchor="center", width=110)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns"); self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self.scan_on_row_double_click)
        if self.read_only: self.tree.unbind("<Double-1>")

        self.control_frame = CTkFrame(self, fg_color="transparent")
        self.control_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.add_student_button = CTkButton(self.control_frame, text="Add Student", command=self._on_add_student_flow)
        self.add_student_button.pack(side="left")
        if self.read_only: self.control_frame.pack_forget()

    def scan_focus_cancel_timer(self):
        if self.scan_focus_timer is not None:
            try: self.after_cancel(self.scan_focus_timer)
            except Exception: pass
            self.scan_focus_timer = None

    def scan_focus_schedule_clear(self, delay=3000):
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
        mapped = [labels.get(task, task.title()) for task in tasks]
        return mapped[0] if len(mapped) == 1 else " & ".join(mapped)

    def scan_append_notes(self, original, addition):
        original_clean, addition_clean = self._clean_value(original), self._clean_value(addition)
        if not addition_clean: return original_clean
        if not original_clean: return addition_clean
        return f"{original_clean.rstrip()}\n{addition_clean}"

    def scan_collect_new_note(self):
        if not hasattr(self, "scan_focus_notes"): return ""
        try: typed = self.scan_focus_notes.get("1.0", "end-1c").strip()
        except Exception: return ""
        if typed == "Add notes here...": return ""
        return self._clean_value(typed)

    def scan_now_tag(self):
        return f"[{datetime.now():%H:%M:%S}]"

    def scan_determine_status(self, scan_ctx):
        if scan_ctx.get("status") in {"not_found", "duplicate"}: return scan_ctx["status"]
        if not scan_ctx.get("found", True): return "not_found"
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
        except Exception as exc: messagebox.showwarning("Attendance Update Failed", str(exc), parent=self); return False

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
        action_note = f"{tag} Attended with override (missing {desc})."
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
        self.stats_frame = CTkFrame(self, fg_color=("#f1f5f9", "#12263a"), corner_radius=10)
        self.stats_frame.pack(fill="x", padx=12, pady=(0, 10))
        stats = [("Total Rows", self.stats_vars["total"]),( "Attended", self.stats_vars["attended"]),( "Attendance %", self.stats_vars["percent"])]
        if self.restrictions.get("exam"): stats.append(("Missing Exam", self.stats_vars["missing_exam"]))
        if self.restrictions.get("homework"): stats.append(("Missing H.W.", self.stats_vars["missing_hw"]))
        for idx, (label, var) in enumerate(stats):
            block = CTkFrame(self.stats_frame, fg_color="transparent")
            block.grid(row=0, column=idx, sticky="w", padx=(12 if idx == 0 else 8, 8), pady=10)
            CTkLabel(block, text=label, font=("Arial", 11, "bold")).pack(anchor="w")
            CTkLabel(block, textvariable=var, font=("Arial", 16)).pack(anchor="w")
            self.stats_frame.grid_columnconfigure(idx, weight=1)

    def _apply_treeview_style(self):
        mode = ctk.get_appearance_mode()
        style = ttk.Style(self)
        style.theme_use("default")
        bg, fg, heading_bg, heading_fg = ("#ffffff", "#1a1a1a", "#e1efff", "#1a1a1a") if mode == "Light" else ("#1e1e1e", "#f2f2f2", "#1f6aa5", "#ffffff")
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
            values = [self._clean_value(rec.get(col) if rec and col in rec else row.get(self.mapping.get(col, col), "")) for col in cols]
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

    def _finalize_and_close(self, status_message=None):
        if status_message is None: status_message = f"Session '{self.sm.name}' saved and closed."
        summary, session_name, session_path, parent, read_only = self._build_summary_payload(), self.sm.name, getattr(self.sm, "session_path", None), self.parent, getattr(self, "read_only", False)
        
        if getattr(self, "scan_focus_window", None): self.scan_focus_window.destroy()
        if self.winfo_exists(): self.destroy()
        
        if hasattr(parent, "_refresh_recent_sessions"): parent._refresh_recent_sessions()
        if getattr(parent, "past_sessions_window", None) and parent.past_sessions_window.winfo_exists(): parent.past_sessions_window.refresh()
        if hasattr(parent, "set_status"): parent.set_status(status_message)
        if hasattr(parent, "show_session_summary"):
            parent.after(160, lambda: parent.show_session_summary(session_name=session_name, summary=summary, session_path=session_path, read_only=read_only))

    def _on_search_change(self, *_): self._filter_all()

    def _filter_all(self):
        query = self._clean_value(self.search_var.get()).lower() if self.search_var else ""
        terms = [term for term in query.split() if term]
        for iid in self._all_iids:
            if not self.tree.exists(iid): continue
            if not terms:
                self.tree.reattach(iid, '', 'end')
                continue
            values = [self._clean_value(self.tree.set(iid, col)).lower() for col in self.tree['columns']] + [str(iid).lower()]
            haystack = ' '.join(values)
            if all(term in haystack for term in terms): self.tree.reattach(iid, '', 'end')
            else: self.tree.detach(iid)

    def _set_attendance(self, code, attendance, notes, *, warn_on_duplicate=True, timestamp_override=None):
        if self.read_only or not self.tree.exists(code): return False
        target_attendance = self._clean_value(attendance)
        if warn_on_duplicate and target_attendance.lower() == "attend" and self.scan_tree_get(code, "attendance").lower() == "attend":
            messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
            return False
        
        timestamp = timestamp_override or datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        rec = self._build_record_payload(code, target_attendance, self._clean_value(notes), timestamp)
        
        try: self.sm.add_record(rec)
        except Exception as exc: messagebox.showwarning("Attendance Update Failed", str(exc), parent=self); return False
        
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
        dialog.bind("<Destroy>", lambda e: self._resume_focus_guard(), add="+ ")

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
        if widget in {self.scan_entry, *self._search_entries}: return
        
        parent = getattr(widget, "master", None)
        while parent is not None:
            if parent == getattr(self, "scan_focus_window", None): return
            parent = getattr(parent, "master", None)
            
        self._focus_reset_job = self.after_idle(self._focus_scan_entry)

    def _student_id_or_phone_exists(self, student_id, phone):
        df = read_data(self.sm.session_path)
        sid_col, phone_col = self.mapping.get("student_id", "student_id"), self.mapping.get("phone", "phone")
        id_exists = student_id in df[sid_col].astype(str).values if sid_col in df.columns else False
        phone_exists = phone in df[phone_col].astype(str).values if phone_col in df.columns else False
        return id_exists, phone_exists