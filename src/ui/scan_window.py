"""All changes are isolated to Scan Attendance. Other windows unchanged. The Scan window now presents a dedicated Focus View after each scan, instantly filtering to the active student while leaving the rest of the UI untouched.

Green-path scans auto-attend, restriction flags call for guided decisions, and unknown cards route into the existing add-student dialog while persisting through the session manager.

Key helpers: scan_focus_create_ui (Focus View creation), scan_determine_status (status computation), scan_on_scan & scan_on_open_row (action handlers), scan_filter_for_focus & scan_restore_from_focus (filter/restore helpers).
No code in Main, Settings, Past Sessions, core, or utils was modified.
"""
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk
from customtkinter import CTkButton, CTkEntry, CTkFrame, CTkLabel, CTkProgressBar, CTkTextbox, CTkToplevel
from PIL import Image, ImageTk
import pandas as pd

from ui.dialogs.add_student_dialog import AddStudentDialog
from ui.dialogs.session_summary_dialog import SessionSummaryDialog
from utils.helpers import (
    HOME_BG_FILE,
    MIN_SCAN_SIZE,
    SETTINGS,
    bring_window_to_front,
    ensure_initial_size,
    read_data,
    write_data,
)

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
        self._search_entries = []
        self.search_var = None
        self._manual_additions = 0
        self._cancellations = 0
        self._focus_reset_job = None
        self._focus_guard_depth = 0
        self.scan_focus_ctx = None
        self.scan_focus_visible_cache = []
        self.scan_focus_timer = None
        self.scan_focus_name_var = ctk.StringVar(value="")
        self.scan_focus_id_var = ctk.StringVar(value="")
        self.scan_focus_status_var = ctk.StringVar(value="")
        self.scan_focus_message_var = ctk.StringVar(value="")
        self.scan_focus_card_var = ctk.StringVar(value="")
        self.scan_focus_window = None
        self.scan_focus_container = None
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
        if self.scan_focus_ctx and self.scan_focus_ctx.get("status") == "not_found":
            try:
                self.tree.selection_set(cid)
                self.tree.focus(cid)
            except Exception:
                pass
            self.scan_focus_set_message("Student added and attended.", level="success")
            self.scan_focus_schedule_clear()
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
        self.scan_entry.bind("<Return>", lambda _e: self.scan_on_scan())

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

        scan_main_content = CTkFrame(self, fg_color="transparent")
        scan_main_content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        scan_main_content.grid_rowconfigure(0, weight=1)
        scan_main_content.grid_columnconfigure(0, weight=1)

        tree_container = CTkFrame(scan_main_content, fg_color="transparent")
        tree_container.grid(row=0, column=0, sticky="nsew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        cols = ["card_id", "student_id", "name", "phone"]
        if self.restrictions.get("exam"):
            cols.append("exam")
        if self.restrictions.get("homework"):
            cols.append("homework")
        cols += ["attendance", "notes", "timestamp"]

        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings", selectmode="browse")
        for col in cols:
            self.tree.heading(col, text=col.replace("_", " ").title())
            self.tree.column(col, anchor="center", width=110)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self.scan_on_row_double_click)
        if self.read_only:
            self.tree.unbind("<Double-1>")

        self.scan_focus_window = None
        self.scan_focus_container = None

        self.control_frame = CTkFrame(self, fg_color="transparent")
        self.control_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.add_student_button = CTkButton(self.control_frame, text="Add Student", command=self._on_add_student_flow)
        self.add_student_button.pack(side="left")
        if self.read_only:
            self.control_frame.pack_forget()

    def scan_focus_create_ui(self, parent):
        self.scan_focus_container = parent
        try:
            parent.grid_propagate(False)
        except Exception:
            pass
        try:
            parent.pack_propagate(False)
        except Exception:
            pass
        parent.configure(width=320)
        focus_colors = ("#e2e8f0", "#1f2933")
        self.scan_focus_frame = CTkFrame(parent, fg_color=focus_colors, corner_radius=12)
        self.scan_focus_frame.pack(fill="both", expand=True)

        header = CTkFrame(self.scan_focus_frame, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1)

        self.scan_focus_name_label = CTkLabel(header, textvariable=self.scan_focus_name_var, font=("Arial", 18, "bold"))
        self.scan_focus_name_label.grid(row=0, column=0, sticky="w")
        self.scan_focus_id_label = CTkLabel(header, textvariable=self.scan_focus_id_var, font=("Arial", 13))
        self.scan_focus_id_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.scan_focus_status_badge = CTkLabel(
            header,
            textvariable=self.scan_focus_status_var,
            font=("Arial", 13, "bold"),
            corner_radius=14,
            padx=12,
            pady=4,
        )
        self.scan_focus_status_badge.grid(row=0, column=1, rowspan=2, sticky="e")

        self.scan_focus_card_label = CTkLabel(header, textvariable=self.scan_focus_card_var, font=("Arial", 12))
        self.scan_focus_card_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        notes_frame = CTkFrame(self.scan_focus_frame, fg_color="transparent")
        notes_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        CTkLabel(notes_frame, text="Notes", font=("Arial", 12, "bold")).pack(anchor="w")
        self.scan_focus_notes = CTkTextbox(notes_frame, height=120)
        self.scan_focus_notes.pack(fill="both", expand=True, pady=(4, 0))
        self.scan_focus_notes.bind("<FocusIn>", lambda _e: self._pause_focus_guard())
        self.scan_focus_notes.bind("<FocusOut>", lambda _e: self._resume_focus_guard())

        self.scan_focus_button_row = CTkFrame(self.scan_focus_frame, fg_color="transparent")
        self.scan_focus_button_row.pack(fill="x", padx=14, pady=(4, 6))

        self.scan_focus_completed_btn = CTkButton(
            self.scan_focus_button_row,
            text="Completed Task & Attend",
            command=self.scan_focus_on_completed,
        )
        self.scan_focus_override_btn = CTkButton(
            self.scan_focus_button_row,
            text="Override & Attend",
            command=self.scan_focus_on_override,
        )
        self.scan_focus_deny_btn = CTkButton(
            self.scan_focus_button_row,
            text="Deny Entry",
            command=self.scan_focus_on_deny,
            fg_color=("#f87171", "#7f1d1d"),
            hover_color=("#ef4444", "#991b1b"),
        )
        self.scan_focus_add_btn = CTkButton(
            self.scan_focus_button_row,
            text="Add New Student",
            command=self.scan_focus_on_add_student,
        )
        self.scan_focus_cancel_btn = CTkButton(
            self.scan_focus_button_row,
            text="Cancel Attendance",
            command=self.scan_focus_on_cancel_attendance,
        )

        for btn in (
            self.scan_focus_completed_btn,
            self.scan_focus_override_btn,
            self.scan_focus_deny_btn,
            self.scan_focus_add_btn,
            self.scan_focus_cancel_btn,
        ):
            btn.pack(side="left", padx=6)
            btn.pack_forget()

        self.scan_focus_buttons = [
            self.scan_focus_completed_btn,
            self.scan_focus_override_btn,
            self.scan_focus_deny_btn,
            self.scan_focus_add_btn,
            self.scan_focus_cancel_btn,
        ]

        self.scan_focus_message_label = CTkLabel(
            self.scan_focus_frame,
            textvariable=self.scan_focus_message_var,
            font=("Arial", 12),
        )
        self.scan_focus_message_label.pack(fill="x", padx=14, pady=(0, 12))
        self.scan_focus_message_default_color = self.scan_focus_message_label.cget("text_color")

        if self.read_only:
            self.scan_focus_notes.configure(state="disabled")
            for btn in self.scan_focus_buttons:
                btn.configure(state="disabled")

    def _ensure_scan_focus_window(self):
        window = getattr(self, "scan_focus_window", None)
        if window is not None and window.winfo_exists():
            return window
        window = CTkToplevel(self)
        window.withdraw()
        window.title("Focus View")
        window.geometry("360x520")
        window.minsize(320, 420)
        window.resizable(False, True)
        window.transient(self)
        window.protocol("WM_DELETE_WINDOW", self.scan_focus_clear)
        window.bind("<Destroy>", self._handle_focus_window_destroy, add="+")
        container = CTkFrame(window, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=12, pady=12)
        self.scan_focus_window = window
        self.scan_focus_create_ui(container)
        return window

    def _position_focus_window(self, window):
        try:
            self.update_idletasks()
            window.update_idletasks()
            width = max(window.winfo_width(), window.winfo_reqwidth(), 360)
            height = max(window.winfo_height(), window.winfo_reqheight(), 520)
            x = self.winfo_rootx() + self.winfo_width() + 16
            y = self.winfo_rooty() + 48
            window.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")
        except Exception:
            pass

    def _handle_focus_window_destroy(self, _event=None):
        self.scan_focus_window = None
        self.scan_focus_container = None

    def _destroy_scan_focus_window(self):
        window = getattr(self, "scan_focus_window", None)
        if window is None:
            return
        try:
            window.destroy()
        except Exception:
            pass
        self.scan_focus_window = None
        self.scan_focus_container = None


    def scan_focus_hide_buttons(self):
        for scan_btn in getattr(self, "scan_focus_buttons", []):
            try:
                scan_btn.pack_forget()
            except Exception:
                pass

    def scan_focus_show_buttons(self, scan_buttons):
        self.scan_focus_hide_buttons()
        for scan_btn in scan_buttons:
            scan_btn.pack(side="left", padx=6)

    def scan_focus_set_message(self, message="", level="info"):
        palette = {
            "info": self.scan_focus_message_default_color,
            "success": ("#166534", "#bbf7d0"),
            "warning": ("#92400e", "#facc15"),
            "error": ("#b91c1c", "#f87171"),
        }
        color = palette.get(level, self.scan_focus_message_default_color)
        if hasattr(self, "scan_focus_message_label"):
            try:
                self.scan_focus_message_label.configure(text_color=color)
            except Exception:
                pass
        self.scan_focus_message_var.set(message or "")

    def scan_focus_cancel_timer(self):
        if self.scan_focus_timer is not None:
            try:
                self.after_cancel(self.scan_focus_timer)
            except Exception:
                pass
            self.scan_focus_timer = None

    def scan_focus_schedule_clear(self, delay=950):
        self.scan_focus_cancel_timer()
        self.scan_focus_timer = self.after(delay, self.scan_focus_clear)

    def scan_restore_from_focus(self):
        if not self.scan_focus_visible_cache:
            return
        for scan_iid in self.scan_focus_visible_cache:
            if not self.tree.exists(scan_iid):
                continue
            try:
                self.tree.reattach(scan_iid, "", "end")
            except Exception:
                pass
        self.scan_focus_visible_cache = []

    def scan_filter_for_focus(self, target_iids):
        self.scan_restore_from_focus()
        if not target_iids:
            return
        current_visible = []
        for scan_iid in self._all_iids:
            if not self.tree.exists(scan_iid):
                continue
            if self.tree.parent(scan_iid):
                continue
            current_visible.append(scan_iid)
        self.scan_focus_visible_cache = current_visible
        for scan_iid in current_visible:
            if scan_iid in target_iids:
                continue
            try:
                self.tree.detach(scan_iid)
            except Exception:
                pass
        for scan_iid in target_iids:
            if not self.tree.exists(scan_iid):
                continue
            try:
                self.tree.reattach(scan_iid, "", "end")
            except Exception:
                pass
        primary = target_iids[0]
        if self.tree.exists(primary):
            self.tree.selection_set(primary)
            self.tree.focus(primary)

    def scan_focus_clear(self):
        self.scan_focus_cancel_timer()
        self.scan_focus_hide_buttons()
        self.scan_focus_ctx = None
        if hasattr(self, "scan_focus_message_label"):
            try:
                self.scan_focus_message_label.configure(text_color=self.scan_focus_message_default_color)
            except Exception:
                pass
        self.scan_focus_message_var.set("")
        self.scan_focus_status_var.set("")
        self.scan_focus_name_var.set("")
        self.scan_focus_id_var.set("")
        self.scan_focus_card_var.set("")
        if not self.read_only and hasattr(self, "scan_focus_notes"):
            try:
                self.scan_focus_notes.configure(state="normal")
            except Exception:
                pass
        if hasattr(self, "scan_focus_notes"):
            try:
                self.scan_focus_notes.delete("1.0", "end")
            except Exception:
                pass
        self.scan_restore_from_focus()
        window = getattr(self, "scan_focus_window", None)
        if window is not None and window.winfo_exists():
            try:
                window.withdraw()
            except Exception:
                pass
        self.after(120, self.scan_entry.focus_set)

    def scan_focus_show(self, scan_ctx):
        self.scan_focus_cancel_timer()
        window = self._ensure_scan_focus_window()
        if window is None:
            return
        try:
            window.deiconify()
            bring_window_to_front(window)
            self._position_focus_window(window)
        except Exception:
            pass
        ctx = dict(scan_ctx or {})
        ctx.setdefault("original_notes", ctx.get("existing_notes", ""))
        self.scan_focus_ctx = ctx
        status = ctx.get("status")
        if status is None:
            status = self.scan_determine_status(ctx)
            ctx["status"] = status
        card_display = ctx.get("card_display") or ctx.get("card_id") or ""
        name_display = ctx.get("name") or "Unknown Student"
        student_id_display = ctx.get("student_id") or "--"
        self.scan_focus_name_var.set(name_display)
        self.scan_focus_id_var.set(f"Student ID: {student_id_display}")
        self.scan_focus_card_var.set(f"Card ID: {card_display}" if card_display else "")
        if hasattr(self, "scan_focus_notes"):
            if not self.read_only:
                try:
                    self.scan_focus_notes.configure(state="normal")
                except Exception:
                    pass
            try:
                self.scan_focus_notes.delete("1.0", "end")
                existing = ctx.get("existing_notes", "")
                if existing:
                    self.scan_focus_notes.insert("end", existing)
            except Exception:
                pass
            if self.read_only:
                try:
                    self.scan_focus_notes.configure(state="disabled")
                except Exception:
                    pass
        message_level = ctx.get("message_level", "info")
        message_text = ctx.get("message", "")
        self.scan_focus_set_message(message_text, level=message_level)
        focus_iids = ctx.get("focus_iids") or []
        if focus_iids and not ctx.get("skip_filter") and ctx.get("iid") is not None:
            self.scan_filter_for_focus(focus_iids)
        else:
            if focus_iids:
                primary = focus_iids[0]
                if self.tree.exists(primary):
                    self.tree.selection_set(primary)
                    self.tree.focus(primary)
            self.scan_restore_from_focus()
        self.scan_focus_set_status(status)

    def scan_focus_set_status(self, kind):
        if self.scan_focus_ctx is not None:
            self.scan_focus_ctx["status"] = kind
        status_styles = {
            "ok": {"text": "? OK", "fg_color": ("#bbf7d0", "#14532d"), "text_color": ("#14532d", "#bbf7d0")},
            "missing_exam": {"text": "?? Missing Exam", "fg_color": ("#fee2e2", "#7c2d12"), "text_color": ("#7c2d12", "#fee2e2")},
            "missing_homework": {"text": "?? Missing Homework", "fg_color": ("#fee2e2", "#7c2d12"), "text_color": ("#7c2d12", "#fee2e2")},
            "not_found": {"text": "? Not Found", "fg_color": ("#dbeafe", "#1e3a8a"), "text_color": ("#1e3a8a", "#dbeafe")},
            "duplicate": {"text": "?? Duplicate Card", "fg_color": ("#fef3c7", "#92400e"), "text_color": ("#92400e", "#fef3c7")},
        }
        style = status_styles.get(kind, status_styles["ok"])
        self.scan_focus_status_var.set(style["text"])
        try:
            self.scan_focus_status_badge.configure(fg_color=style["fg_color"], text_color=style["text_color"])
        except Exception:
            pass
        buttons = []
        context = self.scan_focus_ctx or {}
        allow_cancel = context.get("allow_cancel", False)
        if kind in {"missing_exam", "missing_homework"}:
            buttons = [
                self.scan_focus_completed_btn,
                self.scan_focus_override_btn,
                self.scan_focus_deny_btn,
            ]
            if allow_cancel:
                buttons.append(self.scan_focus_cancel_btn)
        elif kind == "not_found":
            buttons = [self.scan_focus_add_btn]
        elif kind == "duplicate":
            buttons = []
        elif kind == "ok" and context.get("already_attended"):
            buttons = [self.scan_focus_cancel_btn]
        self.scan_focus_show_buttons(buttons)

    def scan_normalize_card(self, value):
        text = self._clean_value(value)
        if not text:
            return ""
        return text.zfill(8) if text.isdigit() else text

    def scan_lookup_matches(self, card_id):
        normalized = self.scan_normalize_card(card_id)
        if not normalized:
            return []
        candidates = []
        if self.tree.exists(normalized):
            candidates.append(normalized)
        for scan_iid in self._all_iids:
            if not self.tree.exists(scan_iid):
                continue
            if scan_iid == normalized:
                continue
            if self.scan_normalize_card(scan_iid) == normalized:
                candidates.append(scan_iid)
                continue
            card_cell = self.scan_tree_get(scan_iid, "card_id")
            if card_cell and self.scan_normalize_card(card_cell) == normalized:
                candidates.append(scan_iid)
        unique = []
        for scan_iid in [normalized] + candidates:
            if scan_iid and scan_iid not in unique and self.tree.exists(scan_iid):
                unique.append(scan_iid)
        return unique

    def scan_tree_get(self, iid, column):
        if column not in self.tree["columns"]:
            return ""
        try:
            return self._clean_value(self.tree.set(iid, column))
        except Exception:
            return ""

    def scan_collect_missing_tasks(self, iid):
        missing = []
        if self.restrictions.get("exam") and "exam" in self.tree["columns"]:
            if not self.scan_tree_get(iid, "exam"):
                missing.append("exam")
        if self.restrictions.get("homework") and "homework" in self.tree["columns"]:
            if not self.scan_tree_get(iid, "homework"):
                missing.append("homework")
        return missing

    def scan_describe_tasks(self, tasks):
        if not tasks:
            return ""
        labels = {
            "exam": "Exam",
            "homework": "Homework",
        }
        mapped = [labels.get(task, task.title()) for task in tasks]
        if len(mapped) == 1:
            return mapped[0]
        return " & ".join(mapped)

    def scan_append_notes(self, original, addition):
        original_clean = self._clean_value(original)
        addition_clean = self._clean_value(addition)
        if not addition_clean:
            return original_clean
        if not original_clean:
            return addition_clean
        if addition_clean.startswith(original_clean):
            addition_clean = addition_clean[len(original_clean):].lstrip("\n ")
            if not addition_clean:
                return original_clean
        if original_clean.endswith("\n"):
            return f"{original_clean}{addition_clean}"
        return f"{original_clean}\n{addition_clean}"

    def scan_collect_new_note(self):
        if not hasattr(self, "scan_focus_notes"):
            return ""
        try:
            typed = self.scan_focus_notes.get("1.0", "end").strip()
        except Exception:
            return ""
        typed_clean = self._clean_value(typed)
        original = self._clean_value((self.scan_focus_ctx or {}).get("original_notes", ""))
        if not typed_clean or typed_clean == original:
            return ""
        if original and typed_clean.startswith(original):
            suffix = typed_clean[len(original):].lstrip("\n ").rstrip()
            return suffix
        return typed_clean

    def scan_now_tag(self):
        return f"[{datetime.now():%H:%M:%S}]"

    def scan_determine_status(self, scan_ctx):
        preset = scan_ctx.get("status")
        if preset in {"not_found", "duplicate"}:
            return preset
        if not scan_ctx.get("found", True):
            return "not_found"
        missing = scan_ctx.get("missing_tasks", [])
        if missing:
            if "exam" in missing:
                return "missing_exam"
            return "missing_homework"
        return "ok"

    def scan_build_context_for_iid(self, iid, *, source="manual"):
        context = {
            "iid": iid,
            "card_id": self.scan_normalize_card(iid),
            "card_display": self.scan_tree_get(iid, "card_id") or self.scan_normalize_card(iid),
            "name": self.scan_tree_get(iid, "name"),
            "student_id": self.scan_tree_get(iid, "student_id"),
            "attendance": self.scan_tree_get(iid, "attendance").lower(),
            "existing_notes": self.scan_tree_get(iid, "notes"),
            "timestamp": self.scan_tree_get(iid, "timestamp"),
            "source": source,
            "focus_iids": [iid],
            "found": True,
        }
        context["missing_tasks"] = self.scan_collect_missing_tasks(iid)
        context["already_attended"] = context["attendance"] == "attend"
        context["allow_cancel"] = context["already_attended"]
        context["status"] = self.scan_determine_status(context)
        display_name = context["name"] or context["student_id"] or context["card_display"] or "Student"
        context["display_name"] = display_name
        return context

    def scan_build_not_found_context(self, card_id):
        context = {
            "iid": None,
            "card_id": card_id,
            "card_display": card_id,
            "name": "Card Not Linked",
            "student_id": "",
            "attendance": "",
            "existing_notes": "",
            "timestamp": "",
            "source": "scan",
            "focus_iids": [],
            "found": False,
            "missing_tasks": [],
            "status": "not_found",
            "display_name": card_id or "Card",
        }
        return context
    def scan_on_scan(self):
        if self.read_only:
            return
        raw_value = self._clean_value(self.scan_entry.get())
        self.scan_entry.delete(0, "end")
        normalized = self.scan_normalize_card(raw_value)
        if not normalized:
            return
        matches = self.scan_lookup_matches(normalized)
        if not matches:
            context = self.scan_build_not_found_context(normalized)
            context["message"] = "Card not found. Add the student to attend directly."
            context["message_level"] = "warning"
            self.scan_focus_show(context)
            return
        if len(matches) > 1:
            context = {
                "card_id": normalized,
                "card_display": normalized,
                "name": "Multiple records found",
                "student_id": "",
                "existing_notes": "",
                "timestamp": "",
                "source": "scan",
                "focus_iids": matches,
                "found": False,
                "missing_tasks": [],
                "status": "duplicate",
                "display_name": normalized,
                "skip_filter": True,
                "message": "Multiple records share this card. Double-click the correct row to continue.",
                "message_level": "warning",
            }
            self.scan_focus_show(context)
            return
        self.scan_on_open_row(matches[0], source="scan", card_id=normalized)

    def scan_on_row_double_click(self, event):
        if self.read_only:
            return
        scan_iid = ""
        try:
            scan_iid = self.tree.identify_row(event.y)
        except Exception:
            scan_iid = ""
        if not scan_iid:
            selection = self.tree.selection()
            if selection:
                scan_iid = selection[0]
        if not scan_iid:
            return
        self.scan_on_open_row(scan_iid, source="manual")

    def scan_on_open_row(self, iid, *, source="manual", card_id=None):
        if self.read_only:
            return
        if not self.tree.exists(iid):
            self.scan_focus_set_message("Selected record is no longer available.", level="warning")
            return
        context = self.scan_build_context_for_iid(iid, source=source)
        if card_id:
            context["card_id"] = card_id
            context["card_display"] = card_id
        context["message"] = ""
        context["message_level"] = "info"
        self.scan_focus_show(context)
        status = context.get("status", "ok")
        self.scan_focus_set_status(status)
        if status == "ok":
            if context.get("already_attended"):
                display = context.get("display_name", "Student")
                timestamp = context.get("timestamp")
                if timestamp:
                    message = f"{display} already attended at {timestamp}."
                else:
                    message = f"{display} already attended."
                self.scan_focus_set_message(message, level="info")
                self.scan_focus_schedule_clear()
            else:
                self.scan_handle_auto_attend(context)
        elif status in {"missing_exam", "missing_homework"}:
            desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "requirements"
            self.scan_focus_set_message(f"{desc} pending. Choose how to proceed.", level="warning")
        elif status == "not_found":
            self.scan_focus_set_message("Card not found. Add the student to continue.", level="warning")
        elif status == "duplicate":
            self.scan_focus_set_message("Multiple records share this card. Double-click the correct row to continue.", level="warning")

    def scan_handle_auto_attend(self, context):
        if not context:
            return
        tag = self.scan_now_tag()
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(context.get("existing_notes", ""), typed)
        success = self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag)
        display = context.get("display_name", "Student")
        if success:
            self.scan_focus_set_message(f"{display} attended.", level="success")
            self.scan_focus_schedule_clear()
        else:
            self.scan_focus_set_message("Unable to record attendance.", level="error")

    def scan_commit_attendance(self, iid, attendance, notes, *, timestamp=None, warn_on_duplicate=False):
        try:
            return bool(self._set_attendance(
                iid,
                attendance,
                notes,
                warn_on_duplicate=warn_on_duplicate,
                timestamp_override=timestamp,
            ))
        except Exception as exc:
            messagebox.showwarning("Attendance Update Failed", str(exc), parent=self)
            return False

    def scan_focus_on_completed(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"):
            return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "task"
        action_note = f"{tag} Completed {desc} at center."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag):
            self.scan_focus_set_message("Attendance recorded after completion.", level="success")
            self.scan_focus_schedule_clear()

    def scan_focus_on_override(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"):
            return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "task"
        action_note = f"{tag} attended without {desc}."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "attend", final_note, timestamp=tag):
            self.scan_focus_set_message("Override recorded and attendance updated.", level="success")
            self.scan_focus_schedule_clear()

    def scan_focus_on_deny(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"):
            return
        tag = self.scan_now_tag()
        desc = self.scan_describe_tasks(context.get("missing_tasks", [])) or "requirements"
        action_note = f"{tag} Denied Entry: No {desc}."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        if self.scan_commit_attendance(context["iid"], "", final_note, timestamp=tag):
            self.scan_focus_set_message("Denied entry noted.", level="warning")
            self.scan_focus_schedule_clear()

    def scan_focus_on_add_student(self):
        if self.read_only:
            return
        context = self.scan_focus_ctx or {}
        card_id = context.get("card_id") or context.get("card_display")
        typed = self.scan_collect_new_note()
        default_notes = self.scan_append_notes("Diff group", typed) or "Diff group"
        self._launch_add_student_dialog(card_id=card_id, default_notes=default_notes)

    def scan_focus_on_cancel_attendance(self):
        context = self.scan_focus_ctx or {}
        if not context.get("iid"):
            return
        tag = self.scan_now_tag()
        action_note = f"{tag} Canceled."
        base = self.scan_append_notes(context.get("existing_notes", ""), action_note)
        typed = self.scan_collect_new_note()
        final_note = self.scan_append_notes(base, typed)
        self._cancellations += 1
        if self.scan_commit_attendance(context["iid"], "", final_note, timestamp=tag):
            self.scan_focus_set_message("Attendance canceled.", level="warning")
            self.scan_focus_schedule_clear()

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
        self._destroy_scan_focus_window()
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

    def _set_attendance(self, code, attendance, notes, *, warn_on_duplicate=True, timestamp_override=None):
        if self.read_only:
            return False
        if not self.tree.exists(code):
            return False
        target_attendance = self._clean_value(attendance)
        current_att = self._clean_value(self.tree.set(code, "attendance")).lower()
        if warn_on_duplicate and target_attendance.lower() == "attend" and current_att == "attend":
            messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
            return False
        current_notes = self._clean_value(self.tree.set(code, "notes"))
        if notes is None:
            incoming_notes = ""
        else:
            incoming_notes = self._clean_value(notes)
        if incoming_notes == "":
            final_notes = current_notes
        else:
            final_notes = incoming_notes
        if timestamp_override is not None:
            timestamp = timestamp_override
        else:
            timestamp = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        rec = self._build_record_payload(code, target_attendance, final_notes, timestamp)
        try:
            self.sm.add_record(rec)
        except Exception as exc:
            messagebox.showwarning("Attendance Update Failed", str(exc), parent=self)
            return False
        self._update_record_cache(rec)
        self._update_row(code, target_attendance, final_notes, timestamp)
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
        timestamp = self.scan_now_tag() if hasattr(self, "scan_now_tag") else datetime.now().strftime("%H:%M:%S")
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
        try:
            self.sm.add_record(rec)
        except Exception as exc:
            messagebox.showwarning("Unable to add student", str(exc), parent=self)
            return False
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
        if self.scan_focus_ctx and self.scan_focus_ctx.get("status") == "not_found":
            try:
                self.tree.selection_set(cid)
                self.tree.focus(cid)
            except Exception:
                pass
            self.scan_focus_set_message("Student added and attended.", level="success")
            self.scan_focus_schedule_clear()
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
            if parent == getattr(self, "scan_focus_frame", None):
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
