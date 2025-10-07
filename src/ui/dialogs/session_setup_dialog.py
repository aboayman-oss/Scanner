"""Dialog for configuring basic session metadata."""
import customtkinter as ctk
from customtkinter import CTkButton, CTkEntry, CTkFrame, CTkLabel, CTkToplevel
from typing import Dict, Optional

from ui.components.modern_dropdown import ModernDropdown
from utils.helpers import MIN_SESSION_SETUP_SIZE, bring_window_to_front, ensure_initial_size, set_dark_title_bar


class SessionSetupDialog(CTkToplevel):
    def __init__(self, parent, stages, centers, has_data, session_data, callback):
        super().__init__(parent)
        set_dark_title_bar(self)
        self.parent = parent
        self.stage_placeholder = "Select a Stage"
        self.center_placeholder = "Select a Center"
        self.stages = [self.stage_placeholder] + (stages or [])
        self.centers = [self.center_placeholder] + (centers or [])
        self.session_data: Dict[str, Dict[str, int]] = session_data or {}
        self.callback = callback
        self.has_data = has_data
        self.stage_cb: Optional[ModernDropdown] = None
        self.center_cb: Optional[ModernDropdown] = None
        self.session_ent: Optional[CTkEntry] = None
        self.submit_btn: Optional[CTkButton] = None
        self.session_no_var = ctk.StringVar()
        self._last_suggested_pair = None
        self.title("Start New Session")
        self.resizable(False, False)
        self.minsize(*MIN_SESSION_SETUP_SIZE)
        self.transient(parent)
        self.grid_columnconfigure(0, weight=1)

        notice_text = (
            "Using the imported dataset for this session."
            if has_data
            else "No dataset imported yet. A blank roster will be created."
        )
        self.notice_var = ctk.StringVar(value=notice_text)
        self.error_var = ctk.StringVar(value="")

        self.title_font = ctk.CTkFont(family="Roboto", size=21, weight="bold")
        self.body_font = ctk.CTkFont(family="Roboto", size=13)
        self.label_font = ctk.CTkFont(family="Roboto", size=12, weight="bold")

        self.field_bg_color = ("#E8EDF6", "#2C3039")
        self.field_border_color = ("#CBD5E1", "#3D4452")
        self.icon_color = ("#4B5563", "#A0AEC0")
        self.error_color = "#b00020"
        self.field_hover_bg_color = ("#F2F5FD", "#343B47")
        self.field_text_color = ("#1F2937", "#E2E8F0")
        self.field_placeholder_color = ("#6B7280", "#94A3B8")
        self.dropdown_panel_color = ("#FFFFFF", "#1F242C")
        self.dropdown_option_hover_color = ("#E1E9F9", "#3A4352")
        self.dropdown_option_active_color = ("#D6E2FB", "#3D485B")

        self._build_form()
        ensure_initial_size(self, min_size=MIN_SESSION_SETUP_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Return>", lambda _e: self._on_submit())
        self.bind("<Escape>", lambda _e: self._on_cancel())
        self._focus_after_id = None
        # Give time for widgets to be properly created and mapped
        self._focus_after_id = self.after(100, self._initialize_window)

    def _build_form(self):
        content = CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)
        content.grid_columnconfigure(0, weight=1)

        row = 0
        CTkLabel(
            content,
            text="Start New Session",
            font=self.title_font,
            anchor="w",
        ).grid(row=row, column=0, sticky="ew")
        row += 1
        CTkLabel(
            content,
            textvariable=self.notice_var,
            font=self.body_font,
            justify="left",
            anchor="w",
            wraplength=360,
        ).grid(row=row, column=0, sticky="ew", pady=(8, 20))
        row += 1

        row, self.stage_cb = self._add_combo_field(
            content,
            row,
            "Stage",
            self.stages,
            placeholder=self.stage_placeholder,
            command=self._on_stage_selection,
        )
        row, self.center_cb = self._add_combo_field(
            content,
            row,
            "Center",
            self.centers,
            placeholder=self.center_placeholder,
            command=self._on_center_selection,
        )

        CTkLabel(
            content,
            text="Session No.",
            font=self.label_font,
            anchor="w",
        ).grid(row=row, column=0, sticky="w")
        row += 1
        session_container = CTkFrame(
            content,
            fg_color=self.field_bg_color,
            corner_radius=16,
            border_width=1,
            border_color=self.field_border_color,
        )
        session_container.grid(row=row, column=0, sticky="ew", pady=(6, 20))
        session_container.grid_columnconfigure(1, weight=1)

        CTkLabel(
            session_container,
            text="#",
            font=self.label_font,
            width=28,
            anchor="center",
            text_color=self.icon_color,
        ).grid(row=0, column=0, padx=(12, 8), pady=10)

        validate_cmd = (self.register(self._validate_session_no_input), "%P")
        self.session_ent = CTkEntry(
            session_container,
            border_width=0,
            corner_radius=10,
            fg_color="transparent",
            font=self.body_font,
            textvariable=self.session_no_var,
            validate="key",
            validatecommand=validate_cmd,
        )
        self.session_ent.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=10)
        self.session_no_var.trace_add("write", self._on_session_no_changed)

        row += 1
        CTkLabel(
            content,
            textvariable=self.error_var,
            font=self.body_font,
            text_color=self.error_color,
            justify="left",
            anchor="w",
            wraplength=360,
        ).grid(row=row, column=0, sticky="ew", pady=(0, 4))

        row += 1
        btn_frame = CTkFrame(content, fg_color="transparent")
        btn_frame.grid(row=row, column=0, sticky="e", pady=(24, 0))
        self.submit_btn = CTkButton(
            btn_frame,
            text="Start Session",
            command=self._on_submit,
            corner_radius=18,
            state="disabled",
        )
        self.submit_btn.pack(side="right")
        CTkButton(
            btn_frame,
            text="Cancel",
            command=self._on_cancel,
            fg_color="transparent",
            hover_color=("#E5E7EB", "#2E2E2E"),
            text_color=("#4B5563", "#A0AEC0"),
            border_width=0,
            corner_radius=18,
        ).pack(side="right", padx=(0, 12))

        self._validate_form()

    def _add_combo_field(self, parent, start_row, label_text, values, *, placeholder, command=None):
        CTkLabel(
            parent,
            text=label_text,
            font=self.label_font,
            anchor="w",
        ).grid(row=start_row, column=0, sticky="w")
        combo = ModernDropdown(
            parent,
            values=[v for v in values if v != placeholder],
            placeholder=placeholder,
            command=command,
            font=self.body_font,
            text_color=self.field_text_color,
            placeholder_color=self.field_placeholder_color,
            base_fg_color=self.field_bg_color,
            hover_fg_color=self.field_hover_bg_color,
            border_color=self.field_border_color,
            dropdown_bg_color=self.dropdown_panel_color,
            option_hover_color=self.dropdown_option_hover_color,
            active_option_color=self.dropdown_option_active_color,
            icon_color=self.icon_color,
            option_height=32,
            max_visible_items=6,
            max_dropdown_height=320,
        )
        combo.grid(row=start_row + 1, column=0, sticky="ew", pady=(6, 16))
        combo.set(placeholder)
        return start_row + 2, combo

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

    def _initialize_window(self):
        """Initialize window position and focus after widgets are mapped."""
        self._focus_after_id = None
        if not self.winfo_exists():
            return
        self._center_on_parent()
        if self.session_ent and self.session_ent.winfo_exists():
            self.session_ent.focus_set()

    def _validate_session_no_input(self, proposed: str) -> bool:
        return proposed.isdigit() or proposed == ""

    def _on_session_no_changed(self, *_):
        self.error_var.set("")
        self._validate_form()

    def _on_stage_selection(self, _value):
        if self.stage_cb and self.stage_cb.get() == self.stage_placeholder:
            self.session_no_var.set("")
            self._last_suggested_pair = None
        self.error_var.set("")
        self._validate_form()
        self._suggest_session_number()

    def _on_center_selection(self, _value):
        if self.center_cb and self.center_cb.get() == self.center_placeholder:
            self.session_no_var.set("")
            self._last_suggested_pair = None
        self.error_var.set("")
        self._validate_form()
        self._suggest_session_number()

    def _validate_form(self, *_):
        if not self.submit_btn:
            return
        stage = self.stage_cb.get().strip() if self.stage_cb else ""
        center = self.center_cb.get().strip() if self.center_cb else ""
        session_no = self.session_no_var.get().strip()
        is_valid = (
            stage
            and center
            and session_no
            and stage != self.stage_placeholder
            and center != self.center_placeholder
            and session_no.isdigit()
        )
        self.submit_btn.configure(state="normal" if is_valid else "disabled")
        if is_valid:
            self.error_var.set("")

    def _suggest_session_number(self):
        if not self.stage_cb or not self.center_cb:
            return
        stage = self.stage_cb.get().strip()
        center = self.center_cb.get().strip()
        if (
            not stage
            or not center
            or stage == self.stage_placeholder
            or center == self.center_placeholder
        ):
            return
        pair = (stage, center)
        center_map = self.session_data.get(stage, {})
        if not isinstance(center_map, dict):
            self._last_suggested_pair = pair
            return
        last_session = center_map.get(center)
        if last_session is None:
            self._last_suggested_pair = pair
            return
        suggestion = str(last_session + 1)
        current_value = self.session_no_var.get().strip()
        if pair != self._last_suggested_pair or current_value != suggestion:
            self.session_no_var.set(suggestion)
        self._last_suggested_pair = pair

    def _on_submit(self):
        if not self.stage_cb or not self.center_cb or not self.session_ent:
            self.error_var.set("Dialog not ready. Please reopen and try again.")
            return
        stage = self.stage_cb.get().strip()
        center = self.center_cb.get().strip()
        session_no = self.session_no_var.get().strip()
        if (
            stage == self.stage_placeholder
            or center == self.center_placeholder
            or not session_no.isdigit()
        ):
            self.error_var.set("Select stage, center, and enter a numeric session number.")
            return
        self.error_var.set("")
        payload = {
            "stage": stage,
            "center": center,
            "no": int(session_no),
            "name": f"{stage} {center} session {int(session_no)}"
        }
        if self._focus_after_id:
            self.after_cancel(self._focus_after_id)
            self._focus_after_id = None
        if self.callback:
            self.callback(payload)
            self.callback = None
        self.destroy()

    def _on_cancel(self):
        if self._focus_after_id:
            self.after_cancel(self._focus_after_id)
            self._focus_after_id = None
        if self.callback:
            self.callback(None)
            self.callback = None
        self.destroy()
