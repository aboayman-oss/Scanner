"""Settings window for configuring application preferences."""
import json
import os
from tkinter import filedialog, messagebox

import customtkinter as ctk
from customtkinter import (
    CTkButton,
    CTkEntry,
    CTkFrame,
    CTkImage,
    CTkLabel,
    CTkScrollableFrame,
    CTkSegmentedButton,
    CTkSwitch,
    CTkToplevel,
)
from PIL import Image

from ui.components.modern_dropdown import ModernDropdown
from utils.helpers import (
    FOLDER_OPEN_ICON_FILE,
    MAPPING_FILE,
    PLUS_ICON_FILE,
    REMOVE_ICON_FILE,
    SETTINGS,
    SETTINGS_FILE,
    STATUS_INFO_ICON_FILE,
    STATUS_OK_ICON_FILE,
    bring_window_to_front,
    read_data,
    set_dark_title_bar,
)


class SettingsWindow(CTkToplevel):
    """Toplevel dialog for managing template mapping, stage lists, and preferences."""

    mapping_placeholder = "-- Select --"
    _list_row_base_color = "#3F4550"
    _list_row_hover_color = "#4C5563"

    def __init__(self, parent):
        super().__init__(parent)
        set_dark_title_bar(self)
        self.title("Settings")
        self.geometry("1000x800")
        self.minsize(1000, 800)
        self.configure(fg_color="#2B2B2B")

        self.parent_app = parent
        self.after(50, lambda: bring_window_to_front(self))

        self.column_map = dict(getattr(parent, "column_map", {}) or {})
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
        self.mapping_groups = [
            ("Student Information", self.mapping_fields[:4]),
            ("Session Data", self.mapping_fields[4:7]),
            ("Additional Tracking", self.mapping_fields[7:]),
        ]
        self.mapping_labels = {field_key: label_text for label_text, field_key in self.mapping_fields}
        self.mapping_controls = {}
        self.mapping_hint_labels = {}
        self.mapping_source_path = None
        self.mapping_columns = []
        for value in self.working_mapping.values():
            if value and value not in self.mapping_columns:
                self.mapping_columns.append(value)

        self.stage_items = list(SETTINGS["stage_options"])
        self.center_items = list(SETTINGS["center_options"])
        self.stage_rows = {}
        self.center_rows = {}
        self.stage_entry = None
        self.stage_scroll = None
        self.center_entry = None
        self.center_scroll = None

        self.var_exam = ctk.BooleanVar(value=SETTINGS["restrictions"].get("exam", False))
        self.var_homework = ctk.BooleanVar(value=SETTINGS["restrictions"].get("homework", False))
        self.var_file_type = ctk.StringVar(value=SETTINGS.get("file_type", "xlsx").upper())

        self.template_status_var = ctk.StringVar()
        self.hint_font = ctk.CTkFont(size=12)

        self._icon_cache = {}
        self.status_icons = {
            "ok": self._load_icon(STATUS_OK_ICON_FILE, (20, 20)),
            "info": self._load_icon(STATUS_INFO_ICON_FILE, (20, 20)),
        }
        self.folder_icon = self._load_icon(FOLDER_OPEN_ICON_FILE, (18, 18))
        self.plus_icon = self._load_icon(PLUS_ICON_FILE, (16, 16))
        self.remove_icon = self._load_icon(REMOVE_ICON_FILE, (14, 14))

        assets_dir = os.path.dirname(FOLDER_OPEN_ICON_FILE)
        self.accent_color = "#2F80ED"
        self.surface_color = "#353C44"
        self.surface_muted_color = "#2A3037"
        self.input_surface_color = "#1F2329"
        self.nav_bg_color = "#232323"
        self.nav_button_color = "#2D2D2D"
        self.nav_button_hover_color = "#383838"
        self.template_status_styles = {
            "info": {"bg": "#253446", "text": "#BED0E8"},
            "warn": {"bg": "#3D2C1F", "text": "#F2C089"},
            "ok": {"bg": "#1F3A2F", "text": "#C7EED6"},
        }
        self.nav_icons = {
            "template": self._load_icon(os.path.join(assets_dir, "mapping.png"), (18, 18)),
            "stage": self._load_icon(os.path.join(assets_dir, "stage.png"), (18, 18)),
            "general": self._load_icon(os.path.join(assets_dir, "settings.png"), (18, 18)),
        }

        container = CTkFrame(self, fg_color="#2B2B2B")
        container.pack(fill="both", expand=True, padx=32, pady=(20, 36))
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=0)
        container.grid_columnconfigure(1, weight=1)

        nav_frame = CTkFrame(container, fg_color=self.nav_bg_color, corner_radius=20)
        nav_frame.grid(row=0, column=0, sticky="nsw")
        nav_frame.grid_rowconfigure(2, weight=1)

        nav_header = CTkFrame(nav_frame, fg_color="transparent")
        nav_header.pack(fill="x", padx=20, pady=(20, 12))
        CTkLabel(nav_header, text="Settings", font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        CTkLabel(
            nav_header,
            text="Choose a section to configure.",
            font=ctk.CTkFont(size=14),
            text_color="#BEBEBE",
            justify="left",
            wraplength=180,
        ).pack(anchor="w", pady=(4, 0))

        self.nav_button_container = CTkFrame(nav_frame, fg_color="transparent")
        self.nav_button_container.pack(fill="both", expand=True, padx=12, pady=(0, 20))

        self.content_area = CTkFrame(container, fg_color="#2B2B2B")
        self.content_area.grid(row=0, column=1, sticky="nsew", padx=(24, 0))
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self.section_frames = {}
        self.nav_buttons = {}
        self.active_section = None

        self.section_frames["template"] = self._build_template_section(self.content_area)
        self.section_frames["stage"] = self._build_stage_section(self.content_area)
        self.section_frames["general"] = self._build_general_section(self.content_area)

        for key, label in (
            ("template", "Template Mapping"),
            ("stage", "Stage & Center"),
            ("general", "General"),
        ):
            icon = self.nav_icons.get(key)
            self._create_nav_button(self.nav_button_container, key, label, icon)

        btn_frame = CTkFrame(self, fg_color=self.nav_bg_color, corner_radius=18)
        btn_frame.pack(side="bottom", fill="x", padx=32, pady=(8, 20))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        self.cancel_button = CTkButton(
            btn_frame,
            text="Cancel",
            command=self._cancel,
            fg_color="#2B2B2B",
            hover_color="#353C44",
            border_width=2,
            border_color=self.accent_color,
            text_color=self.accent_color,
            corner_radius=14,
        )
        self.cancel_button.grid(row=0, column=0, sticky="ew", padx=(16, 8), pady=14)

        self.apply_button = CTkButton(
            btn_frame,
            text="Apply",
            command=self._apply_settings,
            fg_color=self.accent_color,
            hover_color="#1C64D1",
            text_color="#FFFFFF",
            corner_radius=14,
            state="disabled",
        )
        self.apply_button.grid(row=0, column=1, sticky="ew", padx=(8, 16), pady=14)

        if self.working_mapping:
            self.template_status_var.set("Using saved mapping. Load a sample file to refresh the template.")
        else:
            self.template_status_var.set("Load a sample file to map template fields.")

        self._populate_template_controls()
        self._show_section("template")

    def _load_icon(self, path, size):
        if not path or not os.path.exists(path):
            return None
        key = (path, size)
        icon = self._icon_cache.get(key)
        if icon is None:
            image = Image.open(path)
            icon = CTkImage(light_image=image, dark_image=image, size=size)
            self._icon_cache[key] = icon
        return icon

    def _create_nav_button(self, parent, key, label, icon):
        button = CTkButton(
            parent,
            text=label,
            image=icon,
            compound="left",
            anchor="w",
            height=48,
            fg_color=self.nav_button_color,
            hover_color=self.nav_button_hover_color,
            text_color="#E6E6E6",
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=16,
            command=lambda item=key: self._show_section(item),
        )
        button.pack(fill="x", pady=4)
        self.nav_buttons[key] = button
        return button

    def _style_nav_button(self, key, active):
        button = self.nav_buttons.get(key)
        if not button:
            return
        if active:
            button.configure(fg_color=self.accent_color, hover_color="#1C64D1", text_color="#FFFFFF")
        else:
            button.configure(
                fg_color=self.nav_button_color,
                hover_color=self.nav_button_hover_color,
                text_color="#E6E6E6",
            )

    def _show_section(self, key):
        if key == getattr(self, "active_section", None):
            return
        if getattr(self, "active_section", None) in self.section_frames:
            current = self.section_frames[self.active_section]
            current.pack_forget()
            self._style_nav_button(self.active_section, active=False)

        frame = self.section_frames.get(key)
        if frame:
            frame.pack(fill="both", expand=True)
        self._style_nav_button(key, active=True)
        self.active_section = key
        if key == "template":
            self._update_template_status_display()
        self._update_apply_state()

    def _build_template_section(self, parent):
        frame = CTkFrame(parent, fg_color=self.surface_color, corner_radius=24)
        frame.pack_propagate(False)

        header = CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(28, 12))
        CTkLabel(header, text="Template Mapping", font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        CTkLabel(
            header,
            text="Assign each template field to a column from a sample data file.",
            font=ctk.CTkFont(size=15),
            text_color="#BEBEBE",
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        initial_style = self.template_status_styles["info"]
        self.template_status_card = CTkFrame(frame, fg_color=initial_style["bg"], corner_radius=18)
        self.template_status_card.pack(fill="x", padx=28, pady=(4, 16))
        self.template_status_card.grid_columnconfigure(1, weight=1)

        info_icon = self.status_icons.get("info")
        self.template_status_icon_label = CTkLabel(self.template_status_card, text="", image=info_icon)
        self.template_status_icon_label.grid(row=0, column=0, padx=(18, 12), pady=18, sticky="n")

        self.template_status_text = CTkLabel(
            self.template_status_card,
            textvariable=self.template_status_var,
            justify="left",
            wraplength=540,
            text_color=initial_style["text"],
            font=ctk.CTkFont(size=14),
        )
        self.template_status_text.grid(row=0, column=1, sticky="w", pady=18)

        actions = CTkFrame(frame, fg_color="transparent")
        actions.pack(fill="x", padx=28, pady=(0, 12))
        actions.grid_columnconfigure(0, weight=1)
        CTkButton(
            actions,
            text="Load Sample File",
            image=self.folder_icon,
            compound="left",
            command=self._prompt_for_columns,
            fg_color=self.accent_color,
            hover_color="#1C64D1",
            text_color="#FFFFFF",
            corner_radius=12,
        ).grid(row=0, column=1, sticky="e")

        form_container = CTkScrollableFrame(frame, fg_color="transparent")
        form_container.pack(fill="both", expand=True, padx=12, pady=(4, 28))
        form_container.grid_columnconfigure(0, weight=1)

        for group_title, fields in self.mapping_groups:
            group_card = CTkFrame(form_container, fg_color=self.surface_muted_color, corner_radius=20)
            group_card.pack(fill="x", expand=True, padx=16, pady=(0, 16))
            group_card.grid_columnconfigure(1, weight=1)

            CTkLabel(
                group_card,
                text=group_title,
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#FFFFFF",
                anchor="w",
            ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(18, 10))

            for idx, (label_text, field_key) in enumerate(fields):
                row = idx * 2 + 1
                CTkLabel(
                    group_card,
                    text=f"{label_text}:",
                    anchor="w",
                    text_color="#E8EAF0",
                ).grid(row=row, column=0, sticky="w", padx=(20, 16), pady=(0, 8))

                combo = ModernDropdown(
                    group_card,
                    placeholder=self.mapping_placeholder,
                    values=[self.mapping_placeholder],
                    command=lambda value, key=field_key: self._on_mapping_change(key, value),
                    base_fg_color=(self.input_surface_color, self.input_surface_color),
                    hover_fg_color=("#2B2E36", "#2B2E36"),
                    border_color=(self.input_surface_color, self.input_surface_color),
                    dropdown_bg_color=(self.input_surface_color, self.input_surface_color),
                )
                combo.grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=(0, 8))
                self.mapping_controls[field_key] = combo

                hint = CTkLabel(
                    group_card,
                    text="",
                    font=self.hint_font,
                    text_color="#F28D35",
                    justify="left",
                    wraplength=480,
                )
                hint.grid(row=row + 1, column=1, sticky="w", padx=(0, 20), pady=(0, 6))
                self.mapping_hint_labels[field_key] = hint

        return frame

    def _build_stage_section(self, parent):
        frame = CTkFrame(parent, fg_color=self.surface_color, corner_radius=24)

        header = CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(28, 12))
        CTkLabel(header, text="Stage & Center", font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        CTkLabel(
            header,
            text="Manage the stage and center choices available when starting a session.",
            font=ctk.CTkFont(size=15),
            text_color="#BEBEBE",
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        cards = CTkFrame(frame, fg_color="transparent")
        cards.pack(fill="both", expand=True, padx=28, pady=(8, 28))
        cards.grid_columnconfigure(0, weight=1)
        cards.grid_columnconfigure(1, weight=1)
        cards.grid_rowconfigure(0, weight=1)

        stage_card = self._build_option_card(
            cards,
            title="Stage Options",
            placeholder="Add stage",
            add_callback=self._add_stage,
            entry_attr="stage_entry",
            scroll_attr="stage_scroll",
            rows_dict=self.stage_rows,
            items=self.stage_items,
            remove_callback=self._remove_stage_value,
        )
        stage_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        center_card = self._build_option_card(
            cards,
            title="Center Options",
            placeholder="Add center",
            add_callback=self._add_center,
            entry_attr="center_entry",
            scroll_attr="center_scroll",
            rows_dict=self.center_rows,
            items=self.center_items,
            remove_callback=self._remove_center_value,
        )
        center_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        return frame

    def _build_option_card(
        self,
        parent,
        *,
        title,
        placeholder,
        add_callback,
        entry_attr,
        scroll_attr,
        rows_dict,
        items,
        remove_callback,
    ):
        card = CTkFrame(parent, fg_color=self.surface_muted_color, corner_radius=20)
        card.grid_rowconfigure(2, weight=1)

        CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(18, 10))

        entry_row = CTkFrame(card, fg_color=self.input_surface_color, corner_radius=16)
        entry_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 12))
        entry_row.grid_columnconfigure(0, weight=1)

        entry = CTkEntry(entry_row, placeholder_text=placeholder, corner_radius=12)
        entry.grid(row=0, column=0, sticky="ew", padx=(12, 8), pady=12)
        entry.bind("<Return>", lambda _event: add_callback())

        CTkButton(
            entry_row,
            text="Add",
            image=self.plus_icon,
            compound="left",
            width=110,
            fg_color=self.accent_color,
            hover_color="#1C64D1",
            text_color="#FFFFFF",
            corner_radius=12,
            command=add_callback,
        ).grid(row=0, column=1, sticky="e", padx=(0, 12), pady=12)

        setattr(self, entry_attr, entry)

        list_shell = CTkFrame(card, fg_color=self.input_surface_color, corner_radius=18)
        list_shell.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=20, pady=(0, 20))

        scroll = CTkScrollableFrame(list_shell, fg_color=self.input_surface_color)
        scroll.pack(fill="both", expand=True, padx=10, pady=12)
        scroll.grid_columnconfigure(0, weight=1)
        setattr(self, scroll_attr, scroll)

        self._populate_option_rows(scroll, rows_dict, items, remove_callback)
        return card

    def _build_general_section(self, parent):
        frame = CTkFrame(parent, fg_color=self.surface_color, corner_radius=24)

        header = CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(28, 12))
        CTkLabel(header, text="General", font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        CTkLabel(
            header,
            text="Configure optional columns and preferred export format.",
            font=ctk.CTkFont(size=15),
            text_color="#BEBEBE",
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        body = CTkScrollableFrame(frame, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=(8, 28))
        body.grid_columnconfigure(0, weight=1)

        toggles_card = CTkFrame(body, fg_color=self.surface_muted_color, corner_radius=20)
        toggles_card.pack(fill="x", expand=True, padx=16, pady=(0, 20))

        CTkLabel(
            toggles_card,
            text="Optional Columns",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).pack(anchor="w", padx=20, pady=(18, 8))

        toggles_inner = CTkFrame(toggles_card, fg_color="transparent")
        toggles_inner.pack(fill="x", padx=16, pady=(0, 18))

        self._add_toggle_row(
            toggles_inner,
            title="Enable Exam Column",
            description="Include exam grades",
            variable=self.var_exam,
        )
        self._add_toggle_row(
            toggles_inner,
            title="Enable Homework Column",
            description="Include homework completion",
            variable=self.var_homework,
        )

        format_card = CTkFrame(body, fg_color=self.surface_muted_color, corner_radius=20)
        format_card.pack(fill="x", expand=True, padx=16)

        CTkLabel(
            format_card,
            text="Data Format",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).pack(anchor="w", padx=20, pady=(18, 8))

        segment_wrapper = CTkFrame(format_card, fg_color=self.input_surface_color, corner_radius=16)
        segment_wrapper.pack(fill="x", padx=20, pady=(0, 20))
        segment_wrapper.grid_columnconfigure(0, weight=1)

        self.filetype_segment = CTkSegmentedButton(
            segment_wrapper,
            values=["CSV", "XLSX"],
            variable=self.var_file_type,
            command=self._on_file_type_change,
            corner_radius=14,
        )
        self.filetype_segment.grid(row=0, column=0, sticky="ew", padx=12, pady=12)
        self.filetype_segment.configure(
            fg_color=self.input_surface_color,
            selected_color=self.accent_color,
            selected_hover_color="#1C64D1",
            unselected_color="#2F333B",
            unselected_hover_color="#3C414B",
            text_color="#E6E6E6",
        )
        self.filetype_segment.set(self.var_file_type.get())

        return frame

    def _add_toggle_row(self, parent, *, title, description, variable):
        row = CTkFrame(parent, fg_color=self.input_surface_color, corner_radius=16)
        row.pack(fill="x", pady=8)
        row.grid_columnconfigure(0, weight=1)
        row.grid_columnconfigure(1, weight=0)

        text_frame = CTkFrame(row, fg_color="transparent")
        text_frame.grid(row=0, column=0, sticky="w", padx=16, pady=(14, 14))

        CTkLabel(
            text_frame,
            text=title,
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).pack(anchor="w")

        CTkLabel(
            text_frame,
            text=description,
            justify="left",
            text_color="#BEBEBE",
            wraplength=420,
        ).pack(anchor="w", pady=(4, 0))

        switch_frame = CTkFrame(row, fg_color="transparent")
        switch_frame.grid(row=0, column=1, sticky="ns", padx=16)
        switch_frame.grid_rowconfigure(0, weight=1)

        CTkSwitch(
            switch_frame,
            text="",
            variable=variable,
            command=self._update_apply_state,
            progress_color=self.accent_color,
            fg_color="#40464F",
            button_color="#FFFFFF",
            button_hover_color="#D7E3F8",
        ).grid(row=0, column=0, sticky="")

    def _populate_option_rows(self, container, rows_dict, items, remove_callback):
        for child in container.winfo_children():
            child.destroy()
        rows_dict.clear()
        for value in items:
            self._create_option_row(container, rows_dict, value, remove_callback, animate=False)

    def _create_option_row(self, container, rows_dict, value, remove_callback, animate=False):
        row = CTkFrame(container, fg_color=self._list_row_base_color, corner_radius=18)
        row.pack(fill="x", pady=6, padx=2)
        row.grid_columnconfigure(0, weight=1)
        row._hide_job = None  # type: ignore[attr-defined]

        CTkLabel(
            row,
            text=value,
            anchor="w",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#FFFFFF",
        ).grid(row=0, column=0, sticky="w", padx=18, pady=10)

        button = CTkButton(
            row,
            text="",
            image=self.remove_icon,
            width=36,
            command=lambda: remove_callback(value),
            fg_color="transparent",
            hover_color="#566070",
        )
        button.grid(row=0, column=1, sticky="e", padx=(0, 14), pady=8)
        button.grid_remove()

        self._bind_row_hover(row, button)
        rows_dict[value] = row

        if animate:
            self._animate_row_in(row)
        else:
            row.configure(fg_color=self._list_row_base_color)
        return row

    def _bind_row_hover(self, row, button):
        widgets = [row, button, *row.winfo_children()]

        def show(_event=None):
            if not row.winfo_exists():
                return
            hide_job = getattr(row, "_hide_job", None)
            if hide_job:
                row.after_cancel(hide_job)
                row._hide_job = None  # type: ignore[attr-defined]
            if not button.winfo_ismapped():
                button.grid()
            row.configure(fg_color=self._list_row_hover_color)

        def schedule_hide(_event=None):
            if not row.winfo_exists():
                return

            def _hide():
                if not row.winfo_exists():
                    return
                pointer_widget = row.winfo_containing(*row.winfo_pointerxy())
                if not self._is_descendant(pointer_widget, row):
                    if button.winfo_exists():
                        button.grid_remove()
                    row.configure(fg_color=self._list_row_base_color)
                row._hide_job = None  # type: ignore[attr-defined]

            hide_job = getattr(row, "_hide_job", None)
            if hide_job:
                row.after_cancel(hide_job)
            row._hide_job = row.after(120, _hide)  # type: ignore[attr-defined]

        for widget in widgets:
            widget.bind("<Enter>", show, add="+")
            widget.bind("<Leave>", schedule_hide, add="+")
        button.bind("<ButtonRelease-1>", lambda _event: schedule_hide(), add="+")

    def _animate_row_in(self, row):
        colors = ["#2E333B", "#363C45", self._list_row_base_color]

        def step(index=0):
            if not row.winfo_exists() or index >= len(colors):
                return
            row.configure(fg_color=colors[index])
            if index < len(colors) - 1:
                row.after(60, lambda: step(index + 1))

        step()

    def _is_descendant(self, widget, ancestor):
        while widget is not None:
            if widget is ancestor:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _flash_existing_row(self, row):
        if not row or not row.winfo_exists():
            return

        base = self._list_row_base_color
        accent = self.accent_color

        def pulse(iteration=0):
            if not row.winfo_exists():
                return
            row.configure(fg_color=accent if iteration % 2 == 0 else base)
            if iteration < 3:
                row.after(120, lambda: pulse(iteration + 1))
            else:
                row.configure(fg_color=base)

        pulse()

    def _add_stage(self):
        if not self.stage_entry:
            return
        value = self.stage_entry.get().strip()
        if not value:
            return
        if value in self.stage_items:
            self._flash_existing_row(self.stage_rows.get(value))
            self.stage_entry.delete(0, "end")
            return
        self.stage_items.append(value)
        self._create_option_row(self.stage_scroll, self.stage_rows, value, self._remove_stage_value, animate=True)
        self.stage_entry.delete(0, "end")
        self._update_apply_state()

    def _remove_stage_value(self, value):
        if value not in self.stage_items:
            return
        self.stage_items.remove(value)
        row = self.stage_rows.pop(value, None)
        if row and row.winfo_exists():
            row.destroy()
        self._update_apply_state()

    def _add_center(self):
        if not self.center_entry:
            return
        value = self.center_entry.get().strip()
        if not value:
            return
        if value in self.center_items:
            self._flash_existing_row(self.center_rows.get(value))
            self.center_entry.delete(0, "end")
            return
        self.center_items.append(value)
        self._create_option_row(
            self.center_scroll, self.center_rows, value, self._remove_center_value, animate=True
        )
        self.center_entry.delete(0, "end")
        self._update_apply_state()

    def _remove_center_value(self, value):
        if value not in self.center_items:
            return
        self.center_items.remove(value)
        row = self.center_rows.pop(value, None)
        if row and row.winfo_exists():
            row.destroy()
        self._update_apply_state()

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
        self._refresh_mapping_hints()

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
            filetypes=[(f"{file_type.upper()} files", ext)],
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
        # All fields except optional exam/homework must be mapped
        required_fields = [fk for _, fk in self.mapping_fields if fk not in ("exam", "homework")]
        for field_key in required_fields:
            if not mapping.get(field_key):
                return False

        # All mapped values must be unique
        mapped_values = [v for v in mapping.values() if v]
        if len(mapped_values) != len(set(mapped_values)):
            return False

        return True

    def _refresh_mapping_hints(self):
        mapping = self._collect_mapping()
        value_to_fields = {}
        for field_key, column in mapping.items():
            if column:
                value_to_fields.setdefault(column, []).append(field_key)

        conflicts = {
            field
            for columns in value_to_fields.values()
            if len(columns) > 1
            for field in columns
        }

        for field_key, hint_label in self.mapping_hint_labels.items():
            combo = self.mapping_controls[field_key]
            current_value = mapping.get(field_key, "")
            is_required = field_key not in ("exam", "homework")

            if current_value and field_key in conflicts:
                others = [
                    self.mapping_labels[other]
                    for other in value_to_fields.get(current_value, [])
                    if other != field_key
                ]
                hint_label.configure(
                    text=f"Already used by {', '.join(others)}." if others else "Duplicate selection.",
                    text_color="#F28D35",
                )
                combo.configure(border_color="#F28D35", border_width=2)
            elif not current_value and is_required and self.mapping_columns:
                hint_label.configure(text="This field is required.", text_color="#F28D35")
                combo.configure(border_color="#F28D35", border_width=2)
            else:
                hint_label.configure(text="")
                combo.configure(border_width=1, border_color=self.input_surface_color)

    def _update_template_status_display(self, is_valid=None):
        if is_valid is None:
            is_valid = self._is_mapping_valid()

        if not self.mapping_columns:
            style = self.template_status_styles["info"]
            icon = self.status_icons["info"]
        elif is_valid:
            style = self.template_status_styles["ok"]
            icon = self.status_icons["ok"]
        else:
            style = self.template_status_styles["warn"]
            icon = self.status_icons["info"]

        self.template_status_card.configure(fg_color=style["bg"])
        self.template_status_text.configure(text_color=style["text"])
        if icon:
            self.template_status_icon_label.configure(image=icon)

    def _has_changes(self):
        # Check mapping
        if self.working_mapping != self.column_map:
            return True
        # Check stage/center options
        if set(self.stage_items) != set(SETTINGS["stage_options"]):
            return True
        if set(self.center_items) != set(SETTINGS["center_options"]):
            return True
        # Check restrictions
        if self.var_exam.get() != SETTINGS["restrictions"].get("exam", False):
            return True
        if self.var_homework.get() != SETTINGS["restrictions"].get("homework", False):
            return True
        # Check file type
        if self.var_file_type.get().lower() != SETTINGS.get("file_type", "xlsx"):
            return True
        return False

    def _update_apply_state(self, *_args):
        self._refresh_mapping_hints()
        is_valid = self._is_mapping_valid()
        has_changes = self._has_changes()

        # Don't change state if the button is in the "Saved!" state
        if self.apply_button.cget("text") == "Saved!":
            return

        self.apply_button.configure(state="normal" if is_valid and has_changes else "disabled")
        if self.active_section == "template":
            self._update_template_status_display(is_valid)

    def _on_file_type_change(self, value):
        self.var_file_type.set(value)
        self._update_apply_state()

    def _apply_settings(self):
        if not self._is_mapping_valid():
            messagebox.showerror("Invalid Mapping", "Each required field must map to a unique column.", parent=self)
            self._show_section("template")
            return

        mapping = self._collect_mapping()
        stage_options = sorted(list(self.stage_items))
        center_options = sorted(list(self.center_items))
        restrictions = {
            "exam": bool(self.var_exam.get()),
            "homework": bool(self.var_homework.get()),
        }
        file_type = self.var_file_type.get().lower()

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

        if hasattr(self.parent_app, "set_status"):
            self.parent_app.set_status("Settings saved.")

        self.apply_button.configure(text="Saved!", state="disabled")

        def reset_button():
            if self.apply_button.winfo_exists():
                self.apply_button.configure(text="Apply")
                self._update_apply_state()

        self.after(2000, reset_button)

    def _cancel(self):
        self.on_close()

    def on_close(self):
        if getattr(self.parent_app, "settings_window", None) is self:
            self.parent_app.settings_window = None
        self.destroy()
