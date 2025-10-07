"""Custom drop-down widget with animated menu and modern styling."""

from __future__ import annotations

import tkinter as tk
from typing import Callable, List, Optional, Sequence, Union

import customtkinter as ctk

from utils.helpers import set_dark_title_bar


class ModernDropdown(ctk.CTkFrame):
    """Composite drop-down control with smooth animations and modern styling."""

    _ANIMATION_STEPS = 8
    _OPEN_DURATION_MS = 180
    _CLOSE_DURATION_MS = 140
    _DROPDOWN_OFFSET_Y = 6
    _SHADOW_OFFSET_Y = 4

    def __init__(
        self,
        master,
        values: Sequence[str],
        placeholder: str,
        *,
        command: Optional[Callable[[str], None]] = None,
        font: Optional[ctk.CTkFont] = None,
        text_color: tuple[str, str] = ("#1E293B", "#E2E8F0"),
        placeholder_color: tuple[str, str] = ("#64748B", "#94A3B8"),
        base_fg_color: tuple[str, str] = ("#E8EDF6", "#2C3039"),
        hover_fg_color: tuple[str, str] = ("#F3F6FF", "#353C47"),
        border_color: tuple[str, str] = ("#CBD5E1", "#3D4452"),
        dropdown_bg_color: tuple[str, str] = ("#FFFFFF", "#1F242C"),
        option_hover_color: tuple[str, str] = ("#E1E9F9", "#3A4352"),
        active_option_color: tuple[str, str] = ("#D6E2FB", "#3D485B"),
        icon_color: tuple[str, str] = ("#4B5563", "#A0AEC0"),
        shadow_color: str = "#000000",
        option_height: int = 30,
        max_visible_items: Optional[int] = 6,
        max_dropdown_height: Optional[int] = 280,
    ):
        super().__init__(master, corner_radius=16, border_width=1)

        provided_values = list(values) if values is not None else []

        self._placeholder = placeholder
        self._values: List[str] = [v for v in provided_values if v != placeholder]
        self._command = command
        self._font = font or ctk.CTkFont(size=13)
        self._text_color = text_color
        self._placeholder_color = placeholder_color
        self._base_fg_color = base_fg_color
        self._hover_fg_color = hover_fg_color
        self._border_color = border_color
        self._dropdown_bg_color = dropdown_bg_color
        self._option_hover_color = option_hover_color
        self._active_option_color = active_option_color
        self._icon_color = icon_color
        self._shadow_color = shadow_color
        self._option_height = option_height
        self._max_visible_items = max_visible_items if (max_visible_items and max_visible_items > 0) else 0
        self._max_dropdown_height = max_dropdown_height if (max_dropdown_height and max_dropdown_height > 0) else None
        self._content_padding = 48

        self._current_value: Optional[str] = None
        self._hovered = False
        self._measured_width = 0

        self._dropdown_window: Optional[ctk.CTkToplevel] = None
        self._shadow_window: Optional[tk.Toplevel] = None
        self._dropdown_container: Optional[ctk.CTkFrame] = None
        self._options_frame: Optional[Union[ctk.CTkFrame, ctk.CTkScrollableFrame]] = None
        self._option_buttons: List[ctk.CTkButton] = []
        self._global_click_bind_id: Optional[str] = None
        self._animation_after_id: Optional[str] = None
        self._target_geometry: Optional[tuple[int, int, int, int]] = None
        self._closing = False

        self.configure(fg_color=self._base_fg_color, border_color=self._border_color)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._value_var = ctk.StringVar(value=self._placeholder)
        self._label = ctk.CTkLabel(
            self,
            textvariable=self._value_var,
            font=self._font,
            anchor="w",
            text_color=self._placeholder_color,
        )
        self._label.grid(row=0, column=0, sticky="nsew", padx=(16, 8), pady=10)

        self._arrow_canvas = ctk.CTkCanvas(self, width=18, height=18, highlightthickness=0)
        self._arrow_canvas.grid(row=0, column=1, padx=(0, 16))
        self._draw_arrow()

        for widget in (self, self._label, self._arrow_canvas):
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<Button-1>", self._on_press)

        self.bind("<Configure>", self._on_resize)

        self._apply_base_color()
        self._apply_text_color()

    # Public API ---------------------------------------------------------

    def set(self, value: str) -> None:
        if not value or value == self._placeholder:
            self._current_value = None
            self._value_var.set(self._placeholder)
        else:
            if value not in self._values:
                self._values.append(value)
            self._current_value = value
            self._value_var.set(value)
        self._apply_text_color()
        self._update_option_states()

    def get(self) -> str:
        return self._placeholder if self._current_value is None else self._current_value

    def configure(self, require_redraw: bool = False, **kwargs):
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "values" in kwargs:
            updated_values = kwargs.pop("values") or []
            self._values = [v for v in updated_values if v != self._placeholder]
            if self._current_value and self._current_value not in self._values:
                self._current_value = None
                self._value_var.set(self._placeholder)
            self._apply_text_color()
            self._update_option_states()
            if self._dropdown_window:
                self._close_dropdown(animated=False)
        super().configure(require_redraw=require_redraw, **kwargs)

    def cget(self, attribute_name: str):
        if attribute_name == "values":
            return [self._placeholder, *self._values]
        if attribute_name == "command":
            return self._command
        return super().cget(attribute_name)

    # Event handlers -----------------------------------------------------

    def _on_enter(self, _event) -> None:
        self._hovered = True
        self._apply_base_color()

    def _on_leave(self, _event) -> None:
        self._hovered = False
        self._apply_base_color()

    def _on_press(self, _event) -> None:
        if self._dropdown_window:
            self._close_dropdown(animated=True)
        else:
            self._open_dropdown()

    def _on_resize(self, _event) -> None:
        width = self.winfo_width()
        if width > 0:
            self._measured_width = width
        if self._dropdown_window and self._target_geometry:
            _, height, _, _ = self._target_geometry
            self._target_geometry = self._compute_target_geometry(width, height)
            self._apply_geometry()

    # Appearance helpers -------------------------------------------------

    def _apply_base_color(self) -> None:
        target = self._hover_fg_color if (self._hovered or self._dropdown_window) else self._base_fg_color
        self.configure(fg_color=target, border_color=self._border_color)
        self._refresh_arrow()

    def _apply_text_color(self) -> None:
        color = self._placeholder_color if self._current_value is None else self._text_color
        self._label.configure(text_color=color)

    def _refresh_arrow(self) -> None:
        background = self._hover_fg_color if (self._hovered or self._dropdown_window) else self._base_fg_color
        self._arrow_canvas.configure(bg=self._resolve_color(background))
        self._draw_arrow()

    def _draw_arrow(self) -> None:
        self._arrow_canvas.delete("arrow")
        shade = self._resolve_color(self._icon_color)
        self._arrow_canvas.create_polygon(
            4,
            6,
            9,
            12,
            14,
            6,
            fill=shade,
            outline=shade,
            tags="arrow",
        )

    def _resolve_color(self, value):
        if isinstance(value, (list, tuple)):
            return value[0] if ctk.get_appearance_mode().lower() == "light" else value[1]
        return value

    def _determine_content_width(self) -> int:
        text_samples = [self._placeholder, *self._values]
        if not text_samples:
            return self._measured_width
        measure = getattr(self._font, "measure", None)
        if callable(measure):
            # measure() returns a float, but we need to ensure it's convertible to int
            measurements = []
            for sample in text_samples:
                result = measure(sample)
                if result is not None and isinstance(result, (int, float)):
                    measurements.append(int(float(result)))
            longest = max(measurements, default=0)
        else:
            longest = max([len(sample) for sample in text_samples], default=0) * 7
        return longest + self._content_padding

    def _calculate_total_height(self, item_count: int) -> int:
        if item_count <= 0:
            return self._option_height + 16
        padding = 10 + max(0, item_count - 1) * 5
        return item_count * self._option_height + padding

    def _calculate_visible_height(self, total_height: int) -> int:
        candidates = [total_height]
        if self._max_dropdown_height:
            candidates.append(self._max_dropdown_height)
        if self._max_visible_items:
            candidates.append(self._calculate_total_height(self._max_visible_items))
        visible = min(candidates) if candidates else total_height
        return max(visible, self._option_height + 10)

    # Dropdown lifecycle -------------------------------------------------

    def _open_dropdown(self) -> None:
        if self._dropdown_window or not self._values:
            return

        self.update_idletasks()
        width = max(
            self._measured_width,
            self.winfo_width(),
            self.winfo_reqwidth(),
            self._determine_content_width(),
        )

        total_height = self._calculate_total_height(len(self._values))
        visible_height = self._calculate_visible_height(total_height)
        needs_scroll = total_height > visible_height
        container_height = visible_height + (12 if needs_scroll else 6)

        self._shadow_window = tk.Toplevel(self)
        self._shadow_window.withdraw()
        self._shadow_window.overrideredirect(True)
        self._shadow_window.attributes("-alpha", 0.0)
        self._shadow_window.configure(bg=self._shadow_color)

        self._dropdown_window = ctk.CTkToplevel(self)
        set_dark_title_bar(self._dropdown_window)
        self._dropdown_window.withdraw()
        self._dropdown_window.overrideredirect(True)
        self._dropdown_window.transient(self.winfo_toplevel())
        self._dropdown_window.attributes("-alpha", 0.0)
        self._dropdown_window.configure(fg_color=self._dropdown_bg_color)

        self._dropdown_container = ctk.CTkFrame(
            self._dropdown_window,
            fg_color=self._dropdown_bg_color,
            corner_radius=14,
            border_width=1,
            border_color=self._border_color,
        )
        self._dropdown_container.pack(fill="both", expand=True, padx=2, pady=2)
        self._dropdown_container.pack_propagate(False)
        self._dropdown_container.grid_columnconfigure(0, weight=1)

        frame_width = max(width - 12, 1)
        if needs_scroll:
            frame_height = max(visible_height, self._option_height + 12)
            self._options_frame = ctk.CTkScrollableFrame(
                self._dropdown_container,
                fg_color="transparent",
                width=frame_width,
                height=frame_height,
            )
            frame_expand = True
        else:
            self._options_frame = ctk.CTkFrame(
                self._dropdown_container,
                fg_color="transparent",
                width=frame_width,
            )
            frame_expand = False

        if self._options_frame:
            self._options_frame.pack(fill="both", expand=frame_expand, padx=6, pady=6)
            self._options_frame.grid_columnconfigure(0, weight=1)
            if not needs_scroll:
                self._options_frame.configure(width=frame_width)

        self._option_buttons = []
        if self._options_frame:
            for index, option in enumerate(self._values):
                button = ctk.CTkButton(
                    self._options_frame,
                    text=option,
                    anchor="w",
                    height=self._option_height,
                    font=self._font,
                    fg_color=self._active_option_color if option == self._current_value else "transparent",
                    hover_color=self._option_hover_color,
                    text_color=self._text_color,
                    corner_radius=10,
                    border_width=0,
                    command=lambda value=option: self._select(value),
                )
                pady = (4 if index == 0 else 2, 4 if index == len(self._values) - 1 else 2)
                button.grid(row=index, column=0, sticky="ew", padx=10, pady=pady)
                self._option_buttons.append(button)

        self._dropdown_window.update_idletasks()
        if self._dropdown_container:
            actual_height = max(container_height, self._dropdown_container.winfo_reqheight())
            self._dropdown_container.configure(width=width, height=actual_height)
        else:
            actual_height = container_height
        self._target_geometry = self._compute_target_geometry(width, actual_height)

        self._apply_geometry(initial=True)

        if self._shadow_window:
            self._shadow_window.deiconify()
        self._dropdown_window.deiconify()
        if self._shadow_window:
            self._shadow_window.lift()
        self._dropdown_window.lift()

        self._closing = False
        self._register_global_click()
        self._dropdown_window.bind("<Escape>", lambda _e: self._close_dropdown(animated=True))

        self._run_open_animation(0)
        self._apply_base_color()

        # Bind mouse wheel events to prevent background scrolling
        if self._options_frame:
            widgets_to_bind = [self._options_frame, self._dropdown_container, self._dropdown_window]
            if isinstance(self._options_frame, ctk.CTkScrollableFrame):
                widgets_to_bind.append(self._options_frame._parent_canvas)
                vbar = getattr(self._options_frame, "_vbar", None)
                if vbar:
                    widgets_to_bind.append(vbar)

            for widget in widgets_to_bind:
                if widget:
                    widget.bind("<MouseWheel>", self._on_mouse_wheel, add="+")

            for button in self._option_buttons:
                button.bind("<MouseWheel>", self._on_mouse_wheel, add="+")

    def _close_dropdown(self, *, animated: bool) -> None:
        if not self._dropdown_window:
            return
        if self._animation_after_id:
            self.after_cancel(self._animation_after_id)
            self._animation_after_id = None
        self._closing = True
        if animated:
            self._run_close_animation(0)
        else:
            self._destroy_dropdown()

    def _destroy_dropdown(self) -> None:
        if self._dropdown_window:
            self._dropdown_window.destroy()
            self._dropdown_window = None
        if self._shadow_window:
            self._shadow_window.destroy()
            self._shadow_window = None
        self._dropdown_container = None
        self._options_frame = None
        self._option_buttons = []
        self._target_geometry = None
        self._closing = False
        self._unregister_global_click()
        self._apply_base_color()

    def _compute_target_geometry(self, width: int, height: int) -> tuple[int, int, int, int]:
        width = max(width, 1)
        height = max(height, 1)
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + self._DROPDOWN_OFFSET_Y
        return (width, height, x, y)

    def _apply_geometry(self, initial: bool = False) -> None:
        if not self._target_geometry:
            return
        width, height, x, y = self._target_geometry
        if self._shadow_window:
            shadow_height = height
            self._shadow_window.geometry(f"{width}x{shadow_height}+{x}+{y + self._SHADOW_OFFSET_Y}")
        if self._dropdown_window:
            if initial:
                start_height = max(1, int(height * 0.92))
                self._dropdown_window.geometry(f"{width}x{start_height}+{x}+{y}")
            else:
                self._dropdown_window.geometry(f"{width}x{height}+{x}+{y}")

    # Animation ----------------------------------------------------------

    def _run_open_animation(self, step: int) -> None:
        if not self._dropdown_window or not self._target_geometry:
            return
        steps = self._ANIMATION_STEPS
        progress = min(step / (steps - 1), 1.0)
        width, height, x, y = self._target_geometry
        current_height = max(1, int(height * (0.92 + 0.08 * progress)))
        alpha = 0.1 + 0.9 * progress

        self._dropdown_window.geometry(f"{width}x{current_height}+{x}+{y}")
        self._dropdown_window.attributes("-alpha", alpha)
        if self._shadow_window:
            self._shadow_window.attributes("-alpha", 0.18 * progress)

        if progress < 1.0:
            delay = max(12, self._OPEN_DURATION_MS // self._ANIMATION_STEPS)
            self._animation_after_id = self.after(delay, self._run_open_animation, step + 1)
        else:
            self._animation_after_id = None
            self._dropdown_window.geometry(f"{width}x{height}+{x}+{y}")
            self._dropdown_window.attributes("-alpha", 1.0)
            if self._shadow_window:
                self._shadow_window.attributes("-alpha", 0.18)

    def _run_close_animation(self, step: int) -> None:
        if not self._dropdown_window or not self._target_geometry:
            self._destroy_dropdown()
            return
        steps = self._ANIMATION_STEPS
        progress = min(step / (steps - 1), 1.0)
        width, height, x, y = self._target_geometry
        current_height = max(1, int(height * (0.92 + 0.08 * (1.0 - progress))))
        alpha = max(0.0, 1.0 - progress)

        self._dropdown_window.geometry(f"{width}x{current_height}+{x}+{y}")
        self._dropdown_window.attributes("-alpha", alpha)
        if self._shadow_window:
            self._shadow_window.attributes("-alpha", max(0.0, 0.18 * (1.0 - progress)))

        if progress < 1.0:
            delay = max(12, self._CLOSE_DURATION_MS // self._ANIMATION_STEPS)
            self._animation_after_id = self.after(delay, self._run_close_animation, step + 1)
        else:
            self._animation_after_id = None
            self._destroy_dropdown()

    # Selection ----------------------------------------------------------

    def _select(self, value: str) -> None:
        self._current_value = value
        self._value_var.set(value)
        self._apply_text_color()
        self._update_option_states()
        self._close_dropdown(animated=True)
        if self._command:
            self._command(value)

    def _update_option_states(self) -> None:
        if not self._option_buttons:
            return
        for button in self._option_buttons:
            if button.cget("text") == self._current_value:
                button.configure(fg_color=self._active_option_color)
            else:
                button.configure(fg_color="transparent")

    # Global click handling ----------------------------------------------

    def _register_global_click(self) -> None:
        if self._global_click_bind_id:
            return
        root = self.winfo_toplevel()
        self._global_click_bind_id = root.bind("<Button-1>", self._handle_global_click, add="+")

    def _unregister_global_click(self) -> None:
        if not self._global_click_bind_id:
            return
        root = self.winfo_toplevel()
        root.unbind("<Button-1>", self._global_click_bind_id)
        self._global_click_bind_id = None

    def _handle_global_click(self, event) -> None:
        if not self._dropdown_window:
            return
        widget = event.widget
        try:
            if widget.winfo_toplevel() is self._dropdown_window:
                return
        except tk.TclError:
            pass
        if self._is_descendant(widget, self):
            return
        self.after(0, lambda: self._close_dropdown(animated=True))

    def _on_mouse_wheel(self, event) -> str:
        if not self._options_frame or not isinstance(self._options_frame, ctk.CTkScrollableFrame):
            return "break"  # Block scroll even if not scrollable frame

        scrollable_frame = self._options_frame
        start, end = scrollable_frame._parent_canvas.yview()

        # If the scrollable area is not at its boundary, scroll it.
        if (event.delta > 0 and start > 0.0) or (event.delta < 0 and end < 1.0):
            scrollable_frame._parent_canvas.yview_scroll(int(-1 * (event.delta / 20)), "units")

        # Always stop the event from propagating to the parent, even at boundaries.
        return "break"

    @staticmethod
    def _is_descendant(widget, ancestor) -> bool:
        current = widget
        while current is not None:
            if current is ancestor:
                return True
            current = getattr(current, "master", None)
        return False

    # Cleanup ------------------------------------------------------------

    def destroy(self) -> None:
        self._close_dropdown(animated=False)
        super().destroy()