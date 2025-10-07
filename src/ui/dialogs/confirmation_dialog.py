"""A simple, reusable confirmation dialog."""
import customtkinter as ctk
from customtkinter import CTkButton, CTkFrame, CTkLabel, CTkToplevel

from utils.helpers import set_dark_title_bar


class ConfirmationDialog(CTkToplevel):
    """A modal dialog to ask for user confirmation."""

    def __init__(self, parent, title="Confirm", message="", confirm_text="OK", cancel_text="Cancel"):
        super().__init__(parent)
        set_dark_title_bar(self)
        self.transient(parent)
        self.grab_set()
        self.title(title)
        self.result = None  # Store the result: True for confirm, False for cancel

        self.geometry("380x150")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        main_frame = CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=24, pady=20)

        msg_label = CTkLabel(main_frame, text=message, font=("Roboto", 14), wraplength=330)
        msg_label.pack(fill="x", expand=True)

        btn_frame = CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        self.cancel_btn = CTkButton(
            btn_frame,
            text=cancel_text,
            command=self._on_cancel,
            fg_color="transparent",
            border_width=1,
            border_color="#888"
        )
        self.cancel_btn.grid(row=0, column=0, sticky="e", padx=(0, 8))

        self.confirm_btn = CTkButton(btn_frame, text=confirm_text, command=self._on_confirm)
        self.confirm_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.after(50, self.lift) # Ensure it's on top

    def _on_confirm(self):
        self.result = True
        self.destroy()

    def _on_cancel(self):
        self.result = False
        self.destroy()

    def get_result(self):
        self.wait_window()
        return self.result