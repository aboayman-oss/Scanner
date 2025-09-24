'''A simple password dialog.'''
import customtkinter as ctk

class PasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Password Required")
        self.transient(parent)
        self.grab_set()
        self.result = None

        self.label = ctk.CTkLabel(self, text="Enter the password to access settings:")
        self.label.pack(padx=20, pady=(20, 10))

        self.password_entry = ctk.CTkEntry(self, show="*")
        self.password_entry.pack(padx=20, pady=10)
        self.password_entry.bind("<Return>", self._on_ok)

        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(padx=20, pady=20)

        self.ok_button = ctk.CTkButton(self.button_frame, text="OK", command=self._on_ok)
        self.ok_button.pack(side="left", padx=10)

        self.cancel_button = ctk.CTkButton(self.button_frame, text="Cancel", command=self._on_cancel)
        self.cancel_button.pack(side="right", padx=10)

        self.password_entry.focus_set()
        self.wait_window(self)

    def _on_ok(self, event=None):
        self.result = self.password_entry.get()
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()

    def get_input(self):
        return self.result
