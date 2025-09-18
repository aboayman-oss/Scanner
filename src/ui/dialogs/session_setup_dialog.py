"""Dialog for configuring basic session metadata."""
import customtkinter as ctk
from customtkinter import CTkButton, CTkComboBox, CTkEntry, CTkFrame, CTkLabel, CTkToplevel

from utils.helpers import MIN_SESSION_SETUP_SIZE, bring_window_to_front, ensure_initial_size

class SessionSetupDialog(CTkToplevel):
    def __init__(self, parent, stages, centers, has_data, callback):
        super().__init__(parent)
        self.parent = parent
        self.stages = stages
        self.centers = centers
        self.callback = callback
        self.has_data = has_data
        self.title("Start New Session")
        self.resizable(False, False)
        self.minsize(*MIN_SESSION_SETUP_SIZE)
        self.transient(parent)
        self.grid_columnconfigure(1, weight=1)

        notice_text = (
            "Using the imported dataset for this session."
            if has_data
            else "No dataset imported yet. A blank roster will be created."
        )
        self.notice_var = ctk.StringVar(value=notice_text)
        self.error_var = ctk.StringVar(value="")

        self._build_form()
        ensure_initial_size(self, min_size=MIN_SESSION_SETUP_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Return>", lambda _e: self._on_submit())
        self.bind("<Escape>", lambda _e: self._on_cancel())
        self.after(10, self._center_on_parent)
        self.after(50, lambda: self.stage_cb.focus_set())

    def _build_form(self):
        CTkLabel(
            self,
            textvariable=self.notice_var,
            justify="left",
            anchor="w",
            wraplength=320
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 8))

        CTkLabel(self, text="Stage:").grid(row=1, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.stage_cb = CTkComboBox(self, values=self.stages, state="readonly")
        self.stage_cb.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(self, text="Center:").grid(row=2, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.center_cb = CTkComboBox(self, values=self.centers, state="readonly")
        self.center_cb.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(self, text="Session No.:").grid(row=3, column=0, sticky="w", padx=(16, 8), pady=(0, 4))
        self.session_ent = CTkEntry(self)
        self.session_ent.grid(row=3, column=1, sticky="ew", padx=(0, 16), pady=(0, 4))

        CTkLabel(
            self,
            textvariable=self.error_var,
            text_color=("#ff6b6b", "#b00020"),
            justify="left",
            anchor="w",
            wraplength=320
        ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=16, pady=(4, 0))

        btn_frame = CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(12, 16))
        CTkButton(btn_frame, text="Start Session", command=self._on_submit).pack(side="left", padx=(0, 8))
        CTkButton(btn_frame, text="Cancel", command=self._on_cancel).pack(side="left")

        if self.stages:
            self.stage_cb.set(self.stages[0])
        if self.centers:
            self.center_cb.set(self.centers[0])

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

    def _on_submit(self):
        stage = self.stage_cb.get().strip()
        center = self.center_cb.get().strip()
        session_no = self.session_ent.get().strip()
        if not stage or not center or not session_no.isdigit():
            self.error_var.set("Select stage, center, and enter a numeric session number.")
            return
        self.error_var.set("")
        payload = {
            "stage": stage,
            "center": center,
            "no": int(session_no),
            "name": f"{stage} {center} session {int(session_no)}"
        }
        if self.callback:
            self.callback(payload)
            self.callback = None
        self.destroy()

    def _on_cancel(self):
        if self.callback:
            self.callback(None)
            self.callback = None
        self.destroy()
