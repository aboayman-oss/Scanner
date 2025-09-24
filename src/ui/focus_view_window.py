import os
from PIL import Image
import customtkinter as ctk
from customtkinter import CTkFrame, CTkLabel, CTkTextbox, CTkButton, CTkToplevel

# Import constants from scan_window or define them here if needed
# from .scan_window import LIGHT_BG, DARK_BG, LIGHT_PRIMARY_TEXT, DARK_PRIMARY_TEXT, ...

class FocusViewWindow:
    def __init__(self, parent, read_only=False, icon_cache=None,
                 on_complete=None, on_add_student=None, on_override=None, on_deny=None, on_cancel=None, on_dismiss=None):
        self.parent = parent
        self.read_only = read_only
        self._icon_cache = icon_cache if icon_cache is not None else {}
        self._on_complete = on_complete
        self._on_add_student = on_add_student
        self._on_override = on_override
        self._on_deny = on_deny
        self._on_cancel = on_cancel
        self._on_dismiss = on_dismiss
        self._setup_ui()

    def _load_icon(self, name, size=(24, 24)):
        if (name, size) in self._icon_cache:
            return self._icon_cache[(name, size)]
        try:
            img_path = os.path.join('assets', name)
            img = Image.open(img_path)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=size)
            self._icon_cache[(name, size)] = ctk_img
            return ctk_img
        except FileNotFoundError:
            return ctk.CTkImage(light_image=Image.new("RGBA", size, (0,0,0,0)),
                                dark_image=Image.new("RGBA", size, (0,0,0,0)),
                                size=size)

    def _setup_ui(self):
        parent = self.parent
        parent.configure(fg_color="#222222")

        status_zone = CTkFrame(parent, fg_color="transparent")
        status_zone.pack(fill="x", padx=20, pady=(20, 12))
        status_zone.grid_columnconfigure(1, weight=1)

        self.name_label = CTkLabel(status_zone, text="", font=("Roboto", 32, "bold"), anchor="w")
        self.name_label.grid(row=0, column=1, sticky="w")

        self.id_label = CTkLabel(status_zone, text="", font=("Roboto", 12), anchor="w")
        self.id_label.grid(row=1, column=1, sticky="w", pady=(0, 8))

        self.status_icon = CTkLabel(status_zone, text="")
        self.status_icon.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))

        dismiss_icon = self._load_icon("close.png", size=(20, 20))
        self.btn_dismiss = CTkButton(status_zone, text="", image=dismiss_icon, width=32, height=32, fg_color="transparent", command=self._on_dismiss)
        self.btn_dismiss.grid(row=0, column=2, sticky="ne")

        details_zone = CTkFrame(parent, fg_color="transparent")
        details_zone.pack(fill="both", expand=True, padx=20, pady=8)

        self.hw_card = CTkFrame(details_zone, corner_radius=12)
        self.hw_card.pack(fill="x", pady=(0, 8))
        self.hw_card.grid_columnconfigure(1, weight=1)
        self.hw_icon_label = CTkLabel(self.hw_card, text="")
        self.hw_icon_label.grid(row=0, column=0, padx=(12, 8), pady=12)
        CTkLabel(self.hw_card, text="Homework", font=("Roboto", 16, "bold")).grid(row=0, column=1, sticky="w")
        self.hw_grade_label = CTkLabel(self.hw_card, text="", font=("Roboto", 14))
        self.hw_grade_label.grid(row=0, column=2, sticky="e", padx=12)

        self.exam_card = CTkFrame(details_zone, corner_radius=12)
        self.exam_card.pack(fill="x")
        self.exam_card.grid_columnconfigure(1, weight=1)
        self.exam_icon_label = CTkLabel(self.exam_card, text="")
        self.exam_icon_label.grid(row=0, column=0, padx=(12, 8), pady=12)
        CTkLabel(self.exam_card, text="Exam", font=("Roboto", 16, "bold")).grid(row=0, column=1, sticky="w")
        self.exam_grade_label = CTkLabel(self.exam_card, text="", font=("Roboto", 14))
        self.exam_grade_label.grid(row=0, column=2, sticky="e", padx=12)

        self.notes = CTkTextbox(details_zone, corner_radius=12, border_width=0, font=("Roboto", 22, "bold"), text_color="#FFFFFF")
        self.notes.pack(fill="both", expand=True, pady=(12, 0))
        self.notes.insert("1.0", "Add notes here...")
        self.notes.bind("<FocusIn>", self._on_notes_focus_in)
        self.notes.bind("<FocusOut>", self._on_notes_focus_out)

        actions_zone = CTkFrame(parent, fg_color="transparent")
        actions_zone.pack(fill="x", padx=20, pady=(12, 20))

        self.btn_complete = CTkButton(actions_zone, text="Complete & Attend", image=self._load_icon("task_alt.png"), command=self._on_complete)
        self.btn_add_student = CTkButton(actions_zone, text="Add New Student", image=self._load_icon("person_add.png"), command=self._on_add_student)
        self.btn_override = CTkButton(actions_zone, text="Attend Anyway", image=self._load_icon("gpp_good.png"), command=self._on_override)
        self.btn_deny = CTkButton(actions_zone, text="Deny Entry", image=self._load_icon("block.png"), command=self._on_deny)
        self.btn_cancel = CTkButton(actions_zone, text="Cancel Attendance", image=self._load_icon("block.png"), command=self._on_cancel)

        self.buttons = [self.btn_complete, self.btn_add_student, self.btn_override, self.btn_deny, self.btn_cancel]
        # 1. Configure a 3-column grid with equal weight
        actions_zone.grid_columnconfigure((0, 1, 2), weight=1)

        # 2. Place each button in its specific grid layout cell
        #    These will be hidden/shown by the logic in scan_window.py
        self.btn_deny.grid(row=0, column=0, sticky="ew", padx=2)
        self.btn_override.grid(row=0, column=1, sticky="ew", padx=2)
        self.btn_complete.grid(row=0, column=2, sticky="ew", padx=2)
        self.btn_add_student.grid(row=0, column=0, columnspan=3, sticky="ew", padx=4)
        self.btn_cancel.grid(row=0, column=0, columnspan=3, sticky="ew", padx=4)

        # 3. Hide them all initially
        for btn in self.buttons:
            btn.grid_remove()


        if self.read_only:
            self.notes.configure(state="disabled")
            for btn in self.buttons:
                btn.configure(state="disabled")

    def _on_notes_focus_in(self, event):
        if self.notes.get("1.0", "end-1c") == "Add notes here...":
            self.notes.delete("1.0", "end")
            self.notes.configure(text_color="#FFFFFF")

    def _on_notes_focus_out(self, event):
        if not self.notes.get("1.0", "end-1c"):
            self.notes.configure(text_color="gray")
            self.notes.insert("1.0", "Add notes here...")
