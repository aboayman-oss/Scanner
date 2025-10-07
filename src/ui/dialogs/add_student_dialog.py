"""Dialog for manually adding students during a scan session."""
import customtkinter as ctk
from customtkinter import CTkButton, CTkEntry, CTkFrame, CTkLabel, CTkToplevel, CTkImage
from PIL import Image # Pillow is required for CTkImage

from utils.helpers import MIN_SUMMARY_SIZE, bring_window_to_front, ensure_initial_size, set_dark_title_bar

class AddStudentDialog(CTkToplevel):
    def __init__(self, parent, *, card_id=None, on_submit=None, duplicate_checker=None, default_notes="Manually added"):
        super().__init__(parent)
        set_dark_title_bar(self)
        self.parent = parent
        self.card_id = card_id.zfill(8) if card_id and card_id.isdigit() else card_id
        self._on_submit = on_submit
        self._duplicate_checker = duplicate_checker
        self._default_notes = default_notes or "Manually added"
        self._focus_guard_restored = False

        self.title("Add Student")
        self.minsize(*MIN_SUMMARY_SIZE)
        self.transient(parent)
        self.grab_set()
        self.after(40, self._activate_modal)

        # --- M3 Inspired UI Changes ---
        # Main container with more generous padding
        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=20)
        container.grid_columnconfigure(0, weight=1)

        # 1. Add an icon for visual context and simplify the header
        header_frame = CTkFrame(container, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        header_frame.grid_columnconfigure(0, weight=1)
        
        # Load the icon (ensure person_add.png is in a reachable path like 'assets/')
        try:
            add_icon_image = CTkImage(Image.open("assets/person_add.png"), size=(40, 40))
            icon_label = CTkLabel(header_frame, text="", image=add_icon_image)
            icon_label.grid(row=0, column=0, pady=(0, 12))
        except FileNotFoundError:
            print("Warning: 'assets/person_add.png' not found. Skipping icon.")

        subtitle_text = "Link a scanned card to a student profile." if self.card_id else "Create a manual record for a student."
        subtitle_label = CTkLabel(header_frame, text=subtitle_text, font=("Roboto", 16))
        subtitle_label.grid(row=1, column=0)
        
        if self.card_id:
            CTkLabel(
                header_frame,
                text=f"Card ID: {self.card_id}",
                font=("Roboto", 12, "bold"),
                text_color="#1f6aa5"
            ).grid(row=2, column=0, pady=(12, 0))

        # 2. Refined form with labels above entry fields
        form = CTkFrame(container, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew")
        form.grid_columnconfigure(0, weight=1)

        self.inputs = {}
        field_specs = [
            ("student_id", "Student ID", "e.g. 102534"),
            ("name", "Student Name", "Full name"),
            ("phone", "Phone Number", "e.g. 01012345678"),
        ]
        
        # Create fields with top-aligned labels
        for i, (key, label_text, placeholder) in enumerate(field_specs):
            label = CTkLabel(form, text=label_text, font=("Roboto", 13, "bold"))
            label.grid(row=i*2, column=0, sticky="w", pady=(10 if i > 0 else 0, 4))
            
            entry = CTkEntry(form, placeholder_text=placeholder, height=36) # Slightly taller entry
            entry.grid(row=i*2 + 1, column=0, sticky="ew")
            self.inputs[key] = entry

        # Feedback label remains for validation messages
        self.feedback_var = ctk.StringVar(value="")
        self.feedback_label = CTkLabel(
            container,
            textvariable=self.feedback_var,
            font=("Roboto", 12),
            text_color="#d64b4b",
            wraplength=350 # Prevent feedback from making window too wide
        )
        self.feedback_label.grid(row=2, column=0, sticky="w", pady=(16, 0))

        # 3. Differentiated Primary and Secondary Actions, centered layout
        actions = CTkFrame(container, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(24, 0))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(3, weight=1)

        # Secondary action button (Text Button style)
        self.cancel_button = CTkButton(
            actions,
            text="Cancel",
            command=self._on_cancel,
            fg_color="transparent",
            hover_color="#363a45",
            text_color="#f8fafc"
        )
        self.cancel_button.grid(row=0, column=1, padx=(0, 12))

        # Primary action button (Filled Button style)
        self.confirm_button = CTkButton(actions, text="Add Student", command=self._on_confirm)
        self.confirm_button.grid(row=0, column=2)

        ensure_initial_size(self, min_size=MIN_SUMMARY_SIZE)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._first_entry = next(iter(self.inputs.values()), None)
        self.bind("<Return>", lambda _event: self._on_confirm())
        self.bind("<Escape>", lambda _event: self._on_cancel())

    # --- Methods (_set_feedback, _activate_modal, etc.) remain unchanged ---

    def _set_feedback(self, message, level="error"):
        colors = {
            "error": "#d64b4b",
            "warning": "#d67b2d",
            "info": "#1f6aa5",
        }
        color = colors.get(level, "#d64b4b")
        self.feedback_label.configure(text_color=color)
        self.feedback_var.set(message)

    def _activate_modal(self):
        bring_window_to_front(self)
        try:
            self.grab_set()
        except Exception:
            pass
        try:
            self.focus_force()
        except Exception:
            pass
        if self._first_entry is not None and self._first_entry.winfo_exists():
            self._first_entry.focus_set()

    def _on_confirm(self):
        values = {key: entry.get().strip() for key, entry in self.inputs.items()}
        if not all(values.values()):
            self._set_feedback("Please fill in every field before continuing.", level="warning")
            return
        if self._duplicate_checker:
            try:
                id_exists, phone_exists = self._duplicate_checker(values["student_id"], values["phone"])
            except Exception as exc:
                self._set_feedback(str(exc))
                return
            if id_exists or phone_exists:
                if id_exists and phone_exists:
                    msg = "This student ID and phone number already exist."
                elif id_exists:
                    msg = "This student ID already exists."
                else:
                    msg = "This phone number already exists."
                self._set_feedback(msg, level="warning")
                return
        if not self._on_submit:
            self._finalize()
            return
        try:
            outcome = self._on_submit(card_id=self.card_id, values=values, default_notes=self._default_notes)
        except Exception as exc:
            self._set_feedback(str(exc))
            return
        if outcome is None:
            success, message = True, None
        elif isinstance(outcome, tuple):
            success, message = outcome
        else:
            success, message = bool(outcome), None
        if success:
            self._set_feedback("")
            self._finalize()
        else:
            self._set_feedback(message or "Unable to add student.", level="error")

    def _finalize(self):
        try:
            self.grab_release()
        except Exception:
            pass
        if not getattr(self, "_focus_guard_restored", False):
            if hasattr(self.parent, "_resume_focus_guard"):
                self.parent._resume_focus_guard()
            self._focus_guard_restored = True
        if self.winfo_exists():
            self.destroy()
        if hasattr(self.parent, "scan_entry"):
            self.parent.after(120, self.parent.scan_entry.focus_set)

    def _on_cancel(self):
        self._set_feedback("")
        self._finalize()
