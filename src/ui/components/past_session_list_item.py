# In ui/components/past_session_list_item.py
import os
from datetime import datetime
import customtkinter as ctk

class PastSessionListItem(ctk.CTkFrame):
    """A self-contained widget to display a single past session's info and actions."""
    def __init__(self, master, path, modified_timestamp, size_bytes, format_size_func, open_func, reveal_func):
        # Initialize the frame with a border
        super().__init__(master, corner_radius=8, border_width=1)

        # Store the path and callbacks
        self.path = path

        # --- Data Processing ---
        # Extract and format the data needed for display
        session_name = os.path.splitext(os.path.basename(path))[0]
        modified_str = datetime.fromtimestamp(modified_timestamp).strftime("%d %b %Y %H:%M")
        size_str = format_size_func(size_bytes)

        # --- Layout Configuration ---
        # Configure the grid to space out the elements nicely
        self.grid_columnconfigure(0, weight=3, uniform="group1") # Name (takes most space)
        self.grid_columnconfigure(1, weight=2, uniform="group1") # Modified Date
        self.grid_columnconfigure(2, weight=1, uniform="group1") # File Size
        self.grid_columnconfigure(3, weight=2, uniform="group1") # Action Buttons
        self.grid_rowconfigure(0, weight=1)

        # --- Create and Place Widgets ---
        # Information Labels
        name_label = ctk.CTkLabel(self, text=session_name, anchor="w", font=("Arial", 14, "bold"))
        name_label.grid(row=0, column=0, sticky="w", padx=15, pady=10)

        modified_label = ctk.CTkLabel(self, text=modified_str, anchor="w")
        modified_label.grid(row=0, column=1, sticky="w", padx=10)

        size_label = ctk.CTkLabel(self, text=size_str, anchor="w")
        size_label.grid(row=0, column=2, sticky="w", padx=10)

        # Action Buttons Frame (to group them on the right)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=0, column=3, sticky="e", padx=15)

        open_btn = ctk.CTkButton(
            btn_frame,
            text="Open",
            width=80,
            command=lambda: open_func(self.path, read_only=True)
        )
        open_btn.pack(side="left")

        reveal_btn = ctk.CTkButton(
            btn_frame,
            text="Reveal",
            width=80,
            fg_color="transparent",
            border_width=1,
            command=lambda: reveal_func(self.path)
        )
        reveal_btn.pack(side="left", padx=(10, 0))