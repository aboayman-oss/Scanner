"""Application entry point for the RFID Attendance Manager."""
import customtkinter as ctk

from ui.main_window import App

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def main() -> None:
    """Launch the CustomTkinter application."""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
