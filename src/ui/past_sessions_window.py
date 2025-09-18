"""Window for browsing previously saved sessions."""
import os
from datetime import datetime
from tkinter import messagebox, ttk

from customtkinter import CTkButton, CTkFrame, CTkLabel, CTkToplevel

from utils.helpers import MIN_PAST_SESSIONS_SIZE, SESSIONS_FOLDER, bring_window_to_front, ensure_initial_size

class PastSessionsWindow(CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Past Sessions")
        self.minsize(*MIN_PAST_SESSIONS_SIZE)
        self._paths = {}
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(50, lambda: bring_window_to_front(self))
        header = CTkLabel(
            self,
            text="Past Sessions",
            font=("Arial", 20, "bold")
        )
        header.pack(anchor="w", padx=24, pady=(24, 12))

        container = CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        columns = ("name", "modified", "size")
        self.tree = ttk.Treeview(container, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="Session")
        self.tree.column("name", anchor="w", width=280)
        self.tree.heading("modified", text="Last Modified")
        self.tree.column("modified", anchor="center", width=160)
        self.tree.heading("size", text="Size")
        self.tree.column("size", anchor="center", width=100)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.empty_label = CTkLabel(
            container,
            text="No session files found.",
            font=("Arial", 14)
        )
        self.empty_label.place_forget()

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _e: self._open_selected())

        button_bar = CTkFrame(self, fg_color="transparent")
        button_bar.pack(fill="x", padx=24, pady=(0, 24))
        button_bar.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="past_actions")

        self.open_btn = CTkButton(
            button_bar,
            text="Open in Scanner",
            state="disabled",
            command=self._open_selected
        )
        self.open_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.reveal_btn = CTkButton(
            button_bar,
            text="Reveal in Folder",
            state="disabled",
            command=self._reveal_selected
        )
        self.reveal_btn.grid(row=0, column=1, padx=6, sticky="ew")

        self.refresh_btn = CTkButton(button_bar, text="Refresh", command=self.refresh)
        self.refresh_btn.grid(row=0, column=2, padx=6, sticky="ew")

        self.clear_btn = CTkButton(
            button_bar,
            text="Clear All",
            state="disabled",
            command=self._clear_all_sessions
        )
        self.clear_btn.grid(row=0, column=3, padx=6, sticky="ew")

        self.close_btn = CTkButton(button_bar, text="Close", command=self._on_close)
        self.close_btn.grid(row=0, column=4, padx=(6, 0), sticky="ew")

        self.refresh()
        ensure_initial_size(self, min_size=MIN_PAST_SESSIONS_SIZE)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._paths.clear()
        if not os.path.isdir(SESSIONS_FOLDER):
            self._toggle_empty_state(True)
            self._on_select()
            self._update_clear_state()
            return
        files = []
        for entry in os.listdir(SESSIONS_FOLDER):
            path_entry = os.path.join(SESSIONS_FOLDER, entry)
            if os.path.isfile(path_entry) and entry.lower().endswith((".csv", ".xlsx")):
                stats = os.stat(path_entry)
                files.append((path_entry, stats.st_mtime, stats.st_size))
        files.sort(key=lambda item: item[1], reverse=True)
        for index, (path_entry, modified, size) in enumerate(files):
            name = os.path.splitext(os.path.basename(path_entry))[0]
            stamp = datetime.fromtimestamp(modified).strftime("%d %b %Y %H:%M")
            size_text = self._format_size(size)
            iid = f"past_{index}"
            self.tree.insert("", "end", iid=iid, values=(name, stamp, size_text))
            self._paths[iid] = path_entry
        self._toggle_empty_state(len(self._paths) == 0)
        self._on_select()
        self._update_clear_state()

    def _update_clear_state(self):
        state = "normal" if self._paths else "disabled"
        self.clear_btn.configure(state=state)

    def _toggle_empty_state(self, show):
        if show:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.empty_label.place_forget()

    def _format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"

    def _on_select(self, _event=None):
        selection = self.tree.selection()
        state = "normal" if selection else "disabled"
        self.open_btn.configure(state=state)
        self.reveal_btn.configure(state=state)

    def _get_selected_path(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return self._paths.get(selection[0])

    def _open_selected(self):
        path_entry = self._get_selected_path()
        if not path_entry:
            return
        opened = self.parent._open_session_path(path_entry, read_only=True)
        if opened:
            self.parent._refresh_recent_sessions()

    def _reveal_selected(self):
        path_entry = self._get_selected_path()
        if not path_entry:
            return
        self.parent._reveal_session_path(path_entry)

    def _clear_all_sessions(self):
        if not self._paths:
            return
        confirm = messagebox.askyesno(
            "Clear All Sessions",
            "Delete all saved session files?",
            parent=self
        )
        if not confirm:
            return
        failures = []
        for path_entry in list(self._paths.values()):
            try:
                os.remove(path_entry)
            except Exception as exc:
                failures.append(f"{os.path.basename(path_entry)}: {exc}")
        self.refresh()
        if hasattr(self.parent, "_refresh_recent_sessions"):
            try:
                self.parent._refresh_recent_sessions()
            except Exception:
                pass
        if failures:
            messagebox.showerror(
                "Delete Failed",
                "Some session files could not be deleted:\n" + "\n".join(failures),
                parent=self
            )
            if hasattr(self.parent, "set_status"):
                self.parent.set_status("Some past sessions could not be removed.")
            return
        messagebox.showinfo(
            "Sessions Cleared",
            "All past session files have been deleted.",
            parent=self
        )
        if hasattr(self.parent, "set_status"):
            self.parent.set_status("All past sessions cleared.")

    def _on_close(self):
        if hasattr(self.parent, "past_sessions_window"):
            self.parent.past_sessions_window = None
        self.destroy()
