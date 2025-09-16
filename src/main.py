import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import pandas as pd
from datetime import datetime

# --- Location Check ---
expected_folder = "Session scanner"
base_path = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
expected_asset = os.path.join(base_path, "assets", "logo.png")

# Only enforce install location for the packaged build; allow scripts to run from anywhere
if getattr(sys, 'frozen', False):
    if os.path.basename(base_path) != expected_folder or not os.path.exists(expected_asset):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Wrong Location",
            "Return the app to its original location in the Session scanner folder and create a shortcut instead."
        )
        sys.exit(1)
# ----------------------

# Directories & Files
if getattr(sys, 'frozen', False):
    # Running as compiled exe
    BASE_FOLDER = os.path.dirname(sys.executable)
else:
    # Running as script; fall back to parent directory if assets sit there
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate_base = script_dir
    assets_path = os.path.join(candidate_base, "assets", "logo.png")
    if not os.path.exists(assets_path):
        parent_dir = os.path.dirname(script_dir)
        parent_asset = os.path.join(parent_dir, "assets", "logo.png")
        if os.path.exists(parent_asset):
            candidate_base = parent_dir
    BASE_FOLDER = candidate_base

ASSETS_DIR       = os.path.join(BASE_FOLDER, "assets")
LOGO_FILE        = os.path.join(ASSETS_DIR, "logo.png")
HOME_BG_FILE     = os.path.join(ASSETS_DIR, "background.jpg")
SETTINGS_BG_FILE = os.path.join(ASSETS_DIR, "backgroundnew.jpg")

DATA_FOLDER      = os.path.join(BASE_FOLDER, "Data")
SESSIONS_FOLDER  = os.path.join(BASE_FOLDER, "Sessions")
ARCHIVE_FOLDER   = os.path.join(BASE_FOLDER, "Data archive")
MAPPING_FILE     = os.path.join(ARCHIVE_FOLDER, "column_map.json")
SETTINGS_FILE    = os.path.join(ARCHIVE_FOLDER, "app_settings.json")
LAST_DATA_FILE   = os.path.join(ARCHIVE_FOLDER, "last_data.json")

for folder in (DATA_FOLDER, SESSIONS_FOLDER, ARCHIVE_FOLDER):
    os.makedirs(folder, exist_ok=True)

SETTINGS = {
    "stage_options":  ["2nd", "3rd"],
    "center_options": [
        "October", "Ferdous", "Helwan", "Hadayek Helwan",
        "Zayed", "Haram", "Dokki", "Maadi", "15 May"
    ],
    "restrictions": {"exam": False, "homework": False},
    "file_type": "xlsx"
}

class ColumnMappingDialog(tk.Toplevel):
    def __init__(self, parent, df_columns, current_mapping=None):
        super().__init__(parent)
        self.title("Configure Column Template")
        self.geometry("360x380")
        self.df_columns      = df_columns
        self.current_mapping = current_mapping or {}
        self.widgets         = {}
        self.mapping         = {}
        self.build_ui()
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)

    def build_ui(self):
        ttk.Label(self, text="Map each field to an excel column").pack(pady=10)
        def add_row(label, key):
            frm = ttk.Frame(self); frm.pack(fill="x", padx=20, pady=4)
            ttk.Label(frm, text=f"{label}:", width=12).pack(side="left")
            cb = ttk.Combobox(frm, values=self.df_columns, state="readonly")
            cb.pack(fill="x", expand=True)
            if key in self.current_mapping:
                cb.set(self.current_mapping[key])
            self.widgets[key] = cb

        for lbl, key in [
            ("Card ID","card_id"), ("Student ID","student_id"),
            ("Name","name"), ("Phone no.","phone"),
            ("Attendance","attendance"), ("Notes","notes"),
            ("Timestamp","timestamp"),
            ("Exam","exam"), ("Homework","homework")
        ]:
            add_row(lbl, key)

        ttk.Button(self, text="Confirm", command=self.on_confirm).pack(pady=15)

    def on_confirm(self):
        self.mapping = {k: w.get() for k, w in self.widgets.items()}
        vals = list(self.mapping.values())
        if "" in vals or len(set(vals)) < len(vals):
            messagebox.showerror("Mapping Error", "All fields must be uniquely assigned.")
            return
        self.destroy()

class SettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("600x400")

        self.original_bg = Image.open(SETTINGS_BG_FILE)
        self.bg_photo    = ImageTk.PhotoImage(self.original_bg)
        self.bg_label    = tk.Label(self, image=self.bg_photo)
        self.bg_label.place(relwidth=1, relheight=1)
        self.bind("<Configure>", self._on_resize)

        self.parent_app     = parent
        self.column_map     = parent.column_map or {}
        self._main_btns     = []
        self.params_widgets = []
        self.build_main_buttons()

    def _on_resize(self, e):
        if e.widget is self:
            w, h = e.width, e.height
            resized    = self.original_bg.resize((w, h), Image.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(resized)
            self.bg_label.config(image=self.bg_photo)

    def build_main_buttons(self):
        for w in self.params_widgets: w.destroy()
        for b in self._main_btns:     b.destroy()
        self.params_widgets.clear()
        self._main_btns.clear()

        b1 = ttk.Button(self, text="Configure Template", command=self.on_click_configure)
        b2 = ttk.Button(self, text="Session Parameters", command=self.on_click_params)
        b3 = ttk.Button(self, text="Restrictions",       command=self.on_click_restrictions)
        b4 = ttk.Button(self, text="File Type", command=self.on_click_filetype)
        for i, btn in enumerate((b1, b2, b3, b4)):
            btn.place(relx=0.5, rely=0.35 + i*0.2, anchor="center", width=220)
            self._main_btns.append(btn)


        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self.destroy()

    def on_click_configure(self):
        file_type = SETTINGS.get("file_type", "csv")
        ext = "*.xlsx" if file_type == "xlsx" else "*.csv"
        path = filedialog.askopenfilename(
            title=f"Select {file_type.upper()}",
            filetypes=[(f"{file_type.upper()} files", ext)]
        )
        if not path: return
        try:
            df = read_data(path, nrows=0)
        except Exception as e:
            messagebox.showerror("Error", str(e)); return

        dlg = ColumnMappingDialog(self, list(df.columns), self.column_map)
        if getattr(dlg, "mapping", None):
            self.parent_app.column_map = dlg.mapping
            with open(MAPPING_FILE, "w") as f:
                json.dump(dlg.mapping, f, indent=2)
            messagebox.showinfo("Saved", "Template mapping updated.")
            self.focus_force()  # <-- Add this line to refocus settings window

    def on_click_params(self):
        for b in self._main_btns: b.place_forget()
        self.show_params_frame()

    def show_params_frame(self):
        lbl_s = ttk.Label(self, text="Stage Options:"); lbl_s.place(relx=0.1, rely=0.1)
        lb_s  = tk.Listbox(self, height=6); lb_s.place(relx=0.1, rely=0.15, relwidth=0.35)
        for item in SETTINGS["stage_options"]: lb_s.insert("end", item)
        ent_s = ttk.Entry(self); ent_s.place(relx=0.1, rely=0.45, relwidth=0.25)
        btn_add_s = ttk.Button(self, text="Add", command=lambda: add_stage())
        btn_rem_s = ttk.Button(self, text="Remove", command=lambda: rem_stage())
        btn_add_s.place(relx=0.38, rely=0.45, width=50)
        btn_rem_s.place(relx=0.38, rely=0.50, width=50)

        lbl_c = ttk.Label(self, text="Center Options:"); lbl_c.place(relx=0.55, rely=0.1)
        lb_c  = tk.Listbox(self, height=6); lb_c.place(relx=0.55, rely=0.15, relwidth=0.35)
        for item in SETTINGS["center_options"]: lb_c.insert("end", item)
        ent_c = ttk.Entry(self); ent_c.place(relx=0.55, rely=0.45, relwidth=0.25)
        btn_add_c = ttk.Button(self, text="Add", command=lambda: add_center())
        btn_rem_c = ttk.Button(self, text="Remove", command=lambda: rem_center())
        btn_add_c.place(relx=0.83, rely=0.45, width=50)
        btn_rem_c.place(relx=0.83, rely=0.50, width=50)

        btn_save   = ttk.Button(self, text="Save",   command=self.save_params)
        btn_cancel = ttk.Button(self, text="Cancel", command=self.cancel_params)
        btn_save.place(relx=0.35, rely=0.85, width=80)
        btn_cancel.place(relx=0.55, rely=0.85, width=80)

        self.params_widgets = [
            lbl_s, lb_s, ent_s, btn_add_s, btn_rem_s,
            lbl_c, lb_c, ent_c, btn_add_c, btn_rem_c,
            btn_save, btn_cancel
        ]
        self.lb_stage, self.lb_center   = lb_s, lb_c
        self.ent_stage, self.ent_center = ent_s, ent_c

        def add_stage():
            v = ent_s.get().strip()
            if v and v not in lb_s.get(0, "end"): lb_s.insert("end", v)
            ent_s.delete(0, "end")
        def rem_stage():
            sel = lb_s.curselection()
            if sel: lb_s.delete(sel)
        def add_center():
            v = ent_c.get().strip()
            if v and v not in lb_c.get(0, "end"): lb_c.insert("end", v)
            ent_c.delete(0, "end")
        def rem_center():
            sel = lb_c.curselection()
            if sel: lb_c.delete(sel)

    def save_params(self):
        SETTINGS["stage_options"]  = list(self.lb_stage.get(0, "end"))
        SETTINGS["center_options"] = list(self.lb_center.get(0, "end"))
        with open(SETTINGS_FILE, "w") as f:
            json.dump({
                "stage_options":  SETTINGS["stage_options"],
                "center_options": SETTINGS["center_options"],
                "restrictions":   SETTINGS["restrictions"]
            }, f, indent=2)
        for w in self.params_widgets: w.destroy()
        self.build_main_buttons()

    def cancel_params(self):
        for w in self.params_widgets: w.destroy()
        self.build_main_buttons()

    def on_click_restrictions(self):
        dlg = tk.Toplevel(self); dlg.title("Restrictions"); dlg.geometry("300x180")
        var_exam    = tk.BooleanVar(value=SETTINGS["restrictions"]["exam"])
        var_homewrk = tk.BooleanVar(value=SETTINGS["restrictions"]["homework"])
        ttk.Checkbutton(dlg, text="Enable Exam Column",    variable=var_exam).pack(pady=8)
        ttk.Checkbutton(dlg, text="Enable Homework Column",variable=var_homewrk).pack(pady=8)
        frm = ttk.Frame(dlg); frm.pack(pady=10)
        ttk.Button(frm, text="Save",   command=lambda s=True: apply(s)).pack(side="left", padx=10)
        ttk.Button(frm, text="Cancel", command=lambda s=False: apply(s)).pack(side="left", padx=10)
        def apply(save):
            if save:
                SETTINGS["restrictions"]["exam"]     = var_exam.get()
                SETTINGS["restrictions"]["homework"] = var_homewrk.get()
                with open(SETTINGS_FILE, "w") as f:
                    json.dump({
                        "stage_options":  SETTINGS["stage_options"],
                        "center_options": SETTINGS["center_options"],
                        "restrictions":   SETTINGS["restrictions"]
                    }, f, indent=2)
            dlg.destroy()
    def on_click_filetype(self):
        dlg = tk.Toplevel(self)
        dlg.title("File Type")
        dlg.geometry("300x180")
        var_filetype = tk.StringVar(value=SETTINGS.get("file_type", "csv"))
        ttk.Label(dlg, text="Choose data file type:").pack(pady=10)
        rb_csv = ttk.Radiobutton(dlg, text="CSV", variable=var_filetype, value="csv")
        rb_xlsx = ttk.Radiobutton(dlg, text="XLSX", variable=var_filetype, value="xlsx")
        rb_csv.pack(pady=5)
        rb_xlsx.pack(pady=5)
        frm = ttk.Frame(dlg); frm.pack(pady=10)
        ttk.Button(frm, text="Save", command=lambda: apply(True)).pack(side="left", padx=10)
        ttk.Button(frm, text="Cancel", command=lambda: apply(False)).pack(side="left", padx=10)
        def apply(save):
            if save:
                SETTINGS["file_type"] = var_filetype.get()
                with open(SETTINGS_FILE, "w") as f:
                    json.dump(SETTINGS, f, indent=2)
            dlg.destroy()

class SessionDialog(simpledialog.Dialog):
    def __init__(self, parent, stages, centers):
        self.stages  = stages
        self.centers = centers
        self.result  = None
        super().__init__(parent, title="Session Parameters")

    def body(self, master):
        ttk.Label(master, text="Stage:").grid(row=0, column=0, sticky="e")
        self.stage_cb   = ttk.Combobox(master, values=self.stages, state="readonly")
        self.stage_cb.grid(row=0, column=1, pady=5, padx=5)
        ttk.Label(master, text="Center:").grid(row=1, column=0, sticky="e")
        self.center_cb  = ttk.Combobox(master, values=self.centers, state="readonly")
        self.center_cb.grid(row=1, column=1, pady=5, padx=5)
        ttk.Label(master, text="Session No.:").grid(row=2, column=0, sticky="e")
        self.session_ent = ttk.Entry(master, width=4)
        self.session_ent.grid(row=2, column=1, pady=5, padx=5)
        return self.stage_cb

    def validate(self):
        if (not self.stage_cb.get() or
            not self.center_cb.get() or
            not self.session_ent.get().isdigit()):
            messagebox.showerror("Input Error",
                                 "Select stage, center and numeric session no.")
            return False
        return True

    def apply(self):
        stage  = self.stage_cb.get()
        center = self.center_cb.get()
        sn     = int(self.session_ent.get())
        name   = f"{stage} {center} session {sn}"
        self.result = (name, {"stage": stage, "center": center, "no": sn})

class SessionManager:
    def __init__(self, name, params, column_map, data_df):
        self.params       = params
        self.name         = name
        self.mapping      = column_map
        self.data_df      = data_df
        self.records      = []
        self.restrictions = SETTINGS["restrictions"]
        # Use the correct extension based on SETTINGS
        file_type = SETTINGS.get("file_type", "csv")
        ext = "xlsx" if file_type == "xlsx" else "csv"
        self.session_path = os.path.join(SESSIONS_FOLDER, f"{name}.{ext}")
        if os.path.exists(self.session_path):
            df = read_data(self.session_path)
            # Only keep mapped columns
            mapped_keys = ["card_id", "student_id", "name", "phone", "attendance", "notes"]
            if self.restrictions.get("exam"):
                mapped_keys.append("exam")
            if self.restrictions.get("homework"):
                mapped_keys.append("homework")
            self.records = []
            for _, row in df.iterrows():
                rec = {}
                for k in mapped_keys:
                    col = self.mapping.get(k, k)
                    rec[k] = row.get(col, "")
                self.records.append(rec)

    def add_record(self, rec):
        df = read_data(self.session_path)
        card_col     = self.mapping.get("card_id", "card_id")
        att_col      = self.mapping.get("attendance", "attendance")
        notes_col    = self.mapping.get("notes", "notes")
        timestamp_col= self.mapping.get("timestamp", "timestamp")
        mask = df[card_col].astype(str) == str(rec["card_id"])

        if mask.any():
            # --- Preserve timestamp if already present ---
            existing_timestamp = ""
            if timestamp_col in df.columns:
                existing_timestamp = df.loc[mask, timestamp_col].values[0]
            # Only overwrite if rec["timestamp"] is not empty
            if rec.get("timestamp"):
                df.loc[mask, timestamp_col] = rec["timestamp"]
            else:
                df.loc[mask, timestamp_col] = existing_timestamp
            df.loc[mask, att_col]   = rec["attendance"]
            df.loc[mask, notes_col] = rec["notes"]
        else:
            row = {col: "" for col in df.columns}
            for k in ("card_id", "student_id", "name", "phone"):
                col_name = self.mapping.get(k, k)
                if col_name in df.columns:
                    row[col_name] = rec.get(k, "")
            row[att_col]      = rec["attendance"]
            row[notes_col]    = rec["notes"]
            row[timestamp_col]= rec.get("timestamp", "")
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        write_data(df, self.session_path)

def read_data(path, **kwargs):
    if path.lower().endswith(".xlsx"):
        return pd.read_excel(path, dtype=str, **kwargs)
    else:
        return pd.read_csv(path, dtype=str, **kwargs)

def write_data(df, path, **kwargs):
    if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False, **kwargs)
    else:
        df.to_csv(path, index=False, **kwargs)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RFID Attendance Manager")
        self.geometry("800x600")
        self.column_map = {}
        self.data_df    = None
        self.settings_window = None  # <-- Track settings window

        if os.path.exists(MAPPING_FILE):
            with open(MAPPING_FILE) as f:
                self.column_map = json.load(f)
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE) as f:
                SETTINGS.update(json.load(f))

        self._build_ui()
        self._load_last_data()

    def _build_ui(self):
        self.original_bg = Image.open(HOME_BG_FILE)
        self.bg_photo    = ImageTk.PhotoImage(self.original_bg)
        self.canvas      = tk.Canvas(self, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.bg_id       = self.canvas.create_image(0,0,anchor="nw", image=self.bg_photo)
        self.footer_id   = self.canvas.create_text(
            10, 590, anchor="sw",
            text="Powered by Gawish", fill="white", font=("Arial",9)
        )

        logo = Image.open(LOGO_FILE)
        ratio = 100 / logo.width
        logo = logo.resize((100, int(logo.height*ratio)), Image.LANCZOS)
        self.logo_img = ImageTk.PhotoImage(logo)
        self.logo_id  = self.canvas.create_image(10,10,anchor="nw", image=self.logo_img)
        self.canvas.tag_bind(self.logo_id, "<Button-1>", lambda e: self.open_settings())

        self.import_txt = self.canvas.create_text(
            300, 240, text="📥 Import excel",
            font=("Arial",16,"bold"), fill="white"
        )
        self.canvas.tag_bind(self.import_txt, "<Button-1>", lambda e: self.import_csv())
        self.canvas.tag_bind(self.import_txt, "<Enter>",
                             lambda e: self.canvas.itemconfig(self.import_txt, fill="#ddd"))
        self.canvas.tag_bind(self.import_txt, "<Leave>",
                             lambda e: self.canvas.itemconfig(self.import_txt, fill="white"))

        self.scan_txt = self.canvas.create_text(
            500, 240, text="🔍 Scan",
            font=("Arial",16,"bold"), fill="white"
        )
        self.canvas.tag_bind(self.scan_txt, "<Button-1>", lambda e: self.open_scan_window())
        self.canvas.tag_bind(self.scan_txt, "<Enter>",
                             lambda e: self.canvas.itemconfig(self.scan_txt, fill="#ddd"))
        self.canvas.tag_bind(self.scan_txt, "<Leave>",
                             lambda e: self.canvas.itemconfig(self.scan_txt, fill="white"))

        self.bind("<Configure>", self._on_resize)

    def _on_resize(self, e):
        if e.widget is self:
            w, h = e.width, e.height
            bg = self.original_bg.resize((w,h), Image.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(bg)
            self.canvas.itemconfig(self.bg_id, image=self.bg_photo)
            self.canvas.coords(self.footer_id, 10, h-10)
            self.canvas.coords(self.import_txt, 300, h/2)
            self.canvas.coords(self.scan_txt, 500, h/2)

    def _load_last_data(self):
        self.data_df = None  # Always reset on startup
        if os.path.exists(LAST_DATA_FILE):
            with open(LAST_DATA_FILE) as f:
                info = json.load(f)
            path = info.get("path")
            if path and os.path.exists(path):
                self.data_df = read_data(path)

    def open_settings(self):
        # Only open one settings window at a time
        if self.settings_window is not None and tk.Toplevel.winfo_exists(self.settings_window):
            self.settings_window.focus_force()
            self.settings_window.lift()
            return
        self.settings_window = SettingsWindow(self)
        self.settings_window.protocol("WM_DELETE_WINDOW", self._on_settings_close)

    def _on_settings_close(self):
        if self.settings_window is not None:
            self.settings_window.destroy()
            self.settings_window = None

    def import_csv(self):
        # Check if template is configured
        if not self.column_map:
            messagebox.showwarning("No Template", "Please configure a template first.")
            return

        file_type = SETTINGS.get("file_type", "csv")
        ext = "*.xlsx" if file_type == "xlsx" else "*.csv"
        path = filedialog.askopenfilename(
            title=f"Select {file_type.upper()}",
            filetypes=[(f"{file_type.upper()} files", ext)]
        )
        if not path: return
        try:
            df = read_data(path)
            # Pad card_id column to 8 digits and assign 'null N' for blanks
            card_col = self.column_map.get("card_id", "card_id")
            if card_col in df.columns:
                null_counter = 1
                new_card_ids = []
                for val in df[card_col]:
                    val_str = str(val).strip()
                    if not val_str or val_str.lower() == "nan":
                        new_card_ids.append(f"null {null_counter}")
                        null_counter += 1
                    elif val_str.isdigit():
                        new_card_ids.append(val_str.zfill(8))
                    else:
                        new_card_ids.append(val_str)
                df[card_col] = new_card_ids

            # Clear attendance and timestamp columns for imported data only
            att_col = self.column_map.get("attendance", "attendance")
            ts_col  = self.column_map.get("timestamp", "timestamp")
            if att_col in df.columns:
                df[att_col] = ""
            if ts_col in df.columns:
                df[ts_col] = ""
        except Exception as e:
            return messagebox.showerror("Load Error", str(e))
        self.data_df = df
        with open(LAST_DATA_FILE, "w") as f:
            json.dump({"path": path}, f, indent=2)
        messagebox.showinfo("Loaded", f"{len(df)} records imported.")

    def open_scan_window(self):
        if self.data_df is None:
            if not messagebox.askyesno("No data imported","Import excel first. Continue anyway?"):
                return
            # Create an empty DataFrame with default columns
            cols = ["card_id", "student_id", "name", "phone", "attendance", "notes", "timestamp"]
            if SETTINGS["restrictions"].get("exam"):
                cols.insert(4, "exam")
            if SETTINGS["restrictions"].get("homework"):
                cols.insert(5 if SETTINGS["restrictions"].get("exam") else 4, "homework")
            self.data_df = pd.DataFrame(columns=cols)
        dlg = SessionDialog(self, SETTINGS["stage_options"], SETTINGS["center_options"])
        if not getattr(dlg, "result", None): return
        name, params = dlg.result
        file_type = SETTINGS.get("file_type", "csv")
        ext = "xlsx" if file_type == "xlsx" else "csv"
        session_path = os.path.join(SESSIONS_FOLDER, f"{name}.{ext}")
        if not os.path.exists(session_path):
            write_data(self.data_df, session_path)
        session_df = read_data(session_path)
        sm = SessionManager(name, params, self.column_map, session_df)
        ScanWindow(self, sm)
class ScanWindow(tk.Toplevel):
    def __init__(self, parent, session_mgr):
        super().__init__(parent)
        self.parent       = parent
        self.sm           = session_mgr
        self.restrictions = self.sm.restrictions
        # Always use the session file as the base
        self.df = read_data(self.sm.session_path).fillna("")
        self.mapping = self.sm.mapping
        # --- Ensure mapping always has defaults ---
        if not self.mapping:
            self.mapping = {col: col for col in self.df.columns}

        self.original_bg  = Image.open(HOME_BG_FILE)
        self.bg_photo     = ImageTk.PhotoImage(self.original_bg)
        self.bg_label     = tk.Label(self, image=self.bg_photo)
        self.bg_label.place(relwidth=1, relheight=1)
        self.bind("<Configure>", self._on_bg_resize)
        self.bind("<Button-1>",     self._on_click_anywhere)

        self.title("Scan Attendance")
        self.geometry("800x500")

        # Initialize IID list before loading
        self._all_iids = []

        self._build_ui()
        self._load_existing()

        self.scan_entry.focus_set()

    def _on_bg_resize(self, e):
        if e.widget is self:
            bg = self.original_bg.resize((e.width, e.height), Image.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(bg)
            self.bg_label.config(image=self.bg_photo)

    def _build_ui(self):
        # Scan Entry
        self.scan_entry = ttk.Entry(self)
        self.scan_entry.place(x=10, y=10, width=200)
        self.scan_entry.bind("<Return>", lambda e: self._on_scan())

        # Header
        hdr = ttk.Frame(self); hdr.pack(fill="x", pady=5, padx=10)
        ttk.Label(hdr, text=f"Session: {self.sm.name}",
                  font=("Arial",12,"bold")).pack(side="left")
        # End Scan now runs _on_end_scan
        ttk.Button(hdr, text="End Scan", command=self._on_end_scan).pack(side="right")

        # Search Fields
        sf = ttk.Frame(self); sf.pack(fill="x", padx=10, pady=5)
        for i, field in enumerate(("student_id","name","phone")):
            ph = field.replace("_"," ").title()
            ent = ttk.Entry(sf, foreground="grey")
            ent.insert(0, ph)
            ent.grid(row=0, column=i, sticky="ew", padx=5)
            ent.bind("<FocusIn>",  lambda e,en=ent,ph=ph: self._clear_placeholder(en,ph))
            ent.bind("<FocusOut>", lambda e,en=ent,ph=ph: self._restore_placeholder(en,ph))
            ent.bind("<KeyRelease>", self._filter_all)
            setattr(self, f"search_{field}", ent)
            sf.columnconfigure(i, weight=1)

        # Columns (with restrictions)
        cols = ["card_id","student_id","name","phone"]
        if self.restrictions.get("exam"):     cols.append("exam")
        if self.restrictions.get("homework"): cols.append("homework")
        cols += ["attendance","notes","timestamp"]

        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.tree.heading(c, text=c.replace("_"," ").title())
            self.tree.column(c, anchor="center", width=100)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", self._on_row_double_click)

        # Control Bar
        ctl = ttk.Frame(self); ctl.pack(fill="x", padx=10, pady=5)
        ttk.Button(ctl, text="Add Student", command=self._on_add_student_flow).pack(side="left")
        self.pb = ttk.Progressbar(ctl, mode="determinate", maximum=2000)
        self.pb.pack(side="right", fill="x", expand=True)

        self._search_entries = [
            getattr(self, f"search_{field}") for field in ("student_id", "name", "phone")
        ]
        self.bind_all("<FocusIn>", self._global_focus_in, add="+")

    def _load_existing(self):
        def pad_card_id(val):
            val_str = str(val)
            if val_str.isdigit():
                return val_str.zfill(8)
            return val_str

        def clean(val):
            # Handles None, float nan, and string "nan"
            if val is None:
                return ""
            if isinstance(val, float) and pd.isna(val):
                return ""
            if isinstance(val, str) and val.strip().lower() == "nan":
                return ""
            return str(val)

        cols = self.tree["columns"]
        session_records = {pad_card_id(rec.get("card_id", "")): rec for rec in self.sm.records}
        self._all_iids = []

        for _, row in self.df.iterrows():
            card_id_col = self.mapping.get("card_id", "card_id")
            cid = pad_card_id(row.get(card_id_col, ""))
            rec = session_records.pop(cid, None)
            vals = []
            for col in cols:
                mapped_col = self.mapping.get(col, col)
                if rec is not None and col in rec:
                    val = rec.get(col, "")
                else:
                    val = row.get(mapped_col, "")
                vals.append(clean(val))
            self.tree.insert("", "end", iid=cid, values=tuple(vals))
            self._all_iids.append(cid)

        # Add any manual/canceled records not in imported data
        for rec in self.sm.records:
            cid = pad_card_id(rec.get("card_id", ""))
            if not self.tree.exists(cid):
                vals = []
                for col in cols:
                    val = rec.get(col, "")
                    vals.append(clean(val))
                self.tree.insert("", "end", iid=cid, values=tuple(vals))
                self._all_iids.append(cid)
            # Only update attendance, notes, and timestamp columns
            self._update_row(cid, clean(rec.get("attendance", "")), clean(rec.get("notes", "")))

    def _clear_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, "end")
            entry.config(foreground="black")

    def _restore_placeholder(self, entry, placeholder):
        if not entry.get().strip():
            entry.insert(0, placeholder)
            entry.config(foreground="grey")

    def _on_click_anywhere(self, event):
        if not isinstance(event.widget, ttk.Entry):
            self.focus()
            for field in ("student_id","name","phone"):
                ent = getattr(self, f"search_{field}")
                ph  = field.replace("_"," ").title()
                if not ent.get().strip():
                    ent.delete(0, "end")
                    ent.insert(0, ph)
                    ent.config(foreground="grey")

    def _filter_all(self, *_):
        terms = {}
        for fld in ("student_id","name","phone"):
            ent = getattr(self, f"search_{fld}")
            val = ent.get().strip()
            ph  = fld.replace("_"," ").title()
            if val.lower() == ph.lower(): val = ""
            terms[fld] = val.lower()

        if not any(terms.values()):
            for iid in self._all_iids:
                self.tree.reattach(iid, '', 'end')
            return

        for iid in self._all_iids:
            keep = all(
                (not terms[f]) or (terms[f] in self.tree.set(iid, f).lower())
                for f in terms
            )
            if keep: self.tree.reattach(iid, '', 'end')
            else:    self.tree.detach(iid)

    def _on_scan(self):
        code = self.scan_entry.get().strip()
        self.scan_entry.delete(0, "end")
        # Pad to 8 digits only if numeric
        if code.isdigit():
            code = code.zfill(8)
        if not code:
            return
        if not self.tree.exists(code):
            # Show not found popup
            resp = messagebox.askquestion(
                "Not found",
                f"Card ID '{code}' not found.\nDo you want to add this student?",
                icon="warning", type="yesno", default="no", parent=self
            )
            if resp == "yes":
                self._add_student_with_cardid(code)
            # else: just return to scanning
            return
        self._visible_iids = list(self._all_iids)
        for iid in self._all_iids:
            if iid != code: self.tree.detach(iid)
        self.pb.start(10)
        self.after(2000, lambda: self._finish_scan(code))

    def _add_student_with_cardid(self, card_id):
        card_id = card_id.zfill(8)
        dlg = tk.Toplevel(self)
        dlg.title("Add Student")
        width, height = 320, 200
        dlg.geometry(f"{width}x{height}")
        bg_img = self.original_bg.resize((width, height), Image.LANCZOS)
        bg     = ImageTk.PhotoImage(bg_img)
        lbl    = tk.Label(dlg, image=bg); lbl.place(relwidth=1, relheight=1)
        dlg.bg_photo = bg

        # Show Card ID (read-only)
        ttk.Label(dlg, text="Card ID").place(x=10, y=10)
        cid_lbl = ttk.Label(dlg, text=card_id)
        cid_lbl.place(x=100, y=10)

        inputs = {}
        for i, lab in enumerate(("Student ID","Name","Phone")):
            y = 50 + 40 * i
            ttk.Label(dlg, text=lab).place(x=10, y=y)
            ent = ttk.Entry(dlg); ent.place(x=100, y=y, width=200)
            inputs[lab.lower().replace(" ", "_")] = ent

        def confirm():
            vals = {k: e.get().strip() for k,e in inputs.items()}
            if not (vals["student_id"] and vals["name"] and vals["phone"]):
                return messagebox.showerror("Missing", "Fill all fields", parent=dlg)
            id_exists, phone_exists = self._student_id_or_phone_exists(vals["student_id"], vals["phone"])
            if id_exists or phone_exists:
                if id_exists and phone_exists:
                    msg = "This ID & phone no. already exists"
                elif id_exists:
                    msg = "This ID already exists"
                else:
                    msg = "This phone no. already exists"
                messagebox.showwarning("Duplicate", msg, parent=dlg)
                return  # Do not close the dialog
            rec = {
                "card_id":    card_id,
                "student_id": vals["student_id"],
                "name":       vals["name"],
                "phone":      vals["phone"],
                "attendance": "attend",
                "notes":      "Diff group",  # notes only
                "timestamp":  datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
            }
            if self.restrictions.get("exam"):
                rec["exam"] = ""
            if self.restrictions.get("homework"):
                rec["homework"] = ""
            self.sm.add_record(rec)
            vals_list = [
                rec["card_id"], rec["student_id"],
                rec["name"], rec["phone"]
            ]
            if self.restrictions.get("exam"):     vals_list.append(rec["exam"])
            if self.restrictions.get("homework"): vals_list.append(rec["homework"])
            vals_list += [rec["attendance"], rec["notes"], rec["timestamp"]]
            self.tree.insert("", "end", iid=card_id, values=tuple(vals_list))
            self._all_iids.append(card_id)
            self._update_row(card_id, "attend", rec["notes"])
            dlg.destroy()

        ttk.Button(dlg, text="Confirm", command=confirm).place(x=60, y=height-30, width=80)
        ttk.Button(dlg, text="Cancel",  command=dlg.destroy).place(x=180, y=height-30, width=80)

    def _finish_scan(self, code):
        self.pb.stop()
        for iid in self._visible_iids:
            self.tree.reattach(iid, '', 'end')
        # Only mark attendance if not already attended
        already_attended = False
        for r in self.sm.records:
            if str(r["card_id"]) == code and r.get("attendance", "") == "attend":
                already_attended = True
                break

        # --- Check restrictions before marking attendance ---
        missing = []
        if self.restrictions.get("exam") and not self.tree.set(code, "exam").strip():
            missing.append("exam")
        if self.restrictions.get("homework") and not self.tree.set(code, "homework").strip():
            missing.append("homework")

        if missing:
            self._open_gate_dialog(code, missing)
            return

        if not already_attended:
            self._mark_attendance(code)
            self._update_row(code, "attend", "")

    def _on_row_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        cid = sel[0]

        missing = []
        if self.restrictions.get("exam") and not self.tree.set(cid, "exam").strip():
            missing.append("exam")
        if self.restrictions.get("homework") and not self.tree.set(cid, "homework").strip():
            missing.append("homework")

        if missing:
            self._open_gate_dialog(cid, missing)
            return

        name = self.tree.set(cid, "name")
        if messagebox.askyesno("Confirm Attendance",
                               f"Mark {name} as attended?",
                               parent=self):
            self._mark_attendance(cid)
        else:
            # Only allow cancel if currently attended
            if self.tree.set(cid, "attendance") == "attend":
                self._cancel_attendance(cid)
            else:
                # Just close the popup, do nothing
                return

    def _open_gate_dialog(self, cid, missing):
        if len(missing) == 2:
            msg = "This student didn't take the exam and didn't turn in the H.W."
        else:
            msg = ("This student didn't take the exam."
                   if missing[0] == "exam"
                   else "This student didn't turn in the H.W.")
        dlg = tk.Toplevel(self)
        dlg.title("Gate Restriction")
        dlg.transient(self); dlg.grab_set()

        tk.Label(dlg, text=msg, wraplength=300).pack(pady=10, padx=10)
        btn_frame = tk.Frame(dlg); btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="Cancel",
                  command=lambda: self._on_gate_cancel(cid, dlg)
        ).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Attend",
                  command=lambda: self._on_gate_attend(cid, missing, dlg)
        ).pack(side="left", padx=5)

    def _on_gate_cancel(self, cid, dlg):
        note = f"[{datetime.now():%H:%M:%S}] canceled"
        self.sm.add_record({
            "card_id": cid, "student_id": "", "name": "",
            "phone": "", "attendance": "", "notes": note
        })
        self._update_row(cid, "", note)
        dlg.destroy()

    def _on_gate_attend(self, cid, missing, dlg):
        # Check if already attended
        if self.tree.set(cid, "attendance") == "attend":
            messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
            dlg.destroy()
            return
        # Otherwise, mark attendance and show missing note dialog
        self._mark_attendance(cid)
        dlg.destroy()
        self._open_missing_note_dialog(cid, missing)

    def _open_missing_note_dialog(self, cid, missing):
        if len(missing) == 2:
            btn_text = "No exam & H.W."
        else:
            btn_text = "No exam" if missing[0] == "exam" else "No H.W."

        sub = tk.Toplevel(self)
        sub.title("Add Missing Note")
        sub.transient(self); sub.grab_set()

        tk.Label(sub, text="Append a missing-field note?", wraplength=250).pack(pady=10)
        frame = tk.Frame(sub); frame.pack(pady=5)
        tk.Button(frame, text=btn_text,
                  command=lambda: self._append_missing_note(cid, btn_text, sub)
        ).pack(side="left", padx=5)
        tk.Button(frame, text="Finished Session",
                  command=sub.destroy
        ).pack(side="left", padx=5)

    def _append_missing_note(self, cid, text, sub):
        current = self.tree.set(cid, "notes")
        new_note = f"{current} [{text}]"
        # Always update the session file
        rec = {
            "card_id": cid,
            "student_id": self.tree.set(cid, "student_id"),
            "name": self.tree.set(cid, "name"),
            "phone": self.tree.set(cid, "phone"),
            "attendance": "attend",
            "notes": new_note,
            "timestamp": datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        }
        if self.restrictions.get("exam"):
            rec["exam"] = self.tree.set(cid, "exam")
        if self.restrictions.get("homework"):
            rec["homework"] = self.tree.set(cid, "homework")
        self.sm.add_record(rec)
        self._update_row(cid, "attend", new_note)
        sub.destroy()

    def _mark_attendance(self, code):
        card_col = self.mapping.get("card_id", "card_id")
        att_col = self.mapping.get("attendance", "attendance")
        df = read_data(self.sm.session_path)
        mask = df[card_col].astype(str) == code
        if mask.any():
            # If already attended, warn and return
            if (df.loc[mask, att_col] == "attend").any():
                messagebox.showwarning("Already Attended", "This student is already attended.", parent=self)
                return
            # Otherwise, mark as attended
            idx = df.index[mask][0]
            df.at[idx, att_col] = "attend"
            df.at[idx, self.mapping.get("timestamp", "timestamp")] = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
            write_data(df, self.sm.session_path)
        else:
            # Student not found in file, create new record from tree
            if self.tree.exists(code):
                rec = {
                    "card_id": code,
                    "student_id": self.tree.set(code, "student_id"),
                    "name": self.tree.set(code, "name"),
                    "phone": self.tree.set(code, "phone"),
                    "attendance": "attend",
                    "notes": "",
                    "timestamp": datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
                }
                if self.restrictions.get("exam"):
                    rec["exam"] = self.tree.set(code, "exam")
                if self.restrictions.get("homework"):
                    rec["homework"] = self.tree.set(code, "homework")
                self.sm.add_record(rec)
        self._update_row(code, "attend", "")

    def _cancel_attendance(self, code):
        canceled_note = f"[{datetime.now():%H:%M:%S}] canceled"
        found = False
        for r in self.sm.records:
            if str(r["card_id"]) == code:
                r["attendance"] = ""
                # Clean old notes to avoid 'nan'
                old_notes = r.get("notes", "")
                if old_notes is None or str(old_notes).strip().lower() == "nan":
                    r["notes"] = canceled_note
                else:
                    r["notes"] = f"{old_notes} {canceled_note}"
                found = True
                break
        if not found:
            return

        # Update only attendance and notes in the file, keep timestamp as is
        df = read_data(self.sm.session_path)
        card_col = self.mapping.get("card_id", "card_id")
        att_col = self.mapping.get("attendance", "attendance")
        notes_col = self.mapping.get("notes", "notes")
        mask = df[card_col].astype(str) == str(code)
        if mask.any():
            df.loc[mask, att_col] = ""
            # Clean old notes to avoid 'nan'
            old_notes = df.loc[mask, notes_col].astype(str).fillna("")
            def clean_note(val):
                return "" if val.strip().lower() == "nan" else val
            cleaned_notes = old_notes.apply(clean_note)
            df.loc[mask, notes_col] = cleaned_notes + " " + canceled_note
            df.loc[mask, notes_col] = df.loc[mask, notes_col].str.strip()
            write_data(df, self.sm.session_path)

        self._update_row(code, "", r["notes"])

    def _update_row(self, code, attendance, notes):
        def clean(val):
            if val is None:
                return ""
            if isinstance(val, float) and pd.isna(val):
                return ""
            if isinstance(val, str) and val.strip().lower() == "nan":
                return ""
            return str(val)

        if not self.tree.exists(code): return
        vals = list(self.tree.item(code, "values"))
        # Get the timestamp from the session file (self.sm.session_path)
        df = read_data(self.sm.session_path)
        timestamp_col = self.mapping.get("timestamp", "timestamp")
        card_col = self.mapping.get("card_id", "card_id")
        mask = df[card_col].astype(str) == str(code)
        if mask.any() and timestamp_col in df.columns:
            timestamp = df.loc[mask, timestamp_col].values[0]
        else:
            timestamp = ""
        vals[-3], vals[-2], vals[-1] = clean(attendance), clean(notes), clean(timestamp)
        self.tree.item(code, values=vals)

    def _on_add_student_flow(self):
        if not hasattr(self, "_unknown_counter"):
            # Find the highest Unknown N already used
            unknowns = [
                int(str(r["card_id"]).split("Unknown ")[1])
                for r in self.sm.records
                if str(r["card_id"]).startswith("Unknown ")
                   and str(r["card_id"]).split("Unknown ")[1].isdigit()
            ]
            self._unknown_counter = max(unknowns, default=0)
        dlg = tk.Toplevel(self)
        dlg.title("Add Student")
        width, height = 320, 160
        dlg.geometry(f"{width}x{height}")
        bg_img = self.original_bg.resize((width, height), Image.LANCZOS)
        bg     = ImageTk.PhotoImage(bg_img)
        lbl    = tk.Label(dlg, image=bg); lbl.place(relwidth=1, relheight=1)
        dlg.bg_photo = bg

        inputs = {}
        for i, lab in enumerate(("Student ID","Name","Phone")):
            y = 10 + 40 * i
            ttk.Label(dlg, text=lab).place(x=10, y=y)
            ent = ttk.Entry(dlg); ent.place(x=100, y=y, width=200)
            inputs[lab.lower().replace(" ", "_")] = ent

        def confirm():
            vals = {k: e.get().strip() for k,e in inputs.items()}
            if not (vals["student_id"] and vals["name"] and vals["phone"]):
                return messagebox.showerror("Missing", "Fill all fields", parent=dlg)
            id_exists, phone_exists = self._student_id_or_phone_exists(vals["student_id"], vals["phone"])
            if id_exists or phone_exists:
                if id_exists and phone_exists:
                    msg = "This ID & phone no. already exists"
                elif id_exists:
                    msg = "This ID already exists"
                else:
                    msg = "This phone no. already exists"
                messagebox.showwarning("Duplicate", msg, parent=dlg)
                return  # Do not close the dialog
            self._unknown_counter += 1
            cid = f"Unknown {self._unknown_counter}"
            rec = {
                "card_id":    cid,
                "student_id": vals["student_id"],
                "name":       vals["name"],
                "phone":      vals["phone"],
                "attendance": "attend",
                "notes":      "manual addition",  # notes only
                "timestamp":  datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
            }
            if self.restrictions.get("exam"):
                rec["exam"] = ""
            if self.restrictions.get("homework"):
                rec["homework"] = ""
            self.sm.add_record(rec)
            vals_list = [
                rec["card_id"], rec["student_id"],
                rec["name"], rec["phone"]
            ]
            if self.restrictions.get("exam"):     vals_list.append(rec["exam"])
            if self.restrictions.get("homework"): vals_list.append(rec["homework"])
            vals_list += [rec["attendance"], rec["notes"], rec["timestamp"]]  # <-- add timestamp here
            self.tree.insert("", "end", iid=cid, values=tuple(vals_list))
            self._all_iids.append(cid)
            self._update_row(cid, "attend", rec["notes"])
            dlg.destroy()

        ttk.Button(dlg, text="Confirm", command=confirm).place(x=60, y=height-30, width=80)
        ttk.Button(dlg, text="Cancel",  command=dlg.destroy).place(x=180, y=height-30, width=80)

    def _on_end_scan(self):
        self.destroy()

    def _global_focus_in(self, event):
        widget = self.focus_get()
        # Allow focus if it's on scan_entry or a search entry
        if widget == self.scan_entry or widget in self._search_entries:
            return
        # Allow focus if a dialog (Toplevel) is open and has focus
        for w in self.winfo_children():
            if isinstance(w, tk.Toplevel):
                # If the widget is a child of the dialog, don't steal focus
                if widget is not None and (widget == w or widget.winfo_toplevel() == w):
                    return
        # Otherwise, force focus back to scan_entry
        self.after_idle(self.scan_entry.focus_set)

    def _student_id_or_phone_exists(self, student_id, phone):
        # Always check the session file for up-to-date records
        df = read_data(self.sm.session_path)
        sid_col = self.mapping.get("student_id", "student_id")
        phone_col = self.mapping.get("phone", "phone")
        id_exists = phone_exists = False

        if sid_col in df.columns:
            if student_id in df[sid_col].astype(str).values:
                id_exists = True
        if phone_col in df.columns:
            if phone in df[phone_col].astype(str).values:
                phone_exists = True

        return id_exists, phone_exists

if __name__ == "__main__":
    app = App()
    app.mainloop()
