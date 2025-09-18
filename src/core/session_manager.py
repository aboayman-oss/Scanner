"""Session management utilities."""
import os

import pandas as pd

from utils.helpers import SETTINGS, SESSIONS_FOLDER, read_data, write_data

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
