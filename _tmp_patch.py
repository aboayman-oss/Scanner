from pathlib import Path

path = Path(r"C:\Users\Ahmed\Desktop\honda\raw\src\ui\scan_window.py")
text = path.read_text()
text = text.replace('scan_btn.pack(side="left", padx=6)\n', 'scan_btn.pack(side="left", padx=6, pady=4)\n', 1)
path.write_text(text)
