"""
gui.py — Tkinter GUI for the Test Plan Filler tool.
Imports all processing logic from backend.py.

Run:
    python gui.py
Both gui.py and backend.py must be in the same directory.
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from backend import run_processing


# ─────────────────────────────────────────────────────────────────────────────
# Colour palette
# ─────────────────────────────────────────────────────────────────────────────
BG      = "#1e1e2e"
SURFACE = "#2a2a3e"
ACCENT  = "#7c6af7"
ACCENT2 = "#5a4fd4"
FG      = "#cdd6f4"
FG2     = "#a6adc8"
MUTED   = "#44445a"
GREEN   = "#a6e3a1"
RED     = "#f38ba8"

y = 7+9
z = 10 + 13
# ─────────────────────────────────────────────────────────────────────────────
# Reusable file-picker row widget
# ─────────────────────────────────────────────────────────────────────────────

class FilePickerRow(tk.Frame):
    """
    A labelled row containing:
      - a label
      - a read-only Entry showing the selected path
      - a Browse button
      - a Clear (✕) button
    """

    def __init__(self, parent, label: str, filetypes: list, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.var = tk.StringVar()

        tk.Label(self, text=label, font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=FG).grid(row=0, column=0, columnspan=3,
                                     sticky="w", pady=(8, 2))

        self.entry = tk.Entry(
            self, textvariable=self.var,
            font=("Segoe UI", 9), bg=SURFACE, fg=FG,
            insertbackground=FG, relief="flat", bd=0,
            highlightthickness=1,
            highlightbackground=MUTED, highlightcolor=ACCENT
        )
        self.entry.grid(row=1, column=0, sticky="ew", ipady=6, padx=(0, 6))

        browse_btn = tk.Button(
            self, text="Browse…",
            font=("Segoe UI", 9), bg=MUTED, fg=FG,
            activebackground=ACCENT, activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=10, pady=4,
            command=lambda: self._browse(filetypes)
        )
        browse_btn.grid(row=1, column=1, padx=(0, 4))

        clear_btn = tk.Button(
            self, text="✕",
            font=("Segoe UI", 9), bg=MUTED, fg=FG2,
            activebackground=RED, activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=8, pady=4,
            command=lambda: self.var.set("")
        )
        clear_btn.grid(row=1, column=2)

        self.columnconfigure(0, weight=1)

    def _browse(self, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.var.set(path)

    def get(self) -> str:
        return self.var.get().strip()


# ─────────────────────────────────────────────────────────────────────────────
# Log console widget
# ─────────────────────────────────────────────────────────────────────────────

class LogConsole(tk.Frame):
    """Dark scrolled-text widget with colour-tagged log output."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)

        header = tk.Frame(self, bg=SURFACE, pady=4)
        header.pack(fill="x")
        tk.Label(header, text="  Log", font=("Segoe UI", 9, "bold"),
                 bg=SURFACE, fg=FG2).pack(side="left")

        self._text = scrolledtext.ScrolledText(
            self, height=14, font=("Consolas", 9),
            bg="#12121e", fg=FG, insertbackground=FG,
            relief="flat", bd=0, padx=10, pady=8, wrap="word"
        )
        self._text.pack(fill="both", expand=True)
        self._text.tag_config("ok",  foreground=GREEN)
        self._text.tag_config("err", foreground=RED)
        self._text.tag_config("hdr", foreground=ACCENT,
                               font=("Consolas", 9, "bold"))

    def clear(self):
        self._text.delete("1.0", "end")

    def write(self, msg: str, tag: str = ""):
        self._text.insert("end", msg + "\n", tag)
        self._text.see("end")


# ─────────────────────────────────────────────────────────────────────────────
# Main application window
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Test Plan Filler")
        self.configure(bg=BG)
        self.minsize(640, 540)
        self.resizable(True, True)
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Title bar
        title_frame = tk.Frame(self, bg=ACCENT, pady=10)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="  ⚙  Test Plan Filler",
                 font=("Segoe UI", 14, "bold"),
                 bg=ACCENT, fg="white").pack(side="left", padx=16)
        tk.Label(title_frame,
                 text="Fills Expected & Actual Results from .c file",
                 font=("Segoe UI", 9), bg=ACCENT, fg="#ddd").pack(side="left")

        # File pickers
        picker_frame = tk.Frame(self, bg=BG)
        picker_frame.pack(fill="x", padx=16, pady=6)

        self.xlsx_picker = FilePickerRow(
            picker_frame,
            label="📊  Excel File (.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        self.xlsx_picker.pack(fill="x", pady=(4, 0))

        self.c_picker = FilePickerRow(
            picker_frame,
            label="📄  C Test File (.c)",
            filetypes=[("C files", "*.c"), ("All files", "*.*")]
        )
        self.c_picker.pack(fill="x", pady=(4, 0))

        # Generate button
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(pady=(10, 4))

        self.gen_btn = tk.Button(
            btn_frame, text="  ▶   Generate  ",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white",
            activebackground=ACCENT2, activeforeground="white",
            relief="flat", cursor="hand2", bd=0,
            padx=24, pady=8,
            command=self._on_generate
        )
        self.gen_btn.pack()

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=420)
        self.progress.pack(pady=(4, 6))

        # Log console
        self.console = LogConsole(self)
        self.console.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    # ── Event handlers ────────────────────────────────────────────────────────

    def _on_generate(self):
        xlsx = self.xlsx_picker.get()
        c    = self.c_picker.get()

        if not xlsx:
            messagebox.showwarning("Missing Input",
                                   "Please select an Excel (.xlsx) file.")
            return
        if not c:
            messagebox.showwarning("Missing Input",
                                   "Please select a C test file (.c).")
            return

        self.gen_btn.config(state="disabled")
        self.console.clear()
        self.progress.start(12)
        self.console.write("Starting…\n", "hdr")

        # Run backend in a background thread so the GUI stays responsive
        threading.Thread(
            target=self._worker, args=(xlsx, c), daemon=True
        ).start()

    def _worker(self, xlsx: str, c: str):
        """Background thread: calls backend, then schedules UI update."""
        try:
            run_processing(xlsx, c,
                           log=lambda msg: self.after(0, self.console.write, msg))
            self.after(0, self._on_success)
        except Exception as exc:
            self.after(0, self._on_error, str(exc))

    def _on_success(self):
        self.progress.stop()
        self.gen_btn.config(state="normal")
        messagebox.showinfo("Success", "Excel file updated successfully!")

    def _on_error(self, msg: str):
        self.progress.stop()
        self.gen_btn.config(state="normal")
        self.console.write(f"\n❌  Error: {msg}", "err")
        messagebox.showerror("Error", msg)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()