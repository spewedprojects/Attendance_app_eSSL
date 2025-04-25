import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys, os

# ── HIGH-DPI / crisp text ─────────────────────────────────────────────
# Works on Win 10/11 with Tk 8.6  (silently ignored elsewhere)
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)      # per-monitor aware
except Exception:
    pass                                                # non-Windows or older Tk
# you can optionally tweak overall scaling if fonts still look big/small:
# ttk.tkinter.Tk().tk.call("tk", "scaling", 1.2)

# ──  Add your icon  ───────────────────────────────────────────────────
APP_DIR  = Path(getattr(sys, "_MEIPASS", os.getcwd()))  # _MEIPASS for PyInstaller
ICON_PATH = APP_DIR / "icon.ico"                        # place icon.ico next to .py

# ──  import your processing functions  ───────────────────────────────
from clean_workduration_mod import clean_raw
from cleanup_2_mod import rectify_file
from assign_shifttimes_cleanedup_mod import add_shifts
from fill_master_shiftaware_mod import build_master


class AttendanceGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Master Sheet Filler")
        # ── DPI-aware size ──────────────────────────────────────────────
        base_w, base_h = 390, 140  # the size that looked good before
        try:
            scale = float(self.tk.call("tk", "scaling"))  # e.g. 1.25, 1.5 …
        except Exception:
            scale = 1.0
        self.geometry(f"{int(base_w * scale)}x{int(base_h * scale)}")
        self.resizable(False, False)

        # set the window icon (if file exists)
        if ICON_PATH.exists():
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                # Linux / Mac need .png → fallback using iconphoto
                try:
                    tk_img = tk.PhotoImage(file=ICON_PATH)
                    self.iconphoto(True, tk_img)
                except Exception:
                    pass

        # ── variables & UI
        self.raw_var   = tk.StringVar()
        self.shift_var = tk.StringVar()
        self.status    = tk.StringVar(value="Ready")
        self._build_widgets()

    # ---------- UI ----------
    def _build_widgets(self):
        pad = {"padx": 6, "pady": 4}

        # Raw file
        ttk.Label(self, text="Raw work‑duration file:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.raw_var, width=55, state="readonly"
                  ).grid(row=1, column=0, **pad, sticky="ew")
        ttk.Button(self, text="Browse…", command=self._browse_raw).grid(row=1, column=1, **pad)

        # Shifts file
        ttk.Label(self, text="Shifts definition file:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.shift_var, width=55, state="readonly"
                  ).grid(row=3, column=0, **pad, sticky="ew")
        ttk.Button(self, text="Browse…", command=self._browse_shift).grid(row=3, column=1, **pad)

        # Run button
        ttk.Button(self, text="Run full pipeline", command=self._run, width=25
                  ).grid(row=4, column=0, columnspan=2, pady=(12,4))

        ttk.Separator(self).grid(row=5, columnspan=2, sticky="ew", pady=(4,2))
        ttk.Label(self, textvariable=self.status).grid(row=6, column=0, columnspan=2)

    # ---------- callbacks ----------
    def _browse_raw(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path: self.raw_var.set(path)

    def _browse_shift(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path: self.shift_var.set(path)

    def _run(self):
        raw_path   = Path(self.raw_var.get())
        shifts_path = Path(self.shift_var.get())

        if not raw_path.exists() or not shifts_path.exists():
            messagebox.showwarning("Missing file", "Select both Raw and Shifts files.")
            return

        try:
            # 1  clean
            self.status.set("Cleaning raw file…")
            cleaned = clean_raw(raw_path)

            # 2  rectify blanks
            self.status.set("Fixing blanks…")
            fixed = rectify_file(cleaned)

            # 3  add shifts
            self.status.set("Inserting shifts…")
            shifted = add_shifts(fixed, shifts_path)

            # 4  save master
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"{raw_path.stem}_MasterSheet.xlsx")
            if not save_path:
                self.status.set("Cancelled")
                return

            self.status.set("Building master sheet…")
            build_master(shifted, Path(save_path))
            self.status.set("✔ All done!")
            messagebox.showinfo("Success", f"Saved:\n{save_path}")

        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    AttendanceGUI().mainloop()
