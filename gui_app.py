import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys, os

# ── Hi-DPI / crisp text ──────────────────────────────────────────────
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)   # per-monitor aware
except Exception:
    pass

# ──  Icon path (works after PyInstaller) ─────────────────────────────
APP_DIR   = Path(getattr(sys, "_MEIPASS", os.getcwd()))
ICON_PATH = APP_DIR / "icon.ico"

# ──  Import processing functions  ────────────────────────────────────
from clean_workduration_mod           import clean_raw
from cleanup_2_mod                    import rectify_file
from assign_shifttimes_cleanedup_mod  import add_shifts
from fill_master_shiftaware_mod       import build_master


class AttendanceGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        # ── DPI-aware window size ────────────────────────────────────
        base_w, base_h = 400, 270
        scale = float(self.tk.call("tk", "scaling"))
        self.geometry(f"{int(base_w*scale)}x{int(base_h*scale)}")
        self.resizable(True, True)
        self.title("Attendance Master Sheet Filler")

        if ICON_PATH.exists():
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                try:
                    self.iconphoto(True, tk.PhotoImage(file=ICON_PATH))
                except: pass

        # ── Tk variables ─────────────────────────────────────────────
        self.raw_var      = tk.StringVar()
        self.shift_var    = tk.StringVar()
        self.status       = tk.StringVar(value="Ready")

        # single-step vars
        self.single_inp   = tk.StringVar()
        self.extra_var    = tk.StringVar()
        self.step_choice  = tk.StringVar(value="clean")

        self._build_widgets()

    # ── UI layout ────────────────────────────────────────────────────
    def _build_widgets(self):
        pad = {"padx": 6, "pady": 3}

        # ===  FULL PIPELINE  =================================================
        ttk.Label(self, text="A)  Full pipeline").grid(row=0, column=0, sticky="w", **pad)

        ttk.Label(self, text="Raw work-duration file:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.raw_var, width=55, state="readonly"
                 ).grid(row=2, column=0, sticky="ew", **pad)
        ttk.Button(self, text="Browse…", command=self._browse_raw
                 ).grid(row=2, column=1, **pad)

        ttk.Label(self, text="Shifts definition file:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.shift_var, width=55, state="readonly"
                 ).grid(row=4, column=0, sticky="ew", **pad)
        ttk.Button(self, text="Browse…", command=self._browse_shift
                 ).grid(row=4, column=1, **pad)

        ttk.Button(self, text="Run full pipeline", command=self._run_full, width=25
                 ).grid(row=5, column=0, columnspan=2, pady=(6, 10))

        ttk.Separator(self).grid(row=6, columnspan=2, sticky="ew", pady=4)

        # ===  SINGLE-STEP  ===================================================
        ttk.Label(self, text="B)  Single-step utility").grid(row=7, column=0, sticky="w", **pad)

        ttk.Entry(self, textvariable=self.single_inp, width=55, state="readonly"
                 ).grid(row=8, column=0, sticky="ew", **pad)
        ttk.Button(self, text="Browse…", command=self._browse_single
                 ).grid(row=8, column=1, **pad)

        # radio buttons for step choice
        step_frame = ttk.Frame(self)
        step_frame.grid(row=9, column=0, columnspan=2, sticky="w", **pad)
        for text, val in [("Clean raw", "clean"),
                          ("Rectify blanks", "rectify"),
                          ("Add shifts", "shift"),
                          ("Build master", "master")]:
            ttk.Radiobutton(step_frame, text=text, value=val,
                            variable=self.step_choice,
                            command=self._toggle_extra).pack(side="left", padx=4)

        # extra file (for shifts file or save-path)
        self.extra_lbl = ttk.Label(self, text="Shifts file / Save-to:")
        self.extra_ent = ttk.Entry(self, textvariable=self.extra_var,
                                   width=55, state="readonly")
        self.extra_btn = ttk.Button(self, text="Browse…", command=self._browse_extra)

        # initially hidden
        self._toggle_extra()

        ttk.Button(self, text="Run selected step", command=self._run_single, width=25
                 ).grid(row=12, column=0, columnspan=2, pady=(6, 10))

        ttk.Separator(self).grid(row=13, columnspan=2, sticky="ew", pady=(2,4))
        ttk.Label(self, textvariable=self.status).grid(row=14, column=0, columnspan=2)

    # ── helper: show / hide extra file widgets ───────────────────────
    def _toggle_extra(self):
        need_extra = (self.step_choice.get() == "shift")
        if need_extra:
            self.extra_lbl.grid(row=10, column=0, sticky="w", padx=6, pady=(2,0))
            self.extra_ent.grid(row=11, column=0, sticky="ew", padx=6)
            self.extra_btn.grid(row=11, column=1, padx=6)
        else:
            self.extra_lbl.grid_remove()
            self.extra_ent.grid_remove()
            self.extra_btn.grid_remove()

    # ── browse helpers ───────────────────────────────────────────────
    def _browse_raw(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p: self.raw_var.set(p)

    def _browse_shift(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p: self.shift_var.set(p)

    def _browse_single(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p: self.single_inp.set(p)

    def _browse_extra(self):
        if self.step_choice.get() == "shift":
            p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if p: self.extra_var.set(p)

    # ── FULL PIPELINE callback ───────────────────────────────────────
    def _run_full(self):
        raw   = Path(self.raw_var.get())
        shift = Path(self.shift_var.get())

        if not raw.exists() or not shift.exists():
            messagebox.showwarning("Missing file",
                                   "Select both Raw and Shifts files.")
            return

        try:
            self.status.set("Cleaning raw…")
            cleaned = clean_raw(raw)

            self.status.set("Rectifying blanks…")
            fixed   = rectify_file(cleaned)

            self.status.set("Adding shifts…")
            shifted = add_shifts(fixed, shift)

            save_to = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"{raw.stem}_MasterSheet.xlsx")
            if not save_to:
                self.status.set("Cancelled")
                return

            self.status.set("Building master…")
            build_master(shifted, Path(save_to))

            self.status.set("✔ Full pipeline done")
            messagebox.showinfo("Success", f"Saved:\n{save_to}")
        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))

    # ── SINGLE-STEP callback ────────────────────────────────────────
    def _run_single(self):
        step = self.step_choice.get()
        inp  = Path(self.single_inp.get())

        if not inp.exists():
            messagebox.showwarning("No input file", "Select an input file.")
            return

        try:
            if step == "clean":
                out = clean_raw(inp)
            elif step == "rectify":
                out = rectify_file(inp)
            elif step == "shift":
                shifts = Path(self.extra_var.get())
                if not shifts.exists():
                    messagebox.showwarning("Need shifts file",
                                           "Select a shifts file."); return
                out = add_shifts(inp, shifts)
            elif step == "master":
                save_to = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel", "*.xlsx")],
                    initialfile=f"{inp.stem}_MasterSheet.xlsx")
                if not save_to:
                    self.status.set("Cancelled"); return
                build_master(inp, Path(save_to))
                out = save_to
            else:
                return

            self.status.set(f"✔ Done → {out}")
            if step != "master":
                messagebox.showinfo("Done", f"Output saved:\n{out}")
        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    AttendanceGUI().mainloop()
