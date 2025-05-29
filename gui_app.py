import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys, os

# ── Hi-DPI / crisp text on Windows ───────────────────────────────────
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ──  Icon path (works after PyInstaller) ─────────────────────────────
APP_DIR   = Path(getattr(sys, "_MEIPASS", os.getcwd()))
ICON_PATH = APP_DIR / "icon.ico"

# ──  Import processing functions  ──────────────────────────────────
from clean_workduration_mod           import clean_raw
from cleanup_2_mod                    import rectify_file
from assign_shifttimes_cleanedup_mod  import add_shifts
from fill_master_shiftaware_mod       import build_master


EXCEL_FILETYPES = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]


class AttendanceGUI(tk.Tk):
    """Tkinter front‑end for the attendance pipeline.

        NEW in this version (May 2025)
        ─────────────────────────────
        • Checkbox **“Analyse comments in shift notes”** – when ticked we pass the
          same shifts workbook a second time to `build_master`, enabling the new
          cell‑note parsing logic you added earlier.
        • The extra‑file selector (shown below the radio buttons) now appears when
          either the **Shift** step *or* the **Master** step + checkbox combo
          requires a second workbook.
        """

    def __init__(self):
        super().__init__()

        # ╭─ DPI‑aware window size ───────────────────────────────────╮
        base_w, base_h = 400, 310  # a tad taller for the new checkbox
        scale = float(self.tk.call("tk", "scaling"))
        self.geometry(f"{int(base_w * scale)}x{int(base_h * scale)}")
        self.resizable(True, True)
        version = "v3.2"
        self.title(f"Attendance Master Sheet Filler — {version}")

        if ICON_PATH.exists():
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                pass

        # Tk variables
        self.raw_var      = tk.StringVar()
        self.shift_var    = tk.StringVar()
        self.analyze_comments = tk.BooleanVar(value=False)  # NEW
        self.status       = tk.StringVar(value="Ready")

        self.single_inp   = tk.StringVar()
        self.extra_var    = tk.StringVar()
        self.step_choice  = tk.StringVar(value="clean")

        self._build_widgets()

    # ── UI layout ----------------------------------------------------
    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}

        # === FULL PIPELINE ==========================================
        ttk.Label(self, text="A)  Full pipeline").grid(row=0, column=0, sticky="w", **pad)

        ttk.Label(self, text="Raw work-duration file:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.raw_var, width=55, state="readonly").grid(
            row=2, column=0, sticky="ew", **pad
        )
        ttk.Button(self, text="Browse…", command=self._browse_raw).grid(row=2, column=1, **pad)

        ttk.Label(self, text="Shifts definition file:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.shift_var, width=55, state="readonly").grid(
            row=4, column=0, sticky="ew", **pad
        )
        ttk.Button(self, text="Browse…", command=self._browse_shift).grid(row=4, column=1, **pad)

        # NEW checkbox
        ttk.Checkbutton(
            self,
            text="Analyze comments from shifts file",
            variable=self.analyze_comments,
        ).grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        ttk.Button(self, text="Run full pipeline", command=self._run_full, width=25).grid(
            row=8, column=0, columnspan=2, pady=(6, 4)
        )

        # OT Filter — appears just *below* button and *above* separator
        ttk.Label(self, text="EmpCodes for OT table (comma-separated):").grid(
            row=6, column=0, sticky="w", **pad
        )
        self.filter_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.filter_var, width=55).grid(
            row=7, column=0, columnspan=2, sticky="ew", **pad
        )

        ttk.Separator(self).grid(row=9, columnspan=2, sticky="ew", pady=(2, 4))

        # === SINGLE-STEP ============================================
        ttk.Label(self, text="B)  Single-step utility").grid(row=10, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.single_inp, width=55, state="readonly").grid(
            row=11, column=0, sticky="ew", **pad
        )
        ttk.Button(self, text="Browse…", command=self._browse_single).grid(row=11, column=1, **pad)

        step_frame = ttk.Frame(self)
        step_frame.grid(row=12, column=0, columnspan=2, sticky="w", **pad)
        for text, val in [
            ("Clean raw", "clean"),
            ("Rectify blanks", "rectify"),
            ("Add shifts", "shift"),
            ("Build master", "master"),
        ]:
            ttk.Radiobutton(
                step_frame, text=text, value=val, variable=self.step_choice, command=self._toggle_extra
            ).pack(side="left", padx=4)

        # extra file (for shifts file or save‑path)
        self.extra_lbl = ttk.Label(self, text="Shifts file / Save-to:")
        self.extra_ent = ttk.Entry(self, textvariable=self.extra_var, width=55, state="readonly")
        self.extra_btn = ttk.Button(self, text="Browse…", command=self._browse_extra)
        self._toggle_extra()

        ttk.Button(self, text="Run selected step", command=self._run_single, width=25).grid(
            row=14, column=0, columnspan=2, pady=(6, 10)
        )

        ttk.Separator(self).grid(row=15, columnspan=2, sticky="ew")
        ttk.Label(self, textvariable=self.status).grid(row=16, column=0, columnspan=2)

    # ── Browse helpers ----------------------------------------------
    def _browse_raw(self):
        path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        if path:
            self.raw_var.set(path)

    def _browse_shift(self):
        path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        if path:
            self.shift_var.set(path)

    def _browse_single(self):
        path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        if path:
            self.single_inp.set(path)

    def _browse_extra(self):
        choice = self.step_choice.get()
        if choice in ("shift",):
            path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        elif choice in ("master",):
            path = filedialog.askopenfilename(filetypes=EXCEL_FILETYPES)
        else:  # save destination
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=EXCEL_FILETYPES
            )
        if path:
            self.extra_var.set(path)

    # ── Toggle extra-file row based on radio selection --------------
    def _toggle_extra(self):
        choice = self.step_choice.get()
        shown = choice in ("shift", "master") or (
            choice in ("rectify", "clean") and False
        )

        if shown:
            self.extra_lbl.grid(row=11, column=0, sticky="w", padx=6, pady=3)
            self.extra_ent.grid(row=12, column=0, sticky="ew", padx=6, pady=3)
            self.extra_btn.grid(row=12, column=1, padx=6, pady=3)
        else:
            self.extra_lbl.grid_remove()
            self.extra_ent.grid_remove()
            self.extra_btn.grid_remove()

    # ── Run buttons -------------------------------------------------
    def _run_full(self):
        raw   = Path(self.raw_var.get())
        shift = Path(self.shift_var.get())

        if not raw.exists() or not shift.exists():
            messagebox.showwarning("Missing file", "Please select both Raw and Shifts files.")
            return

        try:
            self.status.set("Cleaning raw…")
            cleaned = clean_raw(raw)

            self.status.set("Rectifying blanks…")
            fixed = rectify_file(cleaned)

            self.status.set("Adding shifts…")
            shifted = add_shifts(fixed, shift)

            save_to = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=EXCEL_FILETYPES,
                initialfile=f"{raw.stem}_MasterSheet.xlsx",
            )
            if not save_to:
                self.status.set("Cancelled")
                return

            self.status.set("Building master…")
            filt = [c.strip() for c in self.filter_var.get().split(",") if c.strip()]
            build_master(
                shifted,
                Path(save_to),
                analyze_comments=self.analyze_comments.get(),
                shifts_path=shift,
                ot_filter=filt or None,  # None ⇒ include everyone
            )

            self.status.set("✔ Full pipeline done")
            messagebox.showinfo("Success", f"Saved:\n{save_to}")
        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))

    def _run_single(self):
        inp = Path(self.single_inp.get())
        if not inp.exists():
            messagebox.showwarning("No file", "Please choose a file.")
            return

        try:
            step = self.step_choice.get()
            if step == "clean":
                clean_raw(inp)
                messagebox.showinfo("Done", "Raw file cleaned.")
            elif step == "rectify":
                rectify_file(inp)
                messagebox.showinfo("Done", "Blanks rectified.")
            elif step == "shift":
                shifts = Path(self.extra_var.get())
                if not shifts.exists():
                    messagebox.showwarning("Missing shifts", "Select the shifts file.")
                    return
                add_shifts(inp, shifts)
                messagebox.showinfo("Done", "Shifts added.")
            elif step == "master":
                save_to = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", filetypes=EXCEL_FILETYPES
                )
                if not save_to:
                    return
                filt = [c.strip() for c in self.filter_var.get().split(",") if c.strip()]
                build_master(
                    inp,
                    Path(save_to),
                    analyze_comments=self.analyze_comments.get(),
                    shifts_path=Path(self.extra_var.get()) if self.analyze_comments.get() else None,
                    ot_filter = filt or None  # ← add here too
                )
                messagebox.showinfo("Done", f"Master sheet saved:\n{save_to}")
            self.status.set("✔ Completed")
        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    AttendanceGUI().mainloop()
