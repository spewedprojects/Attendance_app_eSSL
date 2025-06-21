import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys, os
import re
import json # For saving/loading settings

# This import might not be strictly necessary if styles are not directly used,
# but keeping it as it was in the original code.
from openpyxl.styles.builtins import styles

# ── Hi-DPI / crisp text on Windows ───────────────────────────────────
try:
    import ctypes

    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ──  Icon path (works after PyInstaller) ─────────────────────────────
APP_DIR = Path(getattr(sys, "_MEIPASS", os.getcwd()))
ICON_PATH = APP_DIR / "icon.ico"
SETTINGS_FILE = APP_DIR / "settings.json" # Path for settings file
guide_path = APP_DIR / "MasterSheet_Guide_Knowchem.pdf"

# ──  Import processing functions  ──────────────────────────────────
# Assuming these modules are available in the same directory or Python path
from clean_workduration_mod import clean_raw
from cleanup_2_mod import rectify_file
from assign_shifttimes_cleanedup_mod import add_shifts
from fill_master_shiftaware_mod import build_master  # This build_master needs to accept custom_shift_times

EXCEL_FILETYPES = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]

# --- Default Shift Times for Reset Functionality ---
# These are the defaults that will be used if no settings file is found
# or when the user clicks 'Reset'.
DEFAULT_SHIFT_TIMES = {
    "FS": "08:00 - 17:00",
    "GS": "09:30 - 18:30",
    "SS": "13:00 - 21:30",
    "NS": "20:00 - 08:00",
}

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
        • **Customizable Shift Times:** Added a menu option to allow users to define
          custom start times for FS, GS, SS, and NS shifts, overriding defaults.
        • **Integrated Instructions:** Added a menu option to open a PDF guide for users.
        • **Persistent Shift Settings:** Shift times entered by the user are now saved
          and loaded automatically on app launch.
        """

    def __init__(self):
        super().__init__()

        # ╭─ DPI‑aware window size ───────────────────────────────────╮
        # A tad taller for the new checkbox and potentially the menu bar
        base_w, base_h = 400, 310
        scale = float(self.tk.call("tk", "scaling"))
        self.geometry(f"{int(base_w * scale)}x{int(base_h * scale)}")
        self.resizable(True, True)
        version = "v3.6"  # Updated version to reflect changes
        self.title(f"Attendance Master Sheet Filler — {version}")

        if ICON_PATH.exists():
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                pass

        # Tk variables for GUI elements
        self.raw_var = tk.StringVar()
        self.shift_var = tk.StringVar()
        self.analyze_comments = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Ready")

        self.single_inp = tk.StringVar()
        self.extra_var = tk.StringVar()
        self.step_choice = tk.StringVar(value="clean")
        self.filter_var = tk.StringVar()  # For OT Filter

        # Tk variables for Custom shift time settings
        # Initialize with default values first
        self.shift_fs = tk.StringVar(value=DEFAULT_SHIFT_TIMES["FS"])
        self.shift_gs = tk.StringVar(value=DEFAULT_SHIFT_TIMES["GS"])
        self.shift_ss = tk.StringVar(value=DEFAULT_SHIFT_TIMES["SS"])
        self.shift_ns = tk.StringVar(value=DEFAULT_SHIFT_TIMES["NS"])

        self._load_settings() # Load settings after initializing StringVars

        self._build_menu()  # Add the menu bar
        self._build_widgets()

    # ── Settings Load/Save -------------------------------------------
    def _load_settings(self):
        """Loads shift times from settings.json."""
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
                self.shift_fs.set(settings.get("shift_fs", DEFAULT_SHIFT_TIMES["FS"]))
                self.shift_gs.set(settings.get("shift_gs", DEFAULT_SHIFT_TIMES["GS"]))
                self.shift_ss.set(settings.get("shift_ss", DEFAULT_SHIFT_TIMES["SS"]))
                self.shift_ns.set(settings.get("shift_ns", DEFAULT_SHIFT_TIMES["NS"]))
            except Exception as e:
                print(f"Error loading settings: {e}")
                # Optionally show a messagebox here, but might be annoying on every launch
                self.status.set("Warning: Could not load settings.")
        else:
            # If settings file doesn't exist, ensure default values are used (already set in __init__)
            pass

    def _save_settings(self):
        """Saves current shift times to settings.json."""
        settings = {
            "shift_fs": self.shift_fs.get(),
            "shift_gs": self.shift_gs.get(),
            "shift_ss": self.shift_ss.get(),
            "shift_ns": self.shift_ns.get(),
        }
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save settings: {e}")
            self.status.set("Error: Settings not saved.")

    # ── Menu Bar and Settings ----------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Edit Shift Times", command=self._open_shift_settings)
        settings_menu.add_command(label="View Instructions", command=self._open_instructions)
        menubar.add_cascade(label="Options", menu=settings_menu)

    def _open_shift_settings(self):
        win = tk.Toplevel(self)
        win.title("Edit Shift Start & End Times") # Updated title
        win.geometry("500x300") # Adjusted size for better fit
        win.transient(self)  # Make it appear on top of the main window
        win.grab_set()  # Make it modal

        pad = {"padx": 6, "pady": 4}

        ttk.Label(win, text="Enter times in HH:MM - HH:MM format").grid(
            row=0, column=0, columnspan=2, sticky="w", **pad
        )

        for i, (label_text, var) in enumerate([
            ("FS (First Shift):", self.shift_fs),
            ("GS (General Shift):", self.shift_gs),
            ("SS (Second Shift):", self.shift_ss),
            ("NS (Night Shift):", self.shift_ns),
        ], start=1):  # Start from row 1 to accommodate the instruction label
            ttk.Label(win, text=label_text).grid(row=i, column=0, sticky="w", **pad)
            ttk.Entry(win, textvariable=var, width=20).grid(row=i, column=1, **pad) # Wider entry field

        # Buttons for Apply and Reset
        button_frame = ttk.Frame(win)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="Apply", command=lambda: [self._save_settings(), win.destroy()]).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Reset", command=self._reset_shift_times).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=win.destroy).pack(side="left", padx=5) # 'Close' without saving

        self.wait_window(win)  # Wait for the Toplevel window to close

    def _reset_shift_times(self):
        """Resets shift time entry fields to their default hardcoded values."""
        self.shift_fs.set(DEFAULT_SHIFT_TIMES["FS"])
        self.shift_gs.set(DEFAULT_SHIFT_TIMES["GS"])
        self.shift_ss.set(DEFAULT_SHIFT_TIMES["SS"])
        self.shift_ns.set(DEFAULT_SHIFT_TIMES["NS"])
        messagebox.showinfo("Reset", "Shift times reset to defaults. Click Apply to save.")

    def _open_instructions(self):
        import subprocess
        # Assuming the PDF guide is in the same directory as this script
        guide_path = APP_DIR / "MasterSheet_Guide_Knowchem.pdf"

        if guide_path.exists():
            try:
                # 'start' for Windows, 'open' for macOS, 'xdg-open' for Linux
                if sys.platform == "win32":
                    os.startfile(guide_path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", guide_path])
                else:
                    subprocess.Popen(["xdg-open", guide_path])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open PDF: {e}")
        else:
            messagebox.showinfo("Missing",
                                f"Instructions PDF not found at:\n{guide_path}\nPlease ensure 'MasterSheet_Guide_Knowchem.pdf' is in the same folder as the executable.")

    def _get_custom_shift_dict(self):
        custom_shifts = {}
        # Pass the full HH:MM - HH:MM string to build_master
        for k, var in {
            "FS": self.shift_fs,
            "GS": self.shift_gs,
            "SS": self.shift_ss,
            "NS": self.shift_ns
        }.items():
            val = var.get().strip()
            # Basic validation for HH:MM - HH:MM format or single HH:MM
            if re.match(r"^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$", val) or re.match(r"^\d{1,2}:\d{2}$", val):
                custom_shifts[k] = val
        return custom_shifts if custom_shifts else None  # Return None if no valid custom shifts

    # ── UI layout (unchanged) ----------------------------------------
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

        # NEW checkbox (already in Code 1)
        ttk.Checkbutton(
            self,
            text="Analyze comments from shifts file",
            variable=self.analyze_comments,
        ).grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        # OT Filter (already in Code 1)
        ttk.Label(self, text="EmpCodes for OT table (comma-separated):").grid(
            row=6, column=0, sticky="w", **pad
        )
        ttk.Entry(self, textvariable=self.filter_var, width=55).grid(
            row=7, column=0, columnspan=2, sticky="ew", **pad
        )

        ttk.Button(self, text="Run full pipeline", command=self._run_full, width=25).grid(
            row=8, column=0, columnspan=2, pady=(6, 4)
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
        self._toggle_extra()  # Initialize state

        ttk.Button(self, text="Run selected step", command=self._run_single, width=25).grid(
            row=17, column=0, columnspan=2, pady=(6, 10)
        )

        ttk.Separator(self).grid(row=19, columnspan=2, sticky="ew")
        ttk.Label(self, textvariable=self.status).grid(row=20, column=0, columnspan=2)

    # ── Browse helpers (unchanged) ----------------------------------
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
            # When building master, this extra file is the save path for the output
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=EXCEL_FILETYPES
            )
        else:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=EXCEL_FILETYPES
            )
        if path:
            self.extra_var.set(path)

    # ── Toggle extra-file row based on radio selection (unchanged, but note 'False' for clean/rectify)
    def _toggle_extra(self):
        choice = self.step_choice.get()
        shown = choice in ("shift", "master") or (
                choice in ("rectify", "clean") and False
        )

        if shown:
            self.extra_lbl.grid(row=13, column=0, sticky="w", padx=6, pady=3)
            self.extra_ent.grid(row=14, column=0, sticky="ew", padx=6, pady=3)
            self.extra_btn.grid(row=14, column=1, padx=6, pady=3)
        else:
            self.extra_lbl.grid_remove()
            self.extra_ent.grid_remove()
            self.extra_btn.grid_remove()

    # ── Run buttons (unchanged logic for calling build_master with custom_shift_times) -------------------------------------------------
    def _run_full(self):
        raw = Path(self.raw_var.get())
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

            custom_times = self._get_custom_shift_dict()

            build_master(
                shifted,
                Path(save_to),
                analyze_comments=self.analyze_comments.get(),
                shifts_path=shift,
                ot_filter=filt or None,
                custom_shift_times=custom_times
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
                self.status.set("Cleaning raw…")
                clean_raw(inp)
                messagebox.showinfo("Done", "Raw file cleaned.")
            elif step == "rectify":
                self.status.set("Rectifying blanks…")
                rectify_file(inp)
                messagebox.showinfo("Done", "Blanks rectified.")
            elif step == "shift":
                self.status.set("Adding shifts…")
                shifts = Path(self.extra_var.get())
                if not shifts.exists():
                    messagebox.showwarning("Missing shifts", "Select the shifts file.")
                    return
                add_shifts(inp, shifts)
                messagebox.showinfo("Done", "Shifts added.")
            elif step == "master":
                self.status.set("Building master…")
                save_to = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", filetypes=EXCEL_FILETYPES
                )
                if not save_to:
                    self.status.set("Cancelled")
                    return
                filt = [c.strip() for c in self.filter_var.get().split(",") if c.strip()]

                custom_times = self._get_custom_shift_dict()

                build_master(
                    inp,
                    Path(save_to),
                    analyze_comments=self.analyze_comments.get(),
                    shifts_path=Path(self.extra_var.get()) if self.analyze_comments.get() else None,
                    ot_filter=filt or None,
                    custom_shift_times=custom_times
                )
                messagebox.showinfo("Done", f"Master sheet saved:\n{save_to}")
            self.status.set("✔ Completed")
        except Exception as e:
            self.status.set("❌ Failed")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    AttendanceGUI().mainloop()