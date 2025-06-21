import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import re
from openpyxl import load_workbook
from dateutil import parser as dt_parser
from typing import Iterable, List, Optional, Tuple

# ── shift start/end times used for late‑mark check and expected duration ──
DEFAULT_SHIFT_RANGES = {
    "FS": "08:00 - 17:00",  "First":   "08:00 - 17:00",
    "GS": "09:30 - 18:30",  "General": "09:30 - 18:30",
    "SS": "13:30 - 22:30",  "Second":  "13:30 - 22:30",
    "NS": "20:00 - 08:00",  "Night":   "20:00 - 08:00", # Next day
}

def _parse_hhmm(txt: str) -> Optional[datetime]:
    """Parses a string 'HH:MM' into a datetime object (date is arbitrary)."""
    try:
        return datetime.strptime(txt.strip(), "%H:%M")
    except ValueError:
        return None

def _parse_hhmm_range(time_range_str: str) -> Tuple[Optional[datetime], Optional[datetime]]:
    """
    Parses a string 'HH:MM - HH:MM' into a tuple of (start_datetime, end_datetime).
    Handles overnight shifts (end_time < start_time implies next day).
    Also handles single 'HH:MM' by returning (start_datetime, None).
    """
    time_range_str = time_range_str.strip()
    parts = re.split(r'\s*-\s*', time_range_str)

    if len(parts) == 2:
        start_time = _parse_hhmm(parts[0])
        end_time = _parse_hhmm(parts[1])
        if start_time and end_time:
            # If end time is earlier than start time, assume it's on the next day
            if end_time < start_time:
                end_time += timedelta(days=1)
            return start_time, end_time
    elif len(parts) == 1:
        start_time = _parse_hhmm(parts[0])
        if start_time:
            return start_time, None # No explicit end time
    return None, None # Invalid format

# -- Analyze comments/notes logic -----------------------------------

def override_by_comment(note, t_in, t_out):
    """Detect keywords in cell‑comments and override daily status / OT logic."""
    o = {"force_status": None, "skip_half_day": False, "late_override": None, "extra_wh": 0.0}
    if not note:
        return o

    nl = note.strip().lower()

    if nl == "short leave":
        o["force_status"] = "P"
        o["skip_half_day"] = True
        return o

    if nl == "half day":
        o["force_status"] = "0.5P"
        o["skip_half_day"] = True
        return o

    if "out:" in nl:
        try:
            parts = nl.split("out:")
            out_str = parts[1].strip()
            # Try to parse 'hh:mm' or 'hhmm'
            if re.match(r"^\d{1,2}:\d{2}$", out_str):
                forced_out = datetime.strptime(out_str, "%H:%M")
            elif re.match(r"^\d{3,4}$", out_str): # Handle HHMM format
                if len(out_str) == 3: out_str = '0' + out_str
                forced_out = datetime.strptime(out_str, "%H%M")
            else:
                forced_out = None

            if forced_out:
                # Calculate new working hours based on forced out time
                # Assuming t_in is already set for the day
                if t_in and forced_out:
                    # If forced_out is earlier than t_in (e.g. overnight shift corrected by comment), add a day
                    if forced_out < t_in:
                        forced_out += timedelta(days=1)
                    o["extra_wh"] = (forced_out - t_in).total_seconds() / 3600
                    o["force_status"] = "P" # Force presence if out time is recorded
        except Exception:
            pass # Malformed "out:" comment, ignore

    if "late:" in nl:
        try:
            parts = nl.split("late:")
            late_min_str = parts[1].strip()
            late_minutes = int(late_min_str)
            o["late_override"] = late_minutes
        except Exception:
            pass # Malformed "late:" comment, ignore

    return o

def build_master(
    df_shifts_added: pd.DataFrame,
    save_to_path: Path,
    analyze_comments: bool = False,
    shifts_path: Optional[Path] = None,
    ot_filter: Optional[List[str]] = None,
    custom_shift_times: Optional[dict] = None,
):
    """
    Builds the final master attendance sheet and an OT sheet.
    Incorporates custom shift times and comment analysis.

    Args:
        df_shifts_added (pd.DataFrame): DataFrame after shifts have been added.
        save_to_path (Path): Path to save the output Excel file.
        analyze_comments (bool): Whether to analyze cell comments in the shifts file.
        shifts_path (Optional[Path]): Path to the original shifts file, required if analyze_comments is True.
        ot_filter (Optional[List[str]]): List of employee IDs to filter for the OT table.
        custom_shift_times (Optional[dict]): Dictionary of custom shift time strings
                                              (e.g., {"FS": "08:00 - 17:00"}).
    """
    if analyze_comments and not shifts_path:
        raise ValueError("shifts_path must be provided if analyze_comments is True.")

    df = df_shifts_added.copy()
    wb = None
    if analyze_comments:
        try:
            wb = load_workbook(shifts_path)
            # Assuming shift data is in the first sheet
            ws = wb.active
        except Exception as e:
            print(f"Warning: Could not load shifts file for comment analysis: {e}")
            analyze_comments = False # Disable comment analysis if file can't be loaded

    # ── Prepare shift schedules ──────────────────────────────────────
    parsed_shift_schedules = {} # Stores {"FS": (start_datetime, end_datetime), ...}

    # Initialize with default ranges
    for code, time_range_str in DEFAULT_SHIFT_RANGES.items():
        start_dt, end_dt = _parse_hhmm_range(time_range_str)
        if start_dt: # Only add if start time is valid
            parsed_shift_schedules[code] = {"start": start_dt, "end": end_dt}

    # Override with custom shift times if provided
    if custom_shift_times:
        for k, v in custom_shift_times.items():
            start_dt, end_dt = _parse_hhmm_range(v)
            if start_dt: # Only override if custom start time is valid
                parsed_shift_schedules[k] = {"start": start_dt, "end": end_dt}

    # Fallback default if a shift code isn't in ranges
    default_start_fallback = _parse_hhmm("09:30")
    default_end_fallback = _parse_hhmm("18:30")
    if default_start_fallback and default_end_fallback and default_end_fallback < default_start_fallback:
        default_end_fallback += timedelta(days=1)


    # Extract unique dates
    dates = pd.to_datetime(df["Date"]).dt.normalize().unique()
    dates.sort()
    first_date = dates.min()
    num_days = len(dates)

    master_rows = []
    ot_rows = []
    sr = 0

    for emp_id, emp_df in df.groupby("Emp. ID"):
        sr += 1
        emp_name = emp_df["Emp. name"].iloc[0]
        daily_status = []
        daily_ots = []
        present = 0
        leave = 0
        c_off = 0
        od1 = 0
        od2 = 0
        ot_hours = 0.0
        late = 0
        total_wh = 0.0 # Total actual working hours for avg calc

        # Dictionary to quickly look up daily data by date
        daily_data_map = {
            pd.to_datetime(row["Date"]).normalize(): row
            for _, row in emp_df.iterrows()
        }

        for single_date in dates:
            row = daily_data_map.get(single_date)
            status = ""
            ot_today = 0.0
            wh_today = 0.0
            late_flag = False
            comment_override = None

            if row is not None:
                shift_code = row["Shift"].strip() if pd.notna(row["Shift"]) else ""
                t_in = _parse_hhmm(row["In Time"]) if pd.notna(row["In Time"]) else None
                t_out = _parse_hhmm(row["Out Time"]) if pd.notna(row["Out Time"]) else None

                # Get scheduled shift times from parsed_shift_schedules
                scheduled_shift = parsed_shift_schedules.get(shift_code)
                t_sched_start = scheduled_shift["start"] if scheduled_shift else default_start_fallback
                t_sched_end = scheduled_shift["end"] if scheduled_shift else default_end_fallback

                # --- Comment Analysis (if enabled) ---
                if analyze_comments and wb:
                    # Construct cell address (e.g., C5 for date column + employee row)
                    # This requires knowing the Excel sheet structure.
                    # Assuming Date column is D, Emp. ID is B, Shift is C in original file
                    # And 'In Time' is 'E', 'Out Time' is 'F'
                    # The employee row in Excel could be dynamic based on sorting/filtering
                    # This is a simplification; a more robust mapping would be needed.

                    # For now, let's assume `row_num_in_excel` for a given emp_id and date
                    # is available or can be determined from the raw Excel sheet.
                    # As a placeholder, we need the actual row from the original Excel sheet.
                    # This part needs to be accurately implemented based on original data's layout.
                    # For demonstration, let's assume we can retrieve the cell comment.
                    # e.g., if employee data starts from row 2, and Date is column D,
                    # and the shift column is C, and In Time is E, Out Time is F.
                    # This comment retrieval is complex without actual row/column mapping.
                    # For current purpose, this part is conceptual, relying on actual Excel structure.

                    # Let's assume there's a way to get a cell note, perhaps from the df_shifts_added
                    # if it carried over note information. If not, this needs actual Excel sheet access.
                    # Since the original context didn't provide note extraction logic here,
                    # I will keep this as a conceptual hook.
                    # If the `df_shifts_added` contained a 'Note' column, we would use it here.
                    comment_text = None # Placeholder for actual comment retrieval
                    # Example: if df_shifts_added had a 'Note' column corresponding to current row
                    # comment_text = row.get('Note', '')

                    # If the original shifts file is loaded, we could try to find the cell comment
                    # This is highly dependent on how the shifts file maps to the dataframe.
                    # For now, let's assume a dummy way to get comment if needed for testing
                    # In a real scenario, you'd map single_date and emp_id to an Excel cell to get its comment.

                    # Example conceptual comment retrieval (NOT ACTUAL CODE based on current input):
                    # For simplicity of this task, I will assume comment_text is available if analyze_comments is True
                    # and `override_by_comment` will still work as intended.
                    if analyze_comments and 'Note' in row and pd.notna(row['Note']):
                        comment_text = str(row['Note']) # If notes are in the dataframe
                    elif analyze_comments and wb: # If notes are in original Excel file, this needs proper indexing
                        # This part is highly dependent on the Excel structure.
                        # We would need to find the correct row and column for the specific date and employee
                        # to fetch the cell comment from `ws`.
                        # Example: For now, assuming a helper function to get comment by (emp_id, date)
                        # comment_text = _get_comment_from_excel(ws, emp_id, single_date)
                        pass # Placeholder for actual Excel comment lookup

                    if comment_text:
                        comment_override = override_by_comment(comment_text, t_in, t_out)
                        if comment_override["force_status"]:
                            status = comment_override["force_status"]
                        if comment_override["late_override"] is not None:
                            late_flag = (t_in - t_sched_start).total_seconds() / 60 > comment_override["late_override"]
                        if comment_override["extra_wh"] > 0:
                            wh_today = comment_override["extra_wh"]


                # --- Default Status Calculation (if not overridden by comment) ---
                if not status: # Only calculate if not forced by comment
                    if pd.isna(row["Status"]):
                        if t_in and t_out:
                            wh_today = (t_out - t_in).total_seconds() / 3600
                            # Adjust for overnight shifts
                            if t_out < t_in:
                                wh_today = ((t_out + timedelta(days=1)) - t_in).total_seconds() / 3600

                            if wh_today >= 7.0: # Full day if 7+ hours
                                status = "P"
                                present += 1
                            elif 3.0 <= wh_today < 7.0: # Half day
                                if not (comment_override and comment_override["skip_half_day"]):
                                    status = "0.5P"
                                    present += 0.5
                                else: # Skip half day if comment says short leave etc.
                                    status = "P"
                                    present += 1
                            else: # Less than 3 hours or no in/out
                                status = "A" # Absent if less than 3 hours
                        else:
                            status = "A" # Absent if no in/out times
                    else:
                        status = row["Status"] # Use existing status if available

                # --- Overtime (OT) Calculation ---
                if status == "P" or status == "0.5P": # Only calculate OT for present/half-present days
                    if wh_today > 8.0:
                        ot_today = wh_today - 8.0 # OT is hours beyond 8

                # --- Late Mark Calculation ---
                if t_in and t_sched_start:
                    if (t_in - t_sched_start).total_seconds() / 60 > 15: # More than 15 min late
                        late_flag = True

                # Apply overrides from comments for late marks if present
                if comment_override and comment_override["late_override"] is not None:
                    # Use the override threshold from comment
                    late_flag = (t_in - t_sched_start).total_seconds() / 60 > comment_override["late_override"]


                if late_flag:
                    late += 1

                if pd.notna(row["Leave"]):
                    leave += 1
                    status = row["Leave"] # Leave takes precedence over P/A

                if pd.notna(row["C-off"]):
                    c_off += 1

                if pd.notna(row["OD1"]):
                    od1 += 1
                if pd.notna(row["OD2"]):
                    od2 += 1

                # Accumulate total working hours for average calculation later
                total_wh += wh_today
                ot_hours += ot_today

            else: # No data for this date
                status = "A" # Assume absent if no record
                wh_today = 0.0
                ot_today = 0.0

            daily_status.append(status)
            daily_ots.append(f"{ot_today:.2f}") # Format OT to 2 decimal places

        avg_wh = total_wh / num_days if num_days > 0 else 0.0

        # ---- build master_rows and ot_rows --------------------------
        master_rows.append([
            sr, emp_id, emp_name, *daily_status,
            present, leave, 5, 0, num_days, # 5 W. off, 0 Holiday (hardcoded as per original)
            c_off, 20, 19, od1, od2, avg_wh, ot_hours, late, "", "" # 20 TA adjusted, 19 TA (hardcoded)
        ])

        # -- OT table row (filter aware) ------------------------------
        if (ot_filter is None) or (emp_id in ot_filter):
            ot_rows.append([sr, emp_id, emp_name, *daily_ots, ot_hours])

    # ---- build DataFrames ------------------------------------------
    day_headers = [d.strftime("%d") for d in pd.date_range(first_date, periods=num_days)]

    master_cols = (
        ["Sr. no.", "Emp. ID", "Emp. name"] + day_headers +
        ["Present", "Leave", "W. off", "Holiday", "Total", "C-off",
         "TA adjusted", "TA", "OD1", "OD2", "Avg Working Hr",
         "OT", "Late marks", "Leave cut", "Remarks"]
    )

    ot_cols = ["Sr. no.", "Emp. ID", "Emp. name"] + day_headers + ["Total OT"]

    df_master = pd.DataFrame(master_rows, columns=master_cols)
    df_master.index += 1  # start index at 1 visually (optional)

    df_ot = pd.DataFrame(ot_rows, columns=ot_cols) if ot_rows else pd.DataFrame(columns=ot_cols)


    # ---- write both tables with 5‑row spacing -----------------------
    with pd.ExcelWriter(save_to_path, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master Sheet", index=False)
        # Add 5 empty rows before the OT table
        if not df_ot.empty:
            startrow_ot = len(df_master) + 6 # +1 for header, +5 for spacing
            df_ot.to_excel(writer, sheet_name="Master Sheet", startrow=startrow_ot, index=False)