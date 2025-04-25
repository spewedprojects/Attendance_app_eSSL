import pandas as pd
from pathlib import Path
from datetime import datetime

# ── fixed shift-start lookup ───────────────────────────────────────────
shift_start = {
    "FS": "08:00",  "First":   "08:00",
    "GS": "09:30",  "General": "09:30",
    "SS": "11:30",  "Second":  "11:30",
    "NS": "20:30",  "Night":   "20:30",
}

def _parse_hhmm(val: str) -> datetime | None:
    try:
        return datetime.strptime(val.strip(), "%H:%M")
    except Exception:
        return None


# ── callable wrapper for GUI & pipeline ───────────────────────────────
def build_master(shifted_path: str | Path, save_path: str | Path) -> None:
    """Generate master-sheet with shift-aware late calculation."""
    shifted_path = Path(shifted_path)
    df = pd.read_excel(shifted_path)

    date_cols  = df.columns[3:]
    num_days   = len(date_cols)
    first_date = datetime.strptime(f"{date_cols[0]}-2025", "%d-%b-%Y")

    block_size  = 9     # 8 rows + 1 blank
    master_rows = []

    for sr, blk_top in enumerate(range(0, len(df), block_size), start=1):
        block = df.iloc[blk_top : blk_top + 8]

        if block.empty:      # skip separators at file end
            continue

        emp_id   = block.iloc[0]["EmpCode"]
        emp_name = block.iloc[0]["EmpName"]

        status_row = block.loc[block["metric"] == "Status"].iloc[0, 3:]
        in_row     = block.loc[block["metric"] == "InTime"].iloc[0, 3:]
        out_row    = block.loc[block["metric"] == "OutTime"].iloc[0, 3:]
        shift_row  = block.loc[block["metric"] == "Shift"].iloc[0, 3:]
        late_by    = block.loc[block["metric"] == "Late By"].iloc[0, 3:]

        present = leave = wo = hol = od1 = od2 = late = 0
        ot_hours = 0.0
        working_hours = []
        daily_status  = []

        # ── per-day loop ───────────────────────────────────────────
        for day in status_row.index:
            status = str(status_row[day]).strip()

            # -- shift-aware late check only for "P" ----------------
            late_flag = False
            if status == "P":
                shift_name      = str(shift_row[day]).strip()
                scheduled_start = shift_start.get(shift_name, "09:30")
                t_sched         = _parse_hhmm(scheduled_start)
                t_in            = _parse_hhmm(str(in_row[day]))

                if t_sched and t_in:
                    minutes_late = (t_in - t_sched).total_seconds() / 60
                    late_flag    = minutes_late > 15
                else:
                    # fallback to Late By column
                    late_str = str(late_by[day]).strip()
                    if late_str and late_str != "00:00":
                        lt = _parse_hhmm(late_str)
                        late_flag = lt and (lt.hour*60 + lt.minute) > 15

            # -- attendance + counters ------------------------------
            if status == "P":
                present += 1
                late    += int(late_flag)
                daily_status.append("L" if late_flag else "P")

            elif status == "A":
                leave += 1
                daily_status.append("A")

            elif status == "0.5P":
                present += 0.5
                daily_status.append("0.5P")

            elif status == "L":            # already flagged by biometric
                present += 1
                late    += 1
                daily_status.append("L")

            elif status == "OD1":
                present += 1
                od1 += 1
                daily_status.append("OD1")

            elif status == "OD2":
                present += 1
                od2 += 1
                daily_status.append("OD2")

            else:
                daily_status.append("")

        # ── working-hours / OT ------------------------------------
        for tin, tout in zip(in_row, out_row):
            t_in  = _parse_hhmm(str(tin))
            t_out = _parse_hhmm(str(tout))
            if t_in and t_out and t_out > t_in:
                wh = (t_out - t_in).seconds / 3600
                working_hours.append(wh)
                if wh > 8.5:
                    ot_hours += wh - 8.5

        avg_wh = round(sum(working_hours) / len(working_hours), 2) if working_hours else 0

        master_rows.append([
            sr, emp_id, emp_name, *daily_status,
            present, leave, 5, hol, num_days, "", 20, 19,
            od1, od2, avg_wh, round(ot_hours, 2), late, "", ""
        ])

    # ── header & save ──────────────────────────────────────────────
    cols = (
        ["Sr. no.", "Emp. ID", "Emp. name"] +
        [d.strftime("%d") for d in pd.date_range(first_date, periods=num_days)] +
        ["Present", "Leave", "W. off", "Holiday", "Total", "C-off",
         "TA adjusted", "TA", "OD1", "OD2", "Avg Working Hr",
         "OT", "Late marks", "Leave cut", "Remarks"]
    )

    pd.DataFrame(master_rows, columns=cols).to_excel(save_path, index=False)
    print(f"✅ MasterSheet saved → {save_path}")


# ── stand-alone test run (optional) ──────────────────────────────────
if __name__ == "__main__":
    build_master("April25_WorkDurationReport (1)_cleaned_shiftaware.xlsx",
                 "APRIL25_mastersheet_filled_shiftAwareLate.xlsx")
