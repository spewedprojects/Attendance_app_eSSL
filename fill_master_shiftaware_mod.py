import pandas as pd
from pathlib import Path
from datetime import datetime

# ── scheduled start per shift code (for late marks) ──────────────────
shift_start = {
    "FS": "08:00",  "First":   "08:00",
    "GS": "09:30",  "General": "09:30",
    "SS": "11:30",  "Second":  "11:30",
    "NS": "20:30",  "Night":   "20:30",
}

def _parse_hhmm(txt: str) -> datetime | None:
    try:
        return datetime.strptime(txt.strip(), "%H:%M")
    except Exception:
        return None

# ── main callable ────────────────────────────────────────────────────
def build_master(shifted_path: str | Path, save_path: str | Path) -> None:
    """
    Build master sheet with:
      • shift-aware late marks (>15 min)
      • half-day rule  (4.5 h ≤ WH < 5.5 h  ⇒ 0.5P)
      • OT rounding:  add 1 h if fractional ≥ 0.75 h
    """
    df = pd.read_excel(shifted_path)

    date_cols  = df.columns[3:]
    num_days   = len(date_cols)
    first_date = datetime.strptime(f"{date_cols[0]}-2025", "%d-%b-%Y")

    block_size = 9
    master_rows = []

    for sr, top in enumerate(range(0, len(df), block_size), start=1):
        block = df.iloc[top: top + 8]
        if block.empty:
            continue

        emp_id, emp_name = block.iloc[0][["EmpCode", "EmpName"]]

        status_row = block.loc[block["metric"] == "Status"].iloc[0, 3:]
        in_row     = block.loc[block["metric"] == "InTime" ].iloc[0, 3:]
        out_row    = block.loc[block["metric"] == "OutTime"].iloc[0, 3:]
        shift_row  = block.loc[block["metric"] == "Shift"  ].iloc[0, 3:]
        late_by    = block.loc[block["metric"] == "Late By"].iloc[0, 3:]

        present = leave = od1 = od2 = late = 0
        ot_hours = 0
        working_hours = []
        daily_status  = []

        # ── per-date loop ──────────────────────────────────────────
        for col in status_row.index:
            status = str(status_row[col]).strip()

            # --- worked hours for the day (if any) ---
            t_in  = _parse_hhmm(str(in_row[col]))
            t_out = _parse_hhmm(str(out_row[col]))
            wh = (t_out - t_in).seconds / 3600 if (t_in and t_out and t_out > t_in) else 0

            # --- half-day rule -----------------------
            if status == "P" and 4.5 <= wh < 5.5:
                status = "0.5P"

            # --- late mark (only full P) --------------
            late_flag = False
            if status == "P":
                shift_code = str(shift_row[col]).strip()
                sched_str  = shift_start.get(shift_code, "09:30")
                t_sched = _parse_hhmm(sched_str)
                if t_sched and t_in:
                    late_flag = (t_in - t_sched).total_seconds() / 60 > 15
                else:
                    late_txt = str(late_by[col]).strip()
                    if late_txt and late_txt != "00:00":
                        lt = _parse_hhmm(late_txt)
                        late_flag = lt and (lt.hour*60 + lt.minute) > 15

            # --- set master status & counters ----------
            if status == "P":
                present += 1
                late    += int(late_flag)
                daily_status.append("L" if late_flag else "P")
            elif status == "0.5P":
                present += 0.5
                daily_status.append("0.5P")
            elif status == "A":
                leave += 1
                daily_status.append("A")
            elif status == "L":
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

            # --- OT accumulation (rounded) -------------
            if wh and wh > 8.5:
                raw_ot = wh - 8.5
                ot_today = int(raw_ot) + (1 if raw_ot - int(raw_ot) >= 0.75 else 0)
                ot_hours += ot_today

            if wh:
                working_hours.append(wh)

        avg_wh = round(sum(working_hours) / len(working_hours), 2) if working_hours else 0

        master_rows.append([
            sr, emp_id, emp_name, *daily_status,
            present, leave, 5, 0, num_days, "", 20, 19,
            od1, od2, avg_wh, ot_hours, late, "", ""
        ])

    # ---- header & save -------------------------------------------
    cols = (
        ["Sr. no.", "Emp. ID", "Emp. name"] +
        [d.strftime("%d") for d in pd.date_range(first_date, periods=num_days)] +
        ["Present", "Leave", "W. off", "Holiday", "Total", "C-off",
         "TA adjusted", "TA", "OD1", "OD2", "Avg Working Hr",
         "OT", "Late marks", "Leave cut", "Remarks"]
    )

    pd.DataFrame(master_rows, columns=cols).to_excel(save_path, index=False)
    print(f"✅ MasterSheet saved → {save_path}")


# quick standalone run
if __name__ == "__main__":
    build_master("April25_WorkDurationReport (1)_cleaned_shiftaware.xlsx",
                 "APRIL25_master_filled_OTrounded.xlsx")
