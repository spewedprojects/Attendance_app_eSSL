import pandas as pd
from pathlib import Path
from datetime import datetime

def build_master(shifted_path: str | Path, save_path: str | Path) -> None:
    df = pd.read_excel(shifted_path)
    date_cols = df.columns[3:]
    num_days = len(date_cols)
    start_date = datetime.strptime(date_cols[0], "%d-%b")

    master_rows = []
    block_size = 9
    for i in range(0, len(df), block_size):
        block = df.iloc[i:i+8]
        emp_id = block.iloc[0]["EmpCode"]
        emp_name = block.iloc[0]["EmpName"]
        status_row = block[block["metric"] == "Status"].iloc[0, 3:]
        in_time = block[block["metric"] == "InTime"].iloc[0, 3:]
        out_time = block[block["metric"] == "OutTime"].iloc[0, 3:]
        late_by = block[block["metric"] == "Late By"].iloc[0, 3:]

        present = leave = late = od1 = od2 = 0
        ot_hours = 0.0
        working_hours = []
        daily_status = []

        for day in status_row.index:
            status = str(status_row[day]).strip()
            late_flag = False
            if status == "P":
                try:
                    lt = str(late_by[day]).strip()
                    if lt and lt != "00:00":
                        mins = datetime.strptime(lt, "%H:%M")
                        if mins.minute + mins.hour * 60 > 15:
                            late_flag = True
                except: pass
                if late_flag:
                    late += 1
                    present += 1
                    daily_status.append("L")
                else:
                    present += 1
                    daily_status.append("P")
            elif status == "A":
                leave += 1
                daily_status.append("A")
            elif status == "0.5P":
                present += 0.5
                daily_status.append("0.5P")
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

        for tin, tout in zip(in_time, out_time):
            try:
                tin = datetime.strptime(str(tin), "%H:%M")
                tout = datetime.strptime(str(tout), "%H:%M")
                wh = (tout - tin).seconds / 3600
                working_hours.append(wh)
                if wh > 8.5:
                    ot_hours += wh - 8.5
            except: continue

        avg_wh = round(sum(working_hours) / len(working_hours), 2) if working_hours else 0
        master_row = [
            i // 9 + 1, emp_id, emp_name, *daily_status,
            present, leave, 5, 0, num_days, "", 20, 19, od1, od2, avg_wh,
            round(ot_hours, 2), late, "", ""
        ]
        master_rows.append(master_row)

    columns = (
        ["Sr. no.", "Emp. ID", "Emp. name"] +
        [d.strftime("%d") for d in pd.date_range(start_date, periods=num_days)] +
        ["Present", "Leave", "W. off", "Holiday", "Total", "C-off", "TA adjusted", "TA",
         "OD1", "OD2", "Avg Working Hr", "OT", "Late marks", "Leave cut", "Remarks"]
    )

    pd.DataFrame(master_rows, columns=columns).to_excel(save_path, index=False)
    print(f"✅ MasterSheet saved → {save_path}")

if __name__ == "__main__":
    build_master("April25_WorkDurationReport_cleaned_shiftaware.xlsx",
                 "April25_MasterSheet_Filled.xlsx")
