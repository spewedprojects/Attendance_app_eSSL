import pandas as pd
import xlrd
from pathlib import Path
from datetime import datetime, timedelta

def clean_workduration(input_path: str, output_path: str):
    wb = pd.ExcelFile(input_path)
    sheet = wb.sheet_names[0]
    raw = wb.parse(sheet, header=None)

    emp_rows = raw.index[raw[0] == "Employee:"].tolist()
    window_cell = raw.iat[2, 1]
    start_date = window_cell.split("  ")[0].strip()

    blocks = []
    for top in emp_rows:
        code, name = raw.iat[top, 3].split(":", 1)
        code = code.strip()
        name = name.strip()
        metrics = raw.iloc[top + 1: top + 9, :].reset_index(drop=True)
        metrics.index = ["Status", "InTime", "OutTime", "Duration", "Late By", "Early By", "OT", "Shift"]
        metrics = metrics.loc[:, ~metrics.isna().all()]
        metrics.insert(0, "EmpName", name)
        metrics.insert(0, "EmpCode", code)
        date_cols = [(datetime.strptime(start_date, "%b %d %Y") + timedelta(days=i)).strftime("%d-%b")
                     for i in range(metrics.shape[1] - 3)]
        metrics.columns = ["EmpCode", "EmpName", "metric"] + date_cols
        blocks.append(metrics)
        blocks.append(pd.DataFrame([[""] * metrics.shape[1]], columns=metrics.columns))

    final_df = pd.concat(blocks, ignore_index=True)
    final_df.to_excel(output_path, index=False)

def clean_raw(raw_path: str | Path) -> Path:
    raw_path = Path(raw_path)
    out_path = raw_path.with_name(f"{raw_path.stem}_cleaned.xlsx")
    clean_workduration(str(raw_path), str(out_path))
    return out_path

if __name__ == "__main__":
    clean_raw("April25_WorkDurationReport (1).xlsx")
