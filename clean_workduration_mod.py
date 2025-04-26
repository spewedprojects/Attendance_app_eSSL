import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

def _parse_hhmm(txt: str) -> datetime | None:
    try:
        return datetime.strptime(txt.strip(), "%H:%M")
    except Exception:
        return None

def clean_workduration(input_path: str, output_path: str):
    wb    = pd.ExcelFile(input_path)
    sheet = wb.sheet_names[0]
    raw   = wb.parse(sheet, header=None)

    emp_rows    = raw.index[raw[0] == "Employee:"].tolist()
    start_date  = raw.iat[2, 1].split("  ")[0].strip()    # e.g. "Mar 01 2025"

    blocks = []
    for top in emp_rows:
        code, name = raw.iat[top, 3].split(":", 1)
        code, name = code.strip(), name.strip()

        metrics = raw.iloc[top + 1 : top + 9, :].reset_index(drop=True)
        metrics.index = ["Status","InTime","OutTime","Duration",
                         "Late By","Early By","OT","Shift"]
        metrics = metrics.loc[:, ~metrics.isna().all()]

        # ---- half-day rule --------------------------------------------------
        status_row = metrics.loc["Status"]
        in_row     = metrics.loc["InTime"]
        out_row    = metrics.loc["OutTime"]

        # normalise any literal ½P first
        status_row.replace("½P", "0.5P", inplace=True)

        for col in status_row.index[3:]:                # skip EmpCode, EmpName, metric
            if status_row[col] != "P":
                continue
            t_in  = _parse_hhmm(str(in_row[col]))
            t_out = _parse_hhmm(str(out_row[col]))
            if t_in and t_out and t_out > t_in:
                worked = (t_out - t_in).seconds / 3600
                if 4.0 <= worked < 5.5:                 # half-day window
                    status_row[col] = "0.5P"

        # ---------------------------------------------------------------------
        metrics.insert(0, "EmpName", name)
        metrics.insert(0, "EmpCode", code)

        date_cols = [(datetime.strptime(start_date, "%b %d %Y") + timedelta(days=i)).strftime("%d-%b")
                     for i in range(metrics.shape[1] - 3)]
        metrics.columns = ["EmpCode", "EmpName", "metric"] + date_cols

        blocks.append(metrics)
        blocks.append(pd.DataFrame([[""] * metrics.shape[1]], columns=metrics.columns))

    pd.concat(blocks, ignore_index=True).to_excel(output_path, index=False)

def clean_raw(raw_path: str | Path) -> Path:
    raw_path = Path(raw_path)
    out_path = raw_path.with_name(f"{raw_path.stem}_cleaned.xlsx")
    clean_workduration(str(raw_path), str(out_path))
    return out_path

# stand-alone test -------------------------------------------------------------
if __name__ == "__main__":
    clean_raw("April25_WorkDurationReport (1).xlsx")
