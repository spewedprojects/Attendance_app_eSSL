import pandas as pd
from pathlib import Path

# short-label mapping you used in the original
_shift_labels = {
    "First":   "FS",
    "General": "GS",
    "Second":  "SS",
    "Night":   "NS",
}

def add_shifts(cleaned_path: str | Path, shifts_path: str | Path) -> Path:
    """
    Overwrite every employee's 'Shift' row in the cleaned attendance file
    using the daily shift labels from the shifts definition file.

    Returns the Path of the new *shift-aware* attendance file.
    """
    cleaned_path = Path(cleaned_path)
    shifts_path  = Path(shifts_path)

    # 1. load both files (no headers)
    att_df   = pd.read_excel(cleaned_path, header=None)
    shift_df = pd.read_excel(shifts_path, header=None, skiprows=1)

    # 2. build lookup: EmpCode → [FS, FS, GS, ...]
    emp_codes        = shift_df[1]               # code column in shifts file
    daily_shift_cols = shift_df.iloc[:, 3:]      # daily shifts start at col 3
    short_labels_df  = daily_shift_cols.applymap(
        lambda x: _shift_labels.get(str(x).strip(), str(x).strip())
    )
    shift_map = {
        str(code).strip(): list(short_labels_df.iloc[i])
        for i, code in enumerate(emp_codes)
    }

    # 3. iterate employee blocks in attendance sheet
    date_col_start = 3                # first date column in attendance sheet
    max_dates      = att_df.shape[1] - date_col_start
    block_stride   = 9                # 8 data rows + 1 blank separator
    first_block    = 1                # skip header row

    for block_top in range(first_block, len(att_df), block_stride):
        shift_row = block_top + 7     # "Shift" is the 8-th row inside block
        if shift_row >= len(att_df):
            break

        emp_code = str(att_df.iat[block_top, 0]).strip()
        if emp_code in shift_map:
            labels = shift_map[emp_code][:max_dates]
            att_df.iloc[shift_row,
                         date_col_start : date_col_start + len(labels)
                       ] = labels

    # 4. save & return new filename
    out_path = cleaned_path.with_name(f"{cleaned_path.stem}_shiftaware.xlsx")
    att_df.to_excel(out_path, index=False, header=False)
    print(f"✅ Shift rows updated → {out_path}")
    return out_path


# Stand-alone run (optional)
if __name__ == "__main__":
    add_shifts("April25_WorkDurationReport_cleaned_final.xlsx",
               "APRIL_shifttimes.xlsx")
