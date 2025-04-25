import pandas as pd
from pathlib import Path

def rectify_pre_joining_blanks(cleaned_path: str, output_path: str):
    df = pd.read_excel(cleaned_path, header=None)

    for i in range(0, len(df), 9):
        emp_block = df.iloc[i:i+8, :]
        if emp_block.empty:
            continue
        status_row = emp_block[emp_block[2] == "Status"]
        if status_row.empty:
            continue
        status_idx = status_row.index[0]
        for col in range(3, df.shape[1]):
            col_data = emp_block.iloc[:, col]
            if col_data.isnull().all():
                continue
            if pd.isna(df.at[status_idx, col]):
                continue
            if not isinstance(df.at[status_idx, col], str):
                df.at[status_idx, col] = ""
    df.to_excel(output_path, index=False, header=False)

def rectify_file(cleaned_path: str | Path) -> Path:
    cleaned_path = Path(cleaned_path)
    out_path = cleaned_path.with_name(f"{cleaned_path.stem}_final.xlsx")
    rectify_pre_joining_blanks(str(cleaned_path), str(out_path))
    return out_path

if __name__ == "__main__":
    rectify_file("April25_WorkDurationReport_cleaned.xlsx")
