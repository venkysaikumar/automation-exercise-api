from __future__ import annotations

import pandas as pd


def main() -> None:
    path = "postman_to_bruno.xlsx.xlsx"
    xl = pd.ExcelFile(path)
    print("sheets:", xl.sheet_names)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"\n=== {sheet} rows {len(df)} cols {len(df.columns)}")
        print("columns:", list(df.columns))
        print(df.head(20).to_string(index=False))


if __name__ == "__main__":
    main()

