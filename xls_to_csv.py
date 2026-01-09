"""
Convert the provided `opendata.xls` workbook into a CSV file.

Assumptions about the source file:
- There is a single relevant sheet (first sheet).
- The header row starts on Excel row 6 (zero-indexed row 5).
- Rows that are completely empty should be dropped.

Output:
- Writes `opendata.csv` in the same directory.
"""

from pathlib import Path
import pandas as pd


def xls_to_csv(
    source_path: Path = Path("opendata.xls"),
    output_path: Path = Path("opendata.csv"),
    header_row: int = 5,
) -> None:
    """
    Read an XLS file and export the first sheet to CSV.

    :param source_path: Path to the input .xls file.
    :param output_path: Path where the CSV will be written.
    :param header_row: Zero-based index of the row containing column headers.
    """
    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    # Read the first sheet; pandas will pick xlrd engine for legacy .xls.
    df = pd.read_excel(source_path, sheet_name=0, header=header_row)

    # Drop rows that are entirely empty.
    df = df.dropna(how="all")

    # Save to CSV without the pandas index.
    df.to_csv(output_path, index=False)

    print(f"Wrote {len(df)} rows to {output_path}")


if __name__ == "__main__":
    xls_to_csv()
