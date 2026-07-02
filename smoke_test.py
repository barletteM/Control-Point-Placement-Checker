from pathlib import Path

import pandas as pd

from checker import ColumnMap, analyze_fieldbook, export_excel


def main() -> None:
    df = pd.read_csv(Path("sample_fieldbook.csv"))
    results = analyze_fieldbook(
        df,
        ColumnMap(
            point="Point Name",
            easting="Easting",
            northing="Northing",
            height="Height",
            solution="Solution",
        ),
    )
    assert len(results["Report"]) == 3
    assert len(results["Unmatched Measurements"]) == 1
    assert "OVERALL PASS" in set(results["Report"]["Final Status"])
    workbook = export_excel(results)
    assert workbook[:2] == b"PK"
    Path("sample_output_report.xlsx").write_bytes(workbook)
    print("Smoke test passed. Wrote sample_output_report.xlsx")


if __name__ == "__main__":
    main()
