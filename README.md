# Control Point Placement Checker

A Streamlit desktop/web app for checking measured GPS fieldbook observations against control point rows.

## What it does

- Imports CSV and Excel fieldbooks.
- Classifies rows with `Solution = NONE` as control points.
- Treats all other rows as measured observations.
- Matches measured observations to controls by point name, including repeated measured names.
- Selects the ground position check by smallest horizontal displacement.
- Selects the nail level check by smallest height difference.
- Flags position, height, and overall status against editable tolerances.
- Exports a formatted Excel workbook with report, summary, failed points, unmatched measurements, raw classified data, and row-level errors.
- Optionally exports cleaned/classified CSV data.

## Run locally

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

If you are using the bundled Codex runtime in this workspace, run:

```powershell
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m streamlit run app.py
```

## Expected columns

The app auto-detects common spellings:

- Point name: `Point Name`, `Point`, `Name`, `Pt`, `Station`, `Mark`
- Easting: `Easting`, `E`, `East`, `X`
- Northing: `Northing`, `N`, `North`, `Y`
- Height: `Height`, `Elevation`, `Z`, `RL`, `Level`
- Solution: `Solution`, `Fix Status`, `Fix`, `Status`, `Quality`

If detection is uncertain, select the correct columns in the app before exporting.

## Tolerances

- Position tolerance defaults to `100 mm`.
- Height tolerance defaults to `20 mm`.
- Coordinates keep their original signs unless `Invert for AutoCAD Plotting` is selected.

## Test data

Use `sample_fieldbook.csv` to try the app quickly.
