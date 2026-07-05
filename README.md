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

To run the manhole setting-out report generator:

```powershell
streamlit run manhole_app.py
```

If you are using the bundled Codex runtime in this workspace, run:

```powershell
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m streamlit run app.py
```

For the manhole app with the bundled Codex runtime:

```powershell
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m streamlit run manhole_app.py
```

## Manhole setting-out app

- Detects design manholes from `Solution = NONE` rows with valid coordinates and mostly zero observation metadata.
- Lets you choose which design manholes to include.
- Lets you define flow order by listing one manhole per line before plotting the layout and longsection.
- Compares every surveyed/measured point to the nearest selected design manhole.
- Exports branded Excel, PDF, and DXF reports.
- Layout and DXF plotting use inverted coordinates (`plotted East = -East`, `plotted North = -North`) so negative field coordinates flip to the opposite sign for orientation.
- Spreadsheet reports use 2 cm print margins and compact column widths for A4/A3 printing.

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
