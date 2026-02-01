# How to run

## 1) Prepare input files
Place all standardized report files into one folder, for example:
- `A01.xlsx`
- `B01.xlsx`
- `V01.xlsx`
- `G01.xlsx`

## 2) Set the folder path
Open the Excel workbook that contains the Power Query solution and update the named object `link` with the path to the input folder.

## 3) Refresh queries
In Excel:
- Data → Refresh All

This will refresh each section query and produce updated outputs.

## Outputs
The solution generates separate consolidated tables per section:
- `section_R11` → output for `R1.1`
- `section_R12` → output for `R1.2`
- `section_R3`  → output for `R3`
- `section_R4`  → output for `R4`

Notes:
- No hardcoded paths are used. The workflow is portable as long as `link` points to the correct folder.
- Optional VBA macro can be used for one-click refresh (not required for this repository).
