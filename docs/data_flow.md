# Data flow

## Overview
This project implements an ETL workflow in **Excel Power Query (M language)** to decode semi-structured Excel report templates into analysis-ready tables.

## Inputs
- Multiple Excel report files placed into a single folder (e.g., `A01.xlsx`, `B01.xlsx`).
- The input folder path is stored in the current workbook as a named object `link`.

## Processing steps (per section)
Each section query follows the same orchestration pattern:

1. **Load files**
   - Read folder path from `link`
   - Load file list using `Folder.Files(folderPath)`

2. **Build full file path**
   - Concatenate `[Folder Path] & [Name]` into `FullPath`

3. **Extract metadata**
   - `fnTitle(FullPath)` reads the organization name from the `title` sheet (fixed cell)

4. **Decode section**
   - `fnGetRxx(FullPath)` reads the corresponding sheet (e.g., `Р.1.1`, `Р.1.2`, `Р.3`, `Р.4`)
   - Converts template matrix into a single-row wide table using the pattern:
     - Remove noise columns/rows
     - Skip header rows
     - Transpose
     - Promote code headers (e.g., `1000`, `1100`, …)
     - Expand into decoded metric columns (e.g., `10001`, `11005`, etc.)

5. **Expand results**
   - Section query expands the nested table returned by `fnGetRxx`
   - Removes technical metadata columns to produce a clean output

## Outputs
Separate consolidated tables per section:
- `R1.1`, `R1.2`, `R3`, `R4` (included in this repository)
- `R2.1`, `R2.2` follow the same extraction pattern and are omitted to avoid duplication
