# Ratchet Automation Specification

**Project**: Ratchet
**Date**: 2026-01-22
**Author**: Mohamed Abubakr Hassan

## 1. Objective
Automate the manual ratchet calculations for each model in `data/input` and generate a structured, auditable output. Use a conservative envelope by selecting worst-case values across load cases (pressure, temperature, allowables, etc.) before applying the ratchet formula.

## 2. Context
- **Project Structure**:
    - `data/input`: Source Excel files (one model per file; two sheets in each file).
    - `data/output`: Output files (one model per file; same base name as input).
    - `scripts/`: Python scripts for processing.
    - `readme.md`: Project documentation.

## 3. Data Requirements
### Input Data
- **Location**: `data/input`
- **Format**: Excel (.xlsx)
- **Sheets**: `PresTempPipeID` and `PipeProperties` (case-insensitive; allow extra spaces)
- **Row layout**: Row 1 headers, Row 2 blank, data starts on Row 3.
- **Key Fields**:
    - **PresTempPipeID** (runner fields + case fields)
        - Runner fields: `From`, `To`, `Material`, `Pipe ID`, `Nominal in`
        - Case fields (pattern; where x is case number):
            - `Case x  Pres.  psi`
            - `Case x  Temp.  deg F`
            - `Case x  Auto` (ignore in calculations)
            - `Case x  Expan.  in/100ft`
            - `Case x  Hot Mod.  E6 psi`
            - `Case x  Yield(SY)  psi`
            - `Case x  Allow. Sm  psi`
            - `Case x  Delta T1  deg F`
            - `Case x  Delta T2  deg F`
    - **PipeProperties** fields (used for joins and formulas):
        - `PipeID`, `Tag No.`, `Nominal in`, `Actual O.D.  inch`, `Schedule`, `Wall Thick.  inch`, `Corrosion  inch`,
          `Mill Tol.  inch`, `Insul. Thick.  inch`, `Insul. Matl.`, `Insul. Dens.  lb/cu.ft`, `Clad thickness  inch`,
          `Clad material`, `Clad density  lb/cu.ft`, `Lining Thick.  inch`, `Lining Dens.  lb/cu.ft`, `Line Class`,
          `Spec. Grav.`, `Pipe Material`, `Composition`, `No LT warnings`, `Ratchet C4`, `Long Weld`,
          `Long Weld Type`, `Circ Weld`, `Min. Yield (Sy)  psi`, `Allow. Sm  psi`, `Long Mod.  E6 psi`,
          `Hoop Mod.  E6 psi`, `Shear Mod.  E6 psi`, `Thermal Exp.  E-6in/inF`, `Density  lb/cu.ft`,
          `Poisson's Ratio`, `Fatigue Curve`, `Enviromental Factor`

### Output Data
- **Destination**: `data/output`
- **Format**: Excel (.xlsx)
- **Structure**:
    - Sheet `RatchetInputs`: one row per (Pipe ID, Case). Includes runner fields + case fields, plus joined PipeProperties fields required by the formula.
    - Sheet `RatchetSummary`: one row per Pipe ID with conservative envelope values and the computed ratchet result.
    - Sheet `Errors`: any validation issues (missing columns, missing PipeID matches, non-numeric values).

## 4. Technical Stack
- **Language**: Python 3.10+
- **Libraries**: pandas, openpyxl, pytest
- **Environment**: Windows

## 5. Functional Requirements
1. **Discover inputs**: Read all `.xlsx` files in `data/input` (ignore temporary files beginning with `~$`).
2. **Load sheets**: For each file, read `PresTempPipeID` and `PipeProperties`. Skip the blank row after headers.
3. **Normalize columns**: Trim whitespace, collapse multiple spaces, and standardize column names for matching (case-insensitive).
4. **Parse cases**:
    - Detect case numbers from column headers matching `Case x`.
    - Reshape to long format (one row per case) for easier validation and downstream calculations.
    - Ignore `Case x  Auto` columns.
5. **Type coercion**: Convert numeric columns to floats. Log and flag non-numeric values in the `Errors` sheet.
6. **Join properties**: Merge on `Pipe ID` (normalized). If multiple matches exist, take the first and log; if missing, log and skip calculation.
7. **Conservative envelope** (per Pipe ID):
    - Max pressure, max temperature, max expansion, max Delta T1/Delta T2.
    - Min Allowable Sm, min Yield (Sy), min Hot Modulus.
    - Record the controlling case numbers for each envelope value.
8. **Ratchet calculation**:
    - Apply the same formula as the manual spreadsheet used by Henry Song.
    - The formula must use `Ratchet C4`, `Actual O.D.  inch`, `Wall Thick.  inch`, `Corrosion  inch`, `Mill Tol.  inch`,
      material moduli, and the conservative case values above.
    - If the formula is not yet finalized, output all required inputs in `RatchetSummary` and leave `RatchetValue` blank with a clear note.
9. **Outputs**: Write one output file per input file to `data/output` with the same base name + `_ratchet.xlsx`.

## 6. Non-Functional Requirements
- **Error Handling**: Continue processing other files even if one fails; report all errors in `Errors` sheet and console.
- **Typos in column or sheet names**: Print a clear message with the expected names and the actual names found.
- **Logging**: Log progress per file (loaded rows, cases found, joins matched, output path).
- **Performance**: Must handle typical project-sized files without excessive memory use.
- **Input Processing**: Process files one at a time to limit memory usage.
- **Code Style**: DRY, modular functions, and docstrings for core calculations.

## 7. Implementation Steps for AI
1. Ensure `requirements.txt` includes `pandas`, `openpyxl`, and `pytest`.
2. Implement a loader that:
    - Validates sheet names.
    - Normalizes headers.
    - Skips the blank row after headers.
3. Implement a case parser that builds a long-form dataframe and extracts case numbers.
4. Implement a join module for `PipeProperties` with validation and error logging.
5. Implement the conservative envelope summary per Pipe ID.
6. Implement the ratchet calculation (as a separate function) with unit tests.
7. Implement output writer (3 sheets) and smoke tests using minimal sample data.

## 8. Test Data
1. Use `pytest` for unit tests (column normalization, case parsing, join behavior).
2. Use a small synthetic Excel file for a smoke test:
    - 1 Pipe ID, 2 cases, and matching PipeProperties row.
    - Validate the output file structure and that envelope values match expected mins/maxes.
