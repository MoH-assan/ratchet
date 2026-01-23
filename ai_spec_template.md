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
        - Runner fields: 'From', 'To', 'Material', 'Pipe ID', 'Nominal     in'
        - Case fields (pattern; where x is case number): ['Case 1  Pres.  psi', 'Case 1  Temp.  deg F', 'Case 1  Auto', 'Case 1  Expan.  in/100ft', 'Case 1  Auto    .1', 'Case 1  Hot Mod.  E6 psi', 'Case 1  Auto    .2', 'Case 1  Yield(SY)  psi', 'Case 1  Auto    .3', 'Case 1  Allow. Sm  psi', 'Case 1  Delta T1  deg F', 'Case 1  Delta T2  deg F',]


    - **PipeProperties** fields (used for joins and formulas):
        - ['PipeID', 'Tag No.', 'Nominal     in', 'Actual O.D.  inch', 'Schedule', 'Wall Thick.  inch', 'Corrosion  inch', 'Mill Tol.  inch', 'Insul. Thick.  inch', 'Insul. Matl.', 'Insul. Dens.  lb/cu.ft', 'Clad thickness  inch', 'Clad material', 'Clad density  lb/cu.ft', 'Lining Thick.  inch', 'Lining Dens.  lb/cu.ft', 'Line Class', 'Spec. Grav.', 'Pipe Material', 'Composition', 'No LT warnings', 'Ratchet C4', 'Long Weld', 'Long Weld Type', 'Circ Weld', 'Min. Yield (Sy)  psi', 'Allow. Sm  psi', 'Long Mod.  E6 psi', 'Hoop Mod.  E6 psi', 'Shear Mod.  E6 psi', 'Thermal Exp.  E-6in/inF', 'Density  lb/cu.ft', "Poisson's Ratio", 'Fatigue Curve', 'Enviromental Factor']

### Output Data
- **Destination**: `data/output`
- **Format**: Excel (.xlsx)

- **Structure**:
    - Sheet `PerNodeEnvlope`: one row per envelope values for each row in the PresTempPipeID sheet.
    - Sheet `PerMaterialEnvlope`: one row per **unique pipe_material** in `PerNodeEnvlope`, containing conservative envelope values per material.
    - Add a units row after the header row in both `PerNodeEnvlope` and `PerMaterialEnvlope`.
    - Sheet `Errors`: any validation issues (missing columns, missing PipeID matches, non-numeric values).
    - Use (E) for the youngs modulus in all calculations and column name (E not e).

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
7. **Conservative envelope** (per each row in the PresTempPipeID sheet over all cases for per each row):
    - from the PresTempPipeID and over all case, find 
    `P_max`=max absolute value of 'Case x  Pres.  psi',
    `Sy_min`= min of all 'Case x  Yield(SY)  psi'
    `delta_t1_max`= max of 'Case x  Delta T1  deg F', 
    `E_max` = max of all 'Case x  Hot Mod.  E6 psi' (Use E not e for the youngs modulus in all calculations)
    - Record the controlling case numbers for each envelope value.

    - based on the PipeProperties, find the following values: 
    `D_out`=`Actual O.D.  inch`, 
    `thck`=`Wall Thick.  inch`, 
    `pipe_material`=`Pipe Material`,
    `alpha_room`=`Thermal Exp.  E-6in/inF`
    `c4`=`Ratchet C4`

8. **Ratchet calculation**:
    - Apply the  following calculation for each row in the PresTempPipeID sheet.
    - `x`=`P_max`*`D_out`/(2*`thck`*`Sy_min`)
    - `y`= 1/ `x` for 0<=`x`<=0.5, `y`=4*(1-`x`) for 0.5<`x`<=1,
    - `allowable` = `c4` * `y` * `Sy_min` / (0.7 * `E_max` * `alpha_room`)
    - If the formula is not yet finalized, output all required inputs in `PerNodeEnvlope` and leave `allowable` blank with a clear note.
9. **Per-Material Envelope (PerMaterialEnvlope)**    
    - For **each unique** `pipe_material` in `PerNodeEnvlope`, create one envelope row in `PerMaterialEnvlope`.
    - This sheet **does not include** the `from` or `to` columns.
    - For each material, compute:
        - `P_max` = max absolute value of `p_max` over all rows with that material.
        - `Sy_min` = min of `sy_min` over all rows with that material.
        - `delta_t1_max` = max of `delta_t1_max` over all rows with that material.
        - `E_max` = max of `E_max` over all rows with that material.
        - `D_out` = max of `d_out` over all rows with that material.
        - `thck` = min of `thck` over all rows with that material.
        - `alpha_room` = max of `alpha_room` over all rows with that material.
        - `c4` = min of `c4` over all rows with that material.
    - **Controlling row tracking (From/To)**:
        - For each envelope value above, record the **controlling From->To** in the corresponding `_case` column.
        - Additional columns: `d_out_case`, `thck_case`, `c4_case` (placed immediately after `d_out`, `thck`, `c4` respectively).
    - Apply the same ratchet calculation for each material envelope row:
        - `x`=`P_max`*`D_out`/(2*`thck`*`Sy_min`)
        - `y`= 1/ `x` for 0<=`x`<=0.5, `y`=4*(1-`x`) for 0.5<`x`<=1,
        - `allowable` = `c4` * `y` * `Sy_min` / (0.7 * `E_max` * `alpha_room`)
    - If the formula is not yet finalized, output all required inputs in `PerMaterialEnvlope` and leave `allowable` blank with a clear note.

10. **Formatting and Highlighting**
    - Add a units row after the header row in **both** `PerNodeEnvlope` and `PerMaterialEnvlope`.
    - In `PerNodeEnvlope`, the `_case` columns show the **controlling load case number**.
    - In `PerMaterialEnvlope`, the `_case` columns show the **controlling runner From->To**.
    - Use a **different color per material** to avoid confusion:
        - In `PerMaterialEnvlope`, bold values are colored by material.
        - In `PerNodeEnvlope`, the controlling values are bold and colored using the same material color.
    - Do **not** bold calculated fields (`x`, `y`, `allowable`, `allowable_note`) and do **not** color non-bold cells.

11. **Outputs**: Write one output file per input file to `data/output` with the same base name + `_ratchet.xlsx`.


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
5. Implement the conservative envelope summary per each row in the PresTempPipeID sheet over all cases for per each row.
6. Implement the ratchet calculation (as a separate function) with unit tests.
7. Implement output writer and smoke tests using minimal sample data.

## 8. Test Data
1. Use `pytest` for unit tests (column normalization, case parsing, join behavior).
2. Use a small synthetic Excel file for a smoke test:
    - 1 Pipe ID, 2 cases, and matching PipeProperties row.
    - Validate the output file structure and that envelope values match expected mins/maxes.
