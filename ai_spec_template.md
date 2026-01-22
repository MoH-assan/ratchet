# Ratchet Automation Specification
`
**Project**: ./ Ratchet
**Date**: <!-- 2026-01-22 -->
**Author**: <!-- Mohamed Abubakr Hassan -->

## 1. Objective
<!-- Manually calculate the ratchet values for the given data -->

<!-- We are doing a more conservative analysis. -->

## 2. Context
- **Project Structure**:
    - `data/input`: Contains source files <!-- Each file is one model. There are two sheets in each file.  -->
    - `data/output`: Contains output files <!-- Each file is one model. It should have the same name as the input file. -->
    - `scripts/`: Location for Python scripts.
    - `readme.md`: Project documentation.

## 3. Data Requirements
### Input Data
- **Location**: `data/input`
- **Format**: <!-- e.g., Excel (.xlsx)-->
- **Key Fields**: <!-- First sheet is the PresTempPipeID and PipeProperties. The first row in each sheet is the header. Followed by an empty row. Then the data starts. -->
    - In the PresTempPipeID sheet, there are runner fields (From, To, Material, Pipe ID, Nominal in) and  case fields (Case x  Pres.  psi, Case x  Temp.  deg F, Case x  Auto, Case x  Expan.  in/100ft, Case x  Auto, Case x  Hot Mod.  E6 psi, Case x  Auto, Case x  Yield(SY)  psi, Case x  Auto, Case x  Allow. Sm  psi, Case x  Delta T1  deg F, Case x  Delta T2  deg F). Where x is the case number. 
    -  In the PipeProperties sheet, there are fields (PipeID, Tag No., Nominal in, Actual O.D.  inch, Schedule, Wall Thick.  inch, Corrosion  inch, Mill Tol.  inch, Insul. Thick.  inch, Insul. Matl., Insul. Dens.  lb/cu.ft, Clad thickness  inch, Clad material, Clad density  lb/cu.ft, Lining Thick.  inch, Lining Dens.  lb/cu.ft, Line Class, Spec. Grav., Pipe Material, Composition, No LT warnings, Ratchet C4, Long Weld, Long Weld Type, Circ Weld, Min. Yield (Sy)  psi, Allow. Sm  psi, Long Mod.  E6 psi, Hoop Mod.  E6 psi, Shear Mod.  E6 psi, Thermal Exp.  E-6in/inF, Density  lb/cu.ft, Poisson's Ratio, Fatigue Curve, Enviromental Factor)  
    - 

### Output Data
- **Destination**: <!-- e.g., data/output -->
- **Format**: <!-- e.g., Excel Data Table -->
- **Structure**: <!-- Describe columns or layout of the output -->

## 4. Technical Stack
- **Language**: Python (Recommended)
- **Libraries**: <!-- e.g., pandas, openpyxl, selenium, requests -->
- **Environment**: Windows

## 5. Functional Requirements
1. <!-- Step 1: e.g., Read all files in data/input -->
2. <!-- Step 2: e.g., Remove tailing and leading spaces from the columns names -->
3. <!-- Step 3: e.g., For each row and over all cases find the max and min values for the Pres, Temp, Allow. Sm -->
4. <!-- Step 4: e.g., For each row and based on the Pipe ID find find the corresponding row in the PipeProperties sheet and find the corresponding values for the Wall Thick.  inch, Actual O.D, Pipe Material, Composition,E6 psi, Min. Yield (Sy) >

## 6. Non-Functional Requirements
- **Error Handling**: Gracefully handle missing files or malformed data.
- **Typo in column names or sheet names**: In case of typo in column names or sheet names, the script should print the error and the existing column names or sheet names.
- **Logging**: Log progress and field errors to console/log file.
- **Performance**: <!-- No specific speed requirements? -->
- **Input Data**: Search for all files in the data/input folder and process them one at a time. 
- **Code Style**: DRY principle. Modular code, and well documented code.  
## 7. Implementation Steps for AI
<!-- High-level plan for the AI to follow -->
1. Generate requirements.txt file and install the required libraries.
2. Implement data loading module.
3. Implement processing logic.
4. Implement output generation.
5. Verify with provided test data.

## 8. Test Data
<!-- Provide sample input data for testing -->
1.  use pytest for both unit testing and smoke testing.
2.  