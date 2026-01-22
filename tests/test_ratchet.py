import sys
from pathlib import Path

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from scripts import ratchet


def write_excel_with_blank_row(path: Path, pres_df: pd.DataFrame, props_df: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pres_df.to_excel(writer, sheet_name="PresTempPipeID", index=False)
        props_df.to_excel(writer, sheet_name="PipeProperties", index=False)

    from openpyxl import load_workbook

    wb = load_workbook(path)
    for sheet_name in ["PresTempPipeID", "PipeProperties"]:
        ws = wb[sheet_name]
        ws.insert_rows(2)
    wb.save(path)


def test_normalize_header():
    assert ratchet.normalize_header("  Case 1  Pres.  psi  ") == "case 1 pres. psi"


def test_parse_cases_basic():
    df = pd.DataFrame(
        {
            "From": ["A"],
            "To": ["B"],
            "Material": ["CS"],
            "Pipe ID": ["P-1"],
            "Nominal in": [6],
            "Case 1  Pres.  psi": [100],
            "Case 1  Temp.  deg F": [200],
            "Case 1  Auto": ["Y"],
            "Case 1  Allow. Sm  psi": [10],
            "Case 2  Pres.  psi": [150],
            "Case 2  Temp.  deg F": [250],
            "Case 2  Allow. Sm  psi": [8],
        }
    )
    errors = []
    df = ratchet.normalize_columns(df, errors, "file.xlsx", "PresTempPipeID")
    long_df, cases = ratchet.parse_cases(df, errors, "file.xlsx")

    assert cases == [1, 2]
    assert len(long_df) == 2
    assert "pressure_psi" in long_df.columns
    assert "allow_sm_psi" in long_df.columns


def test_build_envelope():
    df = pd.DataFrame(
        {
            "pipe id": ["P-1", "P-1"],
            "pipe_id_key": ["P-1", "P-1"],
            "case": [1, 2],
            "pressure_psi": [100, 200],
            "temperature_deg_f": [300, 250],
            "allow_sm_psi": [10, 8],
            "yield_sy_psi": [30, 25],
            "hot_mod_e6_psi": [28, 27],
        }
    )
    summary = ratchet.build_envelope(df)
    assert summary.loc[0, "max_pressure_psi"] == 200
    assert summary.loc[0, "max_pressure_psi_case"] == 2
    assert summary.loc[0, "min_allow_sm_psi"] == 8
    assert summary.loc[0, "min_allow_sm_psi_case"] == 2


def test_process_file_smoke(tmp_path: Path):
    pres_df = pd.DataFrame(
        {
            "From": ["A"],
            "To": ["B"],
            "Material": ["CS"],
            "Pipe ID": ["P-1"],
            "Nominal in": [6],
            "Case 1  Pres.  psi": [100],
            "Case 1  Temp.  deg F": [200],
            "Case 1  Allow. Sm  psi": [10],
            "Case 1  Delta T1  deg F": [20],
            "Case 1  Delta T2  deg F": [30],
            "Case 2  Pres.  psi": [150],
            "Case 2  Temp.  deg F": [250],
            "Case 2  Allow. Sm  psi": [8],
            "Case 2  Delta T1  deg F": [25],
            "Case 2  Delta T2  deg F": [35],
        }
    )
    props_df = pd.DataFrame(
        {
            "PipeID": ["P-1"],
            "Actual O.D.  inch": [10.0],
            "Wall Thick.  inch": [0.5],
            "Corrosion  inch": [0.05],
            "Mill Tol.  inch": [0.02],
            "Ratchet C4": [1.1],
            "Min. Yield (Sy)  psi": [30000],
            "Allow. Sm  psi": [15000],
            "Long Mod.  E6 psi": [28],
            "Hoop Mod.  E6 psi": [29],
            "Shear Mod.  E6 psi": [11],
            "Thermal Exp.  E-6in/inF": [6.5],
            "Density  lb/cu.ft": [490],
            "Poisson's Ratio": [0.3],
            "Pipe Material": ["A106"],
            "Composition": ["CS"],
        }
    )

    input_path = tmp_path / "model.xlsx"
    output_dir = tmp_path / "output"
    write_excel_with_blank_row(input_path, pres_df, props_df)

    ratchet.process_file(input_path, output_dir)

    output_path = output_dir / "model_ratchet.xlsx"
    assert output_path.exists()

    xl = pd.ExcelFile(output_path)
    assert set(xl.sheet_names) == {"RatchetInputs", "RatchetSummary", "Errors"}
    summary = pd.read_excel(output_path, sheet_name="RatchetSummary")
    assert summary.loc[0, "max_pressure_psi"] == 150
    assert summary.loc[0, "min_allow_sm_psi"] == 8
