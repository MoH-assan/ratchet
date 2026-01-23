import math
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from scripts.helper import (
    calculate_allowable,
    compute_envelope,
    normalize_columns,
    normalize_key,
    parse_cases,
)
from scripts.ratchet import process_file


def write_sheet_with_blank_row(ws, headers, rows):
    ws.append(headers)
    ws.append([None] * len(headers))
    for row in rows:
        ws.append(row)


def test_normalize_key():
    assert normalize_key(" Nominal     in ") == "nominal in"
    assert normalize_key("Yield(SY)  psi") == "yield sy psi"


def test_parse_cases_and_envelope():
    pres_df = pd.DataFrame(
        {
            "From": ["A"],
            "To": ["B"],
            "Material": ["X"],
            "Pipe ID": ["P-1"],
            "Nominal     in": [4],
            "Case 1  Pres.  psi": [10],
            "Case 1  Yield(SY)  psi": [50],
            "Case 1  Delta T1  deg F": [20],
            "Case 1  Hot Mod.  E6 psi": [30],
            "Case 2  Pres.  psi": [-12],
            "Case 2  Yield(SY)  psi": [45],
            "Case 2  Delta T1  deg F": [10],
            "Case 2  Hot Mod.  E6 psi": [28],
        }
    )
    errors = []
    pres_df = normalize_columns(pres_df, errors, "test.xlsx", "PresTempPipeID")
    pres_df["row_id"] = pres_df.index
    long_df, _ = parse_cases(pres_df, ["from", "to", "material", "pipe id", "nominal in"], errors, "test.xlsx", "PresTempPipeID")
    envelope_df = compute_envelope(long_df, errors, "test.xlsx", "PresTempPipeID")
    row = envelope_df.iloc[0]
    assert row["p_max"] == 12
    assert row["p_max_case"] == 2
    assert row["sy_min"] == 45
    assert row["sy_min_case"] == 2
    assert row["delta_t1_max"] == 20
    assert row["delta_t1_case"] == 1
    assert row["E_max"] == 30
    assert row["E_max_case"] == 1


def test_calculate_allowable():
    row = {
        "p_max": 1.0,
        "sy_min": 100.0,
        "E_max": 10.0,
        "alpha_room": 1.0,
        "c4": 1.0,
        "d_out": 1.0,
        "thck": 1.0,
    }
    allowable, note, x_val, y_val = calculate_allowable(pd.Series(row))
    assert note == ""
    expected_x = (1.0 * 1.0) / (2.0 * 1.0 * 100.0)
    expected_y = 1.0 / expected_x
    expected = 1.0 * expected_y * 100.0 / (0.7 * 10.0 * 1.0)
    assert math.isclose(x_val, expected_x, rel_tol=1e-6)
    assert math.isclose(y_val, expected_y, rel_tol=1e-6)
    assert math.isclose(allowable, expected, rel_tol=1e-6)


def test_process_file_smoke(tmp_path: Path):
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    input_dir.mkdir()
    output_dir.mkdir()

    pres_headers = [
        "From",
        "To",
        "Material",
        "Pipe ID",
        "Nominal     in",
        "Case 1  Pres.  psi",
        "Case 1  Yield(SY)  psi",
        "Case 1  Delta T1  deg F",
        "Case 1  Hot Mod.  E6 psi",
        "Case 2  Pres.  psi",
        "Case 2  Yield(SY)  psi",
        "Case 2  Delta T1  deg F",
        "Case 2  Hot Mod.  E6 psi",
    ]
    pres_rows = [
        [
            "A",
            "B",
            "X",
            "P1",
            4,
            0.2,
            100,
            10,
            30,
            -0.3,
            95,
            5,
            32,
        ]
    ]

    prop_headers = [
        "PipeID",
        "Actual O.D.  inch",
        "Wall Thick.  inch",
        "Pipe Material",
        "Thermal Exp.  E-6in/inF",
        "Ratchet C4",
    ]
    prop_rows = [["P1", 1.0, 1.0, "Steel", 6.5, 1.0]]

    wb = Workbook()
    ws_pres = wb.active
    ws_pres.title = "PresTempPipeID"
    write_sheet_with_blank_row(ws_pres, pres_headers, pres_rows)

    ws_prop = wb.create_sheet("PipeProperties")
    write_sheet_with_blank_row(ws_prop, prop_headers, prop_rows)

    input_file = input_dir / "sample.xlsx"
    wb.save(input_file)

    output_path, errors = process_file(input_file, output_dir)
    assert output_path is not None
    assert output_path.exists()
    assert len(errors) == 0

    result = pd.read_excel(output_path, sheet_name="PerNodeEnvlope", engine="openpyxl")
    assert "p_max" in result.columns
    data_rows = result[result["from"].notna()].reset_index(drop=True)
    assert data_rows.loc[0, "p_max"] == 0.3
    assert not pd.isna(data_rows.loc[0, "allowable"])
    assert "x" in result.columns
    assert "y" in result.columns
