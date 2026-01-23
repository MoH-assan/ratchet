import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd


SHEET_PRESTEMP = "PresTempPipeID"
SHEET_PROPERTIES = "PipeProperties"

RUNNER_FIELD_KEYS = {
    "from": "from",
    "to": "to",
    "material": "material",
    "pipe id": "pipe_id",
    "pipeid": "pipe_id",
    "nominal in": "nominal_in",
    "nominalin": "nominal_in",
}

PROPERTY_FIELD_KEYS = {
    "pipeid": "pipe_id",
    "pipe id": "pipe_id",
    "actual o d inch": "d_out",
    "wall thick inch": "thck",
    "pipe material": "pipe_material",
    "thermal exp e 6in inf": "alpha_room",
    "ratchet c4": "c4",
}

CASE_FIELD_MAP = {
    "pres psi": "pres_psi",
    "temp deg f": "temp_deg_f",
    "yield sy psi": "yield_sy_psi",
    "allow sm psi": "allow_sm_psi",
    "delta t1 deg f": "delta_t1_deg_f",
    "delta t2 deg f": "delta_t2_deg_f",
    "hot mod e6 psi": "hot_mod_e6_psi",
    "expan in 100ft": "expan_in_100ft",
}


class ErrorLog(List[Dict[str, object]]):
    """Collects errors for output to the Errors sheet."""


def normalize_key(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^0-9a-z\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_sheet_key(value: object) -> str:
    return normalize_key(value).replace(" ", "")


def normalize_pipe_id(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    return text.upper()


def log_error(
    errors: ErrorLog,
    message: str,
    *,
    file_name: Optional[str] = None,
    sheet: Optional[str] = None,
    row: Optional[int] = None,
    column: Optional[str] = None,
    value: Optional[object] = None,
    level: str = "error",
) -> None:
    errors.append(
        {
            "file": file_name,
            "sheet": sheet,
            "level": level,
            "message": message,
            "row": row,
            "column": column,
            "value": value,
        }
    )


def normalize_columns(df: pd.DataFrame, errors: ErrorLog, file_name: str, sheet: str) -> pd.DataFrame:
    new_cols = []
    seen: Dict[str, int] = {}
    for col in df.columns:
        key = normalize_key(col)
        if key in seen:
            seen[key] += 1
            new_key = f"{key}__dup{seen[key]}"
            log_error(
                errors,
                f"Duplicate column after normalization: '{key}'.",
                file_name=file_name,
                sheet=sheet,
                column=str(col),
                level="warning",
            )
        else:
            seen[key] = 0
            new_key = key
        new_cols.append(new_key)
    df = df.copy()
    df.columns = new_cols
    return df


def strip_dup_suffix(col_name: str) -> str:
    return col_name.split("__dup", 1)[0]


def find_sheet_name(sheet_names: Iterable[str], expected: str) -> Optional[str]:
    expected_key = normalize_sheet_key(expected)
    for name in sheet_names:
        if normalize_sheet_key(name) == expected_key:
            return name
    return None


def read_excel_sheet(file_path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=0,
        skiprows=[1],
        dtype=object,
        engine="openpyxl",
    )


def find_column(df: pd.DataFrame, key: str) -> Optional[str]:
    key_norm = normalize_key(key)
    if key_norm in df.columns:
        return key_norm
    key_ns = key_norm.replace(" ", "")
    for col in df.columns:
        if col.replace(" ", "") == key_ns:
            return col
    return None


def extract_runner_columns(
    pres_df: pd.DataFrame, errors: ErrorLog, file_name: str, sheet: str
) -> Tuple[pd.DataFrame, List[str]]:
    runner_cols: Dict[str, str] = {}
    for key, canonical in RUNNER_FIELD_KEYS.items():
        if canonical in runner_cols.values():
            continue
        col = find_column(pres_df, key)
        if col and canonical not in runner_cols:
            runner_cols[canonical] = col
    missing = [
        canonical for canonical in {"from", "to", "material", "pipe_id", "nominal_in"}
        if canonical not in runner_cols
    ]
    for name in missing:
        log_error(
            errors,
            f"Missing required runner column '{name}'.",
            file_name=file_name,
            sheet=sheet,
        )
    if missing:
        print(
            f"[{file_name}] Missing runner columns in {sheet}. "
            f"Expected: {sorted({'from', 'to', 'material', 'pipe_id', 'nominal_in'})}. "
            f"Found: {list(pres_df.columns)}"
        )
    base_df = pd.DataFrame()
    for canonical, col in runner_cols.items():
        base_df[canonical] = pres_df[col]
    return base_df, list(runner_cols.values())


def parse_cases(
    pres_df: pd.DataFrame,
    runner_cols: List[str],
    errors: ErrorLog,
    file_name: str,
    sheet: str,
) -> Tuple[pd.DataFrame, List[int]]:
    case_map: Dict[int, Dict[str, str]] = {}
    for col in pres_df.columns:
        base_col = strip_dup_suffix(col)
        match = re.match(r"^case\s*(\d+)\s+(.*)$", base_col)
        if not match:
            continue
        case_num = int(match.group(1))
        field_key = match.group(2).strip()
        if field_key.startswith("auto"):
            continue
        canonical = CASE_FIELD_MAP.get(field_key, field_key.replace(" ", "_"))
        case_map.setdefault(case_num, {})[canonical] = col

    if not case_map:
        log_error(
            errors,
            "No case columns found in PresTempPipeID sheet.",
            file_name=file_name,
            sheet=sheet,
        )
        return pd.DataFrame(), []

    frames = []
    for case_num in sorted(case_map.keys()):
        cols = runner_cols + ["row_id"] + list(case_map[case_num].values())
        case_df = pres_df.loc[:, cols].copy()
        case_df = case_df.rename(columns={v: k for k, v in case_map[case_num].items()})
        case_df["case_number"] = case_num
        frames.append(case_df)

    long_df = pd.concat(frames, ignore_index=True)
    return long_df, sorted(case_map.keys())


def coerce_numeric(
    df: pd.DataFrame,
    columns: Iterable[str],
    errors: ErrorLog,
    file_name: str,
    sheet: str,
    row_offset: int = 3,
) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            continue
        original = df[col]
        coerced = pd.to_numeric(original, errors="coerce")
        bad_mask = original.notna() & coerced.isna()
        if bad_mask.any():
            for idx in df.index[bad_mask]:
                row_val = df.loc[idx, "row_id"] if "row_id" in df.columns else idx
                log_error(
                    errors,
                    "Non-numeric value found.",
                    file_name=file_name,
                    sheet=sheet,
                    row=int(row_val) + row_offset,
                    column=col,
                    value=original.loc[idx],
                    level="warning",
                )
        df[col] = coerced
    return df


def compute_envelope(
    long_df: pd.DataFrame,
    errors: ErrorLog,
    file_name: str,
    sheet: str,
) -> pd.DataFrame:
    required_cols = ["pres_psi", "yield_sy_psi", "delta_t1_deg_f", "hot_mod_e6_psi"]
    missing_case_cols = [col for col in required_cols if col not in long_df.columns]
    for col in missing_case_cols:
        if col not in long_df.columns:
            log_error(
                errors,
                f"Missing required case column '{col}'.",
                file_name=file_name,
                sheet=sheet,
            )
            long_df[col] = np.nan
    if missing_case_cols:
        print(
            f"[{file_name}] Missing case columns in {sheet}. "
            f"Expected: {required_cols}. Found: {list(long_df.columns)}"
        )

    def max_abs_with_case(group: pd.DataFrame, col: str) -> Tuple[float, Optional[int]]:
        series = group[col]
        if series.dropna().empty:
            return np.nan, None
        abs_series = series.abs()
        max_abs = abs_series.max()
        idx = abs_series[abs_series == max_abs].index[0]
        case = group.loc[idx, "case_number"]
        return float(max_abs), int(case) if pd.notna(case) else None

    def max_with_case(group: pd.DataFrame, col: str) -> Tuple[float, Optional[int]]:
        series = group[col]
        if series.dropna().empty:
            return np.nan, None
        idx = series.idxmax()
        case = group.loc[idx, "case_number"]
        return float(series.loc[idx]), int(case) if pd.notna(case) else None

    def min_with_case(group: pd.DataFrame, col: str) -> Tuple[float, Optional[int]]:
        series = group[col]
        if series.dropna().empty:
            return np.nan, None
        idx = series.idxmin()
        case = group.loc[idx, "case_number"]
        return float(series.loc[idx]), int(case) if pd.notna(case) else None

    rows = []
    for row_id, group in long_df.groupby("row_id"):
        p_max, p_case = max_abs_with_case(group, "pres_psi")
        sy_min, sy_case = min_with_case(group, "yield_sy_psi")
        dt1_max, dt1_case = max_with_case(group, "delta_t1_deg_f")
        e_max, e_case = max_with_case(group, "hot_mod_e6_psi")
        rows.append(
            {
                "row_id": row_id,
                "p_max": p_max,
                "p_max_case": p_case,
                "sy_min": sy_min,
                "sy_min_case": sy_case,
                "delta_t1_max": dt1_max,
                "delta_t1_case": dt1_case,
                "E_max": e_max,
                "E_max_case": e_case,
            }
        )
    return pd.DataFrame(rows)


def extract_properties(
    prop_df: pd.DataFrame, errors: ErrorLog, file_name: str, sheet: str
) -> pd.DataFrame:
    selected: Dict[str, str] = {}
    for key, canonical in PROPERTY_FIELD_KEYS.items():
        if canonical in selected:
            continue
        col = find_column(prop_df, key)
        if col:
            selected[canonical] = col
    required = {"pipe_id", "d_out", "thck", "pipe_material", "alpha_room", "c4"}
    missing = [name for name in required if name not in selected]
    for name in missing:
        log_error(
            errors,
            f"Missing required property column '{name}'.",
            file_name=file_name,
            sheet=sheet,
        )
    if missing:
        print(
            f"[{file_name}] Missing property columns in {sheet}. "
            f"Expected: {sorted(required)}. Found: {list(prop_df.columns)}"
        )
    prop_subset = pd.DataFrame()
    for canonical, col in selected.items():
        prop_subset[canonical] = prop_df[col]
    return prop_subset


def calculate_allowable(row: pd.Series) -> Tuple[Optional[float], str, Optional[float], Optional[float]]:
    required = ["p_max", "sy_min", "E_max", "alpha_room", "c4", "d_out", "thck"]
    if any(pd.isna(row.get(key)) for key in required):
        return None, "Missing inputs", None, None
    if row["thck"] == 0 or row["sy_min"] == 0:
        return None, "Invalid Sy_min or thickness", None, None
    x = (row["p_max"] * row["d_out"]) / (2 * row["thck"] * row["sy_min"])
    if x == 0:
        return None, "Invalid x (division by zero)", float(x), None
    if 0 <= x <= 0.5:
        y = 1 / x
    elif 0.5 < x <= 1:
        y = 4 * (1 - x)
    else:
        return None, "x out of range", float(x), None
    if row["E_max"] == 0 or row["alpha_room"] == 0:
        return None, "Invalid E_max or alpha", float(x), float(y)
    e_actual = row["E_max"] * 1_000_000.0
    alpha_actual = row["alpha_room"] * 0.000001
    allowable = row["c4"] * y * row["sy_min"] / (0.7 * e_actual * alpha_actual)
    return float(allowable), "", float(x), float(y)


def errors_to_dataframe(errors: ErrorLog) -> pd.DataFrame:
    if not errors:
        return pd.DataFrame(columns=["file", "sheet", "level", "message", "row", "column", "value"])
    return pd.DataFrame(errors)
