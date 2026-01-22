from __future__ import annotations

import argparse
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

CASE_REGEX = re.compile(r"^case\s*(\d+)\s+(.*)$", re.IGNORECASE)

RUNNER_ALIASES: Dict[str, List[str]] = {
    "from": ["from"],
    "to": ["to"],
    "material": ["material"],
    "pipe id": ["pipe id", "pipeid"],
    "nominal in": ["nominal in", "nominal in."],
}

PROPERTY_NUMERIC_COLUMNS = [
    "actual o.d. inch",
    "wall thick. inch",
    "corrosion inch",
    "mill tol. inch",
    "ratchet c4",
    "min. yield (sy) psi",
    "allow. sm psi",
    "long mod. e6 psi",
    "hoop mod. e6 psi",
    "shear mod. e6 psi",
    "thermal exp. e-6in/inf",
    "density lb/cu.ft",
    "poisson's ratio",
    "nominal in",
]

CASE_NUMERIC_COLUMNS = [
    "pressure_psi",
    "temperature_deg_f",
    "expan_in_per_100ft",
    "hot_mod_e6_psi",
    "yield_sy_psi",
    "allow_sm_psi",
    "delta_t1_deg_f",
    "delta_t2_deg_f",
]

SUMMARY_ENVELOPE_FIELDS = [
    ("pressure_psi", "max"),
    ("temperature_deg_f", "max"),
    ("expan_in_per_100ft", "max"),
    ("delta_t1_deg_f", "max"),
    ("delta_t2_deg_f", "max"),
    ("allow_sm_psi", "min"),
    ("yield_sy_psi", "min"),
    ("hot_mod_e6_psi", "min"),
]

PROPERTY_FIELDS_FOR_OUTPUT = [
    "ratchet c4",
    "actual o.d. inch",
    "wall thick. inch",
    "corrosion inch",
    "mill tol. inch",
    "min. yield (sy) psi",
    "allow. sm psi",
    "long mod. e6 psi",
    "hoop mod. e6 psi",
    "shear mod. e6 psi",
    "thermal exp. e-6in/inf",
    "density lb/cu.ft",
    "poisson's ratio",
    "pipe material",
    "composition",
]


@dataclass
class ErrorRecord:
    file: str
    sheet: str
    issue_type: str
    message: str
    column: Optional[str] = None
    row: Optional[int] = None

    def as_dict(self) -> Dict[str, Optional[str]]:
        return {
            "file": self.file,
            "sheet": self.sheet,
            "issue_type": self.issue_type,
            "message": self.message,
            "column": self.column,
            "row": self.row,
        }


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def make_unique(names: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []
    for name in names:
        if name not in seen:
            seen[name] = 0
            result.append(name)
        else:
            seen[name] += 1
            result.append(f"{name}.{seen[name]}")
    return result


def log_error(
    errors: List[ErrorRecord],
    file: str,
    sheet: str,
    issue_type: str,
    message: str,
    column: Optional[str] = None,
    row: Optional[int] = None,
) -> None:
    errors.append(
        ErrorRecord(
            file=file,
            sheet=sheet,
            issue_type=issue_type,
            message=message,
            column=column,
            row=row,
        )
    )


def normalize_columns(
    df: pd.DataFrame,
    errors: List[ErrorRecord],
    file: str,
    sheet: str,
) -> pd.DataFrame:
    raw_cols = list(df.columns)
    norm_cols = [normalize_header(col) for col in raw_cols]
    counts = Counter(norm_cols)
    duplicates = [name for name, count in counts.items() if count > 1]
    if duplicates:
        log_error(
            errors,
            file,
            sheet,
            "duplicate_columns",
            f"Duplicate columns after normalization: {duplicates}",
        )
    df = df.copy()
    df.columns = make_unique(norm_cols)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    return df


def find_sheet_name(excel: pd.ExcelFile, expected: str) -> Optional[str]:
    expected_norm = normalize_header(expected)
    for name in excel.sheet_names:
        if normalize_header(name) == expected_norm:
            return name
    return None


def resolve_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    for name in candidates:
        if name in df.columns:
            return name
    return None


def normalize_key(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    return text.upper()


def canonical_case_field(field: str) -> Optional[str]:
    if field.startswith("auto"):
        return None
    if "pres" in field:
        return "pressure_psi"
    if "temp" in field:
        return "temperature_deg_f"
    if "expan" in field:
        return "expan_in_per_100ft"
    if "hot mod" in field:
        return "hot_mod_e6_psi"
    if "yield" in field:
        return "yield_sy_psi"
    if "allow" in field and "sm" in field:
        return "allow_sm_psi"
    if "delta t1" in field:
        return "delta_t1_deg_f"
    if "delta t2" in field:
        return "delta_t2_deg_f"
    return None


def coerce_numeric(
    df: pd.DataFrame,
    columns: Iterable[str],
    errors: List[ErrorRecord],
    file: str,
    sheet: str,
) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            continue
        original = df[col]
        coerced = pd.to_numeric(original, errors="coerce")
        bad_mask = original.notna() & coerced.isna()
        if bad_mask.any():
            count = int(bad_mask.sum())
            log_error(
                errors,
                file,
                sheet,
                "non_numeric",
                f"Column '{col}' has {count} non-numeric values.",
                column=col,
            )
        df[col] = coerced
    return df


def parse_cases(
    pres_df: pd.DataFrame,
    errors: List[ErrorRecord],
    file: str,
) -> Tuple[pd.DataFrame, List[int]]:
    sheet = "PresTempPipeID"
    runner_cols: Dict[str, str] = {}
    for canonical, aliases in RUNNER_ALIASES.items():
        col = resolve_column(pres_df, aliases)
        if col is None:
            log_error(
                errors,
                file,
                sheet,
                "missing_column",
                f"Missing runner column: {canonical}",
            )
            pres_df[canonical] = pd.NA
            runner_cols[canonical] = canonical
        else:
            runner_cols[canonical] = col

    case_map: Dict[int, Dict[str, str]] = {}
    for col in pres_df.columns:
        if col in runner_cols.values():
            continue
        match = CASE_REGEX.match(col)
        if not match:
            continue
        case_num = int(match.group(1))
        field = normalize_header(match.group(2))
        canonical = canonical_case_field(field)
        if canonical is None:
            continue
        case_fields = case_map.setdefault(case_num, {})
        if canonical in case_fields:
            log_error(
                errors,
                file,
                sheet,
                "duplicate_case_field",
                f"Duplicate case field '{canonical}' in case {case_num}",
                column=col,
            )
        case_fields[canonical] = col

    if not case_map:
        log_error(
            errors,
            file,
            sheet,
            "missing_cases",
            "No case columns found in PresTempPipeID sheet.",
        )
        return pres_df.assign(case=pd.NA), []

    long_frames: List[pd.DataFrame] = []
    for case_num, fields in sorted(case_map.items()):
        cols = list(runner_cols.values()) + list(fields.values())
        df_case = pres_df[cols].copy()
        rename_map = {v: k for k, v in fields.items()}
        df_case = df_case.rename(columns=rename_map)
        df_case["case"] = case_num
        long_frames.append(df_case)

    long_df = pd.concat(long_frames, ignore_index=True)
    return long_df, sorted(case_map.keys())


def join_properties(
    long_df: pd.DataFrame,
    props_df: pd.DataFrame,
    errors: List[ErrorRecord],
    file: str,
) -> pd.DataFrame:
    sheet = "PipeProperties"
    pipe_col_props = resolve_column(props_df, ["pipeid", "pipe id"])
    if pipe_col_props is None:
        log_error(
            errors,
            file,
            sheet,
            "missing_column",
            "Missing PipeID column in PipeProperties.",
        )
        long_df["pipe_id_key"] = long_df.get("pipe id", pd.NA).map(normalize_key)
        return long_df

    props_df = props_df.copy()
    props_df = props_df.rename(columns={pipe_col_props: "pipe_id_prop"})
    props_df["pipe_id_key"] = props_df["pipe_id_prop"].map(normalize_key)
    long_df = long_df.copy()
    long_df["pipe_id_key"] = long_df.get("pipe id", pd.NA).map(normalize_key)

    duplicate_mask = props_df["pipe_id_key"].duplicated(keep="first") & props_df["pipe_id_key"].notna()
    if duplicate_mask.any():
        dupes = props_df.loc[duplicate_mask, "pipe_id_key"].unique().tolist()
        log_error(
            errors,
            file,
            sheet,
            "duplicate_pipeid",
            f"Duplicate PipeID values found in PipeProperties: {dupes}. Using first occurrence.",
        )
    props_df = props_df.drop_duplicates(subset=["pipe_id_key"], keep="first")

    merged = pd.merge(
        long_df,
        props_df,
        on="pipe_id_key",
        how="left",
        suffixes=("", "_prop"),
    )

    missing_mask = merged["pipe_id_key"].notna() & merged["pipe_id_prop"].isna()
    if missing_mask.any():
        count = int(missing_mask.sum())
        log_error(
            errors,
            file,
            sheet,
            "missing_pipe_match",
            f"{count} rows in PresTempPipeID have no matching PipeID in PipeProperties.",
        )

    return merged


def select_value_with_case(
    df: pd.DataFrame,
    value_col: str,
    mode: str,
) -> Tuple[Optional[float], Optional[int]]:
    if value_col not in df.columns:
        return None, None
    series = df[value_col]
    series = pd.to_numeric(series, errors="coerce")
    series = series.dropna()
    if series.empty:
        return None, None
    if mode == "max":
        idx = series.idxmax()
    else:
        idx = series.idxmin()
    value = df.loc[idx, value_col]
    case = df.loc[idx, "case"] if "case" in df.columns else None
    return value, case


def build_envelope(joined_df: pd.DataFrame) -> pd.DataFrame:
    summaries: List[Dict[str, object]] = []
    if "pipe_id_key" not in joined_df.columns:
        return pd.DataFrame()

    for pipe_id_key, group in joined_df.groupby("pipe_id_key", dropna=False):
        summary: Dict[str, object] = {"pipe_id_key": pipe_id_key}
        if "pipe id" in group.columns:
            summary["pipe id"] = group["pipe id"].iloc[0]

        for field, mode in SUMMARY_ENVELOPE_FIELDS:
            value, case = select_value_with_case(group, field, mode)
            summary[f"{mode}_{field}"] = value
            summary[f"{mode}_{field}_case"] = case

        for prop in PROPERTY_FIELDS_FOR_OUTPUT:
            if prop in group.columns:
                summary[prop] = group[prop].iloc[0]

        summaries.append(summary)

    return pd.DataFrame(summaries)


def compute_ratchet(summary_df: pd.DataFrame) -> pd.DataFrame:
    summary_df = summary_df.copy()
    summary_df["ratchet_value"] = pd.NA
    summary_df["ratchet_note"] = "Formula not implemented"
    return summary_df


def write_output(
    output_path: Path,
    ratchet_inputs: pd.DataFrame,
    ratchet_summary: pd.DataFrame,
    errors: List[ErrorRecord],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    errors_df = pd.DataFrame([err.as_dict() for err in errors])
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ratchet_inputs.to_excel(writer, sheet_name="RatchetInputs", index=False)
        ratchet_summary.to_excel(writer, sheet_name="RatchetSummary", index=False)
        errors_df.to_excel(writer, sheet_name="Errors", index=False)


def process_file(input_path: Path, output_dir: Path) -> None:
    errors: List[ErrorRecord] = []
    output_path = output_dir / f"{input_path.stem}_ratchet.xlsx"

    try:
        excel = pd.ExcelFile(input_path)
    except Exception as exc:
        log_error(
            errors,
            input_path.name,
            "",
            "file_read_error",
            f"Failed to read Excel file: {exc}",
        )
        write_output(output_path, pd.DataFrame(), pd.DataFrame(), errors)
        return

    pres_sheet = find_sheet_name(excel, "PresTempPipeID")
    prop_sheet = find_sheet_name(excel, "PipeProperties")

    if pres_sheet is None or prop_sheet is None:
        if pres_sheet is None:
            log_error(
                errors,
                input_path.name,
                "",
                "missing_sheet",
                "Missing sheet PresTempPipeID. Found sheets: " + ", ".join(excel.sheet_names),
            )
        if prop_sheet is None:
            log_error(
                errors,
                input_path.name,
                "",
                "missing_sheet",
                "Missing sheet PipeProperties. Found sheets: " + ", ".join(excel.sheet_names),
            )
        write_output(output_path, pd.DataFrame(), pd.DataFrame(), errors)
        return

    pres_df = pd.read_excel(input_path, sheet_name=pres_sheet, header=0, skiprows=[1])
    props_df = pd.read_excel(input_path, sheet_name=prop_sheet, header=0, skiprows=[1])

    pres_df = normalize_columns(pres_df, errors, input_path.name, "PresTempPipeID")
    props_df = normalize_columns(props_df, errors, input_path.name, "PipeProperties")

    long_df, cases = parse_cases(pres_df, errors, input_path.name)
    long_df = coerce_numeric(long_df, CASE_NUMERIC_COLUMNS, errors, input_path.name, "PresTempPipeID")
    props_df = coerce_numeric(props_df, PROPERTY_NUMERIC_COLUMNS, errors, input_path.name, "PipeProperties")

    joined_df = join_properties(long_df, props_df, errors, input_path.name)

    summary_df = build_envelope(joined_df)
    summary_df = compute_ratchet(summary_df)

    ratchet_inputs = joined_df.drop(columns=["pipe_id_key"], errors="ignore")

    write_output(output_path, ratchet_inputs, summary_df, errors)

    print(
        f"Processed {input_path.name}: rows={len(pres_df)}, cases={len(cases)}, "
        f"outputs={output_path}"
    )


def process_all(input_dir: Path, output_dir: Path) -> None:
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    files = [
        path
        for path in input_dir.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not files:
        print(f"No Excel files found in {input_dir}")
        return

    for path in sorted(files):
        process_file(path, output_dir)


def main() -> None:
    parser = argparse.ArgumentParser(description="Ratchet automation processor")
    parser.add_argument(
        "--input",
        dest="input_dir",
        default="data/input",
        help="Input folder containing Excel files",
    )
    parser.add_argument(
        "--output",
        dest="output_dir",
        default="data/output",
        help="Output folder for ratchet results",
    )
    args = parser.parse_args()
    process_all(Path(args.input_dir), Path(args.output_dir))


if __name__ == "__main__":
    main()
