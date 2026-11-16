"""
Core long -> wide conversion logic for optically-read questionnaire data.

The input is "long" format: one row per (submission_id, question, matrix item).
The output is "wide" format: one row per submission_id, one column per variable.

This module is GUI-free so it can be unit-tested and reused.
"""

import os
import re

import openpyxl  # noqa: F401  (pandas loads it lazily; explicit so PyInstaller bundles it)
import pandas as pd

# Columns we expect in a long-format export. A sheet is considered the data
# sheet if it contains at least this minimal subset.
REQUIRED_COLS = {
    "submission_id",
    "question_code",
    "question_type",
    "matrix_row",
    "answer_1",
}

ANSWER_COLS = ["answer_1", "answer_2", "answer_3"]
META_COLS = ["form_id", "form_title"]

# matrix_row values that mean "this question has no sub-item"
_NON_MATRIX_MARKERS = {"row_0", "nan", "none", ""}


def _is_blank(v):
    """True for None / NaN / pd.NA cells."""
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _matrix_key(matrix_row):
    """Normalised matrix item: None for simple questions, else the label.

    'row_0', blank and NaN all mean "no sub-item", so they must map to the
    same variable even when different files encode them differently.
    """
    if _is_blank(matrix_row):
        return None
    s = str(matrix_row)
    return None if s.strip().lower() in _NON_MATRIX_MARKERS else s


def _fmt_answer(v):
    """Cell value -> text, without the '3.0' artefact pandas adds to integer
    answers in columns that also contain blanks (read as float64)."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def matrix_suffix(text):
    """Extract a short item label from a matrix row like 'a.\\tde penser...'.

    Returns e.g. 'a', 'b', '12' or None when no clean prefix is found.
    """
    m = re.match(r"\s*([A-Za-z0-9]{1,3})\s*[.\):\-]", str(text))
    return m.group(1).lower() if m else None


def _base_variable(code, matrix_row):
    """Best-effort variable name for a single (code, matrix_row) pair."""
    code = str(code)
    mr = _matrix_key(matrix_row)
    if mr is None:
        return code
    mr = mr.strip()
    suf = matrix_suffix(mr)
    if suf:
        return f"{code}_{suf}"
    # Fallback: sanitised, truncated text so the variable is still readable.
    clean = re.sub(r"[^A-Za-z0-9]+", "_", mr).strip("_")[:12]
    return f"{code}_{clean}" if clean else code


def assign_variables(df):
    """Add a unique 'variable' column to a copy of df.

    Guarantees that each distinct (question_code, matrix_row) maps to a
    distinct variable, disambiguating collisions with a numeric suffix.
    Returns (df_with_variable, ordered_unique_variables).
    """
    df = df.copy()
    key_to_var = {}
    used = set()
    variables = []
    order = []
    for code, mr in zip(df["question_code"], df["matrix_row"]):
        key = (str(code), _matrix_key(mr))
        if key in key_to_var:
            variables.append(key_to_var[key])
            continue
        base = _base_variable(code, mr)
        var = base
        i = 2
        while var in used:
            var = f"{base}_{i}"
            i += 1
        used.add(var)
        key_to_var[key] = var
        variables.append(var)
        order.append(var)
    df["variable"] = variables
    return df, order


def _combine_answers(row):
    """Join the non-empty answer_1..answer_3 cells with ' | '."""
    parts = []
    for c in ANSWER_COLS:
        if c in row.index:
            v = row[c]
            if not _is_blank(v):
                s = _fmt_answer(v)
                if s != "" and s.lower() != "nan":
                    parts.append(s)
    return " | ".join(parts)


def build_codebook(df, order=None):
    """One row per variable: code, matrix item, type, full text, page."""
    if "variable" not in df.columns:
        df, order = assign_variables(df)
    first = df.drop_duplicates(subset="variable", keep="first").set_index("variable")
    if order is None:
        order = list(first.index)
    rows = []
    for var in order:
        r = first.loc[var]
        rows.append(
            {
                "variable": var,
                "question_code": r.get("question_code", ""),
                "matrix_row": _matrix_key(r.get("matrix_row")) or "",
                "question_type": r.get("question_type", ""),
                "page_number": r.get("page_number", ""),
                "question_text": r.get("question_text", ""),
            }
        )
    return pd.DataFrame(rows)


def to_wide(df):
    """Convert a long-format DataFrame to wide.

    Returns (wide_df, codebook_df). wide_df has one row per submission_id,
    meta columns first, then one column per variable in questionnaire order.
    """
    df, order = assign_variables(df)
    df = df.copy()
    df["value"] = df.apply(_combine_answers, axis=1)

    wide = df.pivot_table(
        index="submission_id",
        columns="variable",
        values="value",
        aggfunc="first",
    )
    wide = wide.reindex(columns=order)

    present_meta = [c for c in META_COLS if c in df.columns]
    if present_meta:
        meta = df.groupby("submission_id")[present_meta].first()
        out = meta.join(wide)
    else:
        out = wide
    out = out.reset_index()
    codebook = build_codebook(df, order)
    return out, codebook


def find_data_sheet(path):
    """Return (sheet_name, DataFrame) of the first sheet that looks like a
    long-format data sheet. Raises ValueError if none match."""
    # Context manager: otherwise the workbook stays locked on Windows.
    with pd.ExcelFile(path) as xl:
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet)
            except Exception:
                continue
            if df.empty:
                continue
            if REQUIRED_COLS.issubset(set(df.columns)):
                return sheet, df
    raise ValueError(
        f"No long-format data sheet found in '{os.path.basename(path)}'. "
        f"Expected columns include: {', '.join(sorted(REQUIRED_COLS))}."
    )


def load_long(path):
    """Load the long-format data sheet from an .xlsx file."""
    _, df = find_data_sheet(path)
    return df


def export_wide(wide, codebook, out_path):
    """Write wide + codebook sheets to an .xlsx file."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        wide.to_excel(xw, sheet_name="wide", index=False)
        codebook.to_excel(xw, sheet_name="codebook", index=False)
    return out_path
