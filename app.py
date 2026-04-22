import io
import re
from typing import Any

import pandas as pd
import streamlit as st

st.set_page_config(page_title="VA Survey Analyzer", layout="wide")

# Recode tables (Table 1)
RECODE_MAP: dict[str, int] = {
    "A": {1: 100, 2: 75, 3: 50, 4: 25, 5: 0},  # 1,2,20,22,34,36
    "B": {1: 0, 2: 50, 3: 100},  # 3-12 (currently excluded)
    "C": {1: 0, 2: 100},  # 13-19
    "D": {1: 100, 2: 80, 3: 60, 4: 40, 5: 20, 6: 0},  # 21,23,26,27,30
    "E": {1: 0, 2: 20, 3: 40, 4: 60, 5: 80, 6: 100},  # 24,25,28,29,31
    "F": {1: 0, 2: 25, 3: 50, 4: 75, 5: 100},  # 32,33,35
}

QUESTION_TO_GROUP = {
    **{q: "A" for q in [1, 2, 20, 22, 34, 36]},
    **{q: "B" for q in [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]},
    **{q: "C" for q in [13, 14, 15, 16, 17, 18, 19]},
    **{q: "D" for q in [21, 23, 26, 27, 30]},
    **{q: "E" for q in [24, 25, 28, 29, 31]},
    **{q: "F" for q in [32, 33, 35]},
}

# Common text-to-category mappings by question set.
TEXT_CATEGORY_MAP: dict[int, dict[str, int]] = {
    1: {
        "excellent": 1,
        "very good": 2,
        "good": 3,
        "fair": 4,
        "poor": 5,
    },
    2: {
        "much better now than one year ago": 1,
        "somewhat better now than one year ago": 2,
        "about the same": 3,
        "somewhat worse now than one year ago": 4,
        "much worse now than one year ago": 5,
    },
    **{
        q: {"yes": 1, "no": 2}
        for q in [13, 14, 15, 16, 17, 18, 19]
    },
}

# Scales from Table 2 (excluding Physical functioning 3-12 per user request)
SCALES: dict[str, list[int]] = {
    "role_limitations_physical_health": [13, 14, 15, 16],
    "role_limitations_emotional_problems": [17, 18, 19],
    "energy_fatigue": [23, 27, 29, 31],
    "emotional_well_being": [24, 25, 26, 28, 30],
    "social_functioning": [20, 32],
    "pain": [21, 22],
    "general_health": [1, 33, 34, 35, 36],
}


def normalize_text(value: Any) -> str:
    return str(value).strip().lower()


def extract_question_number(col_name: str) -> int | None:
    match = re.match(r"\s*(\d+)", str(col_name))
    if not match:
        return None
    return int(match.group(1))


def parse_category(question_number: int, value: Any) -> int | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        int_value = int(value)
        if abs(float(value) - int_value) < 1e-9:
            return int_value

    text = normalize_text(value)
    if text.isdigit():
        return int(text)

    question_map = TEXT_CATEGORY_MAP.get(question_number)
    if question_map:
        return question_map.get(text)

    return None


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out_df = df.copy()

    recoded_cols_by_question: dict[int, str] = {}

    for col in df.columns:
        q_num = extract_question_number(str(col))
        if q_num is None:
            continue

        # Explicitly skip Q3-12 as requested.
        if 3 <= q_num <= 12:
            continue

        group = QUESTION_TO_GROUP.get(q_num)
        if group is None:
            continue

        recode_map = RECODE_MAP[group]
        recoded_col = f"recoded_q{q_num}"

        categories = df[col].apply(lambda val: parse_category(q_num, val))
        out_df[recoded_col] = categories.map(recode_map)
        recoded_cols_by_question[q_num] = recoded_col

    # Scale averages (Table 2), using available recoded items.
    for scale_name, questions in SCALES.items():
        cols = [recoded_cols_by_question[q] for q in questions if q in recoded_cols_by_question]
        if cols:
            out_df[scale_name] = out_df[cols].mean(axis=1, skipna=True)
        else:
            out_df[scale_name] = pd.NA

    return out_df


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="scored_output")
    output.seek(0)
    return output.read()


st.title("VA Survey Analyzer")
st.write(
    "Upload an Excel file. The app copies your original data and appends recoded question scores "
    "plus scale averages on the right."
)

uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded_file is not None:
    try:
        input_df = pd.read_excel(uploaded_file)
        st.subheader("Preview: input")
        st.dataframe(input_df.head(20), use_container_width=True)

        scored_df = process_dataframe(input_df)
        st.subheader("Preview: output")
        st.dataframe(scored_df.head(20), use_container_width=True)

        file_bytes = dataframe_to_excel_bytes(scored_df)

        st.download_button(
            label="Download scored Excel",
            data=file_bytes,
            file_name="scored_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success("Done. Download your scored file above.")
    except Exception as exc:
        st.error(f"Could not process file: {exc}")

st.markdown("---")
st.caption(
    "Notes: (1) Questions 3-12 are intentionally not processed right now. "
    "(2) Text responses are mapped for common options (e.g., Yes/No). "
    "If your wording differs, use numeric categories in the input (1..N)."
)
