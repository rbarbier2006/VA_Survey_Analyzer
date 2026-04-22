import io
import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Survey Analyzer", layout="wide")


# -----------------------------
# Shared helpers
# -----------------------------
def normalize_text(value: Any) -> str:
    return str(value).strip().lower()


def extract_leading_number(col_name: str) -> int | None:
    match = re.match(r"\s*(\d+)", str(col_name))
    if not match:
        return None
    return int(match.group(1))


def find_column_by_question_id(df: pd.DataFrame, question_id: str) -> str | None:
    """Match headers like '5a.', '5a ', '5A ...' by leading question token."""
    qid = question_id.strip().lower()
    pattern = re.compile(rf"^\s*{re.escape(qid)}\b")
    for col in df.columns:
        if pattern.search(str(col).lower()):
            return col
    return None


def parse_first_float(value: Any) -> float | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value)
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def parse_time_to_hours(value: Any) -> float | None:
    """Return hour-of-day in [0,24), supporting 2230, 07:30, 730, 11:00 PM, etc."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        raw = str(int(value)).zfill(4)
        hour = int(raw[-4:-2])
        minute = int(raw[-2:])
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return hour + minute / 60.0

    text = str(value).strip().lower()

    for fmt in ["%H:%M", "%H%M", "%I:%M %p", "%I %p"]:
        try:
            dt = datetime.strptime(text.upper(), fmt)
            return dt.hour + dt.minute / 60.0
        except ValueError:
            pass

    match = re.search(r"(\d{1,2})(?::?(\d{2}))?", text)
    if match:
        h = int(match.group(1))
        m = int(match.group(2) or 0)
        if 0 <= h <= 23 and 0 <= m <= 59:
            return h + m / 60.0

    return None


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="scored_output")
    output.seek(0)
    return output.read()


# -----------------------------
# RAND (modified wheelchair) processor
# -----------------------------
RAND_RECODE_MAP: dict[str, dict[int, int]] = {
    "A": {1: 100, 2: 75, 3: 50, 4: 25, 5: 0},
    "B": {1: 0, 2: 50, 3: 100},
    "C": {1: 0, 2: 100},
    "D": {1: 100, 2: 80, 3: 60, 4: 40, 5: 20, 6: 0},
    "E": {1: 0, 2: 20, 3: 40, 4: 60, 5: 80, 6: 100},
    "F": {1: 0, 2: 25, 3: 50, 4: 75, 5: 100},
}

RAND_QUESTION_TO_GROUP = {
    **{q: "A" for q in [1, 2, 20, 22, 34, 36]},
    **{q: "B" for q in [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]},
    **{q: "C" for q in [13, 14, 15, 16, 17, 18, 19]},
    **{q: "D" for q in [21, 23, 26, 27, 30]},
    **{q: "E" for q in [24, 25, 28, 29, 31]},
    **{q: "F" for q in [32, 33, 35]},
}

RAND_TEXT_MAP: dict[int, dict[str, int]] = {
    1: {"excellent": 1, "very good": 2, "good": 3, "fair": 4, "poor": 5},
    2: {
        "much better now than one year ago": 1,
        "much better now than 8 weeks ago": 1,
        "somewhat better now than one year ago": 2,
        "somewhat better now than 8 weeks ago": 2,
        "about the same": 3,
        "somewhat worse now than one year ago": 4,
        "somewhat worse now than 8 weeks ago": 4,
        "much worse now than one year ago": 5,
        "much worse now than 8 weeks ago": 5,
    },
    **{q: {"yes": 1, "no": 2} for q in [13, 14, 15, 16, 17, 18, 19]},
}

RAND_ALIASES: dict[int, dict[int, list[str]]] = {
    20: {1: ["not at all"], 2: ["slightly"], 3: ["moderately"], 4: ["quite a bit"], 5: ["extremely"]},
    21: {1: ["none"], 2: ["very mild"], 3: ["mild"], 4: ["moderate"], 5: ["severe"], 6: ["very severe"]},
    22: {1: ["not at all"], 2: ["a little bit", "little bit"], 3: ["moderately"], 4: ["quite a bit"], 5: ["extremely"]},
    **{q: {1: ["all of the time", "all of t"], 2: ["most of the time", "most of t"], 3: ["a good bit of the time", "good bit"], 4: ["some of the time", "some of t"], 5: ["a little bit of the time", "a little bit"], 6: ["none of the time", "none of t", "none of th"]} for q in [23, 24, 25, 26, 27, 28, 29, 30, 31]},
    32: {1: ["all of the time", "all of t"], 2: ["most of the time", "most of t"], 3: ["some of the time", "some of t"], 4: ["a little of the time", "a little bit"], 5: ["none of the time", "none of t", "none of th"]},
    **{q: {1: ["definitely true", "all of the time", "all of th"], 2: ["mostly true", "most of the time", "most of th"], 3: ["dont know", "don't know", "some of the time", "some of th", "a good bit of the time"], 4: ["mostly false", "a little bit of the time", "little bit"], 5: ["definitely false", "none of the time", "none of th"]} for q in [33, 34, 35, 36]},
}

RAND_SCALES: dict[str, list[int]] = {
    "role_limitations_physical_health": [13, 14, 15, 16],
    "role_limitations_emotional_problems": [17, 18, 19],
    "energy_fatigue": [23, 27, 29, 31],
    "emotional_well_being": [24, 25, 26, 28, 30],
    "social_functioning": [20, 32],
    "pain": [21, 22],
    "general_health": [1, 33, 34, 35, 36],
}


def parse_rand_category(question_number: int, value: Any) -> int | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)) and float(value).is_integer():
        return int(value)

    text = normalize_text(value)
    if text.isdigit():
        return int(text)

    q_map = RAND_TEXT_MAP.get(question_number)
    if q_map and text in q_map:
        return q_map[text]

    aliases = RAND_ALIASES.get(question_number)
    if aliases:
        for cat, variants in aliases.items():
            if any(text == v or text.startswith(v) or v in text for v in variants):
                return cat
    return None


def process_rand_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out_df = df.copy()
    recoded_cols_by_question: dict[int, str] = {}

    for col in df.columns:
        q_num = extract_leading_number(str(col))
        if q_num is None or 3 <= q_num <= 12:
            continue
        group = RAND_QUESTION_TO_GROUP.get(q_num)
        if group is None:
            continue
        recoded_col = f"recoded_q{q_num}"
        categories = df[col].apply(lambda val: parse_rand_category(q_num, val))
        out_df[recoded_col] = categories.map(RAND_RECODE_MAP[group])
        recoded_cols_by_question[q_num] = recoded_col

    for scale_name, questions in RAND_SCALES.items():
        cols = [recoded_cols_by_question[q] for q in questions if q in recoded_cols_by_question]
        out_df[scale_name] = out_df[cols].mean(axis=1, skipna=True) if cols else pd.NA

    return out_df


# -----------------------------
# PSQI processor
# -----------------------------
PSQI_FREQ_MAP = {
    "not during the past month": 0,
    "less than once a week": 1,
    "once or twice a week": 2,
    "three or more times a week": 3,
}

PSQI_Q6_MAP = {"very good": 0, "fairly good": 1, "fairly bad": 2, "very bad": 3}
PSQI_Q9_MAP = {
    "no problem at all": 0,
    "only a very slight problem": 1,
    "somewhat of a problem": 2,
    "a very big problem": 3,
}


def parse_psqi_choice(value: Any, mapping: dict[str, int]) -> int | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)) and 0 <= int(value) <= 3 and float(value).is_integer():
        return int(value)

    text = normalize_text(value)
    if text.isdigit() and int(text) in [0, 1, 2, 3]:
        return int(text)

    for label, score in mapping.items():
        # Accept exact, truncated, and "label with extra suffix" variants.
        if (
            text == label
            or text.startswith(label)
            or label.startswith(text)
            or label in text
        ):
            return score

    # Extra tolerant handling for common typo/truncation patterns in Q9 exports.
    if "very slight" in text:
        return 1
    if "very big" in text:
        return 3
    if "somewhat" in text and "problem" in text:
        return 2
    if "no problem" in text:
        return 0

    # Extra tolerant handling for truncated frequency labels.
    if "not during" in text:
        return 0
    if "less than once" in text:
        return 1
    if "once or twice" in text:
        return 2
    if "three or more" in text:
        return 3
    return None


def recode_q2_minutes(minutes: float | None) -> int | None:
    if minutes is None:
        return None
    if minutes <= 15:
        return 0
    if minutes <= 30:
        return 1
    if minutes <= 60:
        return 2
    return 3


def recode_q4_hours(hours: float | None) -> int | None:
    if hours is None:
        return None
    if hours > 7:
        return 0
    if hours >= 6:
        return 1
    if hours >= 5:
        return 2
    return 3


def recode_component_2_sum(total: float | None) -> int | None:
    if total is None or pd.isna(total):
        return None
    if total == 0:
        return 0
    if total <= 2:
        return 1
    if total <= 4:
        return 2
    return 3


def recode_component_5_sum(total: float | None) -> int | None:
    if total is None or pd.isna(total):
        return None
    if total == 0:
        return 0
    if total <= 9:
        return 1
    if total <= 18:
        return 2
    return 3


def recode_efficiency(hse_percent: float | None) -> int | None:
    if hse_percent is None or pd.isna(hse_percent):
        return None
    if hse_percent > 85:
        return 0
    if hse_percent >= 75:
        return 1
    if hse_percent >= 65:
        return 2
    return 3


def time_in_bed_hours(bed_time_h: float | None, wake_time_h: float | None) -> float | None:
    if bed_time_h is None or wake_time_h is None:
        return None
    diff = wake_time_h - bed_time_h
    if diff <= 0:
        diff += 24
    return diff


def process_psqi_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out_df = df.copy()

    # locate columns by PSQI question id
    col_q1 = find_column_by_question_id(df, "1")
    col_q2 = find_column_by_question_id(df, "2")
    col_q3 = find_column_by_question_id(df, "3")
    col_q4 = find_column_by_question_id(df, "4")
    col_q6 = find_column_by_question_id(df, "6")
    col_q7 = find_column_by_question_id(df, "7")
    col_q8 = find_column_by_question_id(df, "8")
    col_q9 = find_column_by_question_id(df, "9")
    col_q5a = find_column_by_question_id(df, "5a")

    q5b_to_q5j_cols = {qid: find_column_by_question_id(df, qid) for qid in ["5b", "5c", "5d", "5e", "5f", "5g", "5h", "5i", "5j"]}

    # raw helper scores
    q6_score = df[col_q6].apply(lambda v: parse_psqi_choice(v, PSQI_Q6_MAP)) if col_q6 else pd.Series(pd.NA, index=df.index)
    q2_minutes = df[col_q2].apply(parse_first_float) if col_q2 else pd.Series(pd.NA, index=df.index)
    q2_score = q2_minutes.apply(recode_q2_minutes)
    q5a_score = df[col_q5a].apply(lambda v: parse_psqi_choice(v, PSQI_FREQ_MAP)) if col_q5a else pd.Series(pd.NA, index=df.index)

    q4_hours = df[col_q4].apply(parse_first_float) if col_q4 else pd.Series(pd.NA, index=df.index)
    q4_score = q4_hours.apply(recode_q4_hours)

    bedtime_hours = df[col_q1].apply(parse_time_to_hours) if col_q1 else pd.Series(pd.NA, index=df.index)
    wake_hours = df[col_q3].apply(parse_time_to_hours) if col_q3 else pd.Series(pd.NA, index=df.index)
    hours_in_bed = pd.Series(
        [time_in_bed_hours(b, w) for b, w in zip(bedtime_hours, wake_hours)],
        index=df.index,
    )
    hse_percent = (q4_hours / hours_in_bed) * 100
    component_4 = hse_percent.apply(recode_efficiency)

    q5_scores = {}
    for qid, col in q5b_to_q5j_cols.items():
        if col:
            q5_scores[qid] = df[col].apply(lambda v: parse_psqi_choice(v, PSQI_FREQ_MAP))
        else:
            q5_scores[qid] = pd.Series(pd.NA, index=df.index)

    # Special rule: 5j missing -> 0
    q5_scores["5j"] = q5_scores["5j"].fillna(0)

    q7_score = df[col_q7].apply(lambda v: parse_psqi_choice(v, PSQI_FREQ_MAP)) if col_q7 else pd.Series(pd.NA, index=df.index)
    q8_score = df[col_q8].apply(lambda v: parse_psqi_choice(v, PSQI_FREQ_MAP)) if col_q8 else pd.Series(pd.NA, index=df.index)
    q9_score = df[col_q9].apply(lambda v: parse_psqi_choice(v, PSQI_Q9_MAP)) if col_q9 else pd.Series(pd.NA, index=df.index)

    # Components
    component_1 = q6_score
    component_2 = (q2_score + q5a_score).apply(recode_component_2_sum)
    component_3 = q4_score
    component_5_raw_sum = sum(q5_scores.values())
    component_5 = component_5_raw_sum.apply(recode_component_5_sum)
    component_6 = q7_score
    component_7 = (q8_score + q9_score).apply(recode_component_2_sum)

    out_df["psqi_component_1_subjective_sleep_quality"] = component_1
    out_df["psqi_component_2_sleep_latency"] = component_2
    out_df["psqi_component_3_sleep_duration"] = component_3
    out_df["psqi_component_4_habitual_sleep_efficiency"] = component_4
    out_df["psqi_component_5_sleep_disturbances"] = component_5
    out_df["psqi_component_6_sleep_medication"] = component_6
    out_df["psqi_component_7_daytime_dysfunction"] = component_7

    component_cols = [
        "psqi_component_1_subjective_sleep_quality",
        "psqi_component_2_sleep_latency",
        "psqi_component_3_sleep_duration",
        "psqi_component_4_habitual_sleep_efficiency",
        "psqi_component_5_sleep_disturbances",
        "psqi_component_6_sleep_medication",
        "psqi_component_7_daytime_dysfunction",
    ]
    out_df["psqi_global_score"] = out_df[component_cols].sum(axis=1, min_count=7)
    out_df["psqi_poor_sleep_flag_gt5"] = out_df["psqi_global_score"].apply(
        lambda x: pd.NA if pd.isna(x) else ("Yes" if x > 5 else "No")
    )

    return out_df


# -----------------------------
# UI
# -----------------------------
st.title("Survey Analyzer")
st.write("Upload an Excel file, choose the survey type, and download a scored copy with appended score columns.")

survey_type = st.selectbox(
    "Survey type",
    [
        "RAND Health Survey Questionnaire (modified for wheelchair users)",
        "Pittsburgh Sleep Quality Index (PSQI)",
    ],
)

uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded_file is not None:
    try:
        input_df = pd.read_excel(uploaded_file)
        st.subheader("Preview: input")
        st.dataframe(input_df.head(20), use_container_width=True)

        if survey_type == "Pittsburgh Sleep Quality Index (PSQI)":
            scored_df = process_psqi_dataframe(input_df)
            suggested_name = "psqi_scored_output.xlsx"
        else:
            scored_df = process_rand_dataframe(input_df)
            suggested_name = "rand_scored_output.xlsx"

        st.subheader("Preview: output")
        st.dataframe(scored_df.head(20), use_container_width=True)

        file_bytes = dataframe_to_excel_bytes(scored_df)
        st.download_button(
            label="Download scored Excel",
            data=file_bytes,
            file_name=suggested_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Done. Download your scored file above.")
    except Exception as exc:
        st.error(f"Could not process file: {exc}")

st.markdown("---")
st.caption(
    "PSQI notes: only self-rated items are used for scoring, bed-partner items are ignored, "
    "question 2 supports values like '20', '20 min', etc., and final output appends 7 components + global score."
)
