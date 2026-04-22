# VA Survey Analyzer

Simple Streamlit web app that:
- takes an input `.xlsx` survey file,
- copies the original columns,
- appends recoded question scores,
- appends scale averages,
- lets you download the scored `.xlsx`.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Input format

- Header row should start each question column with the question number, e.g.:
  - `1. In general, would you say...`
  - `13. Cut down the amount of time...`
- The app extracts the leading number from each column name.

## Scoring behavior implemented

- Table 1 recoding rules implemented for questions used by these scales.
- Questions `3-12` are intentionally skipped for now.
- Scale averages implemented from Table 2 except **Physical functioning** (because it depends on `3-12`).

## Output columns appended

- `recoded_q{N}` for each processed question.
- Scale columns:
  - `role_limitations_physical_health`
  - `role_limitations_emotional_problems`
  - `energy_fatigue`
  - `emotional_well_being`
  - `social_functioning`
  - `pain`
  - `general_health`

## Response parsing

- Supports numeric categories (`1`, `2`, `3`, ...).
- Supports common text values for:
  - Question 1 (`Excellent` ... `Poor`)
  - Question 2 (`Much better now ...` to `Much worse now ...`)
  - Questions `13-19` (`Yes` / `No`)

If your source text differs from those phrases, use numeric category values in the input workbook.
