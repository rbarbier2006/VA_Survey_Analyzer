# VA Survey Analyzer

Streamlit web app that scores uploaded `.xlsx` survey files and returns a scored copy with additional columns appended on the right.

## Supported surveys

1. **RAND Health Survey Questionnaire (modified for wheelchair users)**
2. **Pittsburgh Sleep Quality Index (PSQI)**

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Input format

- The first row is treated as headers.
- Each question column should begin with the question identifier:
  - RAND: `1.`, `13.`, `20.`, etc.
  - PSQI: `1.`, `2.`, `5a.`, `5b.`, ..., `5j.`, `6.`, `7.`, `8.`, `9.`

## RAND scoring behavior

- Applies existing recoding + scale averages.
- Questions `3-12` are intentionally skipped.

## PSQI scoring behavior

- Uses only self-rated items (`1-9`, including `5a-5j`) for scoring.
- Bed partner/roommate items (`10`, `10a-10e`) are ignored.
- Appends these output columns:
  - `psqi_component_1_subjective_sleep_quality`
  - `psqi_component_2_sleep_latency`
  - `psqi_component_3_sleep_duration`
  - `psqi_component_4_habitual_sleep_efficiency`
  - `psqi_component_5_sleep_disturbances`
  - `psqi_component_6_sleep_medication`
  - `psqi_component_7_daytime_dysfunction`
  - `psqi_global_score`
  - `psqi_poor_sleep_flag_gt5`

### PSQI parsing notes

- Question 2 accepts values like `20`, `20 min`, `20 minutes` (text after number is ignored).
- Question 4 accepts values like `9`, `9 hrs`, `6.5`.
- Questions 1 and 3 accept common time styles like `2230`, `07:30`, `11:00 PM`.
- For PSQI bedtime (Question 1), if a plain hour like `11` is entered, it is interpreted as `11 PM` (23:00) to avoid military-time mismatch.
- Frequency/problem text matching is tolerant to truncated exports (e.g., `Not during the past mo`) and common minor typos.
- For item `5j`, missing value is treated as `0` per your rule.
