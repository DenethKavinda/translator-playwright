# Transliteration Accuracy Testing

This workspace contains the Playwright automation used to test the Pixelssuite chat transliterator and record results in the assignment workbook.

## Files

- `test_automation.py` runs the browser automation and writes the actual output and status columns.
- `seed_test_cases.py` populates `Assignment 1 - Test cases.xlsx` with the 50 negative test cases.
- `Assignment 1 - Test cases.xlsx` is the workbook submitted for the assignment.

## Setup

Use the workspace virtual environment at `.venv\Scripts\python.exe`.

Install dependencies:

```powershell
python -m pip install -r requirements.txt
python -m playwright install
```

## Regenerate the workbook

```powershell
python seed_test_cases.py
```

## Run the automation

```powershell
python test_automation.py --headless
```

The script reads the Singlish input column, sends each case to the live translator, captures the actual Sinhala output, and writes the status back into the workbook.
