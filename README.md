# excel2pdf

## Program Summary
- Streamlit web app that converts an uploaded `.xlsx` to PDF using the local Excel application via COM automation.
- Lets you preview the sheet, fill blank cells, and optionally overlay a text string at a chosen X/Y position on the generated PDF.

## How to Use
- Install dependencies: `pip install -r requirements.txt`
- Run the app: `streamlit run app.py` (Not verified)
- In the UI: upload an `.xlsx`, edit blank cells if any, optionally set overlay text and coordinates, then click the convert button.

## Completion Status
- **Partial**: Core flow exists (upload → edit blanks → export PDF → optional overlay) but it is Windows/Excel-dependent and lacks tests or error handling, so portability and robustness are limited.
