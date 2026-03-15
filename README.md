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

## Program Summary
- Streamlit UI that loads an uploaded `.xlsx`, allows filling blank cells, and exports to PDF via Excel COM automation on Windows.
- Optional text overlay is rendered onto the generated PDF using ReportLab + PyPDF2.

## How to Use
- Install dependencies: `pip install -r requirements.txt` (Not verified)
- Run: `streamlit run app.py` (Not verified)
- Use the UI to upload an `.xlsx`, fill blanks, optionally set overlay text/X/Y, then convert and download the PDF.

## Completion Status
- **Partial**: The core conversion flow works in code but is Windows + local Excel dependent, with minimal validation and no tests.

## Program Summary
- Streamlit app that loads an uploaded `.xlsx`, lets you fill blank cells, then exports to PDF via Microsoft Excel COM automation.
- Optional overlay text is drawn onto the generated PDF at user-specified X/Y coordinates using ReportLab and PyPDF2.

## How to Use
- Install dependencies: `pip install -r requirements.txt` (Not verified)
- Run the app: `streamlit run app.py` (Not verified)
- In the UI: upload an `.xlsx`, fill blanks if prompted, set overlay text/X/Y if needed, then convert and download the PDF. (Not verified)

## Completion Status
- **Partial**: The main flow is implemented but depends on Windows + a locally installed Excel instance and has no tests or robust error handling.

## Program Summary
- Streamlit app that uploads an `.xlsx`, lets you fill blank cells, then exports to PDF via Microsoft Excel COM automation.
- Optional overlay text can be drawn onto the PDF at user-specified X/Y coordinates using ReportLab and PyPDF2.

## How to Use
- Install dependencies: `pip install -r requirements.txt` (Not verified)
- Run the app: `streamlit run app.py` (Not verified)
- In the UI: upload an `.xlsx`, fill any blanks, optionally set overlay text/X/Y, then convert and download the PDF. (Not verified)
- Requires Windows with locally installed Excel (Not verified).

## Completion Status
- **Partial**: Core conversion and overlay flow exists, but it is Windows/Excel-dependent and has minimal validation and no tests.

## Program Summary
- Streamlit app that uploads an `.xlsx`, lets you fill blank cells, then exports to PDF via Microsoft Excel COM automation.
- Optional overlay text can be drawn onto each PDF page at user-specified X/Y coordinates using ReportLab and PyPDF2.

## How to Use
- Install dependencies: `pip install -r requirements.txt` (Not verified)
- Run the app: `streamlit run app.py` (Not verified)
- In the UI: upload an `.xlsx`, fill blanks if prompted, optionally set overlay text/X/Y, then convert and download the PDF. (Not verified)
- Requires Windows with locally installed Excel (Not verified).

## Completion Status
- **Partial**: Core conversion and overlay flow exists, but it is Windows/Excel-dependent and has minimal validation and no tests.
