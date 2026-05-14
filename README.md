# LL Sizing Report

Local Flask + Chart.js dashboard and Streamlit app for reviewing Calypso CMM `.xls` sizing data.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run Local Flask App

```powershell
python server.py
```

Open:

```text
http://127.0.0.1:5050
```

## Run Streamlit App

```powershell
streamlit run streamlit_app.py
```

For Streamlit Cloud, use:

```text
streamlit_app.py
```

as the app entrypoint.

## Safety Notes

- The server binds to `127.0.0.1` only.
- Folder browse/import/export APIs are restricted to this project folder.
- Report filenames are sanitized and saved only under `Reports`.
- Parsed `.xls` warnings are returned to the UI and included in exported reports.
- The Streamlit app does not save reports to the server. It generates the HTML in memory and serves it through a download button.

## Data Flow

### Flask

1. Import a batch folder under `SIZER STUDY`.
2. The backend parses required Calypso characteristics from each `.xls`.
3. The UI calculates and charts AS/DS upper/lower positions.
4. Export writes an HTML report to `Reports`.

### Streamlit

1. Upload one or more `.xls` files from the sidebar.
2. Assign the batch type and optional sizer size.
3. Review tables, status flags, and parsing warnings.
4. Click `Download HTML Report` to download the generated report locally.
