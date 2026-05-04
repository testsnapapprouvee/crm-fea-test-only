# AGENTS.md

## Cursor Cloud specific instructions

### Overview

CRM Asset Management — a monolithic **Streamlit** (Python) web app for fund distribution / asset management sales teams. Single service, no external dependencies beyond Python packages.

**Source files:** `app.py` (UI, ~3610 lines), `database.py` (SQLite DAL, ~1633 lines), `pdf_generator.py` (ReportLab PDF export, ~766 lines).

### Running the app

```
streamlit run app.py --server.port 8501 --server.headless true
```

The app auto-creates its SQLite database (`crm_asset_management.db`) on first run via `database.init_db()`. No migrations or external DB servers needed.

### Linting / Testing

No lint tools or test frameworks are configured in this repo. Use `python3 -m py_compile <file>` to verify syntax. There are no automated tests.

### Known gotchas

- **`database.py` corruption:** The file was previously overwritten with a copy of `app.py` (identical MD5). If the app fails to start with import errors like `db.get_connection()` or `db.REGIONS_REFERENTIEL` not found, restore it: `git checkout ecdb914 -- database.py`.
- **kaleido version:** `requirements.txt` pins `kaleido==0.2.1` — newer versions may break Plotly static image export.
- The UI is entirely in French (labels, buttons, messages).
