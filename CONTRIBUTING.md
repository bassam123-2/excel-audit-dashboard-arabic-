# Contributing

Thank you for improving the Excel Arabic Dashboard. This guide helps new developers get productive quickly.

## Before you start

1. **Python 3.10+** installed (`python --version`).
2. Clone the repo and create a virtual environment (see [README.md](README.md)).
3. Install dependencies: `pip install -r requirements.txt`.
4. Copy `.env.example` to `.env` if you need custom port or SMTP (optional).
5. Obtain a **sample compliance Excel file** from your team lead — real registers often contain confidential data and are **not** stored in Git.

## Running locally

| Command | Use case |
|---------|----------|
| `python app.py` | Standard dev server (debug on) |
| `python excel_arabic_desktop.py` | Same app, opens browser automatically |
| `run_dashboard.bat` | Windows convenience script |

The app expects uploads under `uploads/` (created automatically, gitignored).

## Excel test data

Workbooks must use sheet **`سجل الالتزام الموحد`** with **header row on Excel row 4** (`header=3` in pandas). Required columns and naming rules are documented in [docs/EXCEL_FORMAT.md](docs/EXCEL_FORMAT.md).

If parsing fails after your change, check the Arabic column names and that you did not break `_rename_columns_to_canonical()` in `app.py`.

## Code conventions

- **Language:** UI strings and column names are Arabic; code comments may be Arabic or English.
- **Scope:** Keep PRs focused — one feature or fix per PR.
- **Style:** Match existing patterns in `app.py` (type hints, private helpers prefixed with `_`).
- **Secrets:** Do not commit `.env`, `smtp.env`, `.xlsx` samples with real data, or `dist/` / `build/` artifacts.

## Project areas (where to edit)

| Change | Files |
|--------|--------|
| Upload / analyze UI | `templates/upload.html`, `templates/analyze.html` |
| Filters, APIs, Excel logic | `app.py` |
| Brand logos by subsidiary | `assets/logos/`, `assets/logos/mapping.json`, `BRAND_LOGO_CODES` in `app.py` |
| Desktop / exe entry | `excel_arabic_desktop.py`, `excel_arabic_dashboard.spec` |
| Dependencies | `requirements.txt` (runtime), `requirements-build.txt` (PyInstaller) |

## Pull requests

1. Create a branch from `main`: `git checkout -b feature/short-description`.
2. Run the app and test with a valid Excel file (upload, filters, legal-text modal, export if touched).
3. For exe-related changes, run `pyinstaller excel_arabic_dashboard.spec` once and smoke-test `dist/excel_arabic_dashboard.exe`.
4. Open a PR with: **what** changed, **why**, and **how you tested**.

## Reporting bugs

Include:

- OS and Python version
- Steps to reproduce
- Whether it happens in `python app.py` only, or also in the built `.exe`
- Redacted screenshot or error message (no confidential Excel content)

## Questions

If something is unclear, open a GitHub Issue labeled `question` or ask the repo maintainer for a sample workbook and SMTP test account.
