# Sample workbook for developers

Production compliance registers are **confidential** and must not be pushed to GitHub.

## What new developers need

1. A `.xlsx` file that uses sheet **سجل الالتزام الموحد** with headers on row 4.
2. At least the [required columns](EXCEL_FORMAT.md) populated on a few rows.
3. Optional: one row with `email` filled to test SMTP.

## How maintainers should share it

- Internal SharePoint / Teams / email, **or**
- GitHub Release asset on a **private** repo, **or**
- Password-protected zip outside the repo.

## Redacted sample (optional)

You may add `samples/redacted-register.xlsx` to the repo only if:

- All personal emails, IDs, and company-sensitive text are removed or fictionalized, and
- Your organization approves publishing it.

Keep `samples/` in `.gitignore` until a redacted file is ready.
