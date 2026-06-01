# Excel workbook format

The dashboard reads the **سجل الالتزام الموحد** sheet from compliance register templates.

## Structure

| Rule | Value |
|------|--------|
| Sheet name | `سجل الالتزام الموحد` (first sheet used if missing) |
| Header row | Excel row **4** (pandas `header=3`) |
| First data row | Excel row **5** |
| File types | `.xlsx`, `.xls` |

## Required columns

After import, these logical columns must be present (physical header names may vary slightly; see matching logic in `app.py` → `_rename_columns_to_canonical`):

| Canonical name | Typical Arabic header tokens |
|----------------|------------------------------|
| تصنيف المخاطر الكامنة | inherent risk |
| تصنيف المخاطر المتبقية | residual risk |
| الحالة | status |
| الإدارة المسؤولة | department |
| المشرع | legislator |
| اسم النظام | system name (or **النظام**) |
| الهيئة التابعة | authority |
| اللائحة | regulation |
| النص النظامي | legal text (or **النص بالكامل**) |
| حالة الالتزام | compliance status |
| فئة الضوابط الرقابية | control category |

## Optional columns (features enabled when present)

| Column | Feature |
|--------|---------|
| السنوات | Year filter |
| تاريخ التصحيح المستهدف | Target date, aging (target mode) |
| تاريخ التصحيح المعدل | Modified date, aging (modified mode) |
| الشركة القابضة / الشركة التابعة | Company filters + brand logo |
| email | Legal-text email send |
| الخطة التصحيحية, ملاحظات, … | Audit plan panel, legal-text detail modal |

## Brand logos (subsidiary codes)

Values in **الشركة التابعة** or **الشركة القابضة** can match logo codes (case-sensitive): `nat`, `aum`, `saco`, `autostar`, `btc`.

Logo files live in `assets/logos/{code}.png`. Aliases are listed in `assets/logos/mapping.json`.

## Images in cells

Embedded images on a row (openpyxl anchors) appear in the legal-text detail view and PowerPoint export.

## Validation on upload

Upload runs `_load_dashboard_df()` — if required columns are missing or the sheet is wrong, the user sees an Arabic error on the upload page.

## Sample files

Do **not** commit production registers to Git. Maintainers should share a redacted `.xlsx` via internal storage or attach it to a private wiki.
