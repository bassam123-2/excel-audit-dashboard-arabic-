# Database schema (Excel → MySQL)

## What is stored today

| Table | What it holds |
|-------|----------------|
| `uploads_uploadsession` | Each Excel upload (filename, path, hash, row count, time) |
| `uploads_compliancerecord` | **Every data row** from sheet `سجل الالتزام الموحد` |
| `uploads_exportlog` | PPTX / HTML export attempts |
| `uploads_emaillog` | Email send attempts |

## ComplianceRecord columns (Excel → database)

| DB field | Excel column (Arabic) | Used for |
|----------|-------------------------|----------|
| `excel_row` | Row number in file | Images, email cell lookup |
| `inherent_risk` | تصنيف المخاطر الكامنة | Filters / charts |
| `residual_risk` | تصنيف المخاطر المتبقية | Filters / aging |
| `status` | الحالة | Filters / aging |
| `target_date` | تاريخ التصحيح المستهدف | Aging |
| `year_value` | السنوات / السنة | Year filter |
| `department` | الإدارة المسؤولة | Filters |
| `legislator` | المشرع | Filters |
| `system_name` | اسم النظام | Filters |
| `authority` | الهيئة التابعة | Filters |
| `regulation` | اللائحة | Filters |
| `legal_text` | النص النظامي | Legal modal, email |
| `compliance_status` | حالة الالتزام | Filters |
| `control_category` | فئة الضوابط الرقابية | Filters |
| `task_owner` | مالك المهمة / مالك الإجراء | Detail / audit panel |
| `responsible_person` | الشخص المسؤول | Detail |
| `corrective_plan` | الخطة التصحيحية | Detail / audit |
| `modified_date` | تاريخ التصحيح المعدل | Aging (modified mode) |
| `management_notes` | ملاحظات الإدارة | Detail |
| `compliance_notes` | ملاحظات الإلتزام | Detail |
| `holding_company` | الشركة القابضة | Company filter / logo |
| `subsidiary_company` | الشركة التابعة | Company filter / logo |
| `email` | email | Send legal text by email |

## When rows are imported

On successful upload (`POST /`), Django:

1. Saves the file
2. Validates Excel structure
3. Inserts all rows into `ComplianceRecord`
4. Sets `UploadSession.row_count`

## Apply new tables

```bash
python manage.py migrate
```

## View data

- Django admin: `http://127.0.0.1:8765/admin/` (create superuser first)
- MySQL Workbench: table `uploads_compliancerecord`

```bash
python manage.py createsuperuser
```

## Not in database yet (future)

- Embedded cell images (still read from Excel file on demand)
- Full switch of dashboard APIs to read from DB instead of file (teammate task)
