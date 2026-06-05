-- Run this in MySQL Workbench or MySQL Command Line if migrate fails.
-- Fixes partial migration 0002 (duplicate row_count or broken compliancerecord table).
-- Then run: python manage.py migrate

USE excel_arabic;

DROP TABLE IF EXISTS uploads_compliancerecord;

-- Remove row_count if a previous failed migrate added it
ALTER TABLE uploads_uploadsession DROP COLUMN IF EXISTS row_count;
