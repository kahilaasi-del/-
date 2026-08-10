-- שאילתות עזר לחקירת מסד הנתונים של חשבשבת ידנית (SSMS)
-- קריאה בלבד. הרץ על מסד הנתונים של החברה הרלוונטית.

-- 1) כל הטבלאות לפי מספר שורות (מהגדולות לקטנות)
SELECT s.name AS schema_name, t.name AS table_name,
       SUM(CASE WHEN p.index_id IN (0,1) THEN p.rows ELSE 0 END) AS row_count
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.partitions p ON t.object_id = p.object_id
GROUP BY s.name, t.name
ORDER BY row_count DESC;

-- 2) חיפוש טבלאות לפי מילת מפתח בשם הטבלה
--    שנה את 'מסמך' למונח הרצוי (חשבון / פריט / תנוע וכו')
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_NAME LIKE N'%מסמך%'
ORDER BY TABLE_NAME;

-- 3) חיפוש עמודות לפי מילת מפתח (בכל הטבלאות)
--    שימושי לאיתור עמודות סכום/תאריך/כמות/חשבון
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMN_NAME LIKE N'%סכום%'
   OR COLUMN_NAME LIKE N'%כמות%'
   OR COLUMN_NAME LIKE N'%תאריך%'
ORDER BY TABLE_NAME, ORDINAL_POSITION;

-- 4) הצצה לתוכן טבלה (החלף TableName)
-- SELECT TOP 50 * FROM dbo.TableName;

-- 5) כל העמודות של טבלה מסוימת (החלף TableName)
-- SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
-- FROM INFORMATION_SCHEMA.COLUMNS
-- WHERE TABLE_NAME = N'TableName'
-- ORDER BY ORDINAL_POSITION;
