-- Prevalidation Queries --
INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL) 
SELECT '1','Number of Empty ID (Key) in Source: ', COUNT(*) FROM source A WHERE ColumnA IS NULL GROUP BY ColumnA HAVING COUNT(*) > 0;

INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL) 
SELECT '2','Number of Empty ID (Key) in Target: ', COUNT(*) FROM target A WHERE ColumnA IS NULL GROUP BY ColumnA HAVING COUNT(*) > 0;

INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL) 
SELECT '3','Duplicate ID (Key) in Source: ', ColumnA FROM source
GROUP BY ColumnA HAVING COUNT(ColumnA)>1;

INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL) 
SELECT '4','Duplicate ID (Key) in Target: ', ColumnA FROM target
GROUP BY ColumnA HAVING COUNT(ColumnA)>1;

INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL) 
SELECT '5','ID (Key) in Source not found in ID (Key) of Target: ', ColumnA FROM source A
WHERE NOT EXISTS (SELECT 1 FROM target B WHERE B.ColumnA=A.ColumnA);

INSERT INTO R_PREVALIDATION_OUTPUT_TBL (UNIQUE_ID, DESC, VAL)  
SELECT '6','ID (Key) in Target not found in ID (Key) of Source: ', ColumnA FROM target A
WHERE NOT EXISTS (SELECT 1 FROM source B WHERE B.ColumnA=A.ColumnA);