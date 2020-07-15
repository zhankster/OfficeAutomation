USE [RXBackend]
GO

INSERT INTO [dbo].[FAC_TRANS]
           ([TRANS_DATE]
           ,[TRANS_TYPE]
           ,[FAC_CODE]
           ,[DOCUMENTS]
           ,[EMAIL_SENT]
           ,[CREATED_BY])
     VALUES
           (GETDATE()
           ,@trans_type
           ,@fac_code
           ,@documents
           ,@email_sent
           ,@created_by
GO

select * from [dbo].[FAC_TRANS]
order by TRANS_DATE desc

SELECT 
CAST(0 as BIT) Send
,A.DCODE as Code
,F.DNAME as Facility
,ISNULL(
STUFF((SELECT ';' + C.DCODE
    FROM
    CIPS.dbo.FAC FS 
    LEFT JOIN CIPS.dbo.FAC_CHG G 
    ON FS.ID = G.FAC_ID
    LEFT JOIN CIPS.dbo.CHG C
    ON G.CHG_ID = C.ID
    WHERE A.DCODE = FS.DCODE
    FOR XML PATH('')), 1, 1, ''), '') [Accounts]
,ISNULL(
STUFF((SELECT ';' + ADDRESS
FROM
FAC_EMAIL E
    WHERE A.DCODE = E.FAC_CODE AND Billing = 1
    FOR XML PATH('')), 1, 1, ''), '') [Email]
,'' as Documents
,ISNULL(O.NOTES, 'NA')
FROM
RXBackend.dbo.FAC_ALT A 
LEFT JOIN CIPS.dbo.FAC F 
    ON A.DCODE = F.DCODE
OUTER APPLY ( 
	SELECT TOP 1 NOTES  FROM FAC_TRANS T 
	WHERE T.FAC_CODE = A.DCODE
	AND NOTES = '2020-05'
) O
ORDER BY F.DNAME

SELECT TOP 1 NOTES, *  FROM FAC_TRANS T WHERE FAC_CODE = 'AUF' AND NOTES = '2020-05'

SELECT * FROM Department D 
OUTER APPLY 
   ( 
   SELECT * FROM Employee E 
   WHERE E.DepartmentID = D.DepartmentID 
   ) A 


