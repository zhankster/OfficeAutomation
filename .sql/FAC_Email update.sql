SELECT * FROM FAC_ALT WHERE DCODE LIKE 'CC%'

SELECT * FROM CIPS.dbo.FAC WHERE DCODE LIKE 'BE%'

UPDATE _FAC_BILL set [Email Address] = '' where [Email Address] is null

UPDATE _FAC_BILL set [Secondary Email Address] = '' where [Secondary Email Address] is null


SELECT * FROM  _FAC_BILL  where [Secondary Email Address] <> ''
SELECT * FROM  _FAC_BILL  where [Email Address] <> ''

UPDATE _FAC_BILL set Addresses = [Email Address] + ';' + [Secondary Email Address] where [Email Address] <> ''
and [Secondary Email Address] <> ''

select * from _FAC_BILL where [Addresses] like '%,%'
select [Group], right([Addresses], 1) from _FAC_BILL
SELECT  * FROM _FAC_BILL WHERE LEN([Addresses]) > 4
select * from FAC_EMAIL where ARX = 0 and IOU = 0
select * from FAC_EMAIL where Billing = 0 and (ARX = 1 or IOU = 1)
update FAC_EMAIL set Billing = 0

delete  from FAC_EMAIL where ARX = 0 and IOU = 0

;with cte as
(
select distinct(e.FAC_CODE) code from 
FAC_EMAIL e left join FAC_ALT a
on e.FAC_CODE = a.DCODE
where a.DCODE is null
)


INSERT INTO [dbo].[FAC_ALT]
           ([DCODE]
           ,[NOTIFY_TYPE]
           ,[EMAIL]
           ,[FAX1]
           ,[PHONE1]
           ,[USER1]
           ,[IOU_EMAIL]
           ,[MNG]
           ,[IOU_NOTIFY])
select code
,''
,''
,''
,''
,''
,''
,''
,0
from cte

USE [RXBackend]
GO

INSERT INTO [dbo].[FAC_ALT]
           ([DCODE]
           ,[NOTIFY_TYPE]
           ,[EMAIL]
           ,[FAX1]
           ,[PHONE1]
           ,[USER1]
           ,[IOU_EMAIL]
           ,[MNG]
           ,[IOU_NOTIFY])
     VALUES
           (<DCODE, varchar(8),>
           ,<NOTIFY_TYPE, varchar(20),>
           ,<EMAIL, varchar(4000),>
           ,<FAX1, varchar(14),>
           ,<PHONE1, varchar(14),>
           ,<USER1, varchar(30),>
           ,<IOU_EMAIL, varchar(4000),>
           ,<MNG, varchar(12),>
           ,<IOU_NOTIFY, bit,>)
GO



