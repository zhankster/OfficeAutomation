update _FAC_ALT set EMAIL = 'dekalb.hda@gmail.com;hank@dekalbal.com;zrefugee@gmail.com'

delete from CIPS.dbo.FAC where DNAME = 'DEKALB (DJ) CO JAIL (AL)'

select * into _FAC_ALT FROM FAC_ALT
select * from RPT_ACTIVITY order by CREATED desc

select * from FAC_ALT
delete from FAC_ALT where dcode is null
select * from FAC_ALT where email is null
SELECT A.DCODE as [Group Code], F.DNAME as [Facility Name], A.EMAIL as Email ,A.FAX1 as Fax, A.NOTIFY_TYPE, A.PHONE1 as Phone, A.USER1 FROM FAC_ALT A LEFT JOIN FAC F ON A.DCODE = F.DCODE  ORDER BY A.DCODE

select * from FAC where DCODE in  ('ZTF','ZT')

select * from FAC where DCODE like 'F'

UPDATE [FAC_ALT]
   SET [DCODE] = <DCODE, varchar(8),>
      ,[NOTIFY_TYPE] = <NOTIFY_TYPE, varchar(20),>
      ,[EMAIL] = <EMAIL, varchar(4000),>
      ,[FAX1] = <FAX1, varchar(14),>
      ,[PHONE1] = <PHONE1, varchar(14),>
      ,[USER1] = <USER1, varchar(30),>
 WHERE [DCODE] = 
GO

USE [RXBackend]
GO

INSERT INTO [dbo].[FAC_ALT]
           ([DCODE]
           ,[NOTIFY_TYPE]
           ,[EMAIL]
           ,[FAX1]
           ,[PHONE1]
           ,[USER1])
     VALUES
           (<DCODE, varchar(8),>
           ,<NOTIFY_TYPE, varchar(20),>
           ,<EMAIL, varchar(4000),>
           ,<FAX1, varchar(14),>
           ,<PHONE1, varchar(14),>
           ,<USER1, varchar(30),>)
GO

INSERT INTO [FAC_ALT]
                        ([DCODE]
                        ,[NOTIFY_TYPE]
                        ,[EMAIL]
                        ,[FAX1]
                        ,[PHONE1]
                        ,[USER1]) VALUES ('1F'
						,'Email'
						,'mgage@chaffeesheriff.org;tchenoweth@chaffeesheriff.org;ewiepking@chaffeesheriff.
						org;kbassett@chaffeesheriff.org;bmascarenas@chaffeesheriff.org',
						'Fax',
						'Phone',
						'User')

UPDATE[FAC_ALT]SET [NOTIFY_TYPE] = 'Email' ,[EMAIL] ='gfffh' ,[FAX1] = 'hfh' ,[PHONE1] = 'hfhf' ,[USER1] = 'hfhff' WHERE[DCODE] = 'IA'

m@h.com:h@h.com
m@h.com;h@h.com

SELECT DNAME
, F.DCODE
, ISNULL(A.EMAIL, '') as EMAIL
, ISNULL(A.NOTIFY_TYPE, '') as NOTIFY_TYPE
, ISNULL(A..PHONE1, '') as PHONE
, ISNULL(A.FAX1, '') as FAX
FROM CIPS.dbo.FAC F 
LEFT JOIN RXBackend.dbo.FAC_ALT A
	ON F.DCODE = A.DCODE
WHERE F.DCODE = 'DJ'

