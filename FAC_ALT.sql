USE [RXBackend]
GO

/****** Object:  Table [dbo].[FAC_ALT]    Script Date: 5/4/2020 10:24:25 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[FAC_ALT](
	[DCODE] [varchar](8) NOT NULL,
	[NOTIFY_TYPE] [varchar](20) NULL,
	[EMAIL] [varchar](4000) NOT NULL,
	[FAX1] [varchar](14) NULL,
	[PHONE1] [varchar](14) NULL,
	[USER1] [varchar](30) NULL,
 CONSTRAINT [PK_FAC_ALT] PRIMARY KEY CLUSTERED 
(
	[DCODE] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO



