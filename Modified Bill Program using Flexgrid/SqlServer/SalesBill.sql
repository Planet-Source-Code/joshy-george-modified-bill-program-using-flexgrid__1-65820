if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ItemMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ItemMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[STUD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[STUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sales]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Stock]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Stock]
GO

CREATE TABLE [dbo].[ItemMaster] (
	[SLNo] [numeric](18, 0) NULL ,
	[Itmcode] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ItmName] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Rate] [numeric](18, 0) NULL ,
	[OpStock] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[STUD] (
	[studId] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[studName] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sales] (
	[InvoiceNo] [numeric](18, 0) NULL ,
	[Invoicedate] [datetime] NULL ,
	[ItmCode] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[qty] [numeric](18, 0) NULL ,
	[Trxtype] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Total] [numeric](18, 0) NULL ,
	[GTotal] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Stock] (
	[itmcode] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[qty] [numeric](18, 0) NULL ,
	[Trxtype] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Rate] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

