SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TblLog_Imports]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[TblLog_Imports](
	[LocalIndex_imp] [int] IDENTITY(1,1) NOT NULL,
	[TableName] [varchar](255) NOT NULL,
	[SourcePath] [varchar](255) NOT NULL,
	[SourceName] [varchar](255) NOT NULL,
	[SourceDateStamp] [smalldatetime] NULL,
	[ImportStartDateTime] [smalldatetime] NULL,
	[ImportEndDateTime] [smalldatetime] NULL,
	[ImportDurationMinutes] [float] NULL,
	[Records] [int] NULL,
	[FlagNewColumnsReceived] [varchar](1) NULL,
	[Comments] [varchar](8000) NULL,
 CONSTRAINT [PK_TblLog_Imports] PRIMARY KEY CLUSTERED 
(
	[LocalIndex_imp] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TblPar_Parameters]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[TblPar_Parameters](
	[ParameterName] [varchar](255) NOT NULL,
	[ParameterString] [varchar](255) NULL,
	[ParameterDate1] [datetime] NULL,
	[ParameterDate2] [datetime] NULL,
	[ParameterInt] [int] NULL,
	[Comments] [varchar](255) NULL,
 CONSTRAINT [PK_TblPar_Parameters] PRIMARY KEY CLUSTERED 
(
	[ParameterName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
END
GO
INSERT [dbo].[TblPar_Parameters] ([ParameterName], [ParameterString], [ParameterDate1], [ParameterDate2], [ParameterInt], [Comments]) VALUES (N'Email recipients (import procedures)', N'God <yourname@domain.net>', NULL, NULL, NULL, NULL)
GO
INSERT [dbo].[TblPar_Parameters] ([ParameterName], [ParameterString], [ParameterDate1], [ParameterDate2], [ParameterInt], [Comments]) VALUES (N'Filepath_UserLoadFolderRoot', N'D:\Data\UserLoadFolders\', NULL, NULL, NULL, NULL)
GO
INSERT [dbo].[TblPar_Parameters] ([ParameterName], [ParameterString], [ParameterDate1], [ParameterDate2], [ParameterInt], [Comments]) VALUES (N'ReportingFMonth', N'10', NULL, NULL, NULL, NULL)
GO
INSERT [dbo].[TblPar_Parameters] ([ParameterName], [ParameterString], [ParameterDate1], [ParameterDate2], [ParameterInt], [Comments]) VALUES (N'ReportingFyear', N'2022/23', NULL, NULL, NULL, NULL)
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spParLoadAllUserFiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spParLoadAllUserFiles] AS' 
END
GO
ALTER PROCEDURE [dbo].[spParLoadAllUserFiles] (@usr varchar(128)) AS

-- exec spParLoadAllUserFiles @usr='INFOHUB\EC005'

/*

VERSION CONTROL

Modified    Modifyee	Modification

----------- -----------	-----------------------------------------------

13-Mar-2023	JCEH		Created procedure.

----------- -----------	-----------------------------------------------

NOTES:

*/

----------------------------------------------------------------------------
--- Declare variables
DECLARE @db					varchar(255)		SET @db = DB_NAME()
DECLARE @resulttable		TABLE(result int NULL)
--DECLARE @usr				varchar(128)		SET @usr = ''--SUSER_NAME() -- Username but not required here.

----------------------------------------------------------------------------

DELETE FROM @resulttable

EXECUTE AS LOGIN = @usr 
INSERT INTO @resulttable
SELECT HAS_DBACCESS(@db)
REVERT;

IF((select TOP 1 result from @resulttable)=0)
BEGIN
PRINT 'USER DOES NOT EXIST IN THIS DATABASE'
RETURN
END

IF((select TOP 1 result from @resulttable)=1)
BEGIN --- User has access to database

----------------------------------------------------------------------------
--- Set default parameter values if not present ...

IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineator','‡',NULL,'Delineator used to separate columns. Default = double-dagger = ‡. csv with text qualifier = CSV if using SQL 2014 or later.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineatorOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineatorOriginal',',',NULL,'Delineator originally used used to separate columns that needs replacing with something sensible. Default = comma = '',''. csv with text qualifier = CSV for RFC4180 compliant.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineator','NO',NULL,'NO = Leave file alone; SQL = convert in SQL using scalar function (slow but can handle no header); Powershell = Powershell conversion (fast but cannot handle no header or duplicate columns).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorSkipRows')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorSkipRows',NULL,0,'Default = 0. Number of rows to delete from the top of the raw file before the header row.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorChunkFiles','N',1000000,'N = Leave file alone; Y = split the file into multiple files containing up to 1,000,000 rows.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DefaultColumnWidth')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DefaultColumnWidth',NULL,-1,'Only use to force larger columns than the default (255 unless lots of columns). Set to minus 1 to allow proc to determine most appropriate width.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DestinationTable')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DestinationTable','Default',NULL,'Core destination table name. Set to "Default" or '''' in order to use the default behaviour of using the filename.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable01','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable02','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable03','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable04','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable05','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable06','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_Filepath','D:\DO NOT LOAD INTO THIS DATABASE\',NULL,'Filepath to data to load.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn01',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn02',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn03',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn04',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn05',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn06',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumnsRequired')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumnsRequired','N',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FirstRow')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FirstRow','1',NULL,'First row in the file to load. If header is in row 1 then leave as 1.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ForceTextIMEX')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ForceTextIMEX','1',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_IncludeRangeFilterDatabase')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_IncludeRangeFilterDatabase','N',NULL,'Set to Y to include the system range #FileterDatabase (might duplicate data).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable01','',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable02','''',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable03','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable04','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable05','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable06','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_RowTerminator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_RowTerminator','\n',NULL,'Delineator used to indicate a new daat row follows. Windows default is \n (=line feed + carriage return). Tab is \t and linux (and Oracle) just use line feed = char(10).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_TextQualifierOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_TextQualifierOriginal','"',NULL,'Text qualifier used when text contains delimiter. Default = double-quotes = ''"''. csv with text qualifier = ",".')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseHeaderForColumnNames')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseHeaderForColumnNames','YES',NULL,'YES for well structured file with header at top and no duplicate column names, otherwise use NO.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseOPENROWSET')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseOPENROWSET','N',NULL,'Set ParameterString to Y to use OPENROWSET instead of linked server.')
END

----------------------------------------------------------------------------

END   --- User has access to database

----------------------------------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[spSubConvertAllUserTextFiles]    Script Date: 17/03/2023 15:58:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spSubConvertAllUserTextFiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spSubConvertAllUserTextFiles] AS' 
END
GO
ALTER PROCEDURE [dbo].[spSubConvertAllUserTextFiles] (@usrI varchar(128)='',@debug varchar(1)='N') AS

-- exec spSubConvertAllUserTextFiles @usrI='OXNET\Simon.Nash',@debug='Y'

/*

VERSION CONTROL

Modified    Modifyee	Modification

----------- -----------	-----------------------------------------------

27-Oct-2020	JCEH		Created procedure.

----------- -----------	-----------------------------------------------

*/

------------------------------------------
--- Declare variables

--DECLARE @debug			varchar(1)		SET @debug = 'Y'
--DECLARE @usrI			varchar(128)	SET @usrI = SUSER_NAME()
DECLARE @ChunkFiles		char(1)
DECLARE @ChunkRows		int
DECLARE @command		varchar(8000)
DECLARE @convert		varchar(10)
DECLARE @db				varchar(255)	SET @db = (SELECT DB_NAME())
DECLARE @delimiter		varchar(10)
DECLARE @delimiterOrig	varchar(10)
DECLARE @error			varchar(1024)
DECLARE @filename		varchar(255)
DECLARE @filenameOrig	varchar(255)
DECLARE @filepath		varchar(255)
DECLARE @filepathRoot	varchar(255)	SET @filepathRoot = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName = 'Filepath_UserLoadFolderRoot')
DECLARE @folder			varchar(255)
DECLARE @header			varchar(3)
DECLARE @message		varchar(8000)
DECLARE @RowTerminator	varchar(10)
DECLARE @server			varchar(255)	SET @server = @@SERVERNAME
DECLARE @SkipRows		int
DECLARE @TextQualifier  varchar(5)
DECLARE @tableTempf		varchar(255)	SET @tableTempf	= '##TblTempFormatFile_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SUSER_NAME()+CONVERT(varchar(24),GETDATE(),113),'\',''),'/',''),':',''),' ',''),'-',''),'$',''),'.','')
DECLARE @tableTempsci	varchar(128)	SET @tableTempsci = '##TblTempImportAsSingleColumnWithIndex_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SUSER_NAME()+CONVERT(varchar(24),GETDATE(),113),'\',''),'/',''),':',''),' ',''),'-',''),'$',''),'.','')
DECLARE @usr			varchar(128)	SET @usr = ''--SUSER_NAME() -- Username but not required here.

IF(@debug='N')
BEGIN
SET NOCOUNT ON
END

---------------------------------------------------------------
--- Create temporary table with single column for contents and an index column.

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempsci))
BEGIN
SET @command = 'CREATE TABLE '+@tableTempsci+'(LocalIndex bigint IDENTITY(1,1) NOT NULL,FileContentsIncludingHeader varchar(max) NULL, CONSTRAINT[PK_'+@tableTempsci+'] PRIMARY KEY CLUSTERED (LocalIndex ASC))'
END
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

-----------------------------------------------------------
---  Create temporary table for format file

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempf))
BEGIN
SET @command = '
CREATE TABLE ['+@tableTempf+'](FileContent varchar(1024) NULL)
'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

---------------------------------------------------------------
--- Delcare cursor to loop through providers

DECLARE return_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
	 S.Username
	,ParameterString as [Folder] 
	FROM dbo.TblPar_Parameters as P
		RIGHT OUTER JOIN SystemTracking.dbo.TblSysUsers as S
			ON ISNULL(P.ParameterName,@filepathRoot+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usrI,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')) = S.Username+'_Filepath'
	WHERE S.SQLaccountDisabled = 'N'
		AND S.WindowsAccountDisabled = 'N'
		AND S.EmailAddress LIKE '%@%'
		AND S.Username NOT IN('INFOHUB-UHC2\sqlAgent','sa')
		AND ISNULL(p.ParameterString,'') <> 'D:\DO NOT LOAD INTO THIS DATABASE\'
		AND CASE WHEN @usrI = '' THEN S.Username ELSE @usrI END = S.Username
        
OPEN return_cursor 
        FETCH NEXT FROM return_cursor into @usr,@folder
WHILE @@FETCH_STATUS = 0 
        BEGIN 

IF(@debug='Y')
BEGIN
PRINT ''
PRINT @usr
PRINT @folder
END

--- Set default parameter values if not present ...

IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineator','‡',NULL,'Delineator used to separate columns. Default = double-dagger = ‡. csv with text qualifier = CSV if using SQL 2014 or later.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineatorOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineatorOriginal',',',NULL,'Delineator originally used used to separate columns that needs replacing with something sensible. Default = comma = '',''. csv with text qualifier = CSV for RFC4180 compliant.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineator','NO',NULL,'NO = Leave file alone; SQL = convert in SQL using scalar function (slow but can handle no header); Powershell = Powershell conversion (fast but cannot handle no header or duplicate columns).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorSkipRows')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorSkipRows',NULL,0,'Default = 0. Number of rows to delete from the top of the raw file before the header row.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorChunkFiles','N',1000000,'N = Leave file alone; Y = split the file into multiple files containing up to 1,000,000 rows.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DefaultColumnWidth')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DefaultColumnWidth',NULL,-1,'Only use to force larger columns than the default (255 unless lots of columns). Set to minus 1 to allow proc to determine most appropriate width.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DestinationTable')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DestinationTable','Default',NULL,'Core destination table name. Set to "Default" or '''' in order to use the default behaviour of using the filename.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable01','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable02','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable03','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable04','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable05','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable06','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_Filepath','D:\DO NOT LOAD INTO THIS DATABASE\',NULL,'Filepath to data to load.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn01',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn02',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn03',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn04',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn05',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn06',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumnsRequired')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumnsRequired','N',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FirstRow')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FirstRow','1',NULL,'First row in the file to load. If header is in row 1 then leave as 1.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ForceTextIMEX')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ForceTextIMEX','1',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_IncludeRangeFilterDatabase')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_IncludeRangeFilterDatabase','N',NULL,'Set to Y to include the system range #FileterDatabase (might duplicate data).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable01','',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable02','''',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable03','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable04','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable05','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable06','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_RowTerminator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_RowTerminator','\n',NULL,'Delineator used to indicate a new daat row follows. Windows default is \n (=line feed + carriage return). Tab is \t and linux (and Oracle) just use line feed = char(10).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_TextQualifierOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_TextQualifierOriginal','"',NULL,'Text qualifier used when text contains delimiter. Default = double-quotes = ''"''. csv with text qualifier = ",".')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseHeaderForColumnNames')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseHeaderForColumnNames','YES',NULL,'YES for well structured file with header at top and no duplicate column names, otherwise use NO.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseOPENROWSET')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseOPENROWSET','N',NULL,'Set ParameterString to Y to use OPENROWSET instead of linked server.')
END

BEGIN TRY --- RETURN

----------------------------------------------------------
--- Note: if column delimiter is set to delimiter original then file will converted to a preferred delimiter.

SET @convert = (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_ConvertColumnDelineator')

IF(@convert<>'NO')
BEGIN --- CONVERT FILE

SET @filepath		= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_Filepath')
SET @header			= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_UseHeaderForColumnNames')
SET @ChunkFiles		= (SELECT TOP 1 LEFT(ParameterString,1) FROM dbo.TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
SET @ChunkRows		= (SELECT TOP 1 ParameterInt FROM dbo.TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
SET @SkipRows		= (SELECT TOP 1 ParameterInt FROM dbo.TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorSkipRows')
SET @delimiter		= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters where ParameterName = @usr+'_ColumnDelineator')
SET @delimiterOrig	= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters where ParameterName = @usr+'_ColumnDelineatorOriginal')
SET @TextQualifier	= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters where ParameterName = @usr+'_TextQualifierOriginal')
SET @RowTerminator	= (SELECT TOP 1 ParameterString FROM dbo.TblPar_Parameters where ParameterName = @usr+'_RowTerminator')

IF(@header='NO')
BEGIN
SET @ChunkFiles='N' -- Do not chunk the files with no header as Powershell uses the first row as the header
END
IF(@delimiterOrig IS NULL)
BEGIN
SET @delimiterOrig=@delimiter
END
SET @filepath = CASE WHEN RIGHT(@filepath,1)='\' THEN @filepath ELSE @filepath+'\' END

IF(@debug='Y')
BEGIN
PRINT 'Filepath = '+@filepath
PRINT 'Number of rows to skip = '+CAST(@SkipRows as varchar(20))
PRINT 'Chunk files = '+@ChunkFiles
PRINT 'Column delimiter = '+@delimiter
PRINT 'Original delimiter = '+@delimiterOrig
PRINT 'header = '+@header
PRINT 'Text Qualifier = '+@TextQualifier
PRINT 'Row Terminator = '+@RowTerminator
END

IF(@delimiter NOT IN('‡' -- CHAR(135) double dagger
					,'§' -- CHAR(167) section-sign
					,'¬' -- CHAR(172) 
					,'±' -- CHAR(177) plus or minus
					,'¥' -- CHAR(165) yen
					,'©' -- CHAR(169) copyright
					,'¡' -- CHAR(161) spanish opening exclamation
					)
	) -- Preferred delimiters
BEGIN
SET @delimiter='‡'

IF(@debug='Y')
BEGIN
PRINT 'Delimiter will be changed from '+@delimiterOrig+' to '+@delimiter+'.'
END
END

----------------------------------------------------------
--- Obtain directory (folder) contents

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TempDir'))
BEGIN
TRUNCATE TABLE #TempDir
END

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TempDir'))
BEGIN
CREATE TABLE #TempDir(DirList varchar(8000))
END

SET @command = 'DIR /B "'+@filepath+'"'
IF(@debug='Y')
BEGIN
PRINT @command
END

INSERT INTO #TempDir
exec xp_cmdshell @command

IF EXISTS(select top 1 * from #TempDir where DirList LIKE 'Access is denied%')
BEGIN
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR ('Error: access permissions.', -- Message text.
               11, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
IF EXISTS(select top 1 * from #TempDir where DirList LIKE 'The system cannot find the file specified.%')
BEGIN
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR ('Error: folder does not exist.', -- Message text.
               11, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
IF EXISTS(select top 1 * from #TempDir where DirList NOT LIKE 'Access is denied%' AND DirList NOT LIKE 'The system cannot find the file specified.%') AND @debug='Y'
BEGIN
PRINT 'Completed successfully.'
END

----------------------------------------------------------------------------
----------------------------------------------------------------------------

IF EXISTS(SELECT 1 FROM #TempDir WHERE (DirList LIKE '%.csv' OR DirList LIKE '%.txt'))
BEGIN --- Files to load

IF(@debug='Y')
BEGIN
SELECT
 'Files to convert ...' as Comment
,DirList 
FROM #TempDir
WHERE DirList<>'DirectoryListing.txt'
	AND (DirList LIKE '%.csv' OR DirList LIKE '%.txt')
ORDER BY DirList
END

----------------------------------------------------------------------------
--- POWERSHELL

IF(@convert='Powershell' AND @header='YES')
BEGIN -- POWERSHELL

IF(@debug='Y')
BEGIN
PRINT 'EXEC Lookups.dbo.spDOS_ConvertDelimiterAllTextfiles @filepath='''+@filepath+''',@ChunkFiles='''+@ChunkFiles+''',@ChunkRows='''+CAST(@ChunkRows as varchar(20))+''',@colDelineator='''+@delimiterOrig+''',@ColDelineatorI='''+@delimiter+''',@SkipRows='''+CAST(@SkipRows as varchar(20))+''',@debug=''Y'''
END
EXEC Lookups.dbo.spDOS_ConvertDelimiterAllTextfiles @filepath=@filepath,@ChunkFiles=@ChunkFiles,@ChunkRows=@ChunkRows,@colDelineator=@delimiterOrig,@ColDelineatorI=@delimiter,@SkipRows=@SkipRows,@debug=@debug

END -- POWERSHELL

----------------------------------------------------------------------------
--- SQL

IF(@convert='SQL' /*OR @header='NO'*/)
BEGIN -- SQL CONVERSION


SET @command = 'TRUNCATE TABLE ['+@tableTempf+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

/*
Version
NumberOfColumns
FileColumnNumber	DataType	DataHandling	SQL_FieldLength		Terminator	DbColumnNumber	DbColumnName	Collation

9.0
1
1                   SQLCHAR     0               8060                "\n"		2				S1               ""
*/
SET @command = '
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''9.0'')
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''1'')
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''1                   SQLCHAR     0               8060                "'+@RowTerminator+'"    2              FileContentsIncludingHeader               ""'')
'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

IF(@debug='Y')
BEGIN
PRINT ''
PRINT '@filepath = ' + @filepath
PRINT '@tableTempf = ' + @tableTempf
PRINT '@server = ' + @server
PRINT ''
END

SET @command = 'bcp "tempdb..['+@tableTempf+']" out "'+@filepath+@tableTempf+'.fmt" /c /t "|" /S "' + @server + '" -T'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT @command
PRINT ''
PRINT ''
EXEC master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @command,no_output
END

---------------------------------------------------------------
--- Loop through files

DECLARE file_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DirList 
	FROM #TempDir
	WHERE DirList<>'DirectoryListing.txt'
		AND (DirList LIKE '%.csv' OR DirList LIKE '%.txt')
	ORDER BY DirList
        
OPEN file_cursor 
        FETCH NEXT FROM file_cursor into @filename
WHILE @@FETCH_STATUS = 0 
        BEGIN 

SET @filenameOrig = @filename

--- If length of filepath+filename > 128 then xp_cmdshell or BCP or something starts throwing a wobbly.
IF(LEN(@filepath+@filenameOrig)>128)
BEGIN -- Filename > 128 characters
SET @command = 'XCOPY /Y "'+@filepath+@filename+'" "'+@filepath+'file1.txt*"'

IF(@debug='Y')
BEGIN
PRINT @command
exec xp_cmdshell @command--,no_output
END
IF(@debug='N')
BEGIN
exec xp_cmdshell @command,no_output
END

SET @filename = 'file1.txt'

END -- Filename > 128 characters

-----------------------------------------------------------
---Insert new data from file - column delineated by commas with newline for end of row and no text delineators - i.e. standard csv.

SET @command = '
TRUNCATE TABLE ['+@tableTempsci+']
BULK INSERT ['+@tableTempsci+'] FROM "'+@filepath+@filename+'" WITH (TABLOCK,CODEPAGE=''RAW'',FIRSTROW='+CAST(@SkipRows+1 as varchar(20))+',FORMATFILE='''+@filepath+@tableTempf+'.fmt'')'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec (@command)

IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

--- Clean data
SET @command = '
UPDATE t 
SET
 FileContentsIncludingHeader = CAST(Lookups.dbo.ufnConvertTextQualifiedDelimiter(FileContentsIncludingHeader,'''+@delimiterOrig+''','''+@TextQualifier+''','''+@delimiter+''','''+@RowTerminator+''') as varchar(max)) 
FROM ['+@tabletempsci+'] as t
'
IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

--- Export cleaned data
SET @Command = 'bcp  "SELECT FileContentsIncludingHeader FROM tempdb..['+@tabletempsci+'] ORDER BY LocalIndex" queryout "'+@filepath+CASE WHEN LEFT(@filename,4)='CNV_' THEN '' ELSE 'CNV_' END+REPLACE(@filename,'.csv','.txt')+'" /c /t "£££" /S "'+@server+'" -T -C "RAW"'
IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
exec master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @Command, no_output
END

IF(LEN(@filepath++@filenameOrig)>128)
BEGIN -- Filename > 128 characters

SET @command = 'XCOPY /Y "'+@filepath+CASE WHEN LEFT(@filename,4)='CNV_' THEN '' ELSE 'CNV_' END+@filename+'" "'+@filepath+CASE WHEN LEFT(@filenameOrig,4)='CNV_' THEN '' ELSE 'CNV_' END+REPLACE(@filenameOrig,'.csv','.txt')+'*"'
IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
exec master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @Command, no_output
END

SET @command = 'DEL /Q "'+@filepath+'*file1.txt"'
IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
exec master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @Command, no_output
END

END -- Filename > 128 characters

--- Move raw file to processed folder
SET @command = 'MOVE /Y "'+@filepath+@filenameOrig+'" "'+@filepath+'Processed\'+@filenameOrig+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
exec master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @Command, no_output
END

---------------------------------------------------------------
---Close cursor 

        FETCH NEXT FROM file_cursor INTO @filename
END 
CLOSE file_cursor 
DEALLOCATE file_cursor 

---------------------------------------------------------------
--- Clear up any straggler interim files

SET @command = 'DEL /Q "'+@filepath+'*file1.txt"'
IF(@debug='Y')
BEGIN
SET @message = @command
-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
exec master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @Command, no_output
END

END -- SQL CONVERSION

----------------------------------------------------------------------------

END --- Files to load

--- Having converted the files update the delimiter ...

UPDATE P1 SET
 ParameterString = @delimiter
--,Comments=CASE WHEN @delimiterOrig=@delimiter THEN Comments ELSE ISNULL(Comments,'')+ ' '+'Delimiter changed from '+@delimiterOrig+' to '+@delimiter+'.' END
FROM TblPar_Parameters P1
WHERE P1.ParameterName = @usr+'_ColumnDelineator'
	/*AND EXISTS(select 1
			   from TblPar_Parameters P2
			   where P2.ParameterString = @filepath
				and LEFT(P1.ParameterName,9)=LEFT(P2.ParameterName,9)
			  )*/

END --- CONVERT FILE

----------------------------------------------------------------------------
----------------------------------------------------------------------------

END TRY --- RETURN

---Catch error (if any)
BEGIN CATCH --- PROVIDER
  PRINT 'Error detected'
SET @error =
	  (SELECT 
	--ISNULL(CAST(ERROR_NUMBER() as varchar(1000)),'') + ', ' + 
	--ISNULL(CAST(ERROR_SEVERITY() as varchar(1000)),'') + ', ' + 
	--ISNULL(CAST(ERROR_STATE() as varchar(1000)),'') + ', ' + 
	--ISNULL(CAST(ERROR_PROCEDURE() as varchar(1000)),'') + ', ' + 
	ISNULL(CAST(ERROR_LINE() as varchar(1000)),'') + ', ' + 
	ISNULL(CAST(ERROR_MESSAGE() as varchar(7000)),'')
	)

IF(@debug='Y')
BEGIN
PRINT @error
END

END CATCH --- RETURN

---------------------------------------------------------------
--- Delete format file

SET @command = 'DEL "'+@filepath+@tableTempf+'.fmt"'
IF(@debug='Y')
BEGIN
PRINT @command
EXEC master..xp_cmdshell @command
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @command,no_output
END

---------------------------------------------------------------
---Close cursor 

        FETCH NEXT FROM return_cursor INTO @usr,@folder
END 
CLOSE return_cursor 
DEALLOCATE return_cursor 

---------------------------------------------------------------
--- Drop temporary tables

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempf))
BEGIN
SET @command = '
DROP TABLE ['+@tableTempf+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempsci))
BEGIN
SET @command = 'DROP TABLE ['+@tableTempsci+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TempDir'))
BEGIN
IF(@debug='Y')
BEGIN
PRINT 'DROP TABLE #TempDir'
END
DROP TABLE #TempDir
END

---------------------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[spSubLoadAllUserMicrosoftOfficeFiles]    Script Date: 17/03/2023 15:58:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spSubLoadAllUserMicrosoftOfficeFiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spSubLoadAllUserMicrosoftOfficeFiles] AS' 
END
GO
ALTER PROCEDURE [dbo].[spSubLoadAllUserMicrosoftOfficeFiles] (@usrI varchar(128)='',@prefix varchar(14)='_',@tableI varchar(95)='',@suffix varchar(14)='_Import',@version varchar(5)='',@debug varchar(1)='N') AS

/*

VERSION CONTROL

Modified    Modifyee	Modification

----------- -----------	------------------------------------------------

22-Feb-2023	JCEH		Created procedure.

----------- -----------	------------------------------------------------

NOTES:

Loop 1 = User folder cursor
	-> Loop 2 = File cursor
		--> Loop 3 = Table cursor
			---> Loop 4.1 = Column cursor 1
			---> Loop 4.2 = Column cursor 2
			---> Loop 4.3 = Column cursor 3

*/

---------------------------------------------------------------

--DECLARE @debug					varchar(1)		SET @debug   = 'Y'
--DECLARE @prefix					varchar(5)		SET @prefix  = '' -- table name prefix after username and before table name.
--DECLARE @suffix					varchar(5)		SET @suffix  = '' -- table name suffix before and after version.
--DECLARE @tableI					varchar(95)		SET @tableI  = '' -- Intended table to load into. If blank then uses filename.
--DECLARE @usrI					varchar(128)	SET @usrI     = ''	 -- Username but not required here.
--DECLARE @version				varchar(5)		SET @version = ''
---Declare common variables
DECLARE @column					varchar(255)
DECLARE @columnchanges			varchar(255)	SET @columnchanges = ''
DECLARE @command				varchar(max)
DECLARE @command1				varchar(max) -- 13-Oct-2014, JCEH changed from varchar(8000) to varchar(max) as FRA had too many columns!
DECLARE @command2				varchar(max) -- 13-Oct-2014, JCEH changed from varchar(8000) to varchar(max) as FRA had too many columns!
DECLARE @commandDOS				varchar(8000)
DECLARE @comment				varchar(255)
DECLARE @datestamp				varchar(8)		SET @datestamp = cast(CONVERT(varchar(16),getdate(),112) as varchar(8))
DECLARE @datetime				datetime
DECLARE @error					varchar(255)
DECLARE @FilesLoadedList		varchar(max)	--SET @FilesLoadedList = ''
DECLARE @filemodifieddate		smalldatetime
DECLARE @filemodifieddateP		smalldatetime -- Previous modified date
DECLARE @filename				varchar(255)
DECLARE @filepath				varchar(255)
DECLARE @filepathname			varchar(510)
DECLARE @filepathRoot			varchar(255)	SET @filepathRoot = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName = 'Filepath_UserLoadFolderRoot')
DECLARE @folder					varchar(255)
DECLARE @message				varchar(max)	SET @message = ''
DECLARE @quote					varchar(1)		SET @quote = '''' /* Gets round quote problem in a way that makes code readable */
DECLARE @resulttable			TABLE(result int NULL)
DECLARE @rows1					int				SET @rows1 = -1
DECLARE @rows2					int				SET @rows2 = 0
DECLARE @table					varchar(255)
DECLARE @usr					varchar(128)	SET @usr     = ''	 -- Username but not required here.
--- Email details
IF(SELECT COUNT(*) FROM TblPar_Parameters WHERE ParameterName = 'Email recipients (import procedures)')=0
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString) VALUES('Email recipients (import procedures)','yourname@domain.net')
END
DECLARE @instancename			varchar(255)	SET @instancename = @@SERVERNAME -- SQL instance name
DECLARE @db						varchar(255)	SET @db = DB_NAME()
DECLARE @email_profilename		varchar(50)		SET @email_profilename = 'BIU' -- @profile_name
DECLARE @email_recipients		varchar(255)	SET @email_recipients  = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName = 'Email recipients (import procedures)') -- @recipients
DECLARE @email_subjectline		varchar(255)	SET @email_subjectline = 'Load all user office data files (' + @db + ' on ' + @instancename+')' -- @subject 
DECLARE @email_bodytext			varchar(8000)	SET @email_bodytext = 'Results of last load for spSubLoadAllUserMicrosoftOfficeFiles:' -- @body
DECLARE @email_query			varchar(8000)	SET @email_query = ''
DECLARE @email_AttachFilename	varchar(255)	SET @email_AttachFilename = 'FilesNotLoadedMOF.csv'
DECLARE @email_ColumnDelineator varchar(5)		SET @email_ColumnDelineator = char(9)

----------------------------------------------------------------------------
--- Setup e-mail 

--- Excel values required
DECLARE @filldowncolumns		varchar(1)		
DECLARE @filldowncolumn01		tinyint			
DECLARE @filldowncolumn02		tinyint			
DECLARE @filldowncolumn03		tinyint			
DECLARE @filldowncolumn04		tinyint			
DECLARE @filldowncolumn05		tinyint			
DECLARE @filldowncolumn06		tinyint			
DECLARE @header					varchar(10)		
DECLARE @imex					varchar(10)		
--- Sheets / named ranges / tables to import
DECLARE @RangeSheetTable01		varchar(255)	
DECLARE @RangeSheetTable02		varchar(255)	
DECLARE @RangeSheetTable03		varchar(255)	
DECLARE @RangeSheetTable04		varchar(255)	
DECLARE @RangeSheetTable05		varchar(255)	
DECLARE @RangeSheetTable06		varchar(255)	
DECLARE @RangeExclude01			varchar(255)	
DECLARE @RangeExclude02			varchar(255)	
DECLARE @RangeExclude03			varchar(255)	
DECLARE @RangeExclude04			varchar(255)	
DECLARE @RangeExclude05			varchar(255)	
DECLARE @RangeExclude06			varchar(255)	
DECLARE @RangeIncludeFilterDatabase char(1)='N'
DECLARE @UseOPENROWSET			char(1)			SET @UseOPENROWSET = 'N'
---Linked server parameters
DECLARE @RC						int
DECLARE @server					varchar(128)	-- Linked server name
DECLARE @srvproduct 			varchar(128)
DECLARE @provider				varchar(128)
DECLARE @datasrc				varchar(4000)	-- Source file to link to
DECLARE @location				varchar(4000)
DECLARE @provstr				varchar(4000)	-- Connection string (driver)
DECLARE @catalog				varchar(128)
DECLARE @JetOrACE				varchar(3)		-- Microsoft.Jet.4.0 driver or Microsoft.ACE.12.0 driver to be used.
DECLARE @asvt 					varchar(10)			-- Address space version type - 32bit or 64bit
---Linked server table schema (SP_TABLES_EX results)
DECLARE @tab_cat 				varchar(255)
DECLARE @tab_schem 				varchar(255)
DECLARE @tab_name 				varchar(255)
DECLARE @tab_type 				varchar(255)
DECLARE @tab_rem 				varchar(255)

---------------------------------------------------------------
--- Set default values
--- Work out whether ACE drivers loaded. If not use ACE - unless 64bit!
SET @asvt = (SELECT CASE WHEN CAST(SERVERPROPERTY('Edition') as varchar(1000)) LIKE '%64_bit%' THEN '64bit' ELSE '32bit' END as AddressSpaceVersionType)
IF @asvt = '64bit'
BEGIN
SET @JetOrACE = 'ACE'	-- 64bit doesn't support Jet so has to be ACE
END

IF (@asvt = '32bit')
BEGIN -- Start of 32bit check

SET @JetOrACE = 'Jet'	-- Jet driver or ACE driver if loaded

--- Create temporary table to put results in
CREATE TABLE #TblOLEDB(Provider_Name	varchar(255) NULL
					 ,Parse_Name varchar(255) NULL
					 ,Provider_Description varchar(255) NULL
					 )

--- Insert OLE DB provider strings loaded into temporary table
INSERT INTO #TblOLEDB
exec sp_enum_oledb_providers

--- Check Microsoft Database Engine ACE driver loaded
IF EXISTS(SELECT Provider_Name FROM #TblOLEDB WHERE Provider_Name = 'Microsoft.ACE.OLEDB.12.0')
BEGIN
SET @JetOrACE = 'ACE'
END

--- Clear up temporary table
DROP TABLE #TblOLEDB

END --- End of 32bit check

IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = '@JetOrAce = ' + @JetOrAce

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END

SET @filepathRoot = CASE WHEN RIGHT(@filepathRoot,1)='\' THEN @filepathRoot ELSE @filepathRoot+'\' END

---------------------------------------------------------------
--- Delcare cursor to loop through user folders / directories
DECLARE folder_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
	 S.Username
	,ParameterString as [Folder] 
	FROM dbo.TblPar_Parameters as P
		RIGHT OUTER JOIN SystemTracking.dbo.TblSysUsers as S
			ON ISNULL(P.ParameterName,@filepathRoot+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usrI,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')) = S.Username+'_Filepath'
	WHERE S.SQLaccountDisabled = 'N'
		AND S.WindowsAccountDisabled = 'N'
		AND S.EmailAddress LIKE '%@%'
		AND S.Username NOT IN('INFOHUB-UHC2\sqlAgent','sa')
		AND ISNULL(p.ParameterString,'') <> 'D:\DO NOT LOAD INTO THIS DATABASE\'
		AND CASE WHEN @usrI = '' THEN S.Username ELSE @usrI END = S.Username

OPEN folder_cursor 
        FETCH NEXT FROM folder_cursor into @usr,@folder
WHILE @@FETCH_STATUS = 0 
        BEGIN 

----------------------------------------------------------------------------

SET @filepath		  = ISNULL(@folder,@filepathRoot+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_'))
SET @filepath		  = CASE WHEN RIGHT(@filepath,1) = '\' THEN @filepath ELSE @filepath + '\' END
SET @email_recipients = ISNULL((SELECT TOP 1 EmailAddress FROM SystemTracking.dbo.TblSysUsers WHERE Username = @usr),@email_recipients)

DELETE FROM @resulttable

EXECUTE AS LOGIN = @usr 
INSERT INTO @resulttable
SELECT HAS_DBACCESS(@db)
REVERT;

IF((select TOP 1 result from @resulttable)=0)
BEGIN
SET @filepath = 'USER DOES NOT HAVE ACCESS TO DATABASE '+@db
PRINT '@filepath = '+@filepath
END 

----------------------------------------------------------------------------

IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineator','‡',NULL,'Delineator used to separate columns. Default = double-dagger = ‡. csv with text qualifier = CSV if using SQL 2014 or later.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineatorOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineatorOriginal',',',NULL,'Delineator originally used used to separate columns that needs replacing with something sensible. Default = comma = '',''. csv with text qualifier = CSV for RFC4180 compliant.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineator','NO',NULL,'NO = Leave file alone; SQL = convert in SQL using scalar function (slow but can handle no header); Powershell = Powershell conversion (fast but cannot handle no header or duplicate columns).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorSkipRows')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorSkipRows',NULL,0,'Default = 0. Number of rows to delete from the top of the raw file before the header row.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorChunkFiles','N',1000000,'N = Leave file alone; Y = split the file into multiple files containing up to 1,000,000 rows.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DefaultColumnWidth')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DefaultColumnWidth',NULL,-1,'Only use to force larger columns than the default (255 unless lots of columns). Set to minus 1 to allow proc to determine most appropriate width.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DestinationTable')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DestinationTable','Default',NULL,'Core destination table name. Set to "Default" or '''' in order to use the default behaviour of using the filename.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable01','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable02','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable03','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable04','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable05','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable06','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_Filepath','D:\DO NOT LOAD INTO THIS DATABASE\',NULL,'Filepath to data to load.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn01',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn02',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn03',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn04',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn05',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn06',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumnsRequired')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumnsRequired','N',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FirstRow')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FirstRow','1',NULL,'First row in the file to load. If header is in row 1 then leave as 1.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ForceTextIMEX')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ForceTextIMEX','1',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_IncludeRangeFilterDatabase')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_IncludeRangeFilterDatabase','N',NULL,'Set to Y to include the system range #FileterDatabase (might duplicate data).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable01','',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable02','''',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable03','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable04','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable05','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable06','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_RowTerminator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_RowTerminator','\n',NULL,'Delineator used to indicate a new daat row follows. Windows default is \n (=line feed + carriage return). Tab is \t and linux (and Oracle) just use line feed = char(10).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_TextQualifierOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_TextQualifierOriginal','"',NULL,'Text qualifier used when text contains delimiter. Default = double-quotes = ''"''. csv with text qualifier = ",".')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseHeaderForColumnNames')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseHeaderForColumnNames','YES',NULL,'YES for well structured file with header at top and no duplicate column names, otherwise use NO.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseOPENROWSET')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseOPENROWSET','N',NULL,'Set ParameterString to Y to use OPENROWSET instead of linked server.')
END

----------------------------------------------------------------------------
--- Clear up any nasties from testing!

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_tables_ex'))
BEGIN
DROP TABLE #Tbl_sp_tables_ex
END

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_columns_ex'))
BEGIN
DROP TABLE #Tbl_sp_columns_ex
END

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblCols'))
BEGIN
DROP TABLE #TblCols
END

----------------------------------------------------------------------------
---Work out contents of target directory

SET @commandDOS = 'IF EXIST "' + @filepath + '" DIR "' + @filepath + '" > "' + @filepath + 'DirectoryListing.txt"'
IF(@debug='Y')
BEGIN
PRINT @commandDOS
EXEC master..xp_cmdshell @commandDOS--, no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS, no_output
END

-----------------------------------------------------------
---Create table to put results in if it doesn't exist

IF NOT EXISTS (SELECT [name] FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblTempDirList'))
BEGIN
CREATE TABLE #TblTempDirList([DirList] [varchar](8000) NULL)
END

TRUNCATE TABLE #TblTempDirList

-----------------------------------------------------------
---Set variable values - next 3 lines will need editing to specific use.
SET @table		  = '#TblTempDirList'			/* Table to insert data into */
SET @filename	  = 'DirectoryListing.txt'		/* Filename with data to insert into table*/
SET @quote		  = ''''						/* Gets round quote problem in a way that makes code more easily readable */
SET @filepathname = @filepath + @filename		/* Combine path and name of file with data */

-----------------------------------------------------------
---Clear out old data

SET @command = 'TRUNCATE TABLE ['+@table+']'
exec (@command)

-----------------------------------------------------------
---Insert directory listing  from file 

IF EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
SET @command = 'BULK INSERT ['+@table+'] FROM "'+@filepathname+'" WITH (FIELDTERMINATOR = '+@quote+'££££'+@quote+', ROWTERMINATOR = '+@quote+'\n'+@quote+', FIRSTROW = 1)'
IF (@debug='Y')
BEGIN
PRINT(@command) 
END
exec (@command)
END

----------------------------------------------------------------------------
---Cursor to run through each file in turn

IF(SELECT
	 COUNT(*) as Files 
	FROM #TblTempDirList 
	WHERE CASE
			WHEN DirList LIKE '%.xls%' THEN 1
			WHEN DirList LIKE '%.mdb' THEN 1
			WHEN DirList LIKE '%.accdb' THEN 1
			ELSE 0
			END = 1
	) > 0

BEGIN --- Files to load

SET @email_bodytext = @email_bodytext + '

Files in load directory = ' + ISNULL((SELECT TOP 1
								 LEFT(DIR.FileNames,LEN(DIR.FileNames)-1) + '.'
								FROM (SELECT
									 (select
									 ISNULL(RTRIM(SUBSTRING(sc1.DirList,37,255)),'') + ','
									from #TblTempDirList sc1
									WHERE CASE
											WHEN DirList LIKE '%.xls%' THEN 1
											WHEN DirList LIKE '%.mdb' THEN 1
											WHEN DirList LIKE '%.accdb' THEN 1
											ELSE 0
											END = 1
									ORDER BY DirList
									FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS FileNames
									) as DIR
								),'') + '
'

--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'file_cursor')
BEGIN
CLOSE file_cursor
DEALLOCATE file_cursor
END

IF(@debug='Y')
BEGIN
SELECT
 'Files to load ...' as Comment
,RTRIM(SUBSTRING(DirList,37,255)) as [FileName] 
,CAST(LEFT([DirList],20) as datetime) as FileModificationDateTime
--/*American date*/,CAST(SUBSTRING([DirList],4,2) + '/' + LEFT([DirList],2) + '/' + SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
--/*British date*/,CAST(LEFT([DirList],2) + '/' + SUBSTRING([DirList],4,2) + '/' + SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
,ISNULL(iLog.FileModificationDateTime,'05-Jul-1948') as PreviousFileModificationDateTime
FROM #TblTempDirList DL
	LEFT OUTER JOIN (SELECT
					 SourceName as [Filename]
					,ImportStartDateTime
					,SourceDateStamp as FileModificationDateTime
					,ROW_NUMBER() OVER(PARTITION BY SourceName ORDER BY SourceName,ImportStartDateTime DESC) as Seq
					FROM dbo.TblLog_Imports
					) as iLog
		ON RTRIM(SUBSTRING(DL.DirList,40,255)) = iLog.[Filename] AND iLog.Seq = 1
WHERE CASE
		WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.xls[,b,m,x]' THEN 1
		WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.mdb' THEN 1
		WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.accdb' THEN 1
		ELSE 0 
		END = 1
	AND NOT(CAST(LEFT([DirList],20) as datetime) BETWEEN DATEADD(minute,-1,ISNULL(iLog.FileModificationDateTime,'01-Jan-1900 00:01')) AND DATEADD(minute,+1,ISNULL(FileModificationDateTime,'01-Jan-1900'))) --09-Nov-2020, JCEH: Changed cursor select from ISNULL(XXX,01-Jan-1900) to ISNULL(XXX,01-Jan-1900 00:01) to ahndle smalldatetime conversion error - implicit conversion somewhere.
	--AND RTRIM(SUBSTRING(DirList,37,255)) <> 'GRIDALL'
ORDER BY RTRIM(SUBSTRING(DirList,37,255))
END

SET @FilesLoadedList = NULL

DECLARE file_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT RTRIM(SUBSTRING(DirList,37,255)) as [FileName] 
	,CAST(LEFT([DirList],20) as datetime) as FileModificationDateTime
	--/*American date*/,CAST(SUBSTRING([DirList],4,2) + '/' + LEFT([DirList],2) + '/' + SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
	--/*British date*/,CAST(LEFT([DirList],2) + '/' + SUBSTRING([DirList],4,2) + '/' + SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
	,ISNULL(iLog.FileModificationDateTime,'05-Jul-1948') as PreviousFileModificationDateTime
	FROM #TblTempDirList DL
		LEFT OUTER JOIN (SELECT
						 SourceName as [Filename]
						,ImportStartDateTime
						,SourceDateStamp as FileModificationDateTime
						,ROW_NUMBER() OVER(PARTITION BY SourceName ORDER BY SourceName,ImportStartDateTime DESC) as Seq
						FROM dbo.TblLog_Imports
						) as iLog
			ON RTRIM(SUBSTRING(DL.DirList,40,255)) = iLog.[Filename] AND iLog.Seq = 1
	WHERE CASE
			WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.xls[,b,m,x]' THEN 1
			WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.mdb' THEN 1
			WHEN RTRIM(SUBSTRING(DirList,37,255)) LIKE '%.accdb' THEN 1
			ELSE 0 
			END = 1
		--AND NOT(CAST(LEFT([DirList],20) as datetime) BETWEEN DATEADD(minute,-1,ISNULL(iLog.FileModificationDateTime,'01-Jan-1900 00:01')) AND DATEADD(minute,+1,ISNULL(FileModificationDateTime,'01-Jan-1900'))) --09-Nov-2020, JCEH: Changed cursor select from ISNULL(XXX,01-Jan-1900) to ISNULL(XXX,01-Jan-1900 00:01) to ahndle smalldatetime conversion error - implicit conversion somewhere.
		--AND RTRIM(SUBSTRING(DirList,37,255)) <> 'GRIDALL'
	ORDER BY RTRIM(SUBSTRING(DirList,37,255))
	
OPEN file_cursor 
	FETCH NEXT FROM file_cursor INTO @filename,@filemodifieddate,@filemodifieddateP
WHILE @@FETCH_STATUS = 0
	BEGIN --- file cursor

---------------------------------------------------------------

SET @command2 = ''
SET @datetime = (SELECT CONVERT(smalldatetime,CONVERT(varchar(20),GETDATE(),100),100))
SET @FilesLoadedList = ISNULL(@FilesLoadedList+'; ','')+@filename
SET @table = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName=@usr+'_DestinationTable')
SET @table = CASE WHEN @tableI = '' AND @table IN('','Default') THEN REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.xlsb',''),'.xlsm',''),'.xlsx',''),'.xls','') WHEN @tableI = '' AND @table NOT IN('','Default') THEN @table ELSE @tableI END

PRINT ''
PRINT 'Filename: '+ISNULL(@filename,'')
PRINT 'Username: '+ISNULL(@usr,'')
PRINT 'Table prefix: '+ISNULL(@prefix,'')
PRINT 'Table name (core): '+ISNULL(@table,'')
PRINT 'Table suffix: '+ISNULL(@suffix,'')
PRINT 'Table name version: '+ISNULL(@version,'')

--- Excel values required
SET @filldowncolumns   = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_FillDownColumnsRequired') IF(@debug = 'Y') BEGIN PRINT 'Fill down requierd? ' + @filldowncolumns END -- Excel only: are there columns where the value on the previous row needs filling down? Y = yes; N = no.
SET @filldowncolumn01  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn01') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn01 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @filldowncolumn02  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn02') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn02 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @filldowncolumn03  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn03') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn03 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @filldowncolumn04  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn04') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn04 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @filldowncolumn05  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn05') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn05 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @filldowncolumn06  = (select TOP 1 ParameterInt from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn06') IF(@debug = 'Y') BEGIN PRINT 'Fill down column ' + CAST(@filldowncolumn06 as varchar(2)) + ' (0 is default = no fill down)' END -- Excel only: column number of Excel table to fill down (0 is default = none)
SET @header			   = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_UseHeaderForColumnNames') IF(@debug = 'Y') BEGIN PRINT 'Use header for columns? ' + @header END -- Excel only: Table header to be imported, YES or NO.
SET @imex			   = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ForceTextIMEX') IF(@debug = 'Y') BEGIN PRINT 'Force text (IMEX=1)? ' + @imex END -- Excel only: imex = 1 theoretically treats everything as text. 
SET @RangeExclude01	   = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable01') IF(@debug = 'Y') BEGIN PRINT 'Exclude 1 ' + @RangeExclude01 END
SET @RangeExclude02    = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable02') IF(@debug = 'Y') BEGIN PRINT 'Exclude 2 ' + @RangeExclude02 END
SET @RangeExclude03    = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable03') IF(@debug = 'Y') BEGIN PRINT 'Exclude 3 ' + @RangeExclude03 END
SET @RangeExclude04    = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable04') IF(@debug = 'Y') BEGIN PRINT 'Exclude 4 ' + @RangeExclude04 END
SET @RangeExclude05    = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable05') IF(@debug = 'Y') BEGIN PRINT 'Exclude 5 ' + @RangeExclude05 END
SET @RangeExclude06    = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable06') IF(@debug = 'Y') BEGIN PRINT 'Exclude 6 ' + @RangeExclude06 END
SET @RangeIncludeFilterDatabase = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_IncludeRangeFilterDatabase') IF(@debug = 'Y') BEGIN PRINT 'Include Filter Database ' + @RangeExclude06 END
SET @RangeSheetTable01 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable01') IF(@debug = 'Y') BEGIN PRINT 'Import 1 ' + @RangeSheetTable01 END
SET @RangeSheetTable02 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable02') IF(@debug = 'Y') BEGIN PRINT 'Import 2 ' + @RangeSheetTable02 END
SET @RangeSheetTable03 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable03') IF(@debug = 'Y') BEGIN PRINT 'Import 3 ' + @RangeSheetTable03 END
SET @RangeSheetTable04 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable04') IF(@debug = 'Y') BEGIN PRINT 'Import 4 ' + @RangeSheetTable04 END
SET @RangeSheetTable05 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable05') IF(@debug = 'Y') BEGIN PRINT 'Import 5 ' + @RangeSheetTable05 END
SET @RangeSheetTable06 = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable06') IF(@debug = 'Y') BEGIN PRINT 'Import 6 ' + @RangeSheetTable06 END
SET @UseOPENROWSET     = (select TOP 1 ParameterString from TblPar_Parameters where ParameterName = @usr+'_UseOPENROWSET') IF(@debug = 'Y') BEGIN PRINT 'Use OPENROWSET? ' + @UseOPENROWSET END

SET @email_bodytext = @email_bodytext + '

----------------------------------------------------------------------------
Loading file: ' + @filename + '
----------------------------------------------------------------------------

'

----------------------------------------------------------------------------
---Setup file as linked server

SET @server = REPLACE(REPLACE(@filename,'.','_'),'''','')
SET @srvproduct = CASE 
					WHEN @filename LIKE '%.xls' OR @filename LIKE '%.xls[x,m,b]'
						THEN 'Excel'
					WHEN @filename LIKE '%.mdb' OR @filename LIKE '%.accdb'
						THEN 'Access'
					END
SET @provider = CASE 
					WHEN @JetOrACE = 'Jet'
						THEN 'Microsoft.Jet.OLEDB.4.0' 
					ELSE 'Microsoft.ACE.OLEDB.12.0'
					END
SET @datasrc = @filepath + @filename
SET @provstr = CASE 
				WHEN @filename LIKE '%.xls'
					THEN 'Excel 8.0'+CASE WHEN @UseOPENROWSET='N' THEN ';HDR='+@header+';IMEX='+@imex ELSE '' END--Excel
				WHEN @filename LIKE '%.xlsx'
					THEN 'Excel 12.0 xml'+CASE WHEN @UseOPENROWSET='N' THEN ';HDR='+@header+';IMEX='+@imex ELSE '' END--Excel
				WHEN @filename LIKE '%.xlsm'
					THEN 'Excel 12.0 Macro'+CASE WHEN @UseOPENROWSET='N' THEN ';HDR='+@header+';IMEX='+@imex ELSE '' END--Excel
				WHEN @filename LIKE '%.xlsb'
					THEN  'Excel 12.0'+CASE WHEN @UseOPENROWSET='N' THEN ';HDR='+@header+';IMEX='+@imex ELSE '' END--Excel
				ELSE ''
				END
				
SET @command = 
'IF (select srvname from master.dbo.sysservers where srvname = '+@quote+@server+@quote + ') = '+@quote+@server+@quote + '
BEGIN 
PRINT ''Drop linked server if it already exists''
EXEC sp_dropserver '+@quote+@server+@quote + '
END
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

IF(@UseOPENROWSET='N')
BEGIN
PRINT ''
PRINT ''
PRINT ''
PRINT 'Create linked server (to Excel database).'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = 'exec [master].[dbo].[sp_addlinkedserver] '''+@server+''',''' + ISNULL(@srvproduct,'') + ''',''' + ISNULL(@provider,'') +''',''' + ISNULL(@datasrc,'') + ''',' + ISNULL(''+@location+'','NULL') + ',''' + ISNULL(@provstr,'') + ''',' + ISNULL(''+@catalog+'','NULL')

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC @RC = [master].[dbo].[sp_addlinkedserver] @server, @srvproduct, @provider, @datasrc, @location, @provstr, @catalog
END

----------------------------------------------------------------------------
---Create temporary table and place file contents table listing into it

PRINT ''
PRINT ''
PRINT ''
PRINT 'Create temporary table to place table listing of linked Excel database into.'
CREATE TABLE #Tbl_sp_tables_ex([Cat] [varchar](255) NULL,
							   [Schem] [varchar](255) NULL,
							   [Name] [varchar](255) NULL,
							   [Type] [varchar](255) NULL,
							   [Remarks] [varchar](255) NULL
							   )

IF(@UseOPENROWSET='N')
BEGIN
PRINT ''
PRINT ''
PRINT ''
PRINT 'Insert table names into temporary table.'
SET @command = 'EXECUTE SP_TABLES_EX '+@quote+@server+@quote
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
INSERT INTO #Tbl_sp_tables_ex
exec(@command)
END

IF(@UseOPENROWSET='Y')
BEGIN
INSERT INTO #Tbl_sp_tables_ex([Name])
SELECT RangeTable
	FROM (SELECT @RangeSheetTable01 as RangeTable UNION ALL
		  SELECT @RangeSheetTable02 as RangeTable UNION ALL
		  SELECT @RangeSheetTable03 as RangeTable UNION ALL
		  SELECT @RangeSheetTable04 as RangeTable UNION ALL
		  SELECT @RangeSheetTable05 as RangeTable UNION ALL
		  SELECT @RangeSheetTable06 as RangeTable
		  ) as t
	WHERE t.RangeTable NOT IN('','Enter named range or sheet or table name')
END

SET @email_bodytext = @email_bodytext + '

Tables (sheets / ranges) available = ' + ISNULL((SELECT TOP 1
											 LEFT(L.TableNames,LEN(L.TableNames)-1) + '.'
											FROM (SELECT
												  (select ISNULL(T.[name],'') + ',' from #Tbl_sp_tables_ex T ORDER BY [name] FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS TableNames
												 ) as L
											),'') + '
'

IF(@debug='Y')
BEGIN
SELECT * FROM #Tbl_sp_tables_ex
END

PRINT ''
PRINT ''
PRINT ''
PRINT 'Create temporary table to place column listing of linked Excel database into.'
CREATE TABLE #Tbl_sp_columns_ex ([TABLE_CAT] varchar(10) NULL
								,[TABLE_SCHEM] varchar(10) NULL
								,[TABLE_NAME] varchar(255) NULL
								,[COLUMN_NAME] varchar(255) NULL
								,[DATA_TYPE] int NULL
								,[TYPE_NAME] varchar(50) NULL
								,[COLUMN_SIZE] int NULL
								,[BUFFER_LENGTH] int NULL
								,[DECIMAL_DIGITS] int NULL
								,[NUM_PREC_RADIX] int NULL
								,[NULLABLE] int NULL
								,[REMARKS] varchar(50) NULL
								,[COLUMN_DEF] varchar(50) NULL
								,[SQL_DATA_TYPE] int NULL
								,[SQL_DATETIME_SUB] varchar(50) NULL
								,[CHAR_OCTET_LENGTH] int NULL
								,[ORDINAL_POSITION] int NULL
								,[IS_NULLABLE] varchar(10) NULL
								,[SS_DATA_TYPE] int NULL
								)

IF(@UseOPENROWSET='N')
BEGIN
PRINT ''
PRINT ''
PRINT ''
PRINT 'Insert column names into temporary table.'
SET @command = 'EXECUTE SP_COLUMNS_EX '+@quote+@server+@quote
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
INSERT INTO #Tbl_sp_columns_ex
exec(@command)
END

SET @email_bodytext = @email_bodytext + '

Searching for tables (sheets / ranges) with names beginning with = "' + ISNULL(@RangeSheetTable01,'') + '"' +
CASE
	WHEN @RangeSheetTable02 = 'Enter named range or sheet or table name'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeSheetTable02,'') + '"'
	END + 
CASE
	WHEN @RangeSheetTable03 = 'Enter named range or sheet or table name'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeSheetTable03,'') + '"'
	END + 
CASE
	WHEN @RangeSheetTable04 = 'Enter named range or sheet or table name'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeSheetTable04,'') + '"'
	END + 
CASE
	WHEN @RangeSheetTable05 = 'Enter named range or sheet or table name'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeSheetTable05,'') + '"'
	END + 
CASE
	WHEN @RangeSheetTable06 = 'Enter named range or sheet or table name'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeSheetTable06,'') + '"'
	END + 
CASE
	WHEN @RangeIncludeFilterDatabase = 'N'
		THEN 'exclude ...#FilterDatabase range '
	ELSE ' include ...#FilterDatabase range '
	END + 
CASE
	WHEN @RangeExclude01 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' excluding any names beginning with "' + ISNULL(@RangeExclude01,'') + '"'
	END + 
CASE
	WHEN @RangeExclude02 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeExclude02,'') + '"'
	END +
CASE
	WHEN @RangeExclude03 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeExclude03,'') + '"'
	END +
CASE
	WHEN @RangeExclude04 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeExclude04,'') + '"'
	END + 
CASE
	WHEN @RangeExclude05 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeExclude05,'') + '"'
	END +
CASE
	WHEN @RangeExclude06 = 'Enter named range or sheet or table name to exclude here'
		THEN ''
	ELSE ' or "' + ISNULL(@RangeExclude06,'') + '"'
	END + '".'

SET @email_bodytext = @email_bodytext + '

Tables (sheets / ranges) to load = ' + ISNULL((SELECT TOP 1
											 LEFT(L.TableNames,LEN(L.TableNames)-1) + '.'
											FROM (SELECT
												 (select
												 ISNULL(T.[name],'') + ','
												from #Tbl_sp_tables_ex T
												WHERE Type = 'TABLE'
													AND [Name] NOT LIKE '%Print_Titles'
													AND NOT([Name] LIKE '%_FilterDatabase%' AND @RangeIncludeFilterDatabase='N') -- 20210623, JCEH added @RangeIncludeFilterDatabase parameter.
													--AND [Name] NOT LIKE '%xlnm#%'
													AND (CASE
														WHEN [Name] LIKE @RangeSheetTable01+'%' THEN 1
														WHEN [Name] LIKE @RangeSheetTable02+'%' THEN 1
														WHEN [Name] LIKE @RangeSheetTable03+'%' THEN 1
														WHEN [Name] LIKE @RangeSheetTable04+'%' THEN 1
														WHEN [Name] LIKE @RangeSheetTable05+'%' THEN 1
														WHEN [Name] LIKE @RangeSheetTable06+'%' THEN 1
														ELSE 0
														END = 1
														)
													AND (CASE
														WHEN [Name] LIKE @RangeExclude01+'%' THEN 1
														WHEN [Name] LIKE @RangeExclude02+'%' THEN 1
														WHEN [Name] LIKE @RangeExclude03+'%' THEN 1
														WHEN [Name] LIKE @RangeExclude04+'%' THEN 1
														WHEN [Name] LIKE @RangeExclude05+'%' THEN 1
														WHEN [Name] LIKE @RangeExclude06+'%' THEN 1
														ELSE 0
														END = 0
														)
												ORDER BY [name]
												FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS TableNames
												) as L
											),'') + '
'

--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'table_cursor')
BEGIN
CLOSE table_cursor
DEALLOCATE table_cursor
END
/*
--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/
---Open table cursor
DECLARE table_cursor CURSOR FOR 
	SELECT [Cat],[Schem],[Name],[Type],[Remarks]
	FROM #Tbl_sp_tables_ex 
	WHERE Type = 'TABLE'
		AND [Name] NOT LIKE '%Print_Titles'
		AND NOT([Name] LIKE '%_FilterDatabase%' AND @RangeIncludeFilterDatabase='N') -- 20210623, JCEH added @RangeIncludeFilterDatabase parameter.
		AND [Name] NOT LIKE '%$Print_Area'
		--AND [Name] NOT LIKE '%xlnm#%'
		AND (CASE
			WHEN [Name] LIKE @RangeSheetTable01+'%' THEN 1
			WHEN [Name] LIKE @RangeSheetTable02+'%' THEN 1
			WHEN [Name] LIKE @RangeSheetTable03+'%' THEN 1
			WHEN [Name] LIKE @RangeSheetTable04+'%' THEN 1
			WHEN [Name] LIKE @RangeSheetTable05+'%' THEN 1
			WHEN [Name] LIKE @RangeSheetTable06+'%' THEN 1
			ELSE 0
			END = 1
			)
		AND (CASE
			WHEN [Name] LIKE @RangeExclude01+'%' THEN 1
			WHEN [Name] LIKE @RangeExclude02+'%' THEN 1
			WHEN [Name] LIKE @RangeExclude03+'%' THEN 1
			WHEN [Name] LIKE @RangeExclude04+'%' THEN 1
			WHEN [Name] LIKE @RangeExclude05+'%' THEN 1
			WHEN [Name] LIKE @RangeExclude06+'%' THEN 1
			ELSE 0
			END = 0
			)

OPEN table_cursor
	FETCH NEXT FROM table_cursor into @tab_cat,@tab_schem,@tab_name,@tab_type,@tab_rem
WHILE @@FETCH_STATUS = 0
	BEGIN -- Table cursor
	
----------------------------

SET @command = ''
SET @command1 = ''
SET @datetime = (SELECT CONVERT(smalldatetime,CONVERT(varchar(20),GETDATE(),100),100))

----------------------------
---Update log table

PRINT ''
PRINT ''
PRINT ''
PRINT 'Update log table for latest table import.'
PRINT @filename
PRINT @tab_name

INSERT INTO [dbo].[TblLog_Imports] ([TableName],[SourcePath],[SourceName],SourceDateStamp,[ImportStartDateTime]) VALUES(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')',@filepath,@filename,@filemodifieddate,@datetime)

----------------------------------------------------------------------------
---Beginning of try and catch statement

BEGIN TRY -- Import table


----------------------------
---Drop table if it exists

PRINT ''
PRINT ''
PRINT ''
PRINT 'If temporary table exists then drop.'
PRINT @filename
PRINT @tab_name

SET @command = '
IF EXISTS (SELECT [name] from sys.objects WHERE [type] = ''U'' AND [name] = '''+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp'')
BEGIN
DROP TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp]
END'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

-----------------------------
--- Work out columns present so you don't import nulls ..

IF(@UseOPENROWSET='N')
BEGIN
SET @command1 = (SELECT TOP 1
				 LEFT(cols.ColumnNames,LEN(cols.ColumnNames)-3)
				FROM (SELECT
					   sc2.TABLE_NAME
					,(select
					 ISNULL('[' + sc1.COLUMN_NAME + ']','') + ' IS NOT NULL OR '
					from #Tbl_sp_columns_ex sc1
					where sc1.TABLE_NAME = sc2.TABLE_NAME
						AND sc1.COLUMN_NAME <> 'RowNo'
					order by sc1.ORDINAL_POSITION
					FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS ColumnNames
						FROM #Tbl_sp_columns_ex sc2
						group by sc2.TABLE_NAME
						) as cols
				WHERE cols.TABLE_NAME = @tab_name
				)
END

-----------------------------
---Import data into table
PRINT ''
PRINT ''
PRINT ''
PRINT 'Insert data into new table.'
PRINT @filename
PRINT @tab_name

IF(@UseOPENROWSET='N')
BEGIN
SET @command = '
SELECT ' 
+@quote+@filename+@quote+' as Filename 
,'+@quote+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+@quote+' as TableName
,CAST(''' + CONVERT(varchar(20),@DateTime,100) + ''' as smalldatetime) as ImportDate
,* 
INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp] 
FROM ['+@server+']...['+@tab_name+']
WHERE '+@command1+'
'
END

IF(@UseOPENROWSET='Y')
BEGIN
SET @command = '
SELECT ' 
+@quote+@filename+@quote+' as Filename 
,'+@quote+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+@quote+' as TableName
,CAST(''' + CONVERT(varchar(20),@DateTime,100) + ''' as smalldatetime) as ImportDate
,* 
INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp] 
FROM OPENROWSET('''+@provider+''',
    '''+@provstr+';Database='+@datasrc+';HDR='+@header+';IMEX='+@imex+''',
    ''SELECT * FROM ['+@tab_name+']'' --
	)
'
END

IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

SET @rows1 = @@ROWCOUNT

SET @email_bodytext = @email_bodytext + '

' + CAST(@rows1 as varchar(10)) + ' rows in import table.
'

SET @command = '
SELECT *
FROM ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp]
' 
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

UPDATE IMP
SET
 Comments = CAST(@rows1 as varchar(10)) + ' rows put in temporary table.'
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

IF (@rows1 > 0)

BEGIN -- Import table NULL column check

-----------------------------
--- Count NULLs in first 10 columns and e-mail results to check table format hasn't changed and columns need filling down.

IF EXISTS(select [name] from sys.objects where [name] = 'VwTemp_NULLsInImportTable')
BEGIN
DROP VIEW VwTemp_NULLsInImportTable
END

IF NOT EXISTS(select [name] from sys.objects where [name] = 'VwTemp_NULLsInImportTable')
BEGIN
SET @command1 = (SELECT TOP 1
				 'CREATE VIEW VwTemp_NULLsInImportTable AS
				  SELECT CAST(COUNT(*) as varchar(10)) + '' records of which: '''
				 + cols.ColumnNames + ' as NULLcount
				 FROM ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp]' 
				FROM (SELECT
					   OBJECT_NAME(sc2.object_id) as ObjectName
					,(select top 10
					 ' + CAST(SUM(CASE WHEN ' + ISNULL('['+sc1.name+']','') + ' IS NULL THEN 1 ELSE 0 END) as varchar(10)) + ' + ''' NULLs in column = ' + ISNULL('['+sc1.name+']','') + '. '''
					from sys.columns sc1
					where OBJECT_NAME(sc1.object_id) = OBJECT_NAME(sc2.object_id)
						and sc1.[name] not in('Filename','TableName','ImportDate')
					order by sc1.column_id
					FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS ColumnNames
						FROM sys.columns sc2
						group by OBJECT_NAME(sc2.object_id)
						) as cols
				WHERE ObjectName = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp'
				)
IF(@debug='Y')
BEGIN
PRINT @command1
END
exec(@command1)
END

IF NOT EXISTS(select [name] from sys.objects where [name] = 'VwTemp_NULLsInImportTable')
BEGIN
SET @rows1 = -1
END

SET @email_bodytext = @email_bodytext + '

First 10 columns of the temporary import table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp checked for NULLs. If lots of NULLs in the column does the column need filling down where blank? 

					  '	+ CASE WHEN @rows1<0 THEN 'VwTemp_NULLsInImportTable does not exist!' ELSE (SELECT * FROM VwTemp_NULLsInImportTable) END + '
					  
					  ' + CASE WHEN @filldowncolumns = 'N' THEN 'Currently no columns set to fill down.' ELSE '' END
					    + CASE WHEN @filldowncolumns = 'Y' THEN 'Currently the following columns are set to fill down:' ELSE '' END + '
					  '
					   	+ CASE WHEN @filldowncolumn01 > 0 THEN 'Column number ' + CAST(@filldowncolumn01 as varchar(5)) + '.' ELSE '' END + '
					  '	+ CASE WHEN @filldowncolumn02 > 0 THEN 'Column number ' + CAST(@filldowncolumn02 as varchar(5)) + '.' ELSE '' END + '
					  '	+ CASE WHEN @filldowncolumn03 > 0 THEN 'Column number ' + CAST(@filldowncolumn03 as varchar(5)) + '.' ELSE '' END + '
					  '	+ CASE WHEN @filldowncolumn04 > 0 THEN 'Column number ' + CAST(@filldowncolumn04 as varchar(5)) + '.' ELSE '' END + '
					  '	+ CASE WHEN @filldowncolumn05 > 0 THEN 'Column number ' + CAST(@filldowncolumn05 as varchar(5)) + '.' ELSE '' END + '
					  '	+ CASE WHEN @filldowncolumn06 > 0 THEN 'Column number ' + CAST(@filldowncolumn06 as varchar(5)) + '.' ELSE '' END

END ---- Import table NULL column check

IF EXISTS(select [name] from sys.objects where [name] = 'VwTemp_NULLsInImportTable')
BEGIN
DROP VIEW VwTemp_NULLsInImportTable
END

-----------------------------
--- First of all fill down missing values 

IF (@filldowncolumns = 'Y')
BEGIN -- fill down columns

PRINT ''
PRINT ''
PRINT ''
PRINT 'Fill down missing values'

--(select top 1 name from sys.columns where column_id = 3 + @filldowncolumn01 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+@tab_name+'_temp')

SET @command = 
'
' + CASE WHEN @filldowncolumn01 > 0 THEN 'DECLARE @Col01 nvarchar(255)' ELSE '' END + '
' + CASE WHEN @filldowncolumn02 > 0 THEN 'DECLARE @Col02 nvarchar(255)' ELSE '' END + '
' + CASE WHEN @filldowncolumn03 > 0 THEN 'DECLARE @Col03 nvarchar(255)' ELSE '' END + '
' + CASE WHEN @filldowncolumn04 > 0 THEN 'DECLARE @Col04 nvarchar(255)' ELSE '' END + '
' + CASE WHEN @filldowncolumn05 > 0 THEN 'DECLARE @Col05 nvarchar(255)' ELSE '' END + '
' + CASE WHEN @filldowncolumn06 > 0 THEN 'DECLARE @Col06 nvarchar(255)' ELSE '' END + '
UPDATE TMP
SET
' + CASE WHEN @filldowncolumn01 > 0 THEN ' @Col01 = [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn01 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn01 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col01 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn01 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
' + CASE WHEN @filldowncolumn02 > 0 THEN ' @Col02 =,[' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn02 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn02 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col02 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn02 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
' + CASE WHEN @filldowncolumn03 > 0 THEN ' @Col03 =,[' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn03 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn03 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col03 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn03 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
' + CASE WHEN @filldowncolumn04 > 0 THEN ' @Col04 =,[' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn04 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn04 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col04 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn04 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
' + CASE WHEN @filldowncolumn05 > 0 THEN ' @Col05 =,[' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn05 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn05 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col05 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn05 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
' + CASE WHEN @filldowncolumn06 > 0 THEN ' @Col06 =,[' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn06 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] = CASE WHEN [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn06 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] IS NULL THEN @Col06 ELSE [' + (select top 1 name from sys.columns where column_id = 3 + @filldowncolumn06 and object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp') + '] END' ELSE '' END + '
FROM ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp] as TMP
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

SET @rows1 = @@ROWCOUNT

UPDATE IMP
SET
 Comments = CAST(@rows1 as varchar(10)) + ' rows filled down where necessary.' + ISNULL(' ' + Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

END -- fill down columns

-----------------------------
---Insert data into new table if one doesn't already exist

PRINT ''
PRINT ''
PRINT ''
PRINT 'Insert new data into destination table if table doesn''t already exist.'
PRINT @tab_name

IF NOT EXISTS (SELECT [name] FROM sys.objects WHERE [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN -- Start of destination table doesn't exist

SET @command = '
CREATE TABLE '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'(LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' BIGINT IDENTITY(1,1) NOT NULL,
[Filename] varchar(255) NULL,
Tablename varchar(255) NULL,
ImportDate smalldatetime NOT NULL,'
 +
CASE 
	WHEN @header LIKE '%NO' 
		THEN '
[F1] [varchar](255) NULL,
[F2] [varchar](255) NULL,
[F3] [varchar](255) NULL,
[F4] [varchar](255) NULL,
[F5] [varchar](255) NULL,
[F6] [varchar](255) NULL,
[F7] [varchar](255) NULL,
[F8] [varchar](255) NULL,
[F9] [varchar](255) NULL,'
	ELSE ''
	END
 + 
'
	CONSTRAINT [PK_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] PRIMARY KEY CLUSTERED 
	(
	[LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ASC
	)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

END -- End of destination table doesn't exist

----------------------------
---Delete month's data from destination table if it already has been loaded.
PRINT ''
PRINT ''
PRINT ''
PRINT 'Delete data for same filename from destination table.'
PRINT @filename
PRINT @tab_name

IF EXISTS (SELECT [name] FROM sys.objects WHERE [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN -- Start of destination table delete

SET @command = '
DELETE
FROM dbo.['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']
WHERE [Filename] = '+@quote+@filename+@quote+'
	AND [TableName] = '+@quote+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+@quote+'
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

SET @rows1 = @@ROWCOUNT

UPDATE IMP
SET
 Comments = CAST(@rows1 as varchar(10)) + ' rows deleted from table ('+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version + ').' + ISNULL(' ' + Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

END -- End of destination table delete

----------------------------
--- Discover any missing columns

SET @rows2 = 0

select *
INTO #TblCols
FROM (
select
name as ColumnName
from sys.columns
where OBJECT_NAME(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_Temp'

except

select
name
from sys.columns
where OBJECT_NAME(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version 
) as c

SET @rows2 = @@ROWCOUNT

IF @rows2 <> 0
BEGIN -- column changes check

IF(@debug='Y')
BEGIN
PRINT CAST(@rows2 as varchar(10)) + ' columns missing from table ('+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version + ').'
END

UPDATE IMP
SET
 Comments = CAST(@rows2 as varchar(10)) + ' columns missing from table ('+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version + ').' + ISNULL(' ' + Comments,'')
,FlagNewColumnsReceived = CASE WHEN @rows2 > 0 THEN 'Y' ELSE 'N' END
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

SET @columnchanges = '
' + @filename + ' - '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + ' has ' + CAST(@rows2 as varchar(10)) + ' new columns.
'

----------------------------------------------------------------------------
--- E-mail column changes

IF (@columnchanges <> '')
BEGIN -- Column changes e-mail
SET @email_bodytext = @email_bodytext + '
' + @filename + ' - '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + ' has ' + CAST(@rows2 as varchar(10)) + ' new columns.
' + ISNULL((SELECT TOP 1
											 LEFT(L.ColumnNames,LEN(L.ColumnNames)-1) + '.'
											FROM (SELECT
												 (select
												 ISNULL(T.ColumnName,'') + ','
												from #TblCols T
												ORDER BY ColumnName
												FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS ColumnNames
												) as L
											),'') + '
'
END -- column changes e-mail

END -- column changes check

--END -- Import table NULL column check

----------------------------------------------------------------------------
--- add in any missing columns

IF(SELECT COUNT(*) FROM #TblCols) > 0
BEGIN -- Start add in missing columns
/*
--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/
DECLARE column_cursor CURSOR FOR 
select 
ColumnName
from #TblCols
   
OPEN column_cursor 
        FETCH NEXT FROM column_cursor into @column 
WHILE @@FETCH_STATUS = 0 
        BEGIN --- column cursor1

SET @command = 'ALTER TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ADD [' + @column + '] varchar(255) NULL'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command) 

---Close second cursor 

        FETCH NEXT FROM column_cursor INTO @column 
END -- column cursor1

--IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
--BEGIN
CLOSE column_cursor 
DEALLOCATE column_cursor 
--END

END -- End add in missing columns

----------------------------
--- Change all columns to 255 characters

PRINT @tab_name
/*
--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/
DECLARE column_cursor CURSOR FOR 
	select 
	name
	from sys.columns
	where object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version 
		AND NOT(name IN('Filename','TableName') AND system_type_id = 167 AND max_length = 255)
		AND name NOT IN('LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,'ImportDate')
		AND NOT(system_type_id = 167 AND max_length = 255)
		   
OPEN column_cursor 
        FETCH NEXT FROM column_cursor into @column 
WHILE @@FETCH_STATUS = 0 
        BEGIN -- column cursor2

SET @command = 'ALTER TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ALTER COLUMN [' + @column + '] varchar(255) NULL'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command) 

---Close second cursor 

        FETCH NEXT FROM column_cursor INTO @column 
END -- column cursor2

--IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
--BEGIN
CLOSE column_cursor 
DEALLOCATE column_cursor 
--END

----------------------------
--- create insert statement

SET @command = 
'
INSERT INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'](
'
SET @command1 = 
'
SELECT
'
SET @command2 = 
'
WHERE '

/*
--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/

SET @rows1=0

DECLARE column_cursor CURSOR FOR 
	select [name]
	from sys.columns
	where object_name(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp'
	order by column_id
   
OPEN column_cursor 
        FETCH NEXT FROM column_cursor into @column 
WHILE @@FETCH_STATUS = 0 
        BEGIN -- column cursor3

SET @command = @command + '[' + @column + '],'
SET @command1 = @command1 + 'CAST([' + @column + '] as varchar(255)),'
IF(@UseOPENROWSET='Y')
BEGIN
IF(@column NOT IN('LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,'Filename','TableName','ImportDate'))
BEGIN
SET @rows1 = @rows1+1
SET @command2 = @command2 + CASE WHEN @rows1=1 THEN '[' + @column + '] IS NOT NULL' ELSE ' OR [' + @column + '] IS NOT NULL' END
END
END

---Close second cursor 

        FETCH NEXT FROM column_cursor INTO @column 
END -- column cursor3

--IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
--BEGIN
CLOSE column_cursor 
DEALLOCATE column_cursor 
--END

SET @command = LEFT(@command,len(@command)-1) + '
)'
SET @command1 = LEFT(@command1,len(@command1)-1) + '
FROM ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_Temp]
'+CASE WHEN @UseOPENROWSET='N' THEN 'WHERE ' + REPLACE(LEFT(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@command1,'LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,''),'[Filename],',''),'[TableName],',''),'[ImportDate],',''),'SELECT',''),len(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@command1,'LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,''),'[Filename],',''),'[Tablename],',''),'[ImportDate],',''),'SELECT',''))-1),',',' IS NOT NULL OR ') + ' IS NOT NULL 
' ELSE @command2 END

--- execute insert statement
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command+@command1

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
PRINT @message
END
EXEC(@command+@command1)

SET @rows1 = @@ROWCOUNT
IF(@rows1>0)
BEGIN
SET @FilesLoadedList = ISNULL(@FilesLoadedList+'; ','')+@filename
END

SET @email_bodytext = @email_bodytext + '

' + CAST(@rows1 as varchar(10)) + ' rows inserted into table ('+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version + ').'

----------------------------
---Update log table
PRINT ''
PRINT ''
PRINT ''
PRINT 'Update log file where successful.'
PRINT @filename
PRINT @tab_name

UPDATE IMP
SET
 Comments = CAST(@rows1 as varchar(10)) + ' rows inserted into table ('+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version + ').' + ISNULL(' ' + Comments,'')
,Records = @rows1
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

---Clear up temporary table with new columns

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblCols'))
BEGIN
DROP TABLE #TblCols
END

----------------------------
---Drop temporary table
PRINT ''
PRINT ''
PRINT ''
PRINT 'Drop temporary table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp.'
SET @command = '
IF EXISTS (SELECT [name] from sys.objects WHERE [type] = ''U'' AND [name] = '''+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp'')
BEGIN
DROP TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp]
END'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

----------------------------

UPDATE IMP
SET 
 ImportEndDateTime = GETDATE() 
,ImportDurationMinutes = DATEDIFF(second,@datetime,getdate()) / 60
,Comments = 'Success.' + ISNULL(' ' + Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)

----------------------------------------------------------------------------
--- Remove headings from import table into seperate table as this can cause conversion issues

--exec uspSubRTN0600RemoveHeadingsFromDataImport @tableI,@suffix,@prefix,@debug

----------------------------------------------------------------------------
---End of try and catch
END TRY --- Import table

---Catch error
BEGIN CATCH

	PRINT ''
	PRINT ''
	PRINT ''
	PRINT 'Error Detected'

SET @error =
	  (SELECT 
	/*ERROR_NUMBER() ERNumber,
	ERROR_SEVERITY() Error_Severity,
	ERROR_STATE() Error_State,
	ERROR_PROCEDURE() Error_Procedure,
	ERROR_LINE() Error_Line,*/
	ERROR_MESSAGE() Error_Message)

PRINT @error

    -- Test whether the transaction is uncommittable.
    IF (XACT_STATE()) = -1
    BEGIN
    PRINT 'The transaction is in an uncommittable state. ' +
            'Rolling back transaction.'
        ROLLBACK TRANSACTION;
    END;

    -- Test whether the transaction is active and valid.
    IF (XACT_STATE()) = 1
    BEGIN
        PRINT 'The transaction is committable. ' +
            'Committing transaction.'
        COMMIT TRANSACTION;   
    END;

SET @email_bodytext = @email_bodytext + '

Error: ' + ISNULL(@error,'Error!') + '
'

----------------------------

PRINT ''
PRINT ''
PRINT ''
PRINT 'Update log file where unsuccessful.'
PRINT @filename
PRINT @tab_name
PRINT @error

UPDATE IMP
SET 
 ImportEndDateTime = GETDATE() 
,Comments = 'Error' + ISNULL(' = ' + @error,'?') + ISNULL(' ' + Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [TableName] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+' ('+ REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+')'
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND CONVERT(varchar(20),[ImportStartDateTime],100) = CONVERT(varchar(20),@datetime,100)
	
----------------------------
---Drop temporary table
PRINT ''
PRINT ''
PRINT ''
PRINT 'Drop temporary table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp.'
SET @command = '
IF EXISTS (SELECT [name] from sys.objects WHERE [type] = ''U'' AND [name] = '''+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_') + '_temp'')
BEGIN
DROP TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@tab_name,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'.','_')+'_temp]
END'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)

----------------------------
---Update control table with errors

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_tables_ex'))
BEGIN
DROP TABLE #Tbl_sp_tables_ex
END

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_columns_ex'))
BEGIN
DROP TABLE #Tbl_sp_columns_ex
END

--- Drop columns table
IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblCols'))
BEGIN
DROP TABLE #TblCols
END
/*
--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/
END CATCH

----------------------------
---Close table cursor

	FETCH NEXT FROM table_cursor INTO @tab_cat,@tab_schem,@tab_name,@tab_type,@tab_rem
	END -- Table cursor

--IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'table_cursor')
--BEGIN
CLOSE table_cursor
DEALLOCATE table_cursor
--END 

--Clear up temporary tables and linked servers
PRINT ''
PRINT ''
PRINT ''
PRINT 'Delete temporary tables.'

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_tables_ex'))
BEGIN
DROP TABLE #Tbl_sp_tables_ex
END

IF EXISTS (SELECT * FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#Tbl_sp_columns_ex'))
BEGIN
DROP TABLE #Tbl_sp_columns_ex
END

PRINT ''
PRINT ''
PRINT ''
PRINT 'Drop linked server'
EXEC sp_dropserver @server

--- Clear up any previous cursors left in memory
/*
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'table_cursor')
BEGIN
CLOSE table_cursor
DEALLOCATE table_cursor
END
*/
/*
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END
*/
----------------------------------------------------------------------------

SET @comment = @comment + @prefix + ': OK. '

----------------------------------------------------------------------------
--- Copy file to processed folder

SET @commandDOS = '
XCOPY /DY "' + @filepath + @filename + '" "' + @filepath + 'Processed\"
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC master..xp_cmdshell @commandDOS, no_output

--- Delete file
SET @commandDOS = '
DEL "' + @filepath + @filename + '"
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC master..xp_cmdshell @commandDOS, no_output

SET @email_bodytext = @email_bodytext + '
----------------------------------------------------------------------------

'

----------------------------------------------------------------------------
---Close file cursor
	FETCH NEXT FROM file_cursor INTO @filename,@filemodifieddate,@filemodifieddateP
	END --- file cursor

--IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'file_cursor')
--BEGIN
CLOSE file_cursor
DEALLOCATE file_cursor
--END

----------------------------------------------------------------------------
--- Work out failed file loads or tables with no data ...

IF EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN

SELECT
 RTRIM(SUBSTRING(DirList,37,255)) as [Filename] 
INTO dbo.zDirList
FROM #TblTempDirList DL
WHERE CASE
		WHEN DirList LIKE '%.xls[,b,m,x]%' THEN 1
		WHEN DirList LIKE '%.mdb%' THEN 1
		WHEN DirList LIKE '%.accdb%' THEN 1
		ELSE 0 
		END = 1
GROUP BY RTRIM(SUBSTRING(DirList,37,255)) --as [Filename] 

SET @command = 
'
SELECT
 [Filename] 
FROM [' + @db + '].dbo.zDirList

EXCEPT

SELECT
 [Filename]
FROM [' + @db + '].dbo.['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']
GROUP BY [Filename]
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
PRINT @command
END
exec(@command)

SET @email_query = @command

END

END --- Files to load

IF NOT EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN
SET @email_query = 'SELECT ''Table not found!'' as Comment'
END

----------------------------------------------------------------------------
--- Clear up any previous cursors left in memory

IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'column_cursor')
BEGIN
CLOSE column_cursor
DEALLOCATE column_cursor
END

IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'table_cursor')
BEGIN
CLOSE table_cursor
DEALLOCATE table_cursor
END

IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'file_cursor')
BEGIN
CLOSE file_cursor
DEALLOCATE file_cursor
END

----------------------------------------------------------------------------
--- E-mail results

IF(@debug='Y')
BEGIN
PRINT '@profile_name = '+ISNULL(@email_profilename,'missing!')
PRINT '@recipients = '+ISNULL(@email_recipients,'missing!')
PRINT '@subject = '+ISNULL(@email_subjectline,'missing!')
PRINT '@body = '+ISNULL(@email_bodytext,'missing!')
PRINT '@query_result_separator = '+ISNULL(CASE WHEN @email_ColumnDelineator = CHAR(9) THEN 'char(9) = CSV' ELSE @email_ColumnDelineator END,'missing!')
PRINT '@query = '+ISNULL(@email_query,'missing!')
END

--- CAUSING AN ERROR. NOT SURE WHY!
IF(@email_bodytext <> 'Results of last load for spSubLoadAllUserMicrosoftOfficeFiles:')
BEGIN
EXEC msdb.dbo.sp_send_dbmail 
	 @profile_name = @email_profilename					--When no profile_name is specified, sp_send_dbmail uses the default private profile for the current user. If the user does not have a default private profile, sp_send_dbmail uses the default public profile for the msdb database. If the user does not have a default private profile and there is no default public profile for the database, @profile_name must be specified.
	,@recipients = @email_recipients					-- Is a semicolon-delimited list of e-mail addresses to send the message to.
	--,@copy_recipients = 'copy_recipient [ ; ...n ]'	-- Is a semicolon-delimited list of e-mail addresses to carbon copy the message to. 
	--,@blind_copy_recipients = 'blind_copy_recipient [ ; ...n ]' --Is a semicolon-delimited list of e-mail addresses to blind carbon copy the message to. 
	--,@from_address = 'from_address'					-- This is an optional parameter used to override the settings in the mail profile. This parameter is of type varchar(MAX). SMTP security settings determine if these overrides are accepted. If no parameter is specified, the default is NULL.
	--,@reply_to = 'reply_to'							-- This is an optional parameter used to override the settings in the mail profile. This parameter is of type varchar(MAX). SMTP security settings determine if these overrides are accepted. If no parameter is specified, the default is NULL.
	,@subject = @email_subjectline						-- The subject is of type nvarchar(255). If no subject is specified, the default is 'SQL Server Message'.
	,@body = @email_bodytext							-- Is the body of the e-mail message. The message body is of type nvarchar(max), with a default of NULL.
	--,@body_format = 'HTML'							-- Is the format of the message body. The parameter is of type varchar(20), with a default of NULL. TEXT or HTML.
	--,@importance = 'Normal'							-- The parameter is of type varchar(6). The parameter may contain one of the following values: Low; Normal; High.
	--,@sensitivity = 'Normal'							-- The parameter is of type varchar(12). The parameter may contain one of the following values: Normal; Personal; Private; Confidential.
	--,@file_attachments = 'attachment [ ; ...n ]'		-- Is a semicolon-delimited list of file names to attach to the e-mail message. Files in the list must be specified as absolute paths. The attachments list is of type nvarchar(max). By default, Database Mail limits file attachments to 1 MB per file.
	,@query = @email_query								--  Is a query to execute. The results of the query can be attached as a file, or included in the body of the e-mail message. The query is of type nvarchar(max), and can contain any valid Transact-SQL statements. Note that the query is executed in a separate session, so local variables in the script calling sp_send_dbmail are not available to the query.
	--,@execute_query_database = 'db_name'				-- SELECT DB_NAME() -- Is the database context within which the stored procedure runs the query. 
	,@attach_query_result_as_file = 1					--1 = as file attachment. Specifies whether the result set of the query is returned as an attached file. attach_query_result_as_file is of type bit, with a default of 0.
	,@query_attachment_filename = @email_AttachFilename -- Specifies the file name to use for the result set of the query attachment. query_attachment_filename is of type nvarchar(255), with a default of NULL. This parameter is ignored when attach_query_result is 0. When attach_query_result is 1 and this parameter is NULL, Database Mail creates an arbitrary filename.
	,@query_result_header = 1							--Default = 1 = included. Specifies whether the query results include column headers. The query_result_header value is of type bit. When the value is 1, query results contain column headers. When the value is 0, query results do not include column headers. This parameter defaults to 1. This parameter is only applicable if @query is specified.
	,@query_result_width = 8000							-- Without it it starts wrapping text onto new lines!
	,@query_result_separator = @email_ColumnDelineator	-- Is the character used to separate columns in the query output. The separator is of type char(1). Defaults to ' ' (space).
	,@exclude_query_output = 0							-- Specifies whether to return the output of the query execution in the e-mail message. exclude_query_output is bit, with a default of 0. When this parameter is 0, the execution of the sp_send_dbmail stored procedure prints the message returned as the result of the query execution on the console. When this parameter is 1, the execution of the sp_send_dbmail stored procedure does not print any of the query execution messages on the console.
	,@append_query_error = 1							-- Specifies whether to send the e-mail when an error returns from the query specified in the @query argument. append_query_error is bit, with a default of 0. When this parameter is 1, Database Mail sends the e-mail message and includes the query error message in the body of the e-mail message. When this parameter is 0, Database Mail does not send the e-mail message, and sp_send_dbmail ends with return code 1, indicating failure.
	--,@query_no_truncate = 1							-- Specifies whether to execute the query with the option that avoids truncation of large variable length data types (varchar(max), nvarchar(max), varbinary(max), xml, text, ntext, image, and user-defined data types). When set, query results do not include column headers. The query_no_truncate value is of type bit. When the value is 0 or not specified, columns in the query truncate to 256 characters. When the value is 1, columns in the query are not truncated. This parameter defaults to 0.
	--,@query_result_no_padding = 1 --No padding			-- The type is bit. The default is 0. When you set to 1, the query results are not padded, possibly reducing the file size.If you set @query_result_no_padding to 1 and you set the @query_result_width parameter, the @query_result_no_padding parameter overwrites the @query_result_width parameter. If you set the @query_result_no_padding to 1 and you set the @query_no_truncate parameter, an error is raised.
	--,@mailitem_id =  mailitem_id ] [ OUTPUT ]			-- Optional output parameter returns the mailitem_id of the message. The mailitem_id is of type int.
	; 
END

---------------------------------------------------------------
--END -- Directory exists - see line 333.

IF EXISTS (SELECT [name] FROM sys.objects WHERE [name] = N'zDirList')
BEGIN
DROP TABLE dbo.zDirList
END

---------------------------------------------------------------
---Close cursor 

        FETCH NEXT FROM folder_cursor INTO @usr,@folder
END 
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'folder_cursor')
BEGIN
CLOSE folder_cursor 
DEALLOCATE folder_cursor 
END

---------------------------------------------------------------

IF EXISTS (SELECT [name] FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblTempDirList'))
BEGIN
DROP TABLE #TblTempDirList
END

---------------------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[spSubLoadAllUserTextFiles]    Script Date: 17/03/2023 15:58:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spSubLoadAllUserTextFiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spSubLoadAllUserTextFiles] AS' 
END
GO
ALTER PROCEDURE [dbo].[spSubLoadAllUserTextFiles] (@usrI varchar(128)='',@prefix varchar(14)='_',@tableI varchar(95)='',@suffix varchar(14)='_Import',@version varchar(5)='',@debug varchar(1)='N') AS

/*

VERSION CONTROL

Modified    Modifyee	Modification

----------- -----------	-----------------------------------------------

10-Mar-2023	JCEH		Created procedure.

----------- -----------	-----------------------------------------------

NOTES:

USER FOLDERS (Loops through)
	--> FILES in destination folder (Loops through)

HEADER = NO. Load as single column. 
	--> Check delineators. Jail inconsistent lines.
	--> Spit out and reload into generic temporary table.

HEADER = YES. Load header if header present (first line to load might not be line 1. If preceeding lines do not have consistent delineators then set header to NO.)
	--> csv (, or ",") then convert to pipe (|) to avoid issues. Create tables.
	--> Not csv, use delineator. Create tables.

CSV (, or ",") then check file for issues. If issues, clean. Otherwise take as is.
Not CSV. Load file into temporary table.

Compare temporary table to final table. 
	--> Add columns as necessary. 
	--> Delete data from final table if filename already exists.
	--> Insert latest data.

Move files to processed folder.
*/

------------------------------------------
--- Declare variables

--DECLARE @debug				varchar(1)			SET @debug = 'Y'
--DECLARE @prefix				varchar(5)			SET @prefix  = '_'		-- table name prefix after username and before table name.
--DECLARE @suffix				varchar(14)			SET @suffix = '_Import'		-- table name suffix before and after version.
--DECLARE @tableI				varchar(95)			SET @tableI  = ''		-- Intended table to load into. If blank then uses filename.
--DECLARE @usrI				varchar(128)		SET @usrI = SUSER_NAME() -- Username but not required here.
--DECLARE @version			varchar(5)			SET @version = ''
DECLARE @command			varchar(max)  --for the CREATE statements>8000
DECLARE @commandDOS			varchar(8000) --bcp doesn't accept varchar(max)
DECLARE @ColDelineator		varchar(5)
DECLARE @ColDelineatorI		varchar(5)
DECLARE @column				varchar(255)
DECLARE @columncount		int = 0
DECLARE @colid				int
DECLARE @columnchanges		varchar(2048)
DECLARE @comment			varchar(8000)		SET @comment = ''
DECLARE @datetime			smalldatetime		SET @datetime = CONVERT(smalldatetime,CONVERT(varchar(20),GETDATE(),100),100)
DECLARE @db					varchar(255)		SET @db = DB_NAME()
DECLARE @DefaultColumnWidth int					SET @DefaultColumnWidth = -1
DECLARE @error				varchar(1024)
DECLARE @filemodifieddate	smalldatetime
DECLARE @filemodifieddateP	smalldatetime -- Previous modified date
DECLARE @filename			varchar(255)
DECLARE @filepath			varchar(2000)
DECLARE @filepathname		varchar(2255)
DECLARE @filepathRoot		varchar(255)	SET @filepathRoot = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName = 'Filepath_UserLoadFolderRoot')
DECLARE @FilesLoadedList	varchar(max)		--SET @FilesLoadedList = ''
DECLARE @FirstRow			varchar(10)			SET @FirstRow = 1
DECLARE @folder				varchar(255)
DECLARE @header				varchar(3)
DECLARE @message			varchar(8000)		SET @message = ''
DECLARE @periodicity		varchar(3)			SET @periodicity = 's'
DECLARE @quote				varchar(1)			SET @quote = ''''
DECLARE @resulttable		TABLE(result int NULL)
DECLARE @rows1				bigint				SET @rows1 = 0
DECLARE @rows2				bigint				SET @rows2 = 0
DECLARE @RowTerminator		varchar(5)			SET @RowTerminator = '\n'
DECLARE @server				varchar(255)		SET @server	  = (select @@servername)	/* SQL instance*/
DECLARE @sqlProductVersion	float				SET @sqlProductVersion = (SELECT CAST(SERVERPROPERTY('ProductMajorVersion') as float) as sqlProductMajorVersion)
DECLARE @table				varchar(255)
DECLARE @tableTempf			varchar(255)		SET @tableTempf	  = '##TblTempFormatFile_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SUSER_NAME()+CONVERT(varchar(24),GETDATE(),113),'\',''),'/',''),':',''),' ',''),'-',''),'$',''),'.','')
DECLARE @tableTempsc		varchar(128)		SET @tableTempsc  = '##TblTempImportAsSingleColumn_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SUSER_NAME()+CONVERT(varchar(24),GETDATE(),113),'\',''),'/',''),':',''),' ',''),'-',''),'$',''),'.','')
DECLARE @tableTempsci		varchar(128)		SET @tableTempsci = '##TblTempImportAsSingleColumnWithIndex_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SUSER_NAME()+CONVERT(varchar(24),GETDATE(),113),'\',''),'/',''),':',''),' ',''),'-',''),'$',''),'.','')
DECLARE @usr				varchar(128)		SET @usr = ''--SUSER_NAME() -- Username but not required here.
--- Email details
IF(SELECT COUNT(*) FROM TblPar_Parameters WHERE ParameterName = 'Email recipients (import procedures)')=0
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString) VALUES('Email recipients (import procedures)','yourname@domain.net')
END

DECLARE @instancename			varchar(255)	SET @instancename = @@SERVERNAME -- SQL instance name
DECLARE @email_profilename		varchar(50)		SET @email_profilename = 'BIU' -- @profile_name
DECLARE @email_recipients		varchar(255)	SET @email_recipients  = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName = 'Email recipients (import procedures)') -- @recipients
DECLARE @email_subjectline		varchar(255)	SET @email_subjectline = 'Load all user (flat) text files (' + @db + ' on ' + @instancename+')' -- @subject 
DECLARE @email_bodytext			varchar(8000)	SET @email_bodytext = 'Results of last import using spSubLoadAllUserTextFiles ('+@usr+'):' -- @body
DECLARE @email_query			varchar(8000)
DECLARE @email_AttachFilename	varchar(255)	SET @email_AttachFilename = 'FilesNotLoadedTXT.csv'
DECLARE @email_ColumnDelineator varchar(5)		SET @email_ColumnDelineator = CHAR(9)

---------------------------------------------------------------
--- Create temporary table with single column for contents.

SET @command = '
CREATE TABLE '+@tableTempsc+'(FileContentsIncludingHeader varchar(max) NULL)
'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

---------------------------------------------------------------
--- Create temporary table with single column for contents and an index column.

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempsci))
BEGIN
SET @command = 'CREATE TABLE '+@tableTempsci+'(LocalIndex bigint IDENTITY(1,1) NOT NULL,FileContentsIncludingHeader varchar(max) NULL, CONSTRAINT[PK_'+@tableTempsci+'] PRIMARY KEY CLUSTERED (LocalIndex ASC))'
END
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

---------------------------------------------------------------
---  Create temporary table for format file and export format file ...

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempf))
BEGIN
SET @command = '
CREATE TABLE ['+@tableTempf+'](FileContent varchar(1024) NULL)
'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

SET @command = 'TRUNCATE TABLE ['+@tableTempf+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

/*
Version
NumberOfColumns
FileColumnNumber	DataType	DataHandling	SQL_FieldLength		Terminator	DbColumnNumber	DbColumnName	Collation

9.0
1
1                   SQLCHAR     0               8060                "\n"		2				S1               ""
*/
SET @command = '
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''9.0'')
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''1'')
INSERT INTO ['+@tableTempf+'](FileContent) VALUES(''1                   SQLCHAR     0               8060                "'+@RowTerminator+'"    2              FileContentsIncludingHeader               ""'')
'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

IF(@debug='Y')
BEGIN
PRINT ''
PRINT '@filepath = '+ISNULL(@filepath,'Unknown')
PRINT '@tableTempf = '+ISNULL(@tableTempf,'Unknown')
PRINT '@server = ' +ISNULL( @server,'Unknown')
PRINT ''
END
--- Filepath unknown at this stage - changes for each folder!

---------------------------------------------------------------

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
CREATE TABLE #TblColumnCounts(ColumnCount int NOT NULL,Records bigint NOT NULL)
END

---------------------------------------------------------------
---------------------------------------------------------------
--- Delcare cursor to loop through folders

IF(@debug='Y')
BEGIN
PRINT '0100 - loop folders'
END

DECLARE user_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
	 S.Username
	,ParameterString as [Folder] 
	FROM dbo.TblPar_Parameters as P
		RIGHT OUTER JOIN SystemTracking.dbo.TblSysUsers as S
			ON ISNULL(P.ParameterName,@filepathRoot+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usrI,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')) = S.Username+'_Filepath'
	WHERE S.SQLaccountDisabled = 'N'
		AND S.WindowsAccountDisabled = 'N'
		AND S.EmailAddress LIKE '%@%'
		AND S.Username NOT IN('INFOHUB-UHC2\sqlAgent','sa')
		AND ISNULL(p.ParameterString,'') <> 'D:\DO NOT LOAD INTO THIS DATABASE\'
		AND CASE WHEN @usrI = '' THEN S.Username ELSE @usrI END = S.Username
        
OPEN user_cursor 
        FETCH NEXT FROM user_cursor into @usr,@folder
WHILE @@FETCH_STATUS = 0 
        BEGIN 

----------------------------------------------------------------------------

IF(@debug='Y')
BEGIN
PRINT ''
PRINT '-----------------------------------------------'
PRINT '***************** NEXT FOLDER *****************'
PRINT '-----------------------------------------------'
PRINT ''
PRINT '@usr = '+@usr
PRINT ''
END

----------------------------------------------------------------------------

SET @email_bodytext = 'Results of last import using spSubLoadAllUserTextFiles ('+@usr+'):' -- @body
SET @email_recipients = ISNULL((SELECT TOP 1 EmailAddress FROM SystemTracking.dbo.TblSysUsers WHERE Username = @usr),@email_recipients)
SET @filepath		  = ISNULL(@folder,@filepathRoot+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_'))
SET @filepath		  = CASE WHEN RIGHT(@filepath,1) = '\' THEN @filepath ELSE @filepath+'\' END
--SET @table			  = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'

IF(@debug='Y')
BEGIN
PRINT '@email_recipients = '+@email_recipients
PRINT '@filepath = '+@filepath
--PRINT '@table = '+@table
PRINT ''
END

DELETE FROM @resulttable

EXECUTE AS LOGIN = @usr 
INSERT INTO @resulttable
SELECT HAS_DBACCESS(@db)
REVERT;

IF((select TOP 1 result from @resulttable)=0)
BEGIN
SET @filepath = 'USER DOES NOT HAVE ACCESS TO DATABASE '+@db
PRINT '@filepath = '+@filepath
END 

----------------------------------------------------------------------------
--- Set default parameter values if not present ...
IF(@debug='Y')
BEGIN
PRINT '0100 - Add missing parameters to TblPar_Parameters'
END

IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineator','‡',NULL,'Delineator used to separate columns. Default = double-dagger = ‡. csv with text qualifier = CSV if using SQL 2014 or later.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ColumnDelineatorOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ColumnDelineatorOriginal',',',NULL,'Delineator originally used used to separate columns that needs replacing with something sensible. Default = comma = '',''. csv with text qualifier = CSV for RFC4180 compliant.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineator','NO',NULL,'NO = Leave file alone; SQL = convert in SQL using scalar function (slow but can handle no header); Powershell = Powershell conversion (fast but cannot handle no header or duplicate columns).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorSkipRows')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorSkipRows',NULL,0,'Default = 0. Number of rows to delete from the top of the raw file before the header row.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ConvertColumnDelineatorChunkFiles')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ConvertColumnDelineatorChunkFiles','N',1000000,'N = Leave file alone; Y = split the file into multiple files containing up to 1,000,000 rows.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DefaultColumnWidth')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DefaultColumnWidth',NULL,-1,'Only use to force larger columns than the default (255 unless lots of columns). Set to minus 1 to allow proc to determine most appropriate width.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_DestinationTable')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_DestinationTable','Default',NULL,'Core destination table name. Set to "Default" or '''' in order to use the default behaviour of using the filename.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable01','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable02','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable03','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable04','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable05','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ExcludeRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ExcludeRangeSheetTable06','Enter named range or sheet or table name to exclude here',NULL,'Enter named range or sheet or table name to exclude in ParameterString field.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_Filepath',@filepath,NULL,'Filepath to data to load.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn01',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn02',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn03',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn04',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn05',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumn06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumn06',NULL,'0',NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FillDownColumnsRequired')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FillDownColumnsRequired','N',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_FirstRow')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_FirstRow','1',NULL,'First row in the file to load. If header is in row 1 then leave as 1.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_ForceTextIMEX')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_ForceTextIMEX','1',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_IncludeRangeFilterDatabase')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_IncludeRangeFilterDatabase','N',NULL,'Set to Y to include the system range #FileterDatabase (might duplicate data).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable01')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable01','',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable02')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable02','''',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable03')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable03','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable04')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable04','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable05')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable05','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_LoadRangeSheetTable06')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_LoadRangeSheetTable06','Enter named range or sheet or table name',NULL,NULL)
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_RowTerminator')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_RowTerminator','\n',NULL,'Delineator used to indicate a new daat row follows. Windows default is \n (=line feed + carriage return). Tab is \t and linux (and Oracle) just use line feed = char(10).')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_TextQualifierOriginal')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_TextQualifierOriginal','"',NULL,'Text qualifier used when text contains delimiter. Default = double-quotes = ''"''. csv with text qualifier = ",".')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseHeaderForColumnNames')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseHeaderForColumnNames','YES',NULL,'YES for well structured file with header at top and no duplicate column names, otherwise use NO.')
END
IF NOT EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_UseOPENROWSET')
BEGIN
INSERT INTO TblPar_Parameters(ParameterName,ParameterString,ParameterInt,Comments) VALUES(@usr+'_UseOPENROWSET','N',NULL,'Set ParameterString to Y to use OPENROWSET instead of linked server.')
END

----------------------------------------------------------------------------

BEGIN TRY --- FOLDER

SET @FilesLoadedList = NULL

------------------------------------------
--- Get rid of any temporary tables left from testing

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblImpAs1Col'))
BEGIN
DROP TABLE #TblImpAs1Col
END

------------------------------------------
--- Create temporary table with single column to place header into

IF NOT EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblImpAs1Col'))
BEGIN
CREATE TABLE #TblImpAs1Col(AllData varchar(max) NULL)
END

SET @commandDOS = 'bcp "tempdb..['+@tableTempf+']" out "'+@filepath+@tableTempf+'.fmt" /c /t "|" /S "'+@server+'" -T'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT @commandDOS
PRINT ''
PRINT ''
EXEC master..xp_cmdshell @commandDOS
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS,no_output
END

----------------------------------------------------------
--- Obtain directory (folder) contents

SET @commandDOS = 'IF EXIST "'+@filepath+'" DIR "'+@filepath+'" > "'+@filepath+'DirectoryListing.txt"'
IF(@debug='Y')
BEGIN
PRINT @commandDOS
EXEC master..xp_cmdshell @commandDOS--, no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS, no_output
END

-----------------------------------------------------------
---Create table to put results in if it doesn't exist

IF NOT EXISTS (SELECT [name] FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblTempDirList'))
BEGIN
CREATE TABLE #TblTempDirList([DirList] [varchar](8000) NULL)
END

TRUNCATE TABLE #TblTempDirList

IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @filepath

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END

-----------------------------------------------------------
---Set variable values - next 3 lines will need editing to specific use.
SET @table		  = '#TblTempDirList'			/* Table to insert data into */
SET @filename	  = 'DirectoryListing.txt'		/* Filename with data to insert into table*/
SET @quote		  = ''''						/* Gets round quote problem in a way that makes code more easily readable */
SET @filepathname = @filepath + @filename		/* Combine path and name of file with data */

-----------------------------------------------------------
---Clear out old data

SET @command = 'TRUNCATE TABLE ['+@table+']'
exec (@command)

-----------------------------------------------------------
---Insert directory listing  from file 

IF EXISTS(select ParameterName from TblPar_Parameters where ParameterName = @usr+'_Filepath')
BEGIN
SET @command = 'BULK INSERT ['+@table+'] FROM "'+@filepathname+'" WITH (FIELDTERMINATOR = '+@quote+'££££'+@quote+', ROWTERMINATOR = '+@quote+'\n'+@quote+', FIRSTROW = 1)'
IF (@debug='Y')
BEGIN
PRINT(@command) 
END
exec (@command)
END

----------------------------------------------------------------------------
----------------------------------------------------------
--- Setup cursor to loop through files

IF(SELECT
	 COUNT(*) as Files 
	FROM #TblTempDirList
	WHERE DirList NOT LIKE '%DirectoryListing.txt%'
		AND CASE
			WHEN DirList LIKE '%.txt%' THEN 1
			WHEN DirList LIKE '%.csv%' THEN 1
			WHEN DirList LIKE '%.vbs%' THEN 1
			ELSE 0
			END = 1
		AND DirList NOT LIKE '%file1.txt%'
	) > 0
BEGIN --- Files to load

SET @email_bodytext = @email_bodytext+'

Files in load directory = '+ISNULL((SELECT TOP 1
								 LEFT(DIR.FileNames,LEN(DIR.FileNames)-1)+'.'
								FROM (SELECT
									 (select
									 ISNULL(RTRIM(SUBSTRING(sc1.DirList,37,255)),'')+','
									from #TblTempDirList sc1
									WHERE DirList NOT LIKE '%DirectoryListing.txt%'
										AND CASE
											WHEN DirList LIKE '%.txt%' THEN 1
											WHEN DirList LIKE '%.csv%' THEN 1
											WHEN DirList LIKE '%.vbs%' THEN 1
											ELSE 0
											END = 1
										AND DirList NOT LIKE '%file1.txt%'
									ORDER BY DirList
									FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS FileNames
									) as DIR
								),'')+'
'
END --- Files to load

--- Clear up any previous cursors left in memory
IF EXISTS(select cursor_name from sys.syscursors where cursor_name = 'file_cursor')
BEGIN
CLOSE file_cursor
DEALLOCATE file_cursor
END

IF(@debug='Y')
BEGIN
SELECT
 'Files to load ...' as Comment
,RTRIM(SUBSTRING(DirList,37,255)) as [FileName] 
,CAST(LEFT([DirList],20) as datetime) as FileModificationDateTime
--/*American date*/,CAST(SUBSTRING([DirList],4,2)+'/'+LEFT([DirList],2)+'/'+SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
--/*British date*/,CAST(LEFT([DirList],2)+'/'+SUBSTRING([DirList],4,2)+'/'+SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
,ISNULL(iLog.FileModificationDateTime,'05-Jul-1948') as PreviousFileModificationDateTime
FROM #TblTempDirList DL
	LEFT OUTER JOIN (SELECT
					 SourceName as [Filename]
					,ImportStartDateTime
					,SourceDateStamp as FileModificationDateTime
					,ROW_NUMBER() OVER(PARTITION BY SourceName ORDER BY SourceName,ImportStartDateTime DESC) as Seq
					FROM dbo.TblLog_Imports
					) as iLog
		ON RTRIM(SUBSTRING(DL.DirList,40,255)) = iLog.[Filename] AND iLog.Seq = 1
WHERE DirList NOT LIKE '%DirectoryListing.txt%'
	AND CASE
		WHEN DirList LIKE '%.txt%' THEN 1
		WHEN DirList LIKE '%.csv%' THEN 1
		WHEN DirList LIKE '%.vbs%' THEN 1
		ELSE 0
		END = 1
	AND DirList NOT LIKE '%file1.txt%'
ORDER BY DirList
END

IF(@debug='Y')
BEGIN
PRINT '0500 loop files'
END

DECLARE file_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT RTRIM(SUBSTRING(DirList,37,255)) as [FileName] 
	,CAST(LEFT([DirList],20) as datetime) as FileModificationDateTime
	--/*American date*/,CAST(SUBSTRING([DirList],4,2)+'/'+LEFT([DirList],2)+'/'+SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
	--/*British date*/,CAST(LEFT([DirList],2)+'/'+SUBSTRING([DirList],4,2)+'/'+SUBSTRING([DirList],7,4) + SUBSTRING([DirList],12,9) as datetime) as FileModificationDateTime
	,ISNULL(iLog.FileModificationDateTime,'05-Jul-1948') as PreviousFileModificationDateTime
	FROM #TblTempDirList DL
		LEFT OUTER JOIN (SELECT
						 SourceName as [Filename]
						,ImportStartDateTime
						,SourceDateStamp as FileModificationDateTime
						,ROW_NUMBER() OVER(PARTITION BY SourceName ORDER BY SourceName,ImportStartDateTime DESC) as Seq
						FROM dbo.TblLog_Imports
						) as iLog
			ON RTRIM(SUBSTRING(DL.DirList,40,255)) = iLog.[Filename] AND iLog.Seq = 1
	WHERE DirList NOT LIKE '%DirectoryListing.txt%'
		--AND DirList = 'Monthly-SITREPs-CC-Extracts-7hs8d_20140331.csv'
		AND CASE
			WHEN DirList LIKE '%.txt%' THEN 1
			WHEN DirList LIKE '%.csv%' THEN 1
			WHEN DirList LIKE '%.vbs%' THEN 1
			ELSE 0
			END = 1
		AND DirList NOT LIKE '%file1.txt%'
		--AND NOT(CAST(LEFT([DirList],20) as datetime) BETWEEN DATEADD(minute,-1,ISNULL(iLog.FileModificationDateTime,'01-Jan-1900 00:01')) AND DATEADD(minute,+1,ISNULL(FileModificationDateTime,'01-Jan-1900'))) --09-Nov-2020, JCEH: Changed cursor select from ISNULL(XXX,01-Jan-1900) to ISNULL(XXX,01-Jan-1900 00:01) to ahndle smalldatetime conversion error - implicit conversion somewhere.
		--AND RTRIM(SUBSTRING(DirList,37,255)) <> 'GRIDALL'
	ORDER BY RTRIM(SUBSTRING(DirList,37,255))

OPEN file_cursor 
        FETCH NEXT FROM file_cursor INTO @filename,@filemodifieddate,@filemodifieddateP
WHILE @@FETCH_STATUS = 0 
        BEGIN 

SET @datetime = CONVERT(smalldatetime,CONVERT(varchar(20),GETDATE(),100),100)
SET @FilesLoadedList = ISNULL(@FilesLoadedList+'; ','')+@filename
SET @filepathname = CASE WHEN RIGHT(@filepath,1) = '\' THEN @filepath + @filename ELSE @filepath+'\'+@filename END
SET @table = (SELECT TOP 1 ParameterString FROM TblPar_Parameters WHERE ParameterName=@usr+'_DestinationTable')
SET @table = CASE WHEN @tableI = '' AND @table IN('','Default') THEN REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.txt',''),'.csv',''),'.vbs',''),'.log','') WHEN @tableI = '' AND @table NOT IN('','Default') THEN @table ELSE @tableI END

IF(@debug='Y')
BEGIN
PRINT ''
PRINT '***************** NEXT FILE *****************'
PRINT ''
PRINT @filename+' ['+CAST(LEN(@filename) as varchar(10))+' characters]'
PRINT @filepathname+' ['+CAST(LEN(@filepathname) as varchar(10))+' characters]'
PRINT @table
END

IF EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp')
BEGIN -- Temp table exists
IF(@debug='Y')
BEGIN
PRINT '0550 DROP TABLE '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'
END
SET @command = 
'
DROP TABLE '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)
END -- Temp table exists

SET @email_bodytext = @email_bodytext+'
----------------------------------------------------------------------------
Loading file: '+@filename+'
----------------------------------------------------------------------------
'

BEGIN TRY

SET @ColDelineator = (SELECT ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_ColumnDelineator') 
SET @ColDelineatorI = (SELECT ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_ColumnDelineator') 
SET @FirstRow = (SELECT ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_FirstRow')
SET @Header = (SELECT ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_UseHeaderForColumnNames')
SET @RowTerminator = (SELECT ParameterString FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_RowTerminator')
SET @DefaultColumnWidth = (SELECT ParameterInt FROM dbo.TblPar_Parameters WHERE ParameterName = @usr+'_DefaultColumnWidth')

IF(@debug='Y')
BEGIN
PRINT '0800 parameters'
PRINT '@filepathname = '+@filepathname
PRINT '@ColDelineator = '+@ColDelineator
PRINT '@FirstRow = '+@FirstRow
PRINT '@RowTerminator = '+@RowTerminator
PRINT '@DefaultColumnWidth = '+CAST(@DefaultColumnWidth as varchar(20))
END

SET @email_bodytext = @email_bodytext+'
Column delineator: '+@ColDelineator+'
Row terminator: '+@RowTerminator+'
First row: '+@FirstRow+'
'

----------------------------
---Update log table

IF(@debug='Y')
BEGIN
PRINT '0900 Update log table for latest table import.'
PRINT @filename
END

INSERT INTO [dbo].[TblLog_Imports] ([Tablename],[Sourcepath],[Sourcename],SourceDateStamp,[ImportStartDateTime]) VALUES(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,@filepath,@filename,@filemodifieddate,@datetime)

----------------------------------------------------------

IF(LEN(@filepathname)>128)
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /Y "'+@filepathname+'" "'+@filepath+'file1.txt*"'

IF(@debug='Y')
BEGIN
PRINT @commandDOS
exec xp_cmdshell @commandDOS--,no_output
END
IF(@debug='N')
BEGIN
exec xp_cmdshell @commandDOS,no_output
END
SET @filepathname = @filepath+'file1.txt'

END  -- Filename > 128 characters

----------------------------------------------------------
--- No reliable header expected

IF(@header='NO')
BEGIN -- No header

IF(@debug='Y')
BEGIN
PRINT @filename+'1000 - no header'
PRINT @filepathname+'1000 - no header'
END

IF(@filepathname NOT LIKE '%file1.txt')
BEGIN -- Filename <= 128 characters

-----------------------------------------------------------
---Insert new data from file - column delineated by commas with newline for end of row and no text delineators - i.e. standard csv.

SET @command = '
TRUNCATE TABLE ['+@tableTempsci+']
BULK INSERT ['+@tableTempsci+'] FROM '''+@filepath+@filename+''' WITH (TABLOCK,CODEPAGE=''RAW'',FORMATFILE='''+@filepath+@tableTempf+'.fmt'')'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec (@command)
END   -- Filename <= 128 characters

-----------------------------------------------------------

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters

-----------------------------------------------------------
---Insert new data from file - column delineated by commas with newline for end of row and no text delineators - i.e. standard csv.

SET @command = '
TRUNCATE TABLE ['+@tableTempsci+']
BULK INSERT ['+@tableTempsci+'] FROM '''+@filepath+'file1.txt'' WITH (TABLOCK,CODEPAGE=''RAW'',FORMATFILE='''+@filepath+@tableTempf+'.fmt'')'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec (@command)

END   -- Filename > 128 characters

IF(@debug='Y')
BEGIN
PRINT '1100 - no header; check column count'
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

SET @command = '
SELECT
 LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 as ColumnCount
,COUNT(*) as Records
FROM ['+@tableTempsci+']
WHERE FileContentsIncludingHeader IS NOT NULL
GROUP BY LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1
ORDER BY COUNT(*) DESC
'
IF(@debug='Y')
BEGIN
PRINT 'INSERT INTO #TblColumnCounts'
PRINT @command
END

INSERT INTO #TblColumnCounts
exec(@command)

SET @rows1 = (SELECT TOP 1 ColumnCount FROM #TblColumnCounts ORDER BY Records DESC) -- number of columns in file

IF(@debug='Y')
BEGIN
PRINT 'Number of columns found = '+CAST(@rows1 as varchar(10))
END

--- PROBLEM! When LOADS of columns it overflows max row length. 
--- If loads of columns then need to drop field width to fit onto a SQL page ...
IF(@debug='Y')
BEGIN
PRINT @filepathname+'1110 - no header; create temporary table - work out max column width feasible'
END

SET @DefaultColumnWidth = CASE
							WHEN @DefaultColumnWidth > 0 THEN @DefaultColumnWidth
							WHEN @rows1 > 800 THEN 10
							WHEN @rows1 > 500 THEN 15
							WHEN @rows1 > 400 THEN 20
							WHEN @rows1 > 160 THEN 50
							WHEN @rows1 > 80 THEN 100
							ELSE 255
							END

SET @columncount = 0
SET @command = 'CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]('

--- Generate generic column headings
IF(@debug='Y')
BEGIN
PRINT @filepathname+'1120 - no header; create temporary table - create column headings'
END

WHILE @columncount < @rows1
BEGIN -- WHILE
SET @columncount = @columncount + 1
SET @command = @command+'Col'+LEFT('0000',4-LEN(@columncount)) + CAST(@columncount as varchar(4))+' varchar('+CAST(@DefaultColumnWidth as varchar(10))+') NULL,'
END -- WHILE

SET @command = LEFT(@command,LEN(@command)-1)+')'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

--- Jail records with different number of delineators
IF EXISTS(select 1 from #TblColumnCounts where Records <> @rows1)
BEGIN -- Jail records

IF(@debug='Y')
BEGIN
PRINT @filename+'1200 - no header; inconsistent column counts; jail inconsistent column counts'
END

--- Create table to jail the records in if it doesn't exist
IF NOT EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords')
BEGIN -- Create jailed records table

IF(@debug='Y')
BEGIN
PRINT '1220 - no header; Create table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords'
END

SET @command = 
'
CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords]([LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords] BIGINT IDENTITY(1,1) NOT NULL,[Filename] varchar(255) NULL,[ImportDate] smalldatetime NULL,RowNumber int NOT NULL,AllColumns varchar(max) NULL)
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)
END -- Create jailed records table

--- Insert jailed records
IF(@debug='Y')
BEGIN
PRINT @filepathname+'1230 - no header; inconsistent column counts; insert jailed records'
END

SET @command = 'INSERT INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords]([Filename],ImportDate,RowNumber,AllColumns)
SELECT
 CAST('''+CAST(@filename as varchar(255))+''' as varchar(1024)) as [Filename]
--,CAST(''' +LEFT(@filename,LEN(@filename)-CHARINDEX('.',REVERSE(@filename),1))+ ''' as varchar(255)) as Tablename
,CONVERT(smalldatetime,'+@quote+CONVERT(varchar(20),@datetime,100)+@quote+',100) as ImportDate
,LocalIndex,FileContentsIncludingHeader
FROM ['+@tableTempsci+']
WHERE LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 <> ' +CAST(@rows1 as varchar(10))+ '
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

--- Delete jailed records
IF(@debug='Y')
BEGIN
PRINT @filepathname+'1240 - no header; inconsistent column counts; delete jailed records'
END

SET @command = '
DELETE
FROM ['+@tableTempsci+']
WHERE LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 <> ' +CAST(@rows1 as varchar(10))+ '
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

----------------------------------------------------------
--- Move raw files to Processed folder and delete from load directory leaving cleaned files behind
IF(@debug='Y')
BEGIN
PRINT '1300 - no header; inconsistent column counts; move raw files to processed folder'
END

SET @commandDOS = 'XCOPY /YD "'+@filepathname+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

SET @commandDOS = 'DEL "'+@filepathname+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
END   -- Filename > 128 characters

----------------------------------------------------------
--- Export cleaned records with new filename and then load.
IF(@debug='Y')
BEGIN -- 
PRINT '1400 no header; inconsistent column counts; export cleaned data'
END   -- 

SET @commandDOS = 'bcp  "SELECT FileContentsIncludingHeader FROM tempdb..['+@tableTempsci+'] ORDER BY LocalIndex" queryout "'+@filepathname+'" /c /t "£££" /S "'+@server+'" -T -C "RAW"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
EXEC master..xp_cmdshell @commandDOS
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS,no_output
END

IF(@debug='Y')
BEGIN -- 
PRINT '1430 no header; inconsistent column counts; import cleaned data'
END   -- 

SET @command = 'BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]
FROM '''+@filepathname+'''
WITH (TABLOCK,CODEPAGE = ''RAW'',
	  FIELDTERMINATOR = '''+@ColDelineator+''',
	  ROWTERMINATOR = ''\n''
	 )
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

----------------------------------------------------------
--- Move cleaned files to Processed folder and delete from load directory 

IF(@debug='Y')
BEGIN -- 
PRINT '1500 - no header; inconsistent column counts; move cleaned files to processed folder'
END   -- 

SET @commandDOS = 'XCOPY /YD "'+@filepathname+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters


SET @commandDOS = 'DEL "'+@filepathname+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

----------------------------------------------------------
END -- Jail records

IF NOT EXISTS(select 1 from #TblColumnCounts where Records <> @rows1)
BEGIN -- consistent columns

IF(@debug='Y')
BEGIN
PRINT @filename+'1600 - consistent column counts'
PRINT @filepathname+'1630 - insert raw data'
END

SET @command = 
'BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]' +
' FROM '''+@filepathname+
''' WITH (TABLOCK,CODEPAGE = ''RAW'',FIELDTERMINATOR = '+@quote+@ColDelineator+@quote 
	+ ', ROWTERMINATOR = '+@quote+@RowTerminator+@quote+')'  
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec (@command)

END   -- consistent columns

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

IF NOT EXISTS(select 1 from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN -- Destination table doesn't exist

IF(@debug='Y')
BEGIN
PRINT ISNULL(@filename,'NUL')+'1700 - create final staging table'
END

SET @command = 
'
CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']([LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] BIGINT IDENTITY(1,1) NOT NULL,[Filename] varchar(255) NULL,[Tablename] varchar(255) NULL,[ImportDate] smalldatetime NULL)
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END

exec(@command)
END -- Destination table doesn't exist

END   -- No header

----------------------------------------------------------
----------------------------------------------------------
--- Create load table from header

IF(@header='YES')
BEGIN -- Header present

IF(@debug='Y')
BEGIN
PRINT '2000 - header present'
PRINT @filename+' ['+CAST(LEN(@filename) as varchar(10))+' characters]'
PRINT @filepathname+' ['+CAST(LEN(@filepathname) as varchar(10))+' characters]'
END

--- Take header to create tables

IF(@debug='Y')
BEGIN
PRINT '2005 - header present; insert as 1 column'
END

SET @command = '
TRUNCATE TABLE ['+@tableTempsc+']

BULK INSERT ['+@tableTempsc+']
FROM '''+@filepath+@filename+'''
WITH (TABLOCK,CODEPAGE=''RAW''
	 ,FIELDTERMINATOR = ''£££$$$£££'', ROWTERMINATOR = '''+@RowTerminator+'''
	 ,FIRSTROW='+@FirstRow+',LASTROW='+@FirstRow+'
	 )
'
IF(@debug='Y')
BEGIN
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec (@command)

TRUNCATE TABLE #TblImpAs1Col

SET @command = 'SELECT * FROM ['+@tableTempsc+']'

INSERT INTO #TblImpAs1Col
exec(@command)

IF(@debug='Y')
BEGIN
PRINT '2010 header present; Create table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'
END

IF(@ColDelineator IN(',','","','CSV')) -- converts @command = header from csv to pipe delineated
BEGIN --- CSV number of columns

IF(@debug='Y')
BEGIN
PRINT '2012 header present; CSV - convert header row to pipe'
END

SET @command = (SELECT TOP 1 Lookups.dbo.ufnConvertTextQualifiedCSVtoVBS(AllData) FROM #TblImpAs1Col)
IF(@debug='Y')
BEGIN
PRINT '2013 header present; CSV - count columns'
SELECT TOP 1 Lookups.dbo.ufnConvertTextQualifiedCSVtoVBS(AllData) FROM #TblImpAs1Col
END
SET @rows1 = (SELECT TOP 1 LEN(@command)-LEN(REPLACE(@command,'|','')))

END   --- CSV number of columns

IF(@ColDelineator NOT IN(',','","','CSV')) -- if not csv then leave @command = header as current delineator
BEGIN  --- Not-CSV number of columns

IF(@debug='Y')
BEGIN
PRINT '2015 header present; not CSV - count columns'
END

SET @command = (SELECT TOP 1 AllData FROM #TblImpAs1Col)

IF(@debug='Y')
BEGIN
SELECT TOP 1 AllData FROM #TblImpAs1Col
END
SET @rows1 = (SELECT TOP 1 ((LEN(@command)-LEN(REPLACE(@command,@ColDelineator,'')))/LEN(@ColDelineator)))

END   --- Not-CSV number of columns

PRINT 'Number of columns found = '+CAST(@rows1 +1 as varchar(10))
--- PROBLEM! When LOADS of columns it overflows max row length. 
--- If loads of columns then need to drop field width to fit onto a SQL page ...

IF(@debug='Y')
BEGIN
PRINT @filepathname+'2110 - header; create temporary table - work out max column width feasible'
END

SET @DefaultColumnWidth = CASE
							WHEN @DefaultColumnWidth > 0 THEN @DefaultColumnWidth
							WHEN @rows1 > 800 THEN 10
							WHEN @rows1 > 500 THEN 15
							WHEN @rows1 > 400 THEN 20
							WHEN @rows1 > 160 THEN 50
							WHEN @rows1 > 80 THEN 100
							ELSE 255
							END

----------------------------------------------

IF(@debug='Y')
BEGIN
PRINT @filepathname+'2120 - header; create temporary table'
END

DECLARE create1_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT LEFT(s.Item,128) as ColumnName,ItemNumber as Colid
	FROM Lookups.dbo.tFnDelimitedSplitMAX(@command,CASE WHEN @ColDelineator IN(',','","','CSV') THEN '|' ELSE @ColDelineator END) as s

OPEN create1_cursor 
        FETCH NEXT FROM create1_cursor into @column,@colid 
WHILE @@FETCH_STATUS = 0 
	BEGIN 

IF NOT EXISTS(select 1 from sys.objects where [name]=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp') AND (@colid=1)
BEGIN
SET @command = 'CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp](['+@column+'] varchar('+CAST(@DefaultColumnWidth as varchar(10))+') NULL)' 
IF(@debug='Y')
BEGIN
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC (@command) 
END
IF NOT EXISTS(select 1 from sys.objects where [name]=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version) AND (@colid=1)
BEGIN
SET @command = 'CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']([LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] BIGINT IDENTITY(1,1) NOT NULL,[Filename] varchar(255) NULL,[Tablename] varchar(255) NULL,[ImportDate] smalldatetime NULL,['+@column+'] varchar('+CAST(@DefaultColumnWidth as varchar(10))+') NULL
,CONSTRAINT [PK_'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] PRIMARY KEY CLUSTERED 
([LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ASC) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
' 
IF(@debug='Y')
BEGIN
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC (@command) 
END
IF NOT EXISTS(select 1 from sys.columns where [name]=@column and OBJECT_NAME(object_id)=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp') AND (@colid>1)
BEGIN
SET @command = 'ALTER TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp] ADD ['+@column+'] varchar('+CAST(@DefaultColumnWidth as varchar(10))+') NULL' 
IF(@debug='Y')
BEGIN
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC (@command) 
END
IF NOT EXISTS(select 1 from sys.columns where [name]=@column and OBJECT_NAME(object_id)=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version) AND (@colid>1)
BEGIN
SET @command = 'ALTER TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ADD ['+@column+'] varchar('+CAST(@DefaultColumnWidth as varchar(10))+') NULL' 
IF(@debug='Y')
BEGIN
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC (@command) 
END

---Close cursor 
        FETCH NEXT FROM create1_cursor INTO @column,@colid  
	END 
CLOSE create1_cursor 
DEALLOCATE create1_cursor 

----------------------------------------------
----------------------------------------------------------
----------------------------------------------------------
-- CSV but sql version > SQL 2017 (PMV = 14) and file RFC4180 compliant.

IF(@ColDelineator='CSV')
BEGIN -- CSV (RFC4180 coimpliuant)

IF(@debug='Y')
BEGIN
PRINT @filepathname+'2200 - LOAD CSV (RFC4180 coimpliuant)'
END

BEGIN TRY

SET @command = 
'BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]' +
' FROM '''+@filepathname+
''' WITH (TABLOCK,CODEPAGE = ''RAW'',FORMAT = '+@quote+@ColDelineator+@quote 
	+ ',FIRSTROW =  '+CAST(CAST(@FirstRow as int)+1 as varchar(10))+')'  
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec (@command)

--- Move files to processed folder
SET @commandDOS = 'XCOPY /YD "'+@filepathname+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters


SET @commandDOS = 'DEL "'+@filepathname+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

GOTO END_HEADER_PRESENT

END TRY

BEGIN CATCH

PRINT 'Error. File cannot conform to CSV (RFC4180 coimpliuant)'
GOTO HEADER_CHECK_DELINEATORS

END CATCH

END -- CSV (RFC4180 coimpliuant)

----------------------------------------------------------
--- Check number of delineators

HEADER_CHECK_DELINEATORS:

IF(@debug='Y')
BEGIN
PRINT '2800 header present; check no of columns present'
END

PRINT @filename

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

------------------------------------------
--- Import file as one column
IF(@debug='Y')
BEGIN
PRINT 'filepath = '+@filepath
PRINT 'filename = '+@filename
PRINT 'filepathname = '+@filepathname
END

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters

-----------------------------------------------------------
---Insert new data from file - column delineated by commas with newline for end of row and no text delineators - i.e. standard csv.

SET @command = '
TRUNCATE TABLE ['+@tableTempsc+']

BULK INSERT ['+@tableTempsc+']
FROM '''+@filepath+'file1.txt''
--FROM "'+@filepath+'file1.txt"
WITH (TABLOCK,CODEPAGE=''RAW''
	 ,FIELDTERMINATOR = ''£££'', ROWTERMINATOR = '''+@RowTerminator+'''
	 ,FIRSTROW='+@FirstRow+'
	 )
'
IF (@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

-----------------------------------------------------------
END -- Filename > 128 characters
IF(@filepathname NOT LIKE '%file1.txt')
BEGIN -- Filename <= 128 characters

-----------------------------------------------------------
---Insert new data from file - column delineated by commas with newline for end of row and no text delineators - i.e. standard csv.

SET @command = '
TRUNCATE TABLE ['+@tableTempsc+']

BULK INSERT ['+@tableTempsc+']
FROM '''+@filepath+@filename+'''
WITH (TABLOCK,CODEPAGE=''RAW''
	 ,FIELDTERMINATOR = ''£££'', ROWTERMINATOR = '''+@RowTerminator+'''
	 ,FIRSTROW='+@FirstRow+'
	 )
'
IF (@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

-----------------------------------------------------------
END   -- Filename <= 128 characters

------------------------------------------
--- Count delineators in file

SET @command = '
SELECT
 LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 as ColumnCount
,COUNT(*) as Records
FROM ['+@tabletempsc+']
WHERE FileContentsIncludingHeader IS NOT NULL
GROUP BY LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1
ORDER BY COUNT(*) DESC
'
IF(@debug='Y')
BEGIN
PRINT @command
END
INSERT INTO #TblColumnCounts
exec (@command)

------------------------------------------
--- Test to see if variable number of delineators

SET @rows1 = (SELECT TOP 1 ColumnCount FROM #TblColumnCounts ORDER BY Records DESC) -- number of columns in file
SET @rows2 = (SELECT COUNT(*) as Records FROM #TblColumnCounts) --- Number of lines with different column counts

--- Clear up temporary table
IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

IF(@debug='Y')
BEGIN
PRINT 'Number of lines with different column counts = '+CAST(@rows2 as varchar(10))
END

SET @command = 'TRUNCATE TABLE ['+@tabletempsc+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)

------------------------------------------

IF(@rows2 > 1)
BEGIN -- Begin cleaning

IF(@debug='Y')
BEGIN
PRINT '2900 header present - jail inconsistent number of columns'
END

--- Insert as one column

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters

SET @command = '
TRUNCATE TABLE ['+@tableTempsci+']
BULK INSERT ['+@tableTempsci+'] FROM '''+@filepath+'file1.txt'' WITH (TABLOCK,CODEPAGE=''RAW'',FORMATFILE='''+@filepath+@tableTempf+'.fmt'')'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec (@command)

END -- Filename > 128 characters
IF(@filepathname NOT LIKE '%file1.txt')
BEGIN -- Filename <= 128 characters

SET @command = '
TRUNCATE TABLE ['+@tableTempsci+']
BULK INSERT ['+@tableTempsci+'] FROM '''+@filepath+@filename+''' WITH (TABLOCK,CODEPAGE=''RAW'',FORMATFILE='''+@filepath+@tableTempf+'.fmt'')'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec (@command)

END   -- Filename <= 128 characters

----------------------------------------------------------
--- If CSV file and number of delineators is variable, try replacing delineator with pipe/vbs (|) and export as TXT file.

IF(@ColDelineator IN(',','","') AND @filename LIKE '%.csv')
BEGIN --- Convert CSV to VBS

IF(@debug='Y')
BEGIN
PRINT '3000 header present - CSV'
PRINT 'Convert file from CSV to VBS ...'
END

--- Drop table containing data using conversion function
IF(@debug='Y')
BEGIN
PRINT '3010 header present - CSV - drop table for conversion'
END
IF EXISTS(select [name] from sys.objects where [name] = RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log','')))
BEGIN
SET @command = 
'
DROP TABLE ['+RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log',''))+']
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)
END

IF(@debug='Y')
BEGIN
PRINT '3020 header present - CSV - create conversion table'
END
--- Create table containing data using conversion function
IF NOT EXISTS(select [name] from sys.objects where [name] = RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log','')))
BEGIN
SET @command = 
'
SELECT 
 Lookups.dbo.ufnConvertTextQualifiedCSVtoVBS(REPLACE(FileContentsIncludingHeader,''|'',''_'')) as FileContentsIncludingHeader
INTO ['+RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log',''))+']
FROM ['+@tableTempsci+']
ORDER BY LocalIndex 
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command) 
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

IF(@debug='Y')
BEGIN
PRINT '3100 header present - CSV - check column counts'
END

SET @command = '
INSERT INTO #TblColumnCounts
SELECT
 LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,''|'','''')) +1 as ColumnCount
,COUNT(*) as Records
FROM ['+RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log',''))+'] 
WHERE FileContentsIncludingHeader IS NOT NULL
GROUP BY LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,''|'','''')) +1
ORDER BY COUNT(*) DESC
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command) 

SET @rows1 = (SELECT COUNT(*) as Records FROM #TblColumnCounts) --- Number of lines with different column counts

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
TRUNCATE TABLE #TblColumnCounts
END

IF(@rows1 = 1) 
BEGIN -- Delieators now consistent

IF(@debug='Y')
BEGIN
PRINT '3200 header present - CSV consistent number of columns (cleaned data) - export data'
END

--- Export view. (BCP OUT!)
SET @commandDOS = 'bcp "['+@db+'].dbo.['+@quote+RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log',''))+@quote+']" out "'+@filepath+@table+'.txt" -c -t "|" /S "'+@server+'" -T'

IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
--exec(@commandDOS)
exec xp_cmdshell @commandDOS

--- Clear up VBS temporary table

SET @command = 
'
DROP TABLE ['+RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(@filename,'.csv',''),'.txt',''),'.vbs',''),'.log',''))+']
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

----------------------------------------------------------
--- Move .csv files to Processed folder and delete from load directory leaving .txt files behind

IF(@debug='Y')
BEGIN
PRINT '3500 header present - CSV - consistent number of columns (cleaned data) - move raw files'
END

SET @commandDOS = 'XCOPY /YD "'+@filepathname+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters


SET @commandDOS = 'DEL "'+@filepathname+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

----------------------------------------------------------
--- load TXT file

IF(@debug='Y')
BEGIN
PRINT '3600 header present - CSV - load cleaned files'
END
SET @filename = REPLACE(@filename,'.csv','.txt')
IF(@debug='Y')
BEGIN
PRINT @filename
END

----------------------------------------------------------

SET @FirstRow = '1' -- Now 1 as header stripped during BCP export

IF(@debug='Y')
BEGIN
PRINT @commandDOS
exec xp_cmdshell @commandDOS--,no_output
exec Lookups.dbo.spDOS_DirFileListOnly @filepath
END
IF(@debug='N')
BEGIN
exec xp_cmdshell @commandDOS,no_output
END

----------------------------------------------------------

IF(@debug='Y')
BEGIN
PRINT '3660 header present - CSV - load cleaned files into temp table'
END

SET @command = 'BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]
FROM '''+@filepath+@filename+'''
WITH (TABLOCK,CODEPAGE = ''RAW'',
	  FIELDTERMINATOR = ''|'',
	  ROWTERMINATOR = ''\n'',
	  FIRSTROW = 1
	 )
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

END  -- Delieators now consistent

END --- Convert CSV to VBS

----------------------------------------------------------
--- Not CSV but inconsistent delineators so jail dodgy records or converted CSV that is still inconsistent

--- Jail records with different number of delineators
IF(((@rows2 > 1) --- inconsitent delineators in raw file
		AND NOT(@ColDelineator IN(',','","','CSV') AND @filename LIKE '%.csv')) -- and not CSV as already handled that
	OR @rows1 > 1) -- CSV conversion still inconsistent
BEGIN -- Jail records header

IF(@debug='Y')
BEGIN
PRINT '3700 header present - not CSV - inconsistent delineators'
END

--- Create table to jail the records in if it doesn't exist
IF NOT EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords')
BEGIN -- Create jailed records table
IF(@debug='Y')
BEGIN
PRINT '3710 header present - not CSV - inconsistent delineators - create jail table'
PRINT 'Create table '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords'
END

SET @command = 
'
CREATE TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords]([LocalIndex'+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords] BIGINT IDENTITY(1,1) NOT NULL,[Filename] varchar(255) NULL,[ImportDate] smalldatetime NULL,RowNumber int NOT NULL,AllColumns varchar(max) NULL)
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)
END -- Create jailed records table

IF(@debug='Y')
BEGIN
PRINT '3720 header present - not CSV - inconsistent delineators - insert into jail table'
END

--- Insert jailed records header
SET @command = '
INSERT INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_JailedRecords]([Filename],ImportDate,RowNumber,AllColumns)
SELECT
 CAST('''+@filename+''' as varchar(1024)) as [Filename]
--,CAST('''+REPLACE(REPLACE(@filename,'.csv',''),'.txt','')+''' as varchar(255)) as Tablename
,CONVERT(smalldatetime,'+@quote+CONVERT(varchar(20),@datetime,100)+@quote+',100) as ImportDate
,LocalIndex,FileContentsIncludingHeader
FROM ['+@tableTempsci+']
WHERE LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 <> ' +CAST(@rows1 as varchar(10))+ '
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

IF(@debug='Y')
BEGIN
PRINT '3730 header present - not CSV - inconsistent delineators - delete jailed records from table'
END

--- Delete jailed records header
SET @command = '
DELETE
FROM ['+@tableTempsci+']
WHERE LEN(FileContentsIncludingHeader) - LEN(REPLACE(FileContentsIncludingHeader,'''+@ColDelineator+''','''')) +1 <> ' +CAST(@rows1 as varchar(10))+ '
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

----------------------------------------------------------
--- Export cleaned records with new filename and then load.

IF(@debug='Y')
BEGIN
PRINT '3780 header present - not CSV - inconsistent delineators - export cleaned records'
END

SET @commandDOS = 'BCP "SELECT FileContentsIncludingHeader FROM tempdb..['+@tableTempsci+'] ORDER BY LocalIndex" queryout "'+@filepathname+'" /c /t "£££" /S "'+@server+'" -T -C "RAW"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
EXEC master..xp_cmdshell @commandDOS--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS,no_output
END

----------------------------------------------------------
--- Load into staging table 

IF(@debug='Y')
BEGIN
PRINT '3800 header present - not CSV - inconsistent delineators - insert cleaned records into temporary table'
END

SET @command = 'BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]
FROM '''+@filepathname+'''
WITH (TABLOCK,CODEPAGE = ''RAW'',
	  FIELDTERMINATOR = ''|'',
	  ROWTERMINATOR = ''\n''
	 )
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

----------------------------------------------------------
--- Move files to Processed folder and delete from load directory 
IF(@debug='Y')
BEGIN
PRINT '3900 header present - not CSV - inconsistent delineators - move cleaned files'
END

SET @commandDOS = 'XCOPY /YD "'+@filepathname+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+@filepath+@filename+'"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

END -- Jail records header

END -- Begin cleaning

----------------------------------------------------------
---- Load raw file

IF(@rows2 = 1) --@rows2 = Number of lines with different column counts
BEGIN -- uncleaned load

IF(@debug='Y')
BEGIN
PRINT '4000 header present - Not CSV - uncleaned raw data load'
END

IF(@debug='Y')
BEGIN
PRINT '4010 header present - Not CSV - truncate temp table'
END

--- Clear out old data
SET @command = 'TRUNCATE TABLE [dbo].['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec (@command)

IF(@debug='Y')
BEGIN
PRINT '4020 header present - Not CSV - uncleaned raw data load insert'
END

-- load original file
SET @command = 
'
BULK INSERT ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]
FROM '''+@filepathname+'''
WITH (CODEPAGE = '+@quote+'RAW'+@quote 
	+ ', FIELDTERMINATOR = '+@quote+@ColDelineatorI+@quote 
	+ ', ROWTERMINATOR = '+@quote+@RowTerminator+@quote 
	+ ', FIRSTROW =  '+CAST(CAST(@FirstRow as int)+1 as varchar(10))+')'  
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

------------------------------------------
--- Move file to Loaded folder and delete from load directory
IF(@debug='Y')
BEGIN
PRINT '4100 header present - Not CSV - move uncleaned raw files to processed folder'
END

SET @commandDOS = 'XCOPY /YD "'+RTRIM(@filepathname)+'" "'+@filepath+'Processed\"'
IF(@debug='Y')
BEGIN
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

SET @commandDOS = 'XCOPY /YD "'+RTRIM(REPLACE(@filepathname,'.csv','.txt'))+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'XCOPY /YD "'+@filepath+@filename+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
SET @commandDOS = 'XCOPY /YD "'+@filepath+RTRIM(REPLACE(@filename,'.csv','.txt'))+'" "'+@filepath+'Processed\"'
exec master..xp_cmdshell @commandDOS
END   -- Filename > 128 characters

SET @commandDOS = 'DEL "'+RTRIM(@filepathname)+'"'
IF(@debug='Y')
BEGIN
PRINT ''
SET @message = @commandDOS

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec master..xp_cmdshell @commandDOS

SET @commandDOS = 'DEL "'+RTRIM(REPLACE(@filepathname,'.csv','.txt'))+'"'
exec master..xp_cmdshell @commandDOS

IF(@filepathname LIKE '%file1.txt')
BEGIN -- Filename > 128 characters
SET @commandDOS = 'DEL "'+RTRIM(@filepath+@filename)+'"'
exec master..xp_cmdshell @commandDOS

SET @commandDOS = 'DEL "'+@filepath+RTRIM(REPLACE(@filename,'.csv','.txt'))+'"'
exec master..xp_cmdshell @commandDOS

END  -- Filename > 128 characters

------------------------------------------

END -- uncleaned load

----------------------------------------------------------
-- Jump here from CSV but sql version > SQL 2017 (PMV = 14) and file RFC4180 compliant.

END_HEADER_PRESENT:
IF(@debug='Y')
BEGIN
PRINT 'END OF HEADER PRESENT SECTION'
END

----------------------------------------------------------

END -- Header present

----------------------------------------------------------
----------------------------------------------------------
--- Load data into final staging table

BEGIN --- load

IF(@debug='Y')
BEGIN
PRINT '5000 - Load from temp table to final staging table'
END

----------------------------------------------------------

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblCols'))
BEGIN
DROP TABLE #TblCols
END

------------------------------------------
--- If destination table already exists then delete out previous data if filename matches 

PRINT ''
PRINT ''
PRINT ''
PRINT 'Delete data for same filename from destination table.'
PRINT @filepathname
PRINT REPLACE(REPLACE(@filename,'.csv',''),'.txt','')

IF EXISTS (SELECT [name] FROM dbo.sysobjects WHERE [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN -- Import table exists so delete
SET @command = '
DELETE
FROM dbo.['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']
WHERE [Filename] = '+@quote+@filename+@quote+'
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command)
END -- Import table exists so delete

------------------------------------------
--- If destination table exists then check columns are the same and add data in.

SET @columnchanges = ''

----------------------------
--- Discover any missing columns

SET @rows2 = 0

select *
INTO #TblCols
FROM (
select
name as ColumnName
from sys.columns
where OBJECT_NAME(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'

except

select
name
from sys.columns
where OBJECT_NAME(object_id) = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version
) as c

SET @rows2 = @@ROWCOUNT

IF @rows2 <> 0
BEGIN -- New columns alert
SET @columnchanges = @columnchanges+'
'+@filename+' '+REPLACE(REPLACE(@filename,'.csv',''),'.txt','')+' has '+CAST(@rows2 as varchar(10))+' new columns.
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @columnchanges

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END

END -- New columns alert

IF (@columnchanges <> '')
BEGIN -- Column changes e-mail
SET @email_bodytext = @email_bodytext+'
'+ISNULL(@columnchanges,'')
END -- column changes e-mail

----------------------------
--- add in any missing columns

IF(SELECT COUNT(*) FROM #TblCols) > 0
BEGIN -- Missing columns

IF(@debug='Y')
BEGIN
PRINT '5100 - Add missing columns'
END

DECLARE column_cursor CURSOR FOR 
select 
ColumnName
from #TblCols
   
OPEN column_cursor 
        FETCH NEXT FROM column_cursor into @column 
WHILE @@FETCH_STATUS = 0 
        BEGIN 

SET @command = 'ALTER TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'] ADD ['+@column+'] varchar('+CAST(@DefaultColumnWidth as varchar(4))+') NULL'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
EXEC(@command) 

---Close column cursor 

        FETCH NEXT FROM column_cursor INTO @column 
END 
CLOSE column_cursor 
DEALLOCATE column_cursor 

END -- Missing columns

----------------------------
--- if destination table exists then insert data where columns are the same

IF EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN -- Insert into import table

IF(@debug='Y')
BEGIN
PRINT '5200 - insert data'
END

SET @command = 
(SELECT 
 SUBSTRING(td.InsertDef,1,LEN(td.InsertDef)-1)+')
 '+SUBSTRING(td.SelectDef,1,LEN(td.SelectDef)-1)+'
 FROM ['+ObjectName+']'
FROM (select 
	 OBJECT_NAME(object_id) as ObjectName
---INSERT
	,(select
	 CASE 
		WHEN sc3.column_id = 1 
			THEN ' INSERT INTO ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']([Filename],[Tablename],[ImportDate],['+sc3.[name]+'],'
		ELSE '['+sc3.[name]+'],'
		END
	from sys.objects so4
		inner join sys.columns sc3
			on so4.object_id = sc3.object_id
			--and is_identity = 0
		inner join (select
					 so3.object_id
					,so3.[name] as ObjectName
					,sc2.[name] as ColumnName
					from sys.objects so3
						inner join sys.columns sc2
							on so3.object_id = sc2.object_id
					where so3.type = 'U') src1
			on so4.[name] = src1.ObjectName
			and sc3.[name] = src1.ColumnName
	where so4.type = 'U'
		AND so4.object_id = so1.object_id
	order by sc3.column_id
		FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS InsertDef
---SELECT
	,(select
	 CASE 
		WHEN sc3.column_id = 1 
			THEN ' SELECT CAST('''+@filename+''' as varchar(1024)) as [Filename] ,CAST('''+REPLACE(REPLACE(@filename,'.csv',''),'.txt','')+''' as varchar(255)) as Tablename ,CONVERT(smalldatetime,'+@quote+CONVERT(varchar(20),@datetime,100)+@quote+',100) as ImportDate ,CAST(REPLACE(['+sc3.[name]+'],''"'','''') as varchar('+CAST(@DefaultColumnWidth as varchar(4))+')) as ['+sc3.[name]+'],'
		ELSE 'CAST(REPLACE(['+sc3.[name]+'],''"'','''') as varchar('+CAST(@DefaultColumnWidth as varchar(4))+')) as ['+sc3.[name]+'],'
		END
	from sys.objects so4
		inner join sys.columns sc3
			on so4.object_id = sc3.object_id
			--and is_identity = 0
		inner join (select
					 so3.object_id
					,so3.[name] as ObjectName
					,sc2.[name] as ColumnName
					from sys.objects so3
						inner join sys.columns sc2
							on so3.object_id = sc2.object_id
					where so3.type = 'U') src1
			on so4.[name] = src1.ObjectName
			and sc3.[name] = src1.ColumnName
	where so4.type = 'U'
		AND so4.object_id = so1.object_id
	order by sc3.column_id
		FOR XML PATH(''),TYPE).value('(./text())[1]','VARCHAR(MAX)') AS SelectDef
	from sys.objects so1
	WHERE so1.type = 'U'
		and so1.[name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'
	group by so1.object_id) as td)
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

SET @rows1 = @@ROWCOUNT
IF(@rows1>0)
BEGIN
SET @FilesLoadedList = ISNULL(@FilesLoadedList+'; ','')+@filename
END

SET @email_bodytext = @email_bodytext+'

'+CAST(@rows1 as varchar(10))+' rows inserted into table ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'].'

END -- Insert into import table

END -- LOAD

------------------------------------------
--- Drop temporary load tables

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblCols'))
BEGIN
DROP TABLE #TblCols
END

IF EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp')
BEGIN -- Drop temporary table
PRINT 'DROP TABLE '+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp'
SET @command = 
'
DROP TABLE ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'_Temp]
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)
END -- Drop temporary table

----------------------------------------------------------
--- Update log file where successful

UPDATE IMP
SET
 ImportEndDateTime = GETDATE()
,ImportDurationMinutes = ROUND(CAST(DATEDIFF(SECOND,ImportStartDateTime,getdate()) as float) / CAST(60.00 as float),0)
,Records = @rows1
,FlagNewColumnsReceived = CASE WHEN @columnchanges <> '' THEN 'Y' ELSE 'N' END
,Comments = 'Success. '+CAST(@rows1 as varchar(10))+' rows inserted into table ['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+'].'+ISNULL(' '+Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [Tablename] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND convert(varchar(20),[ImportStartDateTime],100) = convert(varchar(20),@datetime,100)

----------------------------------------------------------------------------
--- Remove headings from import table into seperate table as this can cause conversion issues

--exec uspSubRTN0600RemoveHeadingsFromDataImport @table=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version,@debug=@debug

----------------------------------------------------------------------------
---End of try and catch
END TRY --- Import file into table

---Catch error
BEGIN CATCH

	PRINT ''
	PRINT ''
	PRINT ''
	PRINT 'Error Detected'

SET @error =
	  (SELECT 
	/*ERROR_NUMBER() ERNumber,
	ERROR_SEVERITY() Error_Severity,
	ERROR_STATE() Error_State,
	ERROR_PROCEDURE() Error_Procedure,
	ERROR_LINE() Error_Line,*/
	ERROR_MESSAGE() Error_Message)

PRINT @error

    -- Test whether the transaction is uncommittable.
    IF (XACT_STATE()) = -1
    BEGIN
    PRINT 'The transaction is in an uncommittable state. ' +
            'Rolling back transaction.'
        ROLLBACK TRANSACTION;
    END;

    -- Test whether the transaction is active and valid.
    IF (XACT_STATE()) = 1
    BEGIN
        PRINT 'The transaction is committable. ' +
            'Committing transaction.'
        COMMIT TRANSACTION;   
    END;

SET @email_bodytext = @email_bodytext+'

Error: '+@error+'
'

----------------------------

PRINT ''
PRINT ''
PRINT ''
PRINT 'Update log file where unsuccessful.'
PRINT @filename
PRINT @error

UPDATE IMP
SET 
 ImportEndDateTime = GETDATE() 
,Comments = 'Error'+isnull(' = '+@error,'?') + ISNULL(' '+Comments,'')
FROM [dbo].[TblLog_Imports] IMP
WHERE [Tablename] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version
	AND [SourcePath] = @filepath
	AND [SourceName] = @filename
	AND convert(varchar(20),[ImportStartDateTime],100) = convert(varchar(20),@datetime,100)

END CATCH


SET @email_bodytext = @email_bodytext+'
----------------------------------------------------------------------------

'

----------------------------------------------------------
--- Loop cursor for as many files as there are

        FETCH NEXT FROM file_cursor INTO @filename,@filemodifieddate,@filemodifieddateP 
---Close second cursor 
END 
CLOSE file_cursor 
DEALLOCATE file_cursor 

----------------------------------------------------------------------------
--- Clear up temporary tables

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblImpAs1Col'))
BEGIN
DROP TABLE #TblImpAs1Col
END

----------------------------------------------------------------------------
--- Work out failed file loads or tables with no data ...
--- Replaced the csv with txt for where files cleaned but it fails where they aren't cleaned. BUGGER.

IF EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN

SELECT
 RTRIM(SUBSTRING(DirList,37,255)) as [Filename] 
INTO dbo.zDirList
FROM #TblTempDirList DL
WHERE DirList NOT LIKE '%DirectoryListing.txt%'
	AND CASE
		WHEN DirList LIKE '%.txt%' THEN 1
		WHEN DirList LIKE '%.csv%' THEN 1
		WHEN DirList LIKE '%.vbs%' THEN 1
		ELSE 0
		END = 1

SET @command = 
'
SELECT
 [Filename] 
FROM [' + @db + '].dbo.zDirList

EXCEPT

SELECT
 [Filename]
FROM ['+@db+'].dbo.['+REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version+']
GROUP BY [Filename]
'
IF(@debug='Y')
BEGIN
PRINT ''
PRINT ''
PRINT ''
SET @message = @command

-- RAISERROR with severity 11-19 will cause execution to jump to the CATCH block.
	RAISERROR (@message, -- Message text.
               9, -- Severity.
               1 -- State.
               )  WITH NOWAIT;
END
exec(@command)

SET @email_query = @command

END

IF NOT EXISTS(select [name] from sys.objects where [name] = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@usr,'''',''),'\','_'),'/','_'),':','_'),' ','_'),'-','_'),'$',''),'.','_')+@prefix+@table+@suffix+@version)
BEGIN
SET @email_query = 'SELECT ''Table not found!'' as Comment'
END

----------------------------------------------------------------------------
--- E-mail results

IF(@email_bodytext <> 'Results of last import using spSubLoadAllUserTextFiles ('+@usr+'):')
BEGIN
EXEC msdb.dbo.sp_send_dbmail 
	 @profile_name = @email_profilename					--When no profile_name is specified, sp_send_dbmail uses the default private profile for the current user. If the user does not have a default private profile, sp_send_dbmail uses the default public profile for the msdb database. If the user does not have a default private profile and there is no default public profile for the database, @profile_name must be specified.
	,@recipients = @email_recipients					-- Is a semicolon-delimited list of e-mail addresses to send the message to.
	--,@copy_recipients = 'copy_recipient [ ; ...n ]'	-- Is a semicolon-delimited list of e-mail addresses to carbon copy the message to. 
	--,@blind_copy_recipients = 'blind_copy_recipient [ ; ...n ]' --Is a semicolon-delimited list of e-mail addresses to blind carbon copy the message to. 
	--,@from_address = 'from_address'					-- This is an optional parameter used to override the settings in the mail profile. This parameter is of type varchar(MAX). SMTP security settings determine if these overrides are accepted. If no parameter is specified, the default is NULL.
	--,@reply_to = 'reply_to'							-- This is an optional parameter used to override the settings in the mail profile. This parameter is of type varchar(MAX). SMTP security settings determine if these overrides are accepted. If no parameter is specified, the default is NULL.
	,@subject = @email_subjectline						-- The subject is of type nvarchar(255). If no subject is specified, the default is 'SQL Server Message'.
	,@body = @email_bodytext							-- Is the body of the e-mail message. The message body is of type nvarchar(max), with a default of NULL.
	--,@body_format = 'HTML'							-- Is the format of the message body. The parameter is of type varchar(20), with a default of NULL. TEXT or HTML.
	--,@importance = 'Normal'							-- The parameter is of type varchar(6). The parameter may contain one of the following values: Low; Normal; High.
	--,@sensitivity = 'Normal'							-- The parameter is of type varchar(12). The parameter may contain one of the following values: Normal; Personal; Private; Confidential.
	--,@file_attachments = 'attachment [ ; ...n ]'		-- Is a semicolon-delimited list of file names to attach to the e-mail message. Files in the list must be specified as absolute paths. The attachments list is of type nvarchar(max). By default, Database Mail limits file attachments to 1 MB per file.
	,@query = @email_query								--  Is a query to execute. The results of the query can be attached as a file, or included in the body of the e-mail message. The query is of type nvarchar(max), and can contain any valid Transact-SQL statements. Note that the query is executed in a separate session, so local variables in the script calling sp_send_dbmail are not available to the query.
	--,@execute_query_database = 'db_name'				-- SELECT DB_NAME() -- Is the database context within which the stored procedure runs the query. 
	,@attach_query_result_as_file = 1					--1 = as file attachment. Specifies whether the result set of the query is returned as an attached file. attach_query_result_as_file is of type bit, with a default of 0.
	,@query_attachment_filename = @email_AttachFilename -- Specifies the file name to use for the result set of the query attachment. query_attachment_filename is of type nvarchar(255), with a default of NULL. This parameter is ignored when attach_query_result is 0. When attach_query_result is 1 and this parameter is NULL, Database Mail creates an arbitrary filename.
	,@query_result_header = 1							--Default = 1 = included. Specifies whether the query results include column headers. The query_result_header value is of type bit. When the value is 1, query results contain column headers. When the value is 0, query results do not include column headers. This parameter defaults to 1. This parameter is only applicable if @query is specified.
	,@query_result_width = 8000							-- Without it it starts wrapping text onto new lines!
	,@query_result_separator = @email_ColumnDelineator	-- Is the character used to separate columns in the query output. The separator is of type char(1). Defaults to ' ' (space).
	,@exclude_query_output = 0							-- Specifies whether to return the output of the query execution in the e-mail message. exclude_query_output is bit, with a default of 0. When this parameter is 0, the execution of the sp_send_dbmail stored procedure prints the message returned as the result of the query execution on the console. When this parameter is 1, the execution of the sp_send_dbmail stored procedure does not print any of the query execution messages on the console.
	,@append_query_error = 1							-- Specifies whether to send the e-mail when an error returns from the query specified in the @query argument. append_query_error is bit, with a default of 0. When this parameter is 1, Database Mail sends the e-mail message and includes the query error message in the body of the e-mail message. When this parameter is 0, Database Mail does not send the e-mail message, and sp_send_dbmail ends with return code 1, indicating failure.
	--,@query_no_truncate = 1							-- Specifies whether to execute the query with the option that avoids truncation of large variable length data types (varchar(max), nvarchar(max), varbinary(max), xml, text, ntext, image, and user-defined data types). When set, query results do not include column headers. The query_no_truncate value is of type bit. When the value is 0 or not specified, columns in the query truncate to 256 characters. When the value is 1, columns in the query are not truncated. This parameter defaults to 0.
	--,@query_result_no_padding = 1 --No padding			-- The type is bit. The default is 0. When you set to 1, the query results are not padded, possibly reducing the file size.If you set @query_result_no_padding to 1 and you set the @query_result_width parameter, the @query_result_no_padding parameter overwrites the @query_result_width parameter. If you set the @query_result_no_padding to 1 and you set the @query_no_truncate parameter, an error is raised.
	--,@mailitem_id =  mailitem_id ] [ OUTPUT ]			-- Optional output parameter returns the mailitem_id of the message. The mailitem_id is of type int.
	; 
END

----------------------------------------------------------------------------

SET @comment = @comment + @usr+': OK. '

---------------------------------------------------------------

END TRY --- FOLDER

---Catch error (if any)
BEGIN CATCH
  PRINT 'Error detected'
SET @error =
	  (SELECT 
	--ISNULL(CAST(ERROR_NUMBER() as varchar(1000)),'')+', '+
	--ISNULL(CAST(ERROR_SEVERITY() as varchar(1000)),'')+', '+
	--ISNULL(CAST(ERROR_STATE() as varchar(1000)),'')+', '+
	--ISNULL(CAST(ERROR_PROCEDURE() as varchar(1000)),'')+', '+
	ISNULL(CAST(ERROR_LINE() as varchar(1000)),'')+', '+
	ISNULL(CAST(ERROR_MESSAGE() as varchar(7000)),'')
	)
	
IF(@debug='Y')
BEGIN
PRINT @error
END

SET @comment = @comment + @usr+': '+@error

END CATCH

---------------------------------------------------------------
--END -- Directory exists - see line 333.

IF EXISTS (SELECT [name] FROM sys.objects WHERE [name] = N'zDirList')
BEGIN
DROP TABLE dbo.zDirList
END

-----------------------------------------------------------
--- Delete format file

SET @commandDOS = 'DEL "'+@filepath+@tableTempf+'.fmt"'
IF(@debug='Y')
BEGIN
PRINT @commandDOS
EXEC master..xp_cmdshell @commandDOS
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandDOS,no_output
END

IF(@debug='Y')
BEGIN
PRINT @comment
END

---------------------------------------------------------------
---Close cursor 

        FETCH NEXT FROM user_cursor INTO @usr,@folder
END 
CLOSE user_cursor 
DEALLOCATE user_cursor 

---------------------------------------------------------------
--- Drop temporary tables

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempf))
BEGIN
SET @command = '
DROP TABLE ['+@tableTempf+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..'+@tableTempsci))
BEGIN
SET @command = 'DROP TABLE ['+@tableTempsci+']'
IF(@debug='Y')
BEGIN
PRINT @command
END
exec(@command)
END

IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N'tempdb..#TblColumnCounts'))
BEGIN
DROP TABLE #TblColumnCounts
END

IF EXISTS (SELECT [name] FROM tempdb.sys.objects WHERE [object_id] = OBJECT_ID(N'tempdb..#TblTempDirList'))
BEGIN
DROP TABLE #TblTempDirList
END

---------------------------------------------------------------

IF(@debug='Y')
BEGIN
PRINT @comment
END

---------------------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[uspImportAllUserFiles]    Script Date: 17/03/2023 15:58:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[uspImportAllUserFiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[uspImportAllUserFiles] AS' 
END
GO
ALTER PROCEDURE [dbo].[uspImportAllUserFiles] (@usr varchar(128)='',@prefix varchar(14)='_',@table varchar(95)='',@suffix varchar(14)='_Import',@version varchar(5)='',@debug char(1)='N') AS

-- exec uspImportAllUserFiles @usr='DOMAIN\USERNAME',@debug='Y'

---------------------------------------------------------------
--- Check the input for SQL Injection and simply exit to deny an attacker any hints. Thanks to Jeff Moden. :)

IF @debug LIKE '%[^N,Y]' ESCAPE '_'
RETURN
IF @prefix LIKE '%[^-__A-Za-z0-9. ]%' ESCAPE '_'
RETURN
IF @suffix LIKE '%[^-__A-Za-z0-9. ]%' ESCAPE '_'
RETURN
IF @table LIKE '%[^-__A-Za-z0-9. ]%' ESCAPE '_'
RETURN
IF @usr LIKE '%[^-__A-Za-z0-9:\. ]%' ESCAPE '_'
RETURN
IF @version LIKE '%[^-__A-Za-z0-9:\. ]%' ESCAPE '_'
RETURN

------------------------------------------------------------
---Declare variables
/*
DECLARE @debug				char(1)				SET @debug = 'Y'		-- Prints various information to allow troubleshooting.
DECLARE @prefix				varchar(5)			SET @prefix  = '_'		-- table name prefix after username and before core table name.
DECLARE @suffix				varchar(14)			SET @suffix = '_Import'	-- table name suffix before version and after core table name.
DECLARE @table				varchar(95)			SET @table  = ''		-- Intended table core name (excluding prefix, suffix and version) to load into. If blank then uses filename.
DECLARE @usr				varchar(128)		SET @usr = SUSER_NAME() -- Username but not required here.
DECLARE @version			varchar(5)			SET @version = ''
*/
IF(@debug='N')
BEGIN
SET NOCOUNT ON
END

---------------------------------------------------------------------------------

PRINT 'OK'

exec dbo.spSubConvertAllUserTextFiles @usrI=@usr,@debug=@debug
exec dbo.spSubLoadAllUserTextFiles @usrI=@usr,@prefix=@prefix,@tableI=@table,@suffix=@suffix,@version=@version,@debug=@debug
exec dbo.spSubLoadAllUserMicrosoftOfficeFiles @usrI=@usr,@prefix=@prefix,@tableI=@table,@suffix=@suffix,@version=@version,@debug=@debug
GO
