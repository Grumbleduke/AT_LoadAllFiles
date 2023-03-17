USE [Lookups]
GO
/****** Object:  UserDefinedFunction [dbo].[ufnConvertTextQualifiedCSVtoVBS]    Script Date: 17/03/2023 16:11:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ufnConvertTextQualifiedCSVtoVBS]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'CREATE FUNCTION [dbo].[ufnConvertTextQualifiedCSVtoVBS](@Text as varchar(max)) RETURNS varchar(max) as

/*

AUTHOR: JCEH

Intellectual rights of JCEH. Not to be reused without permission.

VERSION CONTROL

Modified	Modifyee	Modification
-----------	-----------	--------------------------------------------------

25-Jun-2013	JCEH		Created function.

-----------	-----------	--------------------------------------------------
*/

BEGIN
--DECLARE @text varchar(max) SET @text = (SELECT TOP 1 * FROM Lookups.dbo.TblTempImportAsSingleColumn)
DECLARE @Reset				bit
DECLARE @Ret				varchar(max)
DECLARE @PositionCharacter	int
DECLARE @FlagDelineator		int
DECLARE @FlagTextQualifier	int
DECLARE @character			char(1)

SELECT @Reset = 1, @PositionCharacter=1, @FlagDelineator=0, @FlagTextQualifier=-1, @Ret = ''''

WHILE @PositionCharacter <= LEN(@Text)
SELECT 
--- Scroll through string one character at a time
 @character = SUBSTRING(@Text,@PositionCharacter,1) 
--- Check if character is column delineator and set flag to 1 if it is, otherwise set flag to 0
,@FlagDelineator = CASE 
		WHEN @character = '','' 
			THEN 1
		ELSE 0
		END
--- Check for text qualifier. 1 indicates text qualifier open. -1 indicates closed.
,@FlagTextQualifier = CASE
		WHEN @character = ''"''
			THEN @FlagTextQualifier * (-1)
		ELSE @FlagTextQualifier * (+1)
		END
--- Decide whether circumstances are right to change column delineator character
,@Reset = CASE 
			WHEN @FlagDelineator = 1 AND @FlagTextQualifier = -1 --- Column delineator found and qualifier closed
				THEN 0
			ELSE 1
			END
--- Decide whether to change 
,@Ret = @Ret + CASE 
				WHEN @Reset=0 -- If delineator found change to pipe (|)
					THEN ''|'' 
				WHEN @character = ''"'' -- If text qualifier found then remove (change to zero length string
					THEN ''''
				ELSE @character 
				END 
--- Increment character position by 1
,@PositionCharacter = @PositionCharacter +1
RETURN @Ret 
END
' 
END
GO
/****** Object:  UserDefinedFunction [dbo].[ufnConvertTextQualifiedDelimiter]    Script Date: 17/03/2023 16:11:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ufnConvertTextQualifiedDelimiter]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'CREATE FUNCTION [dbo].[ufnConvertTextQualifiedDelimiter](@Text				varchar(max)
														,@ColumnDelineator	varchar(5)='',''
														,@TextQualifier		varchar(5)=''"''
														,@ColumnDelineatorI varchar(5)=''‡''
														,@RowTerminator		varchar(5)=''\n''
														) RETURNS varchar(max) as

/*

AUTHOR: JCEH

Intellectual rights of JCEH. Not to be reused without permission.

VERSION CONTROL

Modified	Modifyee	Modification
-----------	-----------	--------------------------------------------------

25-Jun-2013	JCEH		Created function.

-----------	-----------	--------------------------------------------------
*/

BEGIN
--DECLARE @text varchar(max) SET @text = (SELECT TOP 1 * FROM Lookups.dbo.TblTempImportAsSingleColumn)
DECLARE @Reset				bit = 1
DECLARE @Ret				varchar(max)
DECLARE @PositionCharacter	int = 1
DECLARE @FlagDelineator		int = 0
DECLARE @FlagTextQualifier	int = -1
DECLARE @ColumnCount		int = 0
DECLARE @ColumnCountTotal	int = 9999
DECLARE @character			char(1)

--SET @Text=REPLACE(@Text,@ColumnDelineatorI,'' '')
SELECT @Reset = 1, @PositionCharacter=1, @FlagDelineator=0, @FlagTextQualifier=-1, @Ret = ''''--,@Text=REPLACE(@Text,@ColumnDelineatorI,'' '')

WHILE @PositionCharacter <= LEN(@Text)
SELECT 
--- Scroll through string one character at a time
 @character = SUBSTRING(@Text,@PositionCharacter,1) 
--- Check if character is column delineator and set flag to 1 if it is, otherwise set flag to 0
,@FlagDelineator = CASE 
					WHEN @character = @ColumnDelineator
						THEN 1
					ELSE 0
					END
--- Check for text qualifier. 1 indicates text qualifier open. -1 indicates closed.
,@FlagTextQualifier = CASE
						WHEN @character = @TextQualifier
							THEN @FlagTextQualifier * (-1)
						ELSE @FlagTextQualifier * (+1)
						END
--- Decide whether circumstances are right to change column delineator character
,@Reset = CASE 
			WHEN @FlagDelineator = 1 AND @FlagTextQualifier = -1 --- Column delineator found and qualifier closed
				THEN 0
			ELSE 1
			END
--- Decide whether to change 
,@Ret = @Ret + CASE 
				WHEN @Reset=0 -- If delineator found change to pipe (|)
					THEN @ColumnDelineatorI
				WHEN @character = @TextQualifier -- If text qualifier found then remove (change to zero length string
					THEN ''''
				WHEN @ColumnCountTotal<9999 AND @ColumnCount<@ColumnCountTotal AND @character=@RowTerminator
					THEN ''''
				ELSE @character 
				END 
--- Count number of columns in row
,@ColumnCount = CASE
				WHEN @character=@RowTerminator
					THEN 1
				WHEN @Reset=0
					THEN @columnCount+1
				ELSE @ColumnCount
				END
--- Take number of columns in row 1 as expected number of columns
,@ColumnCountTotal = CASE
						WHEN @ColumnCountTotal=9999 AND @character=@RowTerminator THEN @ColumnCount
						ELSE @ColumnCountTotal
						END
--- Increment character position by 1
,@PositionCharacter = @PositionCharacter +1
RETURN @Ret 
END
' 
END
GO
/****** Object:  Table [dbo].[TblPar_0AEAW_AT_Parameters]    Script Date: 17/03/2023 16:11:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TblPar_0AEAW_AT_Parameters]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[TblPar_0AEAW_AT_Parameters](
	[ParameterName] [varchar](255) NOT NULL,
	[ParameterString] [varchar](255) NULL,
	[ParameterDate1] [datetime] NULL,
	[ParameterDate2] [datetime] NULL,
	[ParameterInt] [int] NULL,
	[Comments] [varchar](255) NULL,
 CONSTRAINT [PK_TblPar_0AEAW_AT_Parameters] PRIMARY KEY CLUSTERED 
(
	[ParameterName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  UserDefinedFunction [dbo].[tFnDelimitedSplitMAX]    Script Date: 17/03/2023 16:11:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tFnDelimitedSplitMAX]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'CREATE FUNCTION [dbo].[tFnDelimitedSplitMAX]

-- From the legend that is Jeff Moden!

--===== Define I/O parameters
        (@pString VARCHAR(max), @pDelimiter CHAR(1))
--WARNING!!! DO NOT USE MAX DATA-TYPES HERE!  IT WILL KILL PERFORMANCE!
-- DONE TO SORT THE HEADER ISSUE. USE WITH 1 ROW ONLY!
RETURNS TABLE WITH SCHEMABINDING AS
 RETURN
--===== "Inline" CTE Driven "Tally Table" produces values from 1 up to 10,000...
     -- enough to cover VARCHAR(8000)
  WITH E1(N) AS (
                 SELECT 1 UNION ALL SELECT 1 UNION ALL SELECT 1 UNION ALL
                 SELECT 1 UNION ALL SELECT 1 UNION ALL SELECT 1 UNION ALL
                 SELECT 1 UNION ALL SELECT 1 UNION ALL SELECT 1 UNION ALL SELECT 1
                ),                          --10E+1 or 10 rows
       E2(N) AS (SELECT 1 FROM E1 a, E1 b), --10E+2 or 100 rows
       E4(N) AS (SELECT 1 FROM E2 a, E2 b), --10E+4 or 10,000 rows max
        E8(N) AS (SELECT 1 FROM E2 a, E4 b), -- 1,000,000 rows max
cteTally(N) AS (--==== This provides the "base" CTE and limits the number of rows right up front
                     -- for both a performance gain and prevention of accidental "overruns"
                 SELECT TOP (ISNULL(DATALENGTH(@pString),0)) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) FROM E8
                ),
cteStart(N1) AS (--==== This returns N+1 (starting position of each "element" just once for each delimiter)
                 SELECT 1 UNION ALL
                 SELECT t.N+1 FROM cteTally t WHERE SUBSTRING(@pString,t.N,1) = @pDelimiter
                ),
cteLen(N1,L1) AS(--==== Return start and length (for use in substring)
                 SELECT s.N1,
                        ISNULL(NULLIF(CHARINDEX(@pDelimiter,@pString,s.N1),0)-s.N1,80000)
                   FROM cteStart s
                )
--===== Do the actual split. The ISNULL/NULLIF combo handles the length for the final element when no delimiter is found.
 SELECT ItemNumber = ROW_NUMBER() OVER(ORDER BY l.N1),
        Item       = SUBSTRING(@pString, l.N1, l.L1)
   FROM cteLen l
;
' 
END
GO
/****** Object:  StoredProcedure [dbo].[spDOS_ConvertDelimiterAllTextfiles]    Script Date: 17/03/2023 16:11:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDOS_ConvertDelimiterAllTextfiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDOS_ConvertDelimiterAllTextfiles] AS' 
END
GO

ALTER PROCEDURE [dbo].[spDOS_ConvertDelimiterAllTextfiles] (
										   @filepath		varchar(255)   -- Filepath to folder containing files
										  ,@ChunkFiles		char(1)='N'    -- Powershell bombs out with large files. This chunks the files into 1,000,000 row files first.
										  ,@ChunkRows		int=1000000    -- Number of rows in each file
										  ,@ColDelineator	varchar(5)=',' -- column delimiter in file
										  ,@ColDelineatorI	varchar(5)='‡' -- column delimiter ideally wanted
										  ,@SkipRows		int=0		   -- Number of rows to remove from the top of the file
										  ,@debug			char(1)='N' -- PRINT dynamic SQL for debugging purposes
												   ) AS

IF @debug LIKE '%[^-__A-Za-z0-9:\. ]%' ESCAPE '_'
	RETURN
;

--DECLARE @debug			varchar(1)		SET @debug = 'Y'
--DECLARE @filepath		varchar(2000)	SET @filepath = 'D:\Data\SLAM\UniversityCollegeLondon_RRV\_LoadFolder\Flx\Aggregate'
--DECLARE @ChunkFiles		char(1)			SET @ChunkFiles = 'N'
DECLARE @server			varchar(255)	SET @server = (select @@servername)
DECLARE @db				varchar(255)	SET @db = (SELECT DB_NAME())
DECLARE @table			varchar(255)	SET @table = 'TblTempFileExport'
DECLARE @datetime		smalldatetime
DECLARE @commandXP		varchar(8000) --bcp doesn't accept varchar(max)
DECLARE @quote			varchar(1)
--DECLARE @ColDelineatorI	varchar(5)		SET @ColDelineatorI = '§'
DECLARE @RowTerminator	varchar(5)		SET @RowTerminator = '\n'
DECLARE @FirstRow		varchar(10)		SET @FirstRow = '1'
DECLARE @rows1			bigint
--DECLARE @rows2			bigint
DECLARE @error			varchar(1024)
DECLARE @comment		varchar(8000)

IF(@debug='Y')
BEGIN
SET NOCOUNT ON
END

SET @filepath = CASE WHEN RIGHT(@filepath,1)='\' THEN @filepath ELSE @filepath+'\' END

----------------------------------------------------------
--- Obtain directory (folder) contents

IF EXISTS(select name from sys.objects where name = 'TblTempDir')
BEGIN
TRUNCATE TABLE TblTempDir
END

IF NOT EXISTS(select name from sys.objects where name = 'TblTempDir')
BEGIN
CREATE TABLE TblTempDir(DirList varchar(max))
END


-----------------------------------------------------------
---Make temporary table to bung text in

IF EXISTS(select name from sys.objects where name = 'TblTempFileExport')
BEGIN
TRUNCATE TABLE TblTempFileExport
END

IF NOT EXISTS(SELECT name FROM sys.objects WHERE name = 'TblTempFileExport')
BEGIN -- if TempFileExport
CREATE TABLE TblTempFileExport(FileText varchar(max) NULL)
END -- if TempFileExport

IF(@debug='Y')
BEGIN
PRINT @filepath
END

----------------------------------------------------------------------------
--- Check directory contents

INSERT INTO TblTempDir
exec Lookups.dbo.spDOS_DirFileListOnly @filepath

IF NOT EXISTS(select 1 from TblTempDir where DirList='Processed')
BEGIN --If directory doesn't exist then make directory
SET @commandXP = '
IF NOT EXIST "' + @filepath + 'Processed" MKDIR "' + @filepath + 'Processed"'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END
END --If directory doesn't exist then make directory

----------------------------------------------------------------------------

IF EXISTS(SELECT DirList FROM TblTempDir WHERE DirList LIKE '%.[c,t][s,x][t,v]')
BEGIN -- CSV present in folder

----------------------------------------------------------------------------

IF(@ChunkFiles='N' /*AND NOT EXISTS(SELECT DirList FROM TblTempDir WHERE DirList = 'ConvertDelimiter.ps1')*/)
BEGIN -- Powershell conversion file does not exist

IF(@debug='Y')
BEGIN
PRINT 'Powershell conversion file does not exist so needs creating.'
END

INSERT INTO TblTempFileExport VALUES('###-----------------------------------------------------------------------###
# Start stop-watch ...
$sw2 = new-object System.Diagnostics.Stopwatch
$sw2.Start()

# Setup access options
    [Int32] $bufferSize = 16 * 1024;
    [System.Text.Encoding] $defaultEncoding = [System.Text.Encoding]::default; #UTF7 or default seems to handle the ± character
    [System.IO.FileMode] $mode = [System.IO.FileMode]::Open;
    [System.IO.FileAccess] $access = [System.IO.FileAccess]::Read;
    [System.IO.FileShare] $share = [System.IO.FileShare]::Read;
    [System.IO.FileOptions] $options = [System.IO.FileOptions]::SequentialScan;

# Get all text files and convert them ...
$exclusions= ("*directorylisting*","*DirectoryListing*","*DIRECTORYLISTING*","CNV_*");
Get-ChildItem "$PSScriptRoot\*.[c,t][s,x][t,v]" -Exclude $exclusions | ` # $PSScriptRoot only works above PS2. Previously $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#CHAR(167) -- §
#CHAR(172) -- ¬
#CHAR(177) -- ±
#CHAR(165) -- ¥
#CHAR(169) -- ©
#CHAR(161) -- ¡
#CHAR(135) -- ‡
Foreach-Object { # Start of file loop
  $filepathname = $_.FullName
  $filename = $_.Name
  $file = $_.BaseName
  $ext = $_.Extension
  $folder = $_.DirectoryName
  $fileout = $folder+"\CNV_"+$_.BaseName+$ext
  $fileProcessed = $folder+"\Processed\"+$filename

#  $filename = $_.FullName
   "$filepathname"
   "$fileout"
# Import file as CSV and then write out with new delimiter - unfortunately it auto-quotes the delimiter with ""
Get-Content $filepathname'+CASE WHEN @skipRows>0 THEN ' | Select-Object -Skip '+CAST(@SkipRows as varchar(15)) ELSE '' END+' | convertfrom-csv'+CASE WHEN @ColDelineator IN(',','","') THEN '' ELSE ' -delimiter '''+@colDelineator+'''' END+' <#-Encoding ''default''#> | export-csv "$fileout" -delimiter '''+@ColDelineatorI+''' -Encoding ''default'' -NoTypeInformation
$content = [System.IO.File]::ReadAllText($fileout,$defaultEncoding).Replace("""","")
[System.IO.File]::WriteAllText($fileout, $content,$defaultEncoding)

# Move original file to processed folder - Processed subfolder has to exist for this to work!
Move-Item -path $filepathname -destination $fileProcessed -Force # Needs destination folder to exist

} # END of file loop

# Stop stop-watch and print time taken for conversion
$sw2.Stop()
Write-Host "Conversion complete in " $sw2.Elapsed.TotalSeconds "seconds"')


SET @commandXP = 'bcp ' + @db + '.dbo.' + @table + ' out "' + @filepath + 'ConvertDelimiter.ps1" /c /t "," /S "' + @server + '" -T -C RAW' -- or - C 1252'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

END -- Powershell conversion file does not exist

----------------------------------------------------------------------------

IF(@ChunkFiles='Y' /*AND NOT EXISTS(SELECT DirList FROM TblTempDir WHERE DirList = 'ChunkAndConvertDelimiter.ps1')*/)
BEGIN -- Powershell chunk and conversion file does not exist

IF(@debug='Y')
BEGIN
PRINT 'Powershell chunk and conversion file does not exist so needs creating.'
END

INSERT INTO TblTempFileExport VALUES('# SPLIT LARGE FILES INTO SMALLER FILES FOR PROCESSING
# Taken from http://stackoverflow.com/questions/1001776/how-can-i-split-a-text-file-using-powershell

$sw = new-object System.Diagnostics.Stopwatch
$sw.Start()

# Get all text files and convert them ...
$exclusions= ("*directorylisting*","*DirectoryListing*","*DIRECTORYLISTING*","CNV_*");
Get-ChildItem "$PSScriptRoot\*.[c,t][s,x][t,v]" -Exclude $exclusions | ` # $PSScriptRoot only works above PS2. Previously $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Foreach-Object { # Start of file loop
  $filepathname = $_.FullName
  $filename = $_.Name
  $file = $_.BaseName
  $ext = $_.Extension
  $folder = $_.DirectoryName
  $fileProcessed = $folder+"\Processed\"+$filename

$rootName = $folder+"\"+$file

$linesperFile = '+CAST(@ChunkRows as varchar(20))+' # Number of lines in each file batch
$filecount = 1
$reader = $null

# TRY TO GET THE HEADER INTO A VARIABLE
$headers = Get-Content "$filepathname"'+CASE WHEN @skipRows>0 THEN ' | Select-Object -Skip '+CAST(@SkipRows as varchar(15)) ELSE '' END+' | Select -Index 0

# Setup access options
    [Int32] $bufferSize = 16 * 1024;
    [System.Text.Encoding] $defaultEncoding = [System.Text.Encoding]::default; #UTF7 or default seems to handle the ± character
    [System.IO.FileMode] $mode = [System.IO.FileMode]::Open;
    [System.IO.FileAccess] $access = [System.IO.FileAccess]::Read;
    [System.IO.FileShare] $share = [System.IO.FileShare]::Read;
    [System.IO.FileOptions] $options = [System.IO.FileOptions]::SequentialScan;

try{
#    $reader = [io.file]::OpenText($filepathname)
    # FileStream(String, FileMode, FileAccess, FileShare, Int32, FileOptions) constructor
    # http://msdn.microsoft.com/library/d0y914c5.aspx
    [System.IO.FileStream] $input = New-Object -TypeName ''System.IO.FileStream'' -ArgumentList ($filepathname, $mode, $access, $share, $bufferSize, $options);
# https://stackoverflow.com/questions/39755511/cannot-assign-value-to-variable-from-powershell-form#39756704
#    [System.IO.FileStream] $script:input = New-Object -TypeName ''System.IO.FileStream'' -ArgumentList ($filepathname, $mode, $access, $share, $bufferSize, $options);

    # StreamReader(Stream, Encoding, Boolean, Int32) constructor
    # http://msdn.microsoft.com/library/ms143458.aspx
    [System.IO.StreamReader] $reader = New-Object -TypeName ''System.IO.StreamReader'' -ArgumentList ($input, $defaultEncoding, $true, $bufferSize);
    [String] $line = $null;
    [Int32] $currentIndex = 0;

    try{
        # File 1
         $fileout=($rootName,$filecount.ToString("000"),$ext)  # Added by JCEH
        "Creating file number $filecount $fileout"

        $writer = [io.file]::CreateText("{0}{1}.{2}" -f ($rootName,$filecount.ToString("000"),$ext))
        $filecount++
        $linecount = 0

        while($reader.EndOfStream -ne $true) {
            #"Reading next $linesperFile rows"
            # Read n lines of text
            while( ($linecount -lt $linesperFile) -and ($reader.EndOfStream -ne $true)){
            if($filecount -gt 2 -and $linecount -eq 0){
                $writer.WriteLine($headers);
                            }
                $writer.WriteLine($reader.ReadLine()) ` 
                $linecount++
            }

            if($reader.EndOfStream -ne $true) { # Not reached end of file so close file and open the next one
#                "Closing file $filecount"
                $writer.Dispose();

        # Files > 1
                $fileout=($rootName,$filecount.ToString("000"),$ext) # Added by JCEH
                "Creating file number $filecount $fileout"
                $writer = [io.file]::CreateText("{0}{1}.{2}" -f ($rootName,$filecount.ToString("000"),$ext))
                $filecount++
                $linecount = 0
                #$headers | Out-File -FilePath "$fileout" -Append # Added by JCEH, 10-Aug-2016
            }
#"C $filecount"
        }
    } finally {
        $writer.Dispose();
    }
} finally {
    $reader.Dispose();
"$fileout"
}

# Move original file to processed folder - Processed subfolder has to exist for this to work!
Move-Item -path $filepathname -destination $fileProcessed -Force # Needs destination folder to exist

} # END of file loop

$sw.Stop()

Write-Host "Split complete in " $sw.Elapsed.TotalSeconds "seconds"

###-----------------------------------------------------------------------###
# Start stop-watch ...
$sw2 = new-object System.Diagnostics.Stopwatch
$sw2.Start()

# Get all text files and convert them ...
$exclusions= ("*directorylisting*","*DirectoryListing*","*DIRECTORYLISTING*","CNV_*");
Get-ChildItem "$PSScriptRoot\*.[c,t][s,x][t,v]" -Exclude $exclusions | ` # $PSScriptRoot only works above PS2. Previously $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#CHAR(167) -- §
#CHAR(172) -- ¬
#CHAR(177) -- ±
#CHAR(165) -- ¥
#CHAR(169) -- ©
#CHAR(161) -- ¡
#CHAR(135) -- ‡
Foreach-Object { # Start of file loop
  $filepathname = $_.FullName
  $filename = $_.Name
  $file = $_.BaseName
  $folder = $_.DirectoryName
  $fileout = $folder+"\CNV_"+$_.BaseName+$ext
  $fileProcessed = $folder+"\Processed\"+$filename

#  $filename = $_.FullName
   "$filepathname"
   "$fileout"
# Import file as CSV and then write out with pipe delimiter - unfortunately it auto-quotes the delimiter with ""
Get-Content $filepathname'+CASE WHEN @skipRows>0 THEN ' | Select-Object -Skip '+CAST(@SkipRows as varchar(15)) ELSE '' END+' | convertfrom-csv'+CASE WHEN @ColDelineator IN(',','","') THEN '' ELSE ' -delimiter '''+@colDelineator+'''' END+' <#-Encoding ''default''#> | export-csv "$fileout" -delimiter '''+@ColDelineatorI+''' -Encoding ''default'' -NoTypeInformation
# Unfortunately export-csv always sticks double-quotes around the delimiter so now need to strip the “”.
$content = [System.IO.File]::ReadAllText($fileout,$defaultEncoding).Replace("""","")
[System.IO.File]::WriteAllText($fileout, $content,$defaultEncoding)

# Move original file to processed folder - Processed subfolder has to exist for this to work!
Move-Item -path $filepathname -destination $fileProcessed -Force # Needs destination folder to exist

} # END of file loop

# Stop stop-watch and print time taken for conversion
$sw2.Stop()
Write-Host "Conversion complete in " $sw2.Elapsed.TotalSeconds "seconds"')


SET @commandXP = 'bcp ' + @db + '.dbo.' + @table + ' out "' + @filepath + 'ChunkAndConvertDelimiter.ps1" /c /t "," /S "' + @server + '" -T -C RAW' -- or - C 1252'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

END -- Powershell chunk and conversion file does not exist

----------------------------------------------------------------------------

IF(@ChunkFiles='N')
BEGIN --- CSV conversion only

IF(@debug='Y')
BEGIN
PRINT 'Run CSV conversion ...'
END

TRUNCATE TABLE TblTempFileExport

INSERT INTO TblTempFileExport VALUES('@echo off
Powershell.exe -executionpolicy bypass -File "'+@filepath+'ConvertDelimiter.ps1"
echo on')

SET @commandXP = 'bcp ' + @db + '.dbo.' + @table + ' out "' + @filepath + 'RunConvertDelimiter.bat" /c /t "," /S "' + @server + '" -T -C RAW' -- or - C 1252'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

SET @commandXP = 'CALL "'+@filepath+'RunConvertDelimiter.bat"'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

END --- CSV conversion only

IF(@ChunkFiles='Y')
BEGIN --- Chunk and CSV conversion

IF(@debug='Y')
BEGIN
PRINT 'Run Chunk and CSV conversion ...'
END

TRUNCATE TABLE TblTempFileExport

INSERT INTO TblTempFileExport VALUES('@echo off
Powershell.exe -executionpolicy bypass -File "'+@filepath+'ChunkAndConvertDelimiter.ps1"
echo on')

SET @commandXP = 'bcp ' + @db + '.dbo.' + @table + ' out "' + @filepath + 'RunChunkAndConvertDelimiter.bat" /c /t "," /S "' + @server + '" -T -C RAW' -- or - C 1252'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

SET @commandXP = 'CALL "'+@filepath+'RunChunkAndConvertDelimiter.bat"'
IF(@debug='Y')
BEGIN
PRINT @commandXP
EXEC master..xp_cmdshell @commandXP--,no_output
END
IF(@debug='N')
BEGIN
EXEC master..xp_cmdshell @commandXP,no_output
END

END --- Chunk and CSV conversion

----------------------------------------------------------------------------

END -- CSV present in folder

----------------------------------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[spDOS_DirFileListOnly]    Script Date: 17/03/2023 16:11:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDOS_DirFileListOnly]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDOS_DirFileListOnly] AS' 
END
GO

ALTER PROCEDURE [dbo].[spDOS_DirFileListOnly] (@filepath varchar(8000)) AS

-----------------------------------------------------------

-- EXEC spDOS_DirFileListOnly 'C:\'

-----------------------------------------------------------

DECLARE @command varchar(8000)

-----------------------------------------------------------

SET @command = 'DIR /B "' + @filepath 
EXEC master..xp_cmdshell @command

-----------------------------------------------------------
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Function [dbo].[uFnRemoveNonAlphaCharacters](@Temp VarChar(8000))
Returns VarChar(8000)
AS
Begin

    Declare @KeepValues as varchar(50)
    Set @KeepValues = '%[^0-9, ''a-z]%'
    While PatIndex(@KeepValues, @Temp) > 0
        Set @Temp = Stuff(@Temp, PatIndex(@KeepValues, @Temp), 1, '')

    Return @Temp
End
GO