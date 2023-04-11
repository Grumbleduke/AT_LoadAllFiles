# AT_LoadAllFiles
SQL code to load most common file types

Loading all files for a user.

Overview.

This is initiated by running the procedure [dbo].[uspImportAllUserFiles]. This can be done for all users (default) or by specifying a specific user, e.g. exec uspImportAllUserFiles @usr='login'. The use of @debug=’Y’ for any procedure will print information that might help with any troubleshooting.

 In order for the procedure uspImportAllUserFiles to work it requires three tables: TblLog_Imports contains the results of any attempt to load a file; TblPar_Parameters contains the variables required to load the data files; TblSysUsers contains a list of active users together with their email addresses to email the results. The procedure in turn runs 3 sub-procedures in the following order: dbo.spSubConvertAllUserTextFiles, dbo.spSubLoadAllUserTextFiles and dbo.spSubLoadAllUserMicrosoftOfficeFiles. These sub-procedures in turn require the following user functions: ufnConvertTextQualifiedCSVtoVBS, ufnConvertTextQualifiedDelimiter, tFnDelimitedSplitMAX and uFnRemoveNonAlphaCharacters. Those same sub-procedures also require the following procedures: spDOS_ConvertDelimiterAllTextfiles, spDOS_DirFileListOnly. In order to email the results, database mail will need to be enabled with a profile called BIU. You will also need the Microsoft ACE drivers installed and enabled in order to load from Microsoft Excel or Microsoft Access. Flat (text) files are loaded using BULK INSERT. xp_cmdshell will also need to be enabled and setup in a secure manner with access to the folder(s) where the data will be loaded from.

Security

xp_cmdshell will need to be enabled, a proxy account setup and suitable role for it setup. In addition to that the account for the job or the users themselves if they are going to be allowed to run the procedure, will need access to the relevant folder on the server and also permission for bulk operations (GRANT ADMINISTER BULK OPERATIONS TO [username]. They will also require permission to connect to Excel as a linked server (unless the OPENROWSET option is enabled) which requires permission to alter any linked server (GRANT ALTER ANY LINKED SERVER TO [username]) – could be dangerous! The user account will also need to be able to SELECT from the sys.objects and sys.columns views within the database. 


How it works.

spSubConvertAllUserTextFiles

Earlier versions of SQL struggled to BULK INSERT reliably where a text qualifier wasn’t applied to a whole column of data. Therefore I made a user function to clean the file in SQL and change the delimiter to something unlikely to be present in the data, like a double-dagger = ALT+0135 = ‡. However, user function are notoriously slow, especially when scalar rather than a table-function and in particular with bigger files I found Powershell much more reliable. Strangely, Powershell could handle these files with the standard Import-CSV command yet it took years until SQL had the FORMAT = CSV option added to allow it to handle all RFC4180 compliant files. The first sub-procedure, spSubConvertAllUserTextFiles, handles this conversion. In the TblPar_Parameters table there should be a series of parameters with ParameterName made up of the username (login) for each user, followed by an underscore and then the parameter name. 

SELECT *
FROM TblPar_Parameters
WHERE ParameterName LIKE SUSER_NAME()+'%'

The fields ParameterString, ParameterDate1, ParameterDate2 and ParameterInt contain the required values for the sub-procedure to run and the Comments column contains hints on the values used.

The parameter username_ConvertColumnDelineator has three values in the ParameterString column: NO = do not do any conversion; SQL = use a SQL function to do the conversion (can be slow but can handle no header row and duplicate column names); Powershell = use Powershell for the conversion (fast but needs a header row with unique column names). Powershell also allows the file to be split into smaller files which can be useful with very large files. If you are using Powershell then the security settings may need to be adjusted to allow the scripts to run on the server.
The first parameter the sub-procedure requires is the location of the files to convert. This is obtained from the username_Filepath parameter in the ParameterString column. If the ParameterString is set to “D:\DO NOT LOAD INTO THIS DATABASE\” then that folder (user) will be skipped. After that the sub-procedure needs to know how many rows to skip at the top of the file before the header row (or data if no header) starts using the ParameterInt column value for the parameter username_ConvertColumnDelineatorSkipRows (default = 0). In order to know where each row ends it also needs to know the row terminator used in the file using the username_RowTerminator parameter (default = Windows = \n which is a carriage return and a line feed) in the ParameterString column. If the file has a header row containing the column names then the username_UseHeaderForColumnNames parameter in the ParameterString column should be left on the default value of YES, otherwise set to NO if no header is present in the file. The next three parameters required involve the column delimiter. The parameter in the ParameterString column for username_ColumnDelineatorOriginal is the character used as to delineate the columns in the file as originally received, together with the text qualifier parameter, username_TextQualifierOriginal if used, e.g. “,”. The parameter username_ColumnDelineator is the character you wish to use as a delineator and should be a character not present in the file. If the character is present in the file it will be replaced with an underscore (_). The default value for the intended column delineator is a double-dagger, ‡. Finally, if you wish to chunk the original file into smaller files then the username_ConvertColumnDelineatorChunkFiles parameter should have a ParameterString of Y or Yes and the ParameterInt should be set to the number of rows wanted in each smaller file.


spSubLoadAllUserTextFiles

The second sub-procedure, spSubLoadAllUserTextFiles, uses BULK INSERT to load the data from all files ending in .txt (text), .csv (comma separated values) or .vbs (vertical bar separated). It can handle flat files with or without a header and uses the same parameter table: TblPar_Parameters. Depending on whether a destination table core name is supplied or not, a table will be created for each file encountered or if a core name is supplied for the destination table then all files will be loaded into the same table. 

spSubLoadAllUserTextFiles takes the following optional, input parameters when executed:

@usrI = the login (SUSER_NAME()) if it is to be run only for a specific user.
@prefix = an optinal prefix that will be appended to the final table name after the username. The default value is an underscore, “_”.
@tableI = intended destination table core name. If not blank (zero length string) then all files will be placed into this table. If omitted then a new table will be created for each file based on the filename but with any odd characters replaced by underscores.
@suffix = an optional suffix to the core table name. By default this is set to _Import.
@version = an optional tag to put at the end of the table name to indicate which iteration this might be, e.g. V1.0.

The final table name(s) will start with the relevant username with any odd characters replaced by underscores, followed by an optional prefix, core table name or filename if left blank, optional suffix, and finally an optional version tag.

The procedure has two cursors to loop through each user and each file found in each user folder. The main part of the procedure itself is split into three sections: the first section creates a temporary landing table if no header is present; the second section creates a temporary landing table is a header is present and the final section creates the final table, adds any columns as required and then inserts the data into the final table. Because of some of the historic issues with CSV files there are a few workarounds and then there is another workaround for long filenames because of a limitation with the interaction between SQL and the Windows file system. 

Within the files loop (file_cursor) it first checks if the length of the filepath together with the filename is more than 128 characters long. If it is then it renames the file to file1.txt. If the filepath and filename is still more than 128 characters then you will probably encounter a rather non-descript error.

If ParameterString from TblPar_Parameters where ParameterName is username_UseHeaderForColumnNames = NO then it will load the whole file into a temporary table with a single column and then count the number of columns based on the character(s) used as the column delimiter from ParameterString values for username_ColumnDelineator and username_TextQualifierOriginal in the TblPar_Parameters table. Based on this result the procedure will make a table with column names beginning at Col0001 through to whatever the number of columns found was, e.g. Col0102 where 102 columns found. Note that the ParameterString value for username_RowTerminator will be required to know where each row in the file ends although it is most likely the default value of \n or possibly CHAR(10) if from Oracle or UNIX. The width of the columns will be defined by the ParameterInt value for username_DefaultColumnWidth or if that is left as the default of minus one (-1) then it will assume a default of 255 characters unless there are more than 80 columns in which case it will reduce the width of the columns depending on how many columns are present. If any columns are wider than these assumptions then the ParameterInt value for username_DefaultColumnWidth will need setting higher than the widest column in the data. It is possible to start smaller and just adjust the width of a single column in the final table after the event using ALTER TABLE tablename ALTER COLUMN varchar(columnwidth) NULL. Next any rows with a different number of column are removed and placed in a table ending JailedRecords. The remaining rows are then exported to new file. If no records were jailed then the original file is imported into the temporary table or the cleaned file if records were jailed. [NOTE TO SELF: I ought to add the capability to start the import at a certain row specified by ParameterName = username_FirstRow. Maybe make one for username_LastRow as well?] If it doesn’t exist then the final destination table created or if it does exist, any missing columns are added.

If ParameterString from TblPar_Parameters where ParameterName is username_UseHeaderForColumnNames = YES then it loads the first row from the file based on the ParameterString value for username_FirstRow and then based on the ParameterString values for username_RowTerminator, username_ColumnDelineator and username_TextQualifierOriginal it will create the required columns, again using a default of 255 unless there are more than 80 columns or the ParameterInt value for username_DefaultColumnWidth is not equal to minus one (-1). If it doesn’t exist then the final destination table created or if it does exist, any missing columns are added. 

If the SQL server version is 2017 (PMV = 14) or above and the column delineator = ‘CSV’ as opposed to ‘,’ or ‘”,”’, then it will assume the file is RFC4180 compliant and BULK INSERT into the temporary table using the FORMAT = ‘CSV’ option.

If the SQL Server version isn’t 2017 or above or the column delineator is not set to ‘CSV’ then it will load the file as a single column and check if the number of column delineators is consistent throughout the file. Any rows where the number of columns is different from the majority are quarantined in a separate table ending in “…_JailedRecords” and the rows with consistent columns exported (using BCP) ready to be imported into the …_Temp import table.

If the SQL Server version isn’t 2017 or above and the column delineator is set to ‘,’ or ‘”,”’ with a filename ending .csv and the number of columns is not consistent then it will try to clean the file and change the delimiter to a vertical bar (|). Note to self: you might want to add a check that the ParameterString for ParameterName = username_ConvertColumnDelineator in TblPar_Parameters is set to ‘Yes’ and not ‘Powershell’ or ‘No’. Note that this conversion within SQL currently uses a scalar function so will only be single-threaded and is not suitable for large files. A table function would be better but I suspect still not as efficient as the Powershell conversion script. If it does do the conversion in SQL the cleaned file will be exported (using BCP) to the same filename but with an extension of .vbs (for vertical bar separated). It will then load the cleaned file into the …_Temp import table.

If the SQL Server version isn’t 2017 or above or the column delineator is not set to ‘,’ or ‘”,”’ and the number of columns is not consistent then it will jail the inconsistent rows into a “…JailedRecords” table and export the consistent rows to a new file before loading that cleaned file into the temporary table ending …_Temp.

If the number of columns is consistent then the raw file is loaded using BULK INSERT into the table created from the header row ending with …_Temp.

At this stage there should be a landing table ending …_Temp with all the clean raw data in it and a final destination table. All columns will be text to avoid possible conversion errors. Sometimes you can get a rather cryptic error message at this stage where the column width of a row in the file is larger than the column width in the table it is trying to insert into. The fix is to manually alter the column width in the final destination table (ALTER [tablename] ALTER COLUMN columnname varchar(new width) NULL), set the the ParameterInt value for username_DefaultColumnWidth in TblPar_Parameters to the new width required, drop the temporary table and try to load the file again. Note that if the width of the table exceeds 8000 characters it may display a warning message.

The procedure will check sys.columns for any column differences, add any missing columns and then construct an insert statement where the column names are the same. The data will be inserted from the table ending …_Temp and then that table will be dropped leaving only the final destination table. 
The result of the attempt to load the file(s) will be placed in TblLog_Imports and if a suitable email profile is setup and database mail is enabled then an email will be sent to the user with a summary of what has happened.

spSubLoadAllUserMicrosoftOfficeFiles

The third sub-procedure, spSubLoadAllUserMicrosoftOfficeFiles, uses a linked server or OPENROWSET to load the data from all files ending in .xls[,b,m,x] (Excel spreadsheet) or .mdb / .accdb (Access database). It can handle spreadsheet tabs with or without a header and uses the same parameter table as the other procedures: TblPar_Parameters. Depending on whether a destination table core name is supplied or not, a table will be created for each file encountered or if a core name is supplied for the destination table then all files will be loaded into the same table. 

spSubLoadAllUserMicrosoftOfficeFiles takes the following optional, input parameters when executed:

@usrI = the login (SUSER_NAME()) if it is to be run only for a specific user.
@prefix = an optinal prefix that will be appended to the final table name after the username. The default value is an underscore, “_”.
@tableI = intended destination table core name. If not blank (zero length string) then all files will be placed into this table. If omitted then a new table will be created for each file based on the filename but with any odd characters replaced by underscores.
@suffix = an optional suffix to the core table name. By default this is set to _Import.
@version = an optional tag to put at the end of the table name to indicate which iteration this might be, e.g. V1.0.

The final table name(s) will start with the relevant username with any odd characters replaced by underscores, followed by an optional prefix, core table name or filename if left blank, optional suffix, and finally an optional version tag.

The procedure has three cursors to loop through each user, each file found in each user folder and finally each table / view or tab / named range in each file – as revealed by SP_TABLES_EX if the Microsoft Office file is setup as a linked server. If the parameter @UseOPENROWSET (taken from ParameterString value in TblPar_Parameters for the row with ParameterName = username_OPENROWSET) is set to Y[es] then only the specific locations given by ParameterString for rows with ParameterName from username_LoadRangeSheetTable01 through to username_LoadRangeSheetTable06 will be loaded. If If the parameter @UseOPENROWSET is set to n[o] then all tabs / named ranges or tables that begin with the ParameterString values in rows with ParameterName values username_LoadRangeSheetTable01 through to username_LoadRangeSheetTable06 will be loaded. So, to load everything then set ParameterString to a zero length for username_LoadRangeSheetTable01 and a single quote for username_LoadRangeSheetTable02.

 The main part of the procedure itself is split into three sections: the first section creates a temporary landing table by SELECT * INTO if OPENROWSET has been used; the second section creates a temporary landing table using SP_COLUMN_EX to discover which columns are present and the third and final section creates the final table, adds any columns as required (from comparing sys.columns) and then inserts the data into the final table.


The result of the attempt to load the file(s) will be placed in TblLog_Imports and if a suitable email profile is setup and database mail is enabled then an email will be sent to the user with a summary of what has happened.



