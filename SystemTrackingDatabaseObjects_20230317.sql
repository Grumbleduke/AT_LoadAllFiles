USE [SystemTracking]
GO
/****** Object:  Table [dbo].[TblSysUsers]    Script Date: 17/03/2023 16:17:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TblSysUsers]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[TblSysUsers](
	[Username] [varchar](255) NOT NULL,
	[WindowsAccountDisabled] [char](1) NOT NULL,
	[SQLaccountDisabled] [char](1) NOT NULL,
	[EmailAddress] [varchar](255) NULL,
	[SendNotifications] [char](1) NULL,
 CONSTRAINT [PK_TblSysUsers] PRIMARY KEY CLUSTERED 
(
	[Username] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
END
GO
INSERT [dbo].[TblSysUsers] ([Username], [WindowsAccountDisabled], [SQLaccountDisabled], [EmailAddress], [SendNotifications]) VALUES (N'DOMAIN\Username', N'N', N'N', N'Your Name <your.name@domain.net>', N'Y')
GO
INSERT [dbo].[TblSysUsers] ([Username], [WindowsAccountDisabled], [SQLaccountDisabled], [EmailAddress], [SendNotifications]) VALUES (N'DOMAIN\sqlAgent', N'N', N'N', NULL, N'N')
GO
INSERT [dbo].[TblSysUsers] ([Username], [WindowsAccountDisabled], [SQLaccountDisabled], [EmailAddress], [SendNotifications]) VALUES (N'sa', N'N', N'N', N'God <your.name@domain.net>', N'N')
GO
/****** Object:  StoredProcedure [dbo].[spAdmin_SysUsers]    Script Date: 17/03/2023 16:17:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spAdmin_SysUsers]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spAdmin_SysUsers] AS' 
END
GO
ALTER PROCEDURE [dbo].[spAdmin_SysUsers] (@debug char(1)='N') AS

-- exec spAdmin_SysUsers @debug='Y'

------------------------------------------------------------
--- Check the input for SQL Injection and simply exit to deny an attacker any hints. Thanks to Jeff Moden. :)

IF @debug LIKE '%[^N,Y]' ESCAPE '_'
RETURN

------------------------------------------------------------

--- Find Windows user accounts and add missing ones to the table

DECLARE @domain varchar(100)	set @domain = CASE WHEN @@servername LIKE '%/%' THEN LEFT(@@servername,CHARINDEX('\',@@servername,1)-1) ELSE @@servername END --'UUKP8V-MKAIV01'
DECLARE @users  TABLE(Users varchar(8000) NULL)

SET NOCOUNT ON

IF NOT EXISTS(select 1 from sys.objects where [name] = 'TblSysUsers')
BEGIN
CREATE TABLE TblSysUsers(Username varchar(255) NOT NULL
						,WindowsAccountDisabled char(1) NOT NULL
						,SQLaccountDisabled char(1) NOT NULL
						,EmailAddress varchar(255) NULL
						,SendNotifications char(1) NULL
						,CONSTRAINT [PK_TblSysUsers] PRIMARY KEY CLUSTERED 
						(Username ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						) ON [PRIMARY]
END

-----------------------------------------------------

INSERT INTO @users
exec xp_cmdshell 'wmic useraccount get Disabled,Name'
--exec xp_cmdshell 'wmic useraccount'

IF(@debug='Y')
BEGIN
SELECT * FROM @users
END

-----------------------------------------------------
--- Insert missing users

INSERT INTO dbo.TblSysUsers(Username,WindowsAccountDisabled,SQLaccountDisabled,SendNotifications)
SELECT 
 Lookups.dbo.uFnRemoveNonAlphaCharacters(RTRIM(LTRIM(SUBSTRING(U.Users,CHARINDEX(' ',U.Users,1),255)))) as Username
,CASE
	WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE'
		THEN 'Y'
	ELSE 'N'
	END as WindowsAccountDisabled
,CASE 
	WHEN is_disabled = 1 THEN 'Y'
	WHEN is_disabled = 0 THEN 'N'
	ELSE '?'
	END as SQLaccountDisabled
,'Y' as SendNotifications
FROM @users U
	LEFT OUTER JOIN sys.server_principals P
		ON @domain+'\'+Lookups.dbo.uFnRemoveNonAlphaCharacters(RTRIM(LTRIM(SUBSTRING(U.Users,CHARINDEX(' ',U.Users,1),255)))) = P.[name]
WHERE LEFT(Users,CHARINDEX(' ',Users,1)) <> 'Disabled'
	AND RTRIM(LEFT(Users,CHARINDEX(' ',Users,1)))<>''
	AND Lookups.dbo.uFnRemoveNonAlphaCharacters(RTRIM(LTRIM(SUBSTRING(U.Users,CHARINDEX(' ',U.Users,1),255)))) NOT IN(select Username from TblSysUsers)

-----------------------------------------------------
--- Update information for existing users

UPDATE S
SET
 WindowsAccountDisabled = 
	CASE
	WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE'
		THEN 'Y'
	ELSE WindowsAccountDisabled
	END
,SQLaccountDisabled = 
	CASE 
	WHEN is_disabled = 1 THEN 'Y'
	WHEN is_disabled = 0 THEN 'N'
	ELSE SQLaccountDisabled
	END 
,SendNotifications =
 	CASE
	WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE' THEN 'N'
	WHEN is_disabled = 1 THEN 'N'
	ELSE SendNotifications
	END
FROM dbo.TblSysUsers S
	INNER JOIN @users U
		ON S.Username = Lookups.dbo.uFnRemoveNonAlphaCharacters(RTRIM(LTRIM(SUBSTRING(U.Users,CHARINDEX(' ',U.Users,1),255))))
	LEFT OUTER JOIN sys.server_principals P
		ON @domain+'\'+Lookups.dbo.uFnRemoveNonAlphaCharacters(RTRIM(LTRIM(SUBSTRING(U.Users,CHARINDEX(' ',U.Users,1),255)))) = P.[name]
WHERE LEFT(Users,CHARINDEX(' ',Users,1)) <> 'Disabled'
	AND RTRIM(LEFT(Users,CHARINDEX(' ',Users,1)))<>''
	AND ((WindowsAccountDisabled <> CASE WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE' THEN 'Y' ELSE WindowsAccountDisabled END)
		OR (SQLaccountDisabled <> CASE WHEN is_disabled = 1 THEN 'Y' WHEN is_disabled = 0 THEN 'N' ELSE SQLaccountDisabled END)
		OR (SendNotifications <> CASE WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE' THEN 'N' WHEN is_disabled = 1 THEN 'N' ELSE SendNotifications	END)
		)

-----------------------------------------------------
--- Try to work out how users are accessing SQL and update table

declare @user   sysname
declare @logins TABLE (AccountName varchar(255) NOT NULL
					  ,AccountType varchar(255) NOT NULL
					  ,Privilege   varchar(255) NOT NULL
					  ,LoginMapped varchar(255) NOT NULL
					  ,PermissionPath varchar(255) NULL
					  )
declare @command varchar(8000)
 
declare user_cursor cursor for
select @domain+'\'+Username 
from dbo.TblSysUsers
--where USername = 'aorchard'

open user_cursor 
fetch next from user_cursor into @user
 
while @@fetch_status = 0
begin

DELETE
FROM @logins 

    begin try

		insert into @logins
		exec xp_logininfo @user

		update U
		set
			SQLaccountDisabled = CASE 
									WHEN is_disabled = 1 THEN 'Y'
									WHEN is_disabled = 0 THEN 'N'
									ELSE '?'
									END
		from dbo.TblSysUsers U
			inner join @logins L
				on L.AccountName = @domain+'\'+U.Username
			inner join sys.server_principals P
				on L.PermissionPath = P.[name]

    end try
    begin catch
        --Error on xproc because login doesn't exist
        --SET @command = 'drop login '+convert(varchar(255),@user)
		--exec(@command)
    end catch


    fetch next from user_cursor into @user
end
 
close user_cursor
deallocate user_cursor

-----------------------------------------------------
/*
--- Following scripts drops SQL logins not in the Windows Domain. From:
https://sqlstudies.com/2013/03/01/script-to-clean-up-windows-logins-no-longer-in-ad/
*/

declare recscan cursor for
select name 
from sys.server_principals
where type IN('U','G') 
	and name like @domain+'%'
 
open recscan 
fetch next from recscan into @user
 
while @@fetch_status = 0
begin

DELETE
FROM @logins 

    begin try

		insert into @logins
		exec xp_logininfo @user

    end try
    begin catch
        --Error on xproc because login doesn't exist
        SET @command = 'drop login '+convert(varchar(255),@user)
		exec(@command)
    end catch


    fetch next from recscan into @user
end
 
close recscan
deallocate recscan

IF(@debug='Y')
BEGIN
SELECT * FROM dbo.TblSysUsers
END

-----------------------------------------------------
GO
/****** Object:  StoredProcedure [dbo].[spAdmin_SysUsersActive]    Script Date: 17/03/2023 16:17:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spAdmin_SysUsersActive]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spAdmin_SysUsersActive] AS' 
END
GO
ALTER PROCEDURE [dbo].[spAdmin_SysUsersActive] (@debug char(1)='N') AS

-- exec spAdmin_SysUsers @debug='Y'

------------------------------------------------------------
--- Check the input for SQL Injection and simply exit to deny an attacker any hints. Thanks to Jeff Moden. :)

IF @debug LIKE '%[^N,Y]' ESCAPE '_'
RETURN

------------------------------------------------------------
--- Find Windows user accounts and add missing ones to the table
--DECLARE @debug char(1)='Y'
--DECLARE @domain varchar(100) = 'OPTUMUK'--CASE WHEN @@servername LIKE '%/%' THEN LEFT(@@servername,CHARINDEX('\',@@servername,1)-1) ELSE @@servername END --'UUKP8V-MKAIV01'
DECLARE @users  TABLE(Users varchar(8000) NULL)

SET NOCOUNT ON

IF NOT EXISTS(select 1 from sys.objects where [name] = 'TblSysUsers')
BEGIN
CREATE TABLE TblSysUsers(Username varchar(255) NOT NULL
						,WindowsAccountDisabled char(1) NOT NULL
						,SQLaccountDisabled char(1) NOT NULL
						,EmailAddress varchar(255) NULL
						,SendNotifications char(1) NULL
						,CONSTRAINT [PK_TblSysUsers] PRIMARY KEY CLUSTERED 
						(Username ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						) ON [PRIMARY]
END

-----------------------------------------------------


INSERT INTO @users
select
 convert(sysname, rtrim(loginame)) as loginname
from sys.sysprocesses
group by convert(sysname, rtrim(loginame)) --as loginname

IF(@debug='Y')
BEGIN
SELECT * FROM @users
END

-----------------------------------------------------
--- Insert missing users

INSERT INTO dbo.TblSysUsers(Username,WindowsAccountDisabled,SQLaccountDisabled,SendNotifications)
SELECT 
 U.Users as Username
,CASE
	WHEN RTRIM(LEFT(Users,CHARINDEX(' ',Users,1))) = 'TRUE'
		THEN 'Y'
	ELSE 'N'
	END as WindowsAccountDisabled
,CASE 
	WHEN is_disabled = 1 THEN 'Y'
	WHEN ISNULL(is_disabled,0) = 0 THEN 'N'
	ELSE '?'
	END as SQLaccountDisabled
,'Y' as SendNotifications
FROM @users U
	LEFT OUTER JOIN sys.server_principals P
		ON U.Users = P.[name]
WHERE U.Users NOT IN(select Username from TblSysUsers)

-----------------------------------------------------
--- Update information for existing users
/*
UPDATE S
SET
 SQLaccountDisabled = 
	CASE 
	WHEN is_disabled = 1 THEN 'Y'
	WHEN ISNULL(is_disabled,0) = 0 THEN 'N'
	ELSE SQLaccountDisabled
	END 
FROM dbo.TblSysUsers S
	INNER JOIN @users U
		ON S.Username = U.Users
	INNER JOIN sys.server_principals P
		ON U.Users = P.[name]
*/
-----------------------------------------------------
--- Try to work out how users are accessing SQL and update table
/*
declare @user   sysname
declare @logins TABLE (AccountName varchar(255) NOT NULL
					  ,AccountType varchar(255) NOT NULL
					  ,Privilege   varchar(255) NOT NULL
					  ,LoginMapped varchar(255) NOT NULL
					  ,PermissionPath varchar(255) NULL
					  )
declare @command varchar(8000)
 
declare user_cursor cursor for
select Username 
from dbo.TblSysUsers

open user_cursor 
fetch next from user_cursor into @user
 
while @@fetch_status = 0
begin

DELETE
FROM @logins 

    begin try

		insert into @logins
		exec xp_logininfo @user

		update U
		set
			SQLaccountDisabled = CASE 
									WHEN is_disabled = 1 THEN 'Y'
									WHEN is_disabled = 0 THEN 'N'
									ELSE '?'
									END
		from dbo.TblSysUsers U
			inner join @logins L
				on L.AccountName = U.Username
			inner join sys.server_principals P
				on L.PermissionPath = P.[name]

    end try
    begin catch
        --Error on xproc because login doesn't exist
        --SET @command = 'drop login '+convert(varchar(255),@user)
		--exec(@command)
    end catch


    fetch next from user_cursor into @user
end
 
close user_cursor
deallocate user_cursor
*/
-----------------------------------------------------
/*
--- Following scripts drops SQL logins not in the Windows Domain. From:
https://sqlstudies.com/2013/03/01/script-to-clean-up-windows-logins-no-longer-in-ad/
*/
/*
declare recscan cursor for
select name 
from sys.server_principals
where type IN('U','G') 
	and name like @domain+'%'
 
open recscan 
fetch next from recscan into @user
 
while @@fetch_status = 0
begin

DELETE
FROM @logins 

    begin try

		insert into @logins
		exec xp_logininfo @user

    end try
    begin catch
        --Error on xproc because login doesn't exist
        SET @command = 'drop login '+convert(varchar(255),@user)
		exec(@command)
    end catch


    fetch next from recscan into @user
end
 
close recscan
deallocate recscan

IF(@debug='Y')
BEGIN
SELECT * FROM dbo.TblSysUsers
END
*/
-----------------------------------------------------
GO
