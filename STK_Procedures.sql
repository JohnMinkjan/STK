USE [BIA_DEV]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210511
-- Description:	Check SQL Services
-- =============================================
CREATE OR ALTER PROCEDURE [stk].[uspCheckSQLServices]
AS
BEGIN

	SET NOCOUNT ON;
    DECLARE @LoadDT DATETIME = GETDATE();
	DECLARE @WINSCCMD TABLE (ID INT IDENTITY (1,1) PRIMARY KEY NOT NULL, Line VARCHAR(MAX))

	IF OBJECT_ID(N'tempdb..#ServiceStates') IS NOT NULL
	BEGIN
		DROP TABLE #ServiceStates
	END
 
	INSERT INTO @WINSCCMD(Line) EXEC master.dbo.xp_cmdshell 'sc queryex type= service state= all'
 
	SELECT  @LoadDT											AS LoadDT	
			, @@SERVERNAME									AS [ServerName]
			, ltrim(rtrim (SUBSTRING (W1.Line, 15, 100)))	AS ServiceName
			, ltrim(rtrim (SUBSTRING (W2.Line, 15, 100)))	AS DisplayName
			, ltrim(rtrim (SUBSTRING (W3.Line, 33, 100)))	AS ServiceState
			INTO #ServiceStates
	FROM @WINSCCMD W1, @WINSCCMD W2, @WINSCCMD W3
	WHERE W1.ID = W2.ID - 1 AND
			W3.ID - 3 = W1.ID AND
			LTRIM(RTRIM (LOWER (SUBSTRING (W3.Line, 33, 100)))) in ('RUNNING','STOPPED')
	ORDER BY 2

	DECLARE @StoppedServices INT

	SELECT * FROM #ServiceStates
	WHERE ServiceName IN 
	('MSOLAP$TABULAR'
	,'MSSQLSERVER'
	,'MSSQLServerOLAPService'
	,'SQLServerReportingServices'
	,'MSOLAP$TABULAR'
	,'SSASTELEMETRY'
	,'SSASTELEMETRY$TABULAR'
	,'SQLBrowser'
	,'SQLTELEMETRY'
	,'MsDtsServer150'
	,'SSISTELEMETRY150'
	,'MSSQLLaunchpad'
	,'SQLWriter')
	--AND ServiceState = 'STOPPED'

END
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		John Minkjan
-- Create date: 20210326
-- Description:	Export to CSV
-- V0.1	
--		base checks
-- =============================================

CREATE OR ALTER PROCEDURE [stk].[uspExportToCSV]
(
	 @ExportQuery		VARCHAR(2000)
	,@ExportFileName	VARCHAR(250)
	,@DatabaseServer	VARCHAR(250) NULL
)
AS
BEGIN 
	DECLARE @pShell VARCHAR(4000)

	IF (@DatabaseServer IS NULL OR @DatabaseServer = '')
	BEGIN
		SET @DatabaseServer = @@SERVERNAME
	END

	SET @pshell = ' 
$filename = "'+@ExportFileName+'"
Invoke-Sqlcmd -Query '''+@ExportQuery+''' -ServerInstance "'+@DatabaseServer+'" | Export-Csv -Path "$filename" -NoTypeInformation'

	DECLARE @File  varchar(300) = 'c:\Temp\STKExportToCSV.ps1'
	DECLARE @Text  varchar(8000) = @pshell
	DECLARE @OLE            INT 
	DECLARE @FileID         INT

	EXECUTE sp_OACreate 'Scripting.FileSystemObject', @OLE OUT 
	EXECUTE sp_OAMethod @OLE, 'OpenTextFile', @FileID OUT, @File, 2, 1 
	EXECUTE sp_OAMethod @FileID, 'WriteLine', Null, @Text
	EXECUTE sp_OADestroy @FileID 
	EXECUTE sp_OADestroy @OLE 

	EXEC MASTER..xp_cmdshell 'powershell.exe -executionpolicy unrestricted  c:\Temp\STKExportToCSV.ps1'

END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		John Minkjan
-- Create date: 20210326
-- Description:	simple Mail Send procedure
-- V0.1	
--		base checks
-- =============================================

CREATE OR ALTER PROCEDURE [stk].[uspSendMail]
(
	@recipients VARCHAR(500),
	@copy_recipients VARCHAR(500) NULL,
	@subject VARCHAR(500),
	@body VARCHAR (4000)
	
)
AS
BEGIN

	EXEC msdb.dbo.sp_send_dbmail
	@profile_name= 'SQL Toolkit Public Profile',
	@recipients = @recipients,
	@copy_recipients = @copy_recipients,
	@subject = @subject,
	@body =@body
END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210511
-- Description:	Uptime Server 
-- =============================================
CREATE OR ALTER PROCEDURE [stk].[uspServerUptime]
AS
BEGIN

	DECLARE @cmd nvarchar(255) = 'systeminfo|find "Time:"'

	CREATE TABLE #output 
		(id int identity(1,1)
		, feedback nvarchar(255) null)

	INSERT #output (feedback) 
	EXEC [MASTER]..xp_cmdshell @cmd
	SELECT 
		CONVERT(DATETIME, 
			LTRIM(RTRIM(REPLACE(REPLACE(feedback,'System Boot Time:',''),',','')))
		,101)  as server_start_time
		,CONVERT(VARCHAR,  
			GETDATE() -
			CONVERT(DATETIME, 
				LTRIM(RTRIM(REPLACE(REPLACE(feedback,'System Boot Time:',''),',','')))
			,101) 
		,108) AS server_up_time	
		FROM #output 
	WHERE feedback is not null 

	DROP TABLE #output

END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210511
-- Description:	Uptime SQL Server Instance
-- =============================================
CREATE OR ALTER PROCEDURE [stk].[uspSQLServerUptime]
AS
BEGIN
	SELECT 
	   sqlserver_start_time
	 , Convert(varchar,  GETDATE() -sqlserver_start_time, 108)  sqlserver_up_time
	FROM sys.dm_os_sys_info;
END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20220117
-- Description:	Create missing indexes
-- Based on the work of Pinal Dave 
-- =============================================
CREATE OR ALTER PROCEDURE [stk].[uspCreateMissingIndexes]
(
	@DBName VARCHAR(200) = NULL,
	@IDXFileGroup VARCHAR(200) = 'INDEX',
	@IDXIndexHitSince INT = 3,
	@IDXUserImpactTreshold INT = 100,
	@IDXHitCountTreshold INT = 25,
	@PrintOnly BIT = 0
)
AS
BEGIN
	IF @DBName IS NULL
		SET @DBName = (SELECT DB_NAME())

	DECLARE @SQLCreateStatement NVARCHAR(MAX);

	DECLARE C CURSOR FOR
		SELECT 
		--dm_mid.database_id AS DatabaseID,
		--dm_migs.avg_user_impact*(dm_migs.user_seeks+dm_migs.user_scans) Avg_Estimated_Impact,
		--dm_migs.user_seeks+dm_migs.user_scans as SeeksAndScans,
		--dm_migs.last_user_seek AS Last_User_Seek,
		--right('000' + CAST(ABS(CHECKSUM(NewId())) % 1000 as varchar(3)),3),
		--OBJECT_NAME(dm_mid.OBJECT_ID,dm_mid.database_id) AS [TableName],
		'CREATE INDEX [IX_' + OBJECT_NAME(dm_mid.OBJECT_ID,dm_mid.database_id) + '_'
		+ REPLACE(REPLACE(REPLACE(ISNULL(dm_mid.equality_columns,''),', ','_'),'[',''),']','') 
		+ CASE
		WHEN dm_mid.equality_columns IS NOT NULL
		AND dm_mid.inequality_columns IS NOT NULL THEN '_'
		ELSE ''
		END
		+ REPLACE(REPLACE(REPLACE(ISNULL(dm_mid.inequality_columns,''),', ','_'),'[',''),']','')
		+ '_'
		+ CONVERT(VARCHAR(8), GETDATE(),112)
		+ '_'
		+ right('000' + CAST(ABS(CHECKSUM(NewId())) % 1000 as varchar(3)),3)
		+ ']'
		+ ' ON ' + dm_mid.statement
		+ ' (' + ISNULL (dm_mid.equality_columns,'')
		+ CASE WHEN dm_mid.equality_columns IS NOT NULL AND dm_mid.inequality_columns 
		IS NOT NULL THEN ',' ELSE
		'' END
		+ ISNULL (dm_mid.inequality_columns, '')
		+ ')'
		+ ISNULL (' INCLUDE (' + dm_mid.included_columns + ')', '')
		+ ' ON ['+@IDXFileGroup+']'
		 AS Create_Statement
		FROM sys.dm_db_missing_index_groups dm_mig
		INNER JOIN sys.dm_db_missing_index_group_stats dm_migs
		ON dm_migs.group_handle = dm_mig.index_group_handle
		INNER JOIN sys.dm_db_missing_index_details dm_mid
		ON dm_mig.index_handle = dm_mid.index_handle
		INNER JOIN sys.databases db on dm_mid.database_id = db.database_id
		WHERE (1=1)
		AND db.[name] = @DBName
		and dm_migs.last_user_seek >= GETDATE() - @IDXIndexHitSince
		and dm_migs.avg_user_impact*(dm_migs.user_seeks+dm_migs.user_scans) > @IDXUserImpactTreshold
		and dm_migs.user_seeks+dm_migs.user_scans >  @IDXHitCountTreshold 
		--ORDER BY dm_migs.user_seeks+dm_migs.user_scans desc, Avg_Estimated_Impact DESC

	OPEN C
	FETCH NEXT FROM C INTO @SQLCreateStatement

	WHILE @@FETCH_STATUS = 0 
		BEGIN 
			PRINT @SQLCreateStatement

			IF @PrintOnly = 0
			BEGIN 
				EXEC (@SQLCreateStatement)
			END 
			FETCH NEXT FROM C INTO @SQLCreateStatement
		END
	CLOSE C
	DEALLOCATE C
END 
GO 


-- =============================================
-- Author:		John Minkjan
-- Create date: 20220117
-- Description:	Delete Unused indexes
-- Based on the work of Pinal Dave 
-- =============================================
CREATE OR ALTER PROCEDURE [stk].[uspDeleteUnusedIndexes]
(
	@DBName VARCHAR(200) = NULL,
	@IDXUserUpdatesTreshold INT = 1000,
	@PrintOnly BIT = 0
)
AS
BEGIN
	IF @DBName IS NULL
		SET @DBName = (SELECT DB_NAME())

	DECLARE @SQLCreateStatement NVARCHAR(MAX);

	DECLARE C CURSOR FOR
		SELECT
			--  o.name AS ObjectName
			--, i.name AS IndexName
			--, i.index_id AS IndexID
			--, dm_ius.user_seeks AS UserSeek
			--, last_user_seek AS Last_User_Seek
			--, dm_ius.user_scans AS UserScans
			--, dm_ius.user_lookups AS UserLookups
			--, dm_ius.user_updates AS UserUpdates
			--, p.TableRows,
			  'DROP INDEX ' + QUOTENAME(i.name)
			+ ' ON ' + QUOTENAME(s.name) + '.'
			+ QUOTENAME(OBJECT_NAME(dm_ius.OBJECT_ID)) AS 'drop statement'

			FROM sys.dm_db_index_usage_stats dm_ius
			INNER JOIN sys.indexes i ON i.index_id = dm_ius.index_id 
			AND dm_ius.OBJECT_ID = i.OBJECT_ID
			INNER JOIN sys.objects o ON dm_ius.OBJECT_ID = o.OBJECT_ID
			INNER JOIN sys.schemas s ON o.schema_id = s.schema_id
			INNER JOIN sys.databases db on dm_ius.database_id = db.database_id
			INNER JOIN (SELECT SUM(p.rows) TableRows, p.index_id, p.OBJECT_ID
			FROM sys.partitions p GROUP BY p.index_id, p.OBJECT_ID) p
			ON p.index_id = dm_ius.index_id AND dm_ius.OBJECT_ID = p.OBJECT_ID
			WHERE  (1=1)
			AND db.[name] = @DBName
			AND OBJECTPROPERTY(dm_ius.OBJECT_ID,'IsUserTable') = 1
			--AND dm_ius.database_id = DB_ID()
			AND i.type_desc = 'nonclustered'
			AND i.is_primary_key = 0
			AND i.is_unique_constraint = 0
			and  dm_ius.user_updates > @IDXUserUpdatesTreshold
			and (dm_ius.user_seeks + dm_ius.user_scans) = 0
			--ORDER BY (dm_ius.user_seeks + dm_ius.user_scans + dm_ius.user_lookups) ASC

	OPEN C
	FETCH NEXT FROM C INTO @SQLCreateStatement

	WHILE @@FETCH_STATUS = 0 
		BEGIN 
			PRINT @SQLCreateStatement

			IF @PrintOnly = 0
			BEGIN 
				EXEC (@SQLCreateStatement)
			END 
			FETCH NEXT FROM C INTO @SQLCreateStatement
		END
	CLOSE C
	DEALLOCATE C
END 
GO 