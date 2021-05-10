USE [BIA_DEV]
GO

/****** Object:  StoredProcedure [stk].[uspExportToCSV]    Script Date: 10-5-2021 10:05:13 ******/
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

CREATE   PROCEDURE [stk].[uspExportToCSV]
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

/****** Object:  StoredProcedure [stk].[uspSendMail]    Script Date: 10-5-2021 10:05:14 ******/
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

CREATE   PROCEDURE [stk].[uspSendMail]
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


