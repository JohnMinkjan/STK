USE [BIA_DEV]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Strip string from everything but 
--				[AZaz09]
-- V0.1	
--		base checks
-- =============================================
CREATE OR ALTER FUNCTION [stk].[fnCleanString] 
(	
	 @string varchar(1000)
	,@validchar varchar(255)=null
)
RETURNS nvarchar(2000)
AS
BEGIN
	-- Declare the return variable here:
	DECLARE @Result nvarchar(2000)
	
	Declare @fromNumber int=1,
	  @toNumber int=100,
	  @byStep int=1
	 
	-- List of allowed characters:
	if @validchar is null
		begin		
			set @validchar='[-a-z0-9 ]'
		end

	-- CTE with the calculation of string's length:
	;WITH CTE AS (
	  SELECT @fromNumber AS i
	  UNION ALL
	  SELECT i + @byStep
	  FROM CTE
	  WHERE
	  (i + @byStep) <= @toNumber
	)
	
	-- Cleans the string:
	SELECT @Result=cast(cast((select substring(@string,CTE.i,1)
	FROM CTE
	WHERE CTE.i <= len(@string)	
		and substring(@string,CTE.i,1) like @validchar for xml path('')) as xml)as varchar(max))

	RETURN @Result
END
GO

/****** Object:  UserDefinedFunction [stk].[fnCleanStringAccents]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Strips string from accents 
-- V0.1	
--		base checks
-- =============================================
CREATE OR ALTER FUNCTION [stk].[fnCleanStringAccents] 
(	
	 @string varchar(1000)	
)
RETURNS nvarchar(2000)
AS
BEGIN
	RETURN( @string Collate SQL_Latin1_General_CP1253_CI_AI)
END
GO

/****** Object:  UserDefinedFunction [stk].[fnConvertUnixToDateTime]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Convert Unix time to DateTime
-- V0.1	
--		Unix Seconds
-- =============================================


CREATE OR ALTER FUNCTION [stk].[fnConvertUnixToDateTime] 
	(@Datetime BIGINT)
RETURNS DATETIME
AS
BEGIN
 --   DECLARE @LocalTimeOffset BIGINT
 --          ,@AdjustedLocalDatetime BIGINT;
 --   SET @LocalTimeOffset = DATEDIFF(second,GETDATE(),GETUTCDATE())
 --   SET @AdjustedLocalDatetime = @Datetime - @LocalTimeOffset
    RETURN (SELECT DATEADD(second,@Datetime, CAST('1970-01-01 00:00:00' AS datetime)))
END;
GO

/****** Object:  UserDefinedFunction [stk].[fnConvertUnixToDateTimeMS]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Convert Unix time to DateTime
-- V0.1	
--		Unix MiliSeconds
-- =============================================


CREATE OR ALTER FUNCTION [stk].[fnConvertUnixToDateTimeMS] 
	(@Datetime BIGINT)
RETURNS DATETIME
AS
BEGIN
	DECLARE @dtOut DATETIME
	DECLARE @ms INT = CAST(RIGHT( @Datetime, 3) AS INT)
	SET @Datetime = @Datetime / 1000
	SET @dtOut = DATEADD(SECOND,@Datetime, CAST('1970-01-01 00:00:00' AS datetime))
	SET @dtOut = DATEADD(MILLISECOND,@ms,@dtOut)
    RETURN @dtOut
END;
GO

/****** Object:  UserDefinedFunction [stk].[fnCountChar]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Count the number of occurrences
--				of a character in a string
--				CaSe SenSitive
-- V0.1	
--		base checks
-- =============================================
CREATE OR ALTER FUNCTION [stk].[fnCountChar] 
	( @pInput VARCHAR(1000), @pSearchChar CHAR(1) )
RETURNS INT
BEGIN
	RETURN (LEN(@pInput) - LEN(REPLACE(	@pInput      COLLATE SQL_Latin1_General_Cp1_CS_AS, 
										@pSearchChar COLLATE SQL_Latin1_General_Cp1_CS_AS, '')))

END
GO

/****** Object:  UserDefinedFunction [stk].[fnExcelToDateTime]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


--Testing/Usage
--SELECT [stk].[fnInitCap] ('the quick brown fox jumps over the lazy dog..')

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210315
-- Description:	Convert Excel Date to DateTime
-- V0.1	
--		base checks
-- =============================================


CREATE OR ALTER FUNCTION [stk].[fnExcelToDateTime]
	(@ExcelDateTime FLOAT)
RETURNS DATETIME
BEGIN
	DECLARE @DT DATETIME
	DECLARE @MS BIGINT

	SET @MS = ((@ExcelDateTime - FLOOR(@ExcelDateTime))/1.0) * ( 24 * 60 * 60 * 1000)

	SET @DT = dateadd(d,@ExcelDateTime,'1899-12-30')
	SET @DT = dateadd(MILLISECOND,@MS,@DT)

	RETURN @DT
END

GO

/****** Object:  UserDefinedFunction [stk].[fnInitCap]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Every word starts with a Capital
-- V0.1	
--		base checks
-- =============================================
CREATE OR ALTER FUNCTION [stk].[fnInitCap] 
	( @InputString varchar(4000) ) 
RETURNS VARCHAR(4000)
AS
BEGIN

	DECLARE @Index          INT
	DECLARE @Char           CHAR(1)
	DECLARE @PrevChar       CHAR(1)
	DECLARE @OutputString   VARCHAR(255)

	SET @OutputString = LOWER(@InputString)
	SET @Index = 1

	WHILE @Index <= LEN(@InputString)
	BEGIN
		SET @Char     = SUBSTRING(@InputString, @Index, 1)
		SET @PrevChar = CASE WHEN @Index = 1 THEN ' '
							 ELSE SUBSTRING(@InputString, @Index - 1, 1)
						END

		IF @PrevChar IN (' ', ';', ':', '!', '?', ',', '.', '_', '-', '/', '&', '''', '(')
		BEGIN
			IF @PrevChar != '''' OR UPPER(@Char) != 'S'
				SET @OutputString = STUFF(@OutputString, @Index, 1, UPPER(@Char))
		END

		SET @Index = @Index + 1
	END

	RETURN @OutputString

END
GO

/****** Object:  UserDefinedFunction [stk].[fnStripHTML]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210510
-- Description:	Strips string from HTML
-- V0.1	
--		base checks
-- =============================================
CREATE OR ALTER FUNCTION [stk].[fnStripHTML]
(	
	@HTMLText varchar(MAX)
)
RETURNS varchar(MAX)
AS
BEGIN
	DECLARE @Start  int
	DECLARE @End    int
	DECLARE @Length int

	set @HTMLText = replace(@htmlText, '<br>',CHAR(13) + CHAR(10))
	set @HTMLText = replace(@htmlText, '<br/>',CHAR(13) + CHAR(10))
	set @HTMLText = replace(@htmlText, '<br />',CHAR(13) + CHAR(10))
	set @HTMLText = replace(@htmlText, '<li>','- ')
	set @HTMLText = replace(@htmlText, '</li>',CHAR(13) + CHAR(10))

	set @HTMLText = replace(@htmlText, '&rsquo;' collate Latin1_General_CS_AS, ''''  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&quot;' collate Latin1_General_CS_AS, '"'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&amp;' collate Latin1_General_CS_AS, '&'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&euro;' collate Latin1_General_CS_AS, '€'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&lt;' collate Latin1_General_CS_AS, '<'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&gt;' collate Latin1_General_CS_AS, '>'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&oelig;' collate Latin1_General_CS_AS, 'oe'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&nbsp;' collate Latin1_General_CS_AS, ' '  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&copy;' collate Latin1_General_CS_AS, '©'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&laquo;' collate Latin1_General_CS_AS, '«'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&reg;' collate Latin1_General_CS_AS, '®'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&plusmn;' collate Latin1_General_CS_AS, '±'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&sup2;' collate Latin1_General_CS_AS, '²'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&sup3;' collate Latin1_General_CS_AS, '³'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&micro;' collate Latin1_General_CS_AS, 'µ'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&middot;' collate Latin1_General_CS_AS, '·'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ordm;' collate Latin1_General_CS_AS, 'º'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&raquo;' collate Latin1_General_CS_AS, '»'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&frac14;' collate Latin1_General_CS_AS, '¼'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&frac12;' collate Latin1_General_CS_AS, '½'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&frac34;' collate Latin1_General_CS_AS, '¾'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Aelig' collate Latin1_General_CS_AS, 'Æ'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Ccedil;' collate Latin1_General_CS_AS, 'Ç'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Egrave;' collate Latin1_General_CS_AS, 'È'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Eacute;' collate Latin1_General_CS_AS, 'É'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Ecirc;' collate Latin1_General_CS_AS, 'Ê'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&Ouml;' collate Latin1_General_CS_AS, 'Ö'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&agrave;' collate Latin1_General_CS_AS, 'à'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&acirc;' collate Latin1_General_CS_AS, 'â'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&auml;' collate Latin1_General_CS_AS, 'ä'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&aelig;' collate Latin1_General_CS_AS, 'æ'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ccedil;' collate Latin1_General_CS_AS, 'ç'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&egrave;' collate Latin1_General_CS_AS, 'è'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&eacute;' collate Latin1_General_CS_AS, 'é'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ecirc;' collate Latin1_General_CS_AS, 'ê'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&euml;' collate Latin1_General_CS_AS, 'ë'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&icirc;' collate Latin1_General_CS_AS, 'î'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ocirc;' collate Latin1_General_CS_AS, 'ô'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ouml;' collate Latin1_General_CS_AS, 'ö'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&divide;' collate Latin1_General_CS_AS, '÷'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&oslash;' collate Latin1_General_CS_AS, 'ø'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ugrave;' collate Latin1_General_CS_AS, 'ù'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&uacute;' collate Latin1_General_CS_AS, 'ú'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&ucirc;' collate Latin1_General_CS_AS, 'û'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&uuml;' collate Latin1_General_CS_AS, 'ü'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&quot;' collate Latin1_General_CS_AS, '"'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&amp;' collate Latin1_General_CS_AS, '&'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&lsaquo;' collate Latin1_General_CS_AS, '<'  collate Latin1_General_CS_AS)
	set @HTMLText = replace(@htmlText, '&rsaquo;' collate Latin1_General_CS_AS, '>'  collate Latin1_General_CS_AS)


	-- Remove anything between <STYLE> tags
	SET @Start = CHARINDEX('<STYLE', @HTMLText)
	SET @End = CHARINDEX('</STYLE>', @HTMLText, CHARINDEX('<', @HTMLText)) + 7
	SET @Length = (@End - @Start) + 1

	WHILE (@Start > 0 AND @End > 0 AND @Length > 0) BEGIN
	SET @HTMLText = STUFF(@HTMLText, @Start, @Length, '')
	SET @Start = CHARINDEX('<STYLE', @HTMLText)
	SET @End = CHARINDEX('</STYLE>', @HTMLText, CHARINDEX('</STYLE>', @HTMLText)) + 7
	SET @Length = (@End - @Start) + 1
	END

	-- Remove anything between <whatever> tags
	SET @Start = CHARINDEX('<', @HTMLText)
	SET @End = CHARINDEX('>', @HTMLText, CHARINDEX('<', @HTMLText))
	SET @Length = (@End - @Start) + 1

	WHILE (@Start > 0 AND @End > 0 AND @Length > 0) BEGIN
	SET @HTMLText = STUFF(@HTMLText, @Start, @Length, '')
	SET @Start = CHARINDEX('<', @HTMLText)
	SET @End = CHARINDEX('>', @HTMLText, CHARINDEX('<', @HTMLText))
	SET @Length = (@End - @Start) + 1
	END

	RETURN LTRIM(RTRIM(@HTMLText))

END
GO

/****** Object:  UserDefinedFunction [stk].[fnValidateEmail]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Check if email adress is valid
-- V0.1	
--		base checks
-- =============================================

CREATE OR ALTER FUNCTION [stk].[fnValidateEmail] 
	(@email VARCHAR(255))
RETURNS bit
AS
BEGIN
RETURN
(
	SELECT
	CASE
		WHEN	@Email is null then 0	                	--NULL Email is invalid
		WHEN	CHARINDEX(' ', @email) 	<> 0 or		--Check for invalid character
				CHARINDEX('/', @email) 	<> 0 or --Check for invalid character
				CHARINDEX(':', @email) 	<> 0 or --Check for invalid character
				CHARINDEX(';', @email) 	<> 0 then 0 --Check for invalid character
		WHEN	LEN(@Email)-1 <= charindex('.', @Email) then 0--check for '%._' at end of string
		WHEN 	@Email like '%@%@%'or
			   	@Email Not Like '%@%.%'  then 0--Check for duplicate @ or invalid format
		ELSE	1
	END
)
END
GO

/****** Object:  UserDefinedFunction [stk].[fnValidateURL]    Script Date: 10-5-2021 10:02:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Check if url has a respons
-- V0.1	
--		base checks
-- =============================================


CREATE OR ALTER FUNCTION [stk].[fnValidateURL]
	(@URL VARCHAR(300))
RETURNS BIT
AS 
BEGIN
	DECLARE @Object INT
	DECLARE @Return TINYINT
	DECLARE @Valid BIT SET @Valid = 0 --default to false
	
--create the XMLHTTP object
	EXEC @Return = sp_oacreate 'MSXML2.ServerXMLHTTP.3.0', @Object OUTPUT
	
	IF @Return = 0
	BEGIN
	   DECLARE @Method VARCHAR(350)		--define setTimeouts method --Resolve, Connect, Send, Receive
	   SET @Method = 'setTimeouts(45000, 45000, 45000, 45000)'	--set the timeouts
	   EXEC @Return = sp_oamethod @Object, @Method
	   
		IF @Return = 0
			BEGIN			--define open method
				SET @Method = 'open("GET", "' + @URL + '", false)'--Open the connection
				EXEC @Return = sp_oamethod @Object, @Method
			END			
		
		IF @Return = 0
			BEGIN			--SEND the request
				EXEC @Return = sp_oamethod @Object, 'send()'
			END
		IF @Return = 0
			BEGIN
				DECLARE @Output INT
				EXEC @Return = sp_oamethod @Object, 'status', @Output OUTPUT
				
				IF @Output = 200
					BEGIN
						SET @Valid = 1
					END
			END
		END		--destroy the object	EXEC sp_oadestroy @Object
	RETURN (@Valid)
END
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20210314
-- Description:	Create a Validation code
-- V0.1	
--	Default 6 characters
-- =============================================

CREATE OR ALTER FUNCTION [stk].[fnValidationCode]
(
	@size INT = 6
)
RETURNS VARCHAR(50)
AS
BEGIN
	DECLARE @source VARCHAR(50)= 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
	DECLARE @out VARCHAR(50) = ''
	DECLARE @ct INT = 0
	DECLARE @pos INT

	WHILE @ct < @size
	BEGIN
		SET @pos = (select CAST(CEILING(random_value*36) AS INT) from [stk].[vwRandomVal]) 
		SET @out = @out + SUBSTRING(@source, @pos,1);
		SET @ct = @ct + 1;
	END

	RETURN @out
END;
GO

-- =============================================
-- Author:		John Minkjan
-- Create date: 20220117
-- Description:	Get ISO YearWeek
-- V0.1	
--		Default Current date
-- =============================================


CREATE OR ALTER FUNCTION [stk].[fnISOYearWeek]
(
	@dtDate DATETIME = NULL
)
RETURNS INT
AS
BEGIN
	IF @dtDate IS NULL 
		SET @dtDate = GETDATE()

    DECLARE @ISOYearWeek INT

	SET @ISOYearWeek = (	 
	SELECT CASE
		WHEN DATEPART(ISO_WEEK, @dtDate) > 50 AND MONTH(@dtDate) = 1 THEN YEAR(@dtDate) - 1
		WHEN DATEPART(ISO_WEEK, @dtDate) = 1 AND MONTH(@dtDate) = 12 THEN YEAR(@dtDate) + 1
		ELSE YEAR(@dtDate) END ) * 100 +
	    DATEPART(ISO_WEEK,	@dtDate)

	RETURN @ISOYearWeek
END;
GO

--SELECT [stk].[fnISOYearWeek]('01-JAN-2021')