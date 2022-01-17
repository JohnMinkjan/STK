# STK
SQL Toolkit

This is a set of functions and procedures I collected over the years. 
Just some handy stuff everybody can use. 
Most of these are adapted from things I found on the net.

## Functions

### [stk].[fnCleanString]
-- Strip string from everything but [AZaz09]
- @string varchar(1000)
- @validchar varchar(255)=null

### [stk].[fnCleanStringAccents]
-- Strips string from accents
- @string varchar(1000)

[stk].[fnConvertUnixToDateTime] 

[stk].[fnConvertUnixToDateTimeMS] 

[stk].[fnCountChar] 

[stk].[fnExcelToDateTime]

[stk].[fnInitCap] 

[stk].[fnStripHTML]

[stk].[fnValidateEmail] 

[stk].[fnValidateURL]

[stk].[fnValidationCode]

[stk].[fnISOYearWeek]\

[stk].[fnTrimLeadingCharacters]\
--Remove all leading characters, default '0'\
@strIn VARCHAR(500),\
@LeadingCharacter CHAR(1) = '0'


## Procedures

[stk].[uspCheckSQLServices]

[stk].[uspExportToCSV]

[stk].[uspSendMail]

[stk].[uspServerUptime]

[stk].[uspSQLServerUptime]

[stk].[uspCreateMissingIndexes]

