<div align="center">

## Creating Dynamic ODBC DSN

<img src="PIC20017131428452588.jpg">
</div>

### Description

ODBCDSN Project main Idea is Adding a ODBCDSN Dynamically Inside a Program using

the Registry.

It is Written in Generic and nothing is Hard coded.

It is Generic and it has ODBCDSN.ini Configuration File.

User needs to change the Name of the Odbc What every they want to create
 
### More Info
 


User needs to change the Name of the Odbc What every they want to create

Configuration File Read as follows:

'Don't Change this line MAINKEY is must

MainKey="SOFTWARE\ODBC\ODBC.INI\"

'Change the DataSourceName What ever you want to call it

DataSourceName = "MIDAS PRODUCTION"

'Type the DatabaseName

DatabaseName = "MIDAS"

'Type the Description

Description = "MIDAS ON CLUSTER VIRTUAL SQL SERVER TROYSQL"

'Type the SqlServer Name Example TROYSQL, THEN TYPE \\TROYSQL\PIPE\SQL\QUERY

DriverPath = "\\TROYSQL\pipe\sql\query"

'If you want Trusted Connection then Say Yes Otherwise No

Trusted_Connection= "Yes"

'Type the User who you want to connect Say Train

LastUser = "TRAIN"

'Type the Server Name

Server = "TROYSQL"

'Type the Driver Name SQl Server

DriverName = "SQL Server"

'Type Whether you need QuoteId Yes or No

QuotedId = "No"


<span>             |<span>
---                |---
**Submitted On**   |2001-07-13 14:24:00
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Creating D227167132001\.zip](https://github.com/Planet-Source-Code/creating-dynamic-odbc-dsn__1-24987/archive/master.zip)








