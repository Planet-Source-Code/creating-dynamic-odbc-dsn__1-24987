ODBCDSN Project

ODBCDSN Project main Idea is Adding a ODBCDSN Dynamically Inside a Program using
the Registry.

It is Written in Generic and nothing is Hard coded.

It is Generic and it has ODBCDSN.ini Configuration File.
User needs to change the Name of the Odbc What every they want to create


Configuration File Read as follows:


'Don't Change this line  MAINKEY is must
MainKey="SOFTWARE\ODBC\ODBC.INI\"
'Change the  DataSourceName What ever you want to call it
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