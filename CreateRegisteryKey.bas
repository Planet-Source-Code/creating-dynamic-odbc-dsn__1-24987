Attribute VB_Name = "CreateRegisteryKey"
Option Explicit

Public Enum REG_TOPLEVEL_KEYS
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_CURRENT_USER = &H80000001
 HKEY_DYN_DATA = &H80000006
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_USERS = &H80000003
End Enum


Private Declare Function RegCreateKey Lib _
   "advapi32.dll" Alias "RegCreateKeyA" _
   (ByVal Hkey As Long, ByVal lpSubKey As _
   String, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib _
   "advapi32.dll" (ByVal Hkey As Long) As Long

Private Declare Function RegSetValueEx Lib _
   "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal Hkey As Long, ByVal _
   lpValueName As String, ByVal _
   Reserved As Long, ByVal dwType _
   As Long, lpData As Any, ByVal _
   cbData As Long) As Long

Private Const REG_SZ = 1
'PURPOSE:  To create a registry key
'Const HKEY_CLASSES_ROOT = &H80000000
'Const HKEY_CURRENT_USER = &H80000001
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const HKEY_USERS = &H80000003

Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_INVALID_PARAMETER = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259

Const KEY_ALL_ACCESS = &H3F

Const REG_OPTION_NON_VOLATILE = 0




Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
Alias "RegCreateKeyExA" (ByVal Hkey As Long, _
ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal lpClass As String, ByVal dwOptions As Long, _
ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
phkResult As Long, lpdwDisposition As Long) As Long


'PARAMETERS:

'KeyName: Name of key to create
'ParentKey: Top level key under which the new key will be created
'For example: HKEY_CURRENT_USER
'Use one of the constants defined in delcarations

'RETURNS:  True if executes without error, false otherwise

'can also be used to create nested subkeys
'e.g. CreateKey(HKEY_LOCAL_MACHINE/Software/MyApp/Settings)

Public Function CreateKey(KeyName As String, _
ParentKey As Long) As Boolean

Dim lNewKeyHnd As Long
Dim lAns As Long

lAns = RegCreateKeyEx(ParentKey, KeyName, 0&, vbNullString, _
REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lNewKeyHnd, lAns)

RegCloseKey lNewKeyHnd

CreateKey = (lAns = 0)
 'Altenatively, return specific error inormation
 'See declarations for error codes
End Function

Private Function WriteStringToRegistry(Hkey As _
  REG_TOPLEVEL_KEYS, strPath As String, strValue As String, _
  strdata As String) As Boolean
 
'WRITES A STRING VALUE TO REGISTRY:
'PARAMETERS:

'Hkey: Top Level Key as defined by
'REG_TOPLEVEL_KEYS Enum (See Declarations)

'strPath - 'Full Path of Subkey
'if path does not exist it will be created

'strValue ValueName

'strData - Value Data

'Returns: True if successful, false otherwise

'EXAMPLE:
'WriteStringToRegistry(HKEY_LOCAL_MACHINE, _
"Software\Microsoft", "CustomerName", "FreeVBCode.com")

Dim bAns As Boolean

On Error GoTo ErrorHandler
   Dim keyhand As Long
   Dim r As Long
   r = RegCreateKey(Hkey, strPath, keyhand)
   If r = 0 Then
        r = RegSetValueEx(keyhand, strValue, 0, _
           REG_SZ, ByVal strdata, Len(strdata))
        r = RegCloseKey(keyhand)
    End If
    
   WriteStringToRegistry = (r = 0)

Exit Function

ErrorHandler:
    WriteStringToRegistry = False
    Exit Function
    
End Function

Public Sub MAIN()
Dim cKeyName As String
Dim cMainKeyName As String


On Error GoTo err_handler

'Call the MainINIStrings
MainINIStrings
cKeyName = Trim(G_DataSourceName)
cMainKeyName = Trim(G_MainKey) & Trim(cKeyName)


CreateKey cMainKeyName, HKEY_LOCAL_MACHINE
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "DatabaseName", Trim(G_DatabaseName)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "Description", Trim(G_Description)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "DriverPath", Trim(G_DriverPath)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "Trusted_Connection", Trim(G_Trusted_Connection)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "Server", Trim(G_Server)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "DriverName", Trim(G_DriverName)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "QuotedId", Trim(G_QuotedId)
WriteStringToRegistry HKEY_LOCAL_MACHINE, cMainKeyName, "LastUser", Trim(G_LastUser)
MsgBox Trim(G_DataSourceName) & " ODBC Data Source Successfully Created"
Exit Sub
err_handler:
    MsgBox "Error in creating ODBC Data Source " & Trim(G_DataSourceName) & " - Error  Details " & Err.Description


End Sub
 Private Sub Command1_Click()

   Dim DataSourceName As String
   Dim DatabaseName As String
   Dim Description As String
   Dim DriverPath As String
   Dim DriverName As String
   Dim LastUser As String
   Dim Regional As String
   Dim Server As String

   Dim lResult As Long
   Dim hKeyHandle As Long

   'Specify the DSN parameters.

   DataSourceName = "<the name of your new DSN>"
   DatabaseName = "<name of the database to be accessed by the new DSN>"
   Description = "<a description of the new DSN>"
   DriverPath = "<path to your SQL Server driver>"
   LastUser = "<default user ID of the new DSN>"
   Server = "<name of the server to be accessed by the new DSN>"
   DriverName = "SQL Server"

   'Create the new DSN key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

   'Set the values of the new DSN key.

   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
      ByVal DatabaseName, Len(DatabaseName))
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal DriverPath, Len(DriverPath))
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal LastUser, Len(LastUser))
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
      ByVal Server, Len(Server))

   'Close the new DSN key.

   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)

   End Sub

