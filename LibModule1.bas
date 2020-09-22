Attribute VB_Name = "LibModule1"
Option Explicit

' ============================================================
' GLOBAL DEFINITIONS FOR GENERAL PURPOSE FUNCTIONS
' ============================================================

Global Const PJ_Centered = 1
Global Const PJ_Right = 2
Global Const PJ_Left = 3

Global G_PrtCellHeight As Integer
Global G_PrtCellWidth As Integer

Public Const OFS_MAXPATHNAME = 128

Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long


' ============================================================
' dcl 2/28/96 the alias "GetClipboardDataA", as listed in the
' VB API Guide does not seem to be registered in the user32
' library so I removed it from the definition here

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long

Public Const SM_CMETRICS = 44
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXBORDER = 5
Public Const SM_CXCURSOR = 13
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CXFRAME = 32
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CXHSCROLL = 21
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CXICONSPACING = 38
Public Const SM_CXMIN = 28
Public Const SM_CXMINTRACK = 34
Public Const SM_CXSCREEN = 0
Public Const SM_CXSIZE = 30
Public Const SM_CXVSCROLL = 2
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const SM_CYCURSOR = 14
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CYFRAME = 33
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYHSCROLL = 3
Public Const SM_CYICON = 12
Public Const SM_CYICONSPACING = 39
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_CYMENU = 15
Public Const SM_CYMIN = 29
Public Const SM_CYMINTRACK = 35
Public Const SM_CYSCREEN = 1
Public Const SM_CYSIZE = 31
Public Const SM_CYVSCROLL = 20
Public Const SM_CYVTHUMB = 9
Public Const SM_DBCSENABLED = 42
Public Const SM_DEBUG = 22
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_MOUSEPRESENT = 19
Public Const SM_PENWINDOWS = 41
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_SWAPBUTTON = 23

Public G_DataSourceName As String
Public G_DatabaseName  As String
Public G_Description  As String
Public G_DriverPath  As String
Public G_Trusted_Connection  As String
Public G_Server As String
Public G_DriverName  As String
Public G_QuotedId  As String
Public G_LastUser  As String
Public G_MainKey  As String
Function F_Exists(p_fname As String) As Integer
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

F_Exists = (Dir(p_fname) <> "")

End Function


Function ProfilePrivateGetStr(p_sec As String, p_key As String, p_def As String) As String
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim l_apiret As Integer
Dim l_ininame As String
Dim l_returned As String
Dim l_size As Integer

' ---------------------------------------------------------------
'Set up parameters for writing the private profile string
' ---------------------------------------------------------------

l_ininame = App.Path & "\" & App.EXEName & ".INI"
If F_Exists(l_ininame) Then
    l_size = 255
    l_returned = Space(l_size)

    ' ---------------------------------------------------------------
    'Read profile string
    ' ---------------------------------------------------------------

    l_apiret = GetPrivateProfileString(p_sec, p_key, p_def, l_returned, l_size, l_ininame)

    ' ---------------------------------------------------------------
    'Trim string at terminating 0 byte
    ' ---------------------------------------------------------------

    l_returned = Left$(l_returned, InStr(l_returned, Chr$(0)) - 1)

    ' ---------------------------------------------------------------
    ' return the profile string
    ' ---------------------------------------------------------------

    ProfilePrivateGetStr = Trim$(l_returned)
Else
    MsgBox "Initialization File does not exist"
End If

End Function

Function ProfilePrivatePutStr(p_sec As String, p_key As String, p_usrstr As String) As Integer
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim l_ininame As String

' ---------------------------------------------------------------
'Set up parameters for writing the private profile string
' ---------------------------------------------------------------

l_ininame = App.Path & "\" & App.EXEName & ".INI"

' ---------------------------------------------------------------
'Replace string in file
' ---------------------------------------------------------------

If WritePrivateProfileString(p_sec, p_key, p_usrstr, l_ininame) <> 0 Then
    ProfilePrivatePutStr = True
Else
    ProfilePrivatePutStr = False
End If

End Function

Sub ShowBottomRight(F2Show As Form, Fmode As Integer)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim sxdim As Integer
Dim sydim As Integer

sxdim = GetSystemMetrics(SM_CXFULLSCREEN) * Screen.TwipsPerPixelX
sydim = GetSystemMetrics(SM_CYFULLSCREEN)
sydim = sydim + GetSystemMetrics(SM_CYCAPTION)
sydim = sydim * Screen.TwipsPerPixelY

F2Show.Left = (sxdim - F2Show.Width)
F2Show.Top = (sydim - F2Show.Height)

F2Show.Show Fmode
'If Fmode = vbModal Then
' Do
'  DoEvents
'Loop Until F2Show.Visible = False
'End If

End Sub

Sub ShowCenteredDesk(F2Show As Form, Fmode As Integer)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim sxdim As Integer
Dim sydim As Integer

sxdim = GetSystemMetrics(SM_CXFULLSCREEN) * Screen.TwipsPerPixelX
sydim = GetSystemMetrics(SM_CYFULLSCREEN)
sydim = sydim + GetSystemMetrics(SM_CYCAPTION)
sydim = sydim * Screen.TwipsPerPixelY

F2Show.Left = (sxdim - F2Show.Width) / 2
F2Show.Top = (sydim - F2Show.Height) / 2

F2Show.Show Fmode
'If Fmode = vbModal Then
' Do
'  DoEvents
'Loop Until F2Show.Visible = False
'End If

End Sub

Public Sub MainINIStrings()
Dim l_mess As String

G_MainKey = ProfilePrivateGetStr("ODBCDSN", "MainKey", " ")
G_DataSourceName = ProfilePrivateGetStr("ODBCDSN", "DataSourceName", " ")
G_DatabaseName = ProfilePrivateGetStr("ODBCDSN", "DatabaseName", " ")
G_Description = ProfilePrivateGetStr("ODBCDSN", "Description", " ")
G_DriverPath = ProfilePrivateGetStr("ODBCDSN", "DriverPath", " ")
G_Trusted_Connection = ProfilePrivateGetStr("ODBCDSN", "Trusted_Connection", " ")
G_Server = ProfilePrivateGetStr("ODBCDSN", "Server", " ")
G_DriverName = ProfilePrivateGetStr("ODBCDSN", "DriverName", " ")
G_QuotedId = ProfilePrivateGetStr("ODBCDSN", "QuotedId", " ")
G_LastUser = ProfilePrivateGetStr("ODBCDSN", "LastUser", " ")

End Sub
Sub CenterForm(F2Show As Form)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim sxdim As Integer
Dim sydim As Integer

sxdim = GetSystemMetrics(SM_CXFULLSCREEN) * Screen.TwipsPerPixelX
sydim = GetSystemMetrics(SM_CYFULLSCREEN)
sydim = sydim + GetSystemMetrics(SM_CYCAPTION)
sydim = sydim * Screen.TwipsPerPixelY

F2Show.Left = (sxdim - F2Show.Width) / 2
F2Show.Top = (sydim - F2Show.Height) / 2

'If Fmode = vbModal Then
' Do
'  DoEvents
'Loop Until F2Show.Visible = False
'End If

End Sub

Function ISLOWER(v1) As Integer
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim n1 As Integer

n1 = Asc(Left$(v1, 1))

' ---------------------------------------------------------------
' see if it is or a-z
' ---------------------------------------------------------------

If (n1 > 96 And n1 < 123) Then
    ISLOWER = True
Else
    ISLOWER = False
End If

End Function

Function ISUPPER(v1) As Integer
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim n1 As Integer

n1 = Asc(Left$(v1, 1))

' ---------------------------------------------------------------
' see if it is or A-Z
' ---------------------------------------------------------------

If (n1 > 64 And n1 < 91) Then
    ISUPPER = True
Else
    ISUPPER = False
End If

End Function

Sub PRT_AtRowCol(rnum, cnum, prtext)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Printer.CurrentX = cnum * G_PrtCellWidth
Printer.CurrentY = rnum * G_PrtCellHeight

Printer.Print prtext

End Sub

Sub PRT_BoxCols(t_leftr, t_leftc, b_rightr, b_rightc)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim X1, Y1, X2, Y2

X1 = t_leftc * G_PrtCellWidth
Y1 = t_leftr * G_PrtCellHeight

X2 = b_rightc * G_PrtCellWidth
Y2 = b_rightr * G_PrtCellHeight

Printer.DrawWidth = 2

Printer.Line (X1, Y1)-(X2, Y2), , B

End Sub

Sub PRT_BoxColsShade(t_leftr, t_leftc, b_rightr, b_rightc)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim X1, Y1, X2, Y2

Dim SavStyleD, SavStyleF, SavColor

SavStyleF = Printer.FillStyle
SavStyleD = Printer.DrawStyle
SavColor = Printer.FillColor

X1 = t_leftc * G_PrtCellWidth
Y1 = t_leftr * G_PrtCellHeight

X2 = b_rightc * G_PrtCellWidth
Y2 = b_rightr * G_PrtCellHeight

' ---------------------------------------------------------------
'Printer.DrawWidth = 1
'Printer.DrawStyle = 0
'Printer.FillColor = G_LIGHT_GRAY ' Form1.RHeader.BackColor
'Printer.FillStyle = 6 ' Form1.RHeader.BackStyle
' ---------------------------------------------------------------

Printer.Line (X1, Y1)-(X2, Y2), RGB(222, 222, 222), BF

Printer.FillStyle = SavStyleF
Printer.DrawStyle = SavStyleD
Printer.FillColor = SavColor

End Sub

Sub PRT_Init()

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim PW0, PW1, PW2
Dim Afont, tempst As String
Dim I As Integer
' ---------------------------------------------------------------
' set scale mode to get width of selected paper size in inches
' ---------------------------------------------------------------

Printer.ScaleMode = 5

PW0 = Printer.ScaleWidth

' ---------------------------------------------------------------
' set scale mode to get width of selected paper size in twips
' ---------------------------------------------------------------

Printer.ScaleMode = 1

' ---------------------------------------------------------------
' get the twips width of a logical character cell based on
' a typical 80 character line
' ---------------------------------------------------------------

G_PrtCellWidth = Printer.ScaleWidth / 80

G_PrtCellHeight = Printer.ScaleHeight / 66

Printer.FontSize = 8
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False
For I = 0 To Printer.FontCount - 1  ' Determine number of fonts.
    tempst = Screen.Fonts(I)
    If tempst = "Arial" Then
        Afont = Screen.Fonts(I)
        Exit For
    End If
Next I
' Afont = Screen.Fonts(18)
Printer.FontName = Afont

End Sub

Sub PRT_LineCols(t_leftr, t_leftc, b_rightr, b_rightc)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim X1, Y1, X2, Y2

X1 = t_leftc * G_PrtCellWidth
Y1 = t_leftr * G_PrtCellHeight

X2 = b_rightc * G_PrtCellWidth
Y2 = b_rightr * G_PrtCellHeight

Printer.DrawWidth = 2

Printer.Line (X1, Y1)-(X2, Y2)

End Sub

Sub PRT_OnRowAt(rnum, tjust, prtext)
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Select Case tjust
Case PJ_Centered
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(prtext)) / 2
Case PJ_Left
    Printer.CurrentX = 1
Case PJ_Right
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(prtext))
Case Else
    Printer.CurrentX = 1
End Select

Printer.CurrentY = rnum * G_PrtCellHeight

Printer.Print prtext

End Sub

Sub ShowSystemMetrics()
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ---------------------------------------------------------------
' Purpose:
' ---------------------------------------------------------------
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

MsgBox "SM_CMETRICS = " & GetSystemMetrics(SM_CMETRICS)
MsgBox "SM_CMOUSEBUTTONS = " & GetSystemMetrics(SM_CMOUSEBUTTONS)
MsgBox "SM_CXBORDER = " & GetSystemMetrics(SM_CXBORDER)
MsgBox "SM_CXCURSOR = " & GetSystemMetrics(SM_CXCURSOR)
MsgBox "SM_CXDLGFRAME = " & GetSystemMetrics(SM_CXDLGFRAME)
MsgBox "SM_CXDOUBLECLK = " & GetSystemMetrics(SM_CXDOUBLECLK)
MsgBox "SM_CXFRAME = " & GetSystemMetrics(SM_CXFRAME)
MsgBox "SM_CXFULLSCREEN = " & GetSystemMetrics(SM_CXFULLSCREEN)
MsgBox "SM_CXHSCROLL = " & GetSystemMetrics(SM_CXHSCROLL)
MsgBox "SM_CXHTHUMB = " & GetSystemMetrics(SM_CXHTHUMB)
MsgBox "SM_CXICON = " & GetSystemMetrics(SM_CXICON)
MsgBox "SM_CXICONSPACING = " & GetSystemMetrics(SM_CXICONSPACING)
MsgBox "SM_CXMIN = " & GetSystemMetrics(SM_CXMIN)
MsgBox "SM_CXMINTRACK = " & GetSystemMetrics(SM_CXMINTRACK)
MsgBox "SM_CXSCREEN = " & GetSystemMetrics(SM_CXSCREEN)
MsgBox "SM_CXSIZE = " & GetSystemMetrics(SM_CXSIZE)
MsgBox "SM_CXVSCROLL = " & GetSystemMetrics(SM_CXVSCROLL)
MsgBox "SM_CYBORDER = " & GetSystemMetrics(SM_CYBORDER)
MsgBox "SM_CYCAPTION = " & GetSystemMetrics(SM_CYCAPTION)
MsgBox "SM_CYCURSOR = " & GetSystemMetrics(SM_CYCURSOR)
MsgBox "SM_CYDLGFRAME = " & GetSystemMetrics(SM_CYDLGFRAME)
MsgBox "SM_CYDOUBLECLK = " & GetSystemMetrics(SM_CYDOUBLECLK)
MsgBox "SM_CYFRAME = " & GetSystemMetrics(SM_CYFRAME)
MsgBox "SM_CYFULLSCREEN = " & GetSystemMetrics(SM_CYFULLSCREEN)
MsgBox "SM_CYHSCROLL = " & GetSystemMetrics(SM_CYHSCROLL)
MsgBox "SM_CYICON = " & GetSystemMetrics(SM_CYICON)
MsgBox "SM_CYICONSPACING = " & GetSystemMetrics(SM_CYICONSPACING)
MsgBox "SM_CYKANJIWINDOW = " & GetSystemMetrics(SM_CYKANJIWINDOW)
MsgBox "SM_CYMENU = " & GetSystemMetrics(SM_CYMENU)
MsgBox "SM_CYMIN = " & GetSystemMetrics(SM_CYMIN)
MsgBox "SM_CYMINTRACK = " & GetSystemMetrics(SM_CYMINTRACK)
MsgBox "SM_CYSCREEN = " & GetSystemMetrics(SM_CYSCREEN)
MsgBox "SM_CYSIZE = " & GetSystemMetrics(SM_CYSIZE)
MsgBox "SM_CYVSCROLL = " & GetSystemMetrics(SM_CYVSCROLL)
MsgBox "SM_CYVTHUMB = " & GetSystemMetrics(SM_CYVTHUMB)
MsgBox "SM_DBCSENABLED = " & GetSystemMetrics(SM_DBCSENABLED)
MsgBox "SM_DEBUG = " & GetSystemMetrics(SM_DEBUG)
MsgBox "SM_MENUDROPALIGNMENT = " & GetSystemMetrics(SM_MENUDROPALIGNMENT)
MsgBox "SM_MOUSEPRESENT = " & GetSystemMetrics(SM_MOUSEPRESENT)
MsgBox "SM_PENWINDOWS = " & GetSystemMetrics(SM_PENWINDOWS)
MsgBox "SM_SWAPBUTTON = " & GetSystemMetrics(SM_SWAPBUTTON)
MsgBox "SM_RESERVED1 = " & GetSystemMetrics(SM_RESERVED1)
MsgBox "SM_RESERVED2 = " & GetSystemMetrics(SM_RESERVED2)
MsgBox "SM_RESERVED3 = " & GetSystemMetrics(SM_RESERVED3)
MsgBox "SM_RESERVED4 = " & GetSystemMetrics(SM_RESERVED4)

End Sub


Sub delay(timedelay As Long)

Dim pausetime As Long
Dim start

pausetime = timedelay
start = Timer
Do While Timer < start + pausetime
    DoEvents    ' Yield to other processes.
Loop

End Sub

Public Sub KickOut(f As Form)
f.Top = Screen.Height - 1
f.Left = Screen.Width - 1
End Sub

