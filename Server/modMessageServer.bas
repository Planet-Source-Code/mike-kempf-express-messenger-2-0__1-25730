Attribute VB_Name = "modMessageServer"
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGE SERVER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   MODULE NAME:        modMessageServer
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////


Option Explicit

Global Const MaxUsers As Integer = 500

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public UserInfo(MaxUsers) As UserData

Public Type UserData
     NickName As String
     IPAddress As String
     Group As String
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

'Variables to hold the Server and User Information
Public ServerIP As String
Public Serverport As String
Public NickName As String
Public Group As String
'Variables to hold the Last Send Message
Public LastSentMessage As String
Public LastSentMessageRecipients As String

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) _
   As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
    lpBuffer As String, nSize As Long) As Long
   
'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     IsItRunning
'
'  AUTHOR:       Michael J. Kempf        7/30/2001 11:21:15 AM
'
'  PURPOSE:      This function checks to see if a specfic application is running.
'
'  PARAMETERS:
'                strClassName (String) = Application Class Name
'
'  RETURN:       Data Type = Long 0 if application is not running
'                            hwnd of the application window
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function IsItRunning(strClassName As String) As Long
   'Attempt to get window handle
   IsItRunning = FindWindow(strClassName, vbNullString)
End Function

Public Sub UnloadAllForms()
Dim frm As Form

    For Each frm In Forms
        Unload frm
    Next

End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     GetUser
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:48:53 PM
'
'  PURPOSE:      Get the current logged in user of the PC
'
'
'  RETURN:       String - Logged in User
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function GetUser() As String
    Dim lpUserID As String
    Dim nBuffer As Long
    Dim Ret As Long
    lpUserID = String(25, 0)
    nBuffer = 25
    Ret = GetUserName(lpUserID, nBuffer)


    If Ret Then
        GetUser$ = lpUserID$
    End If
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     ClipNull
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:48:53 PM
'
'  PURPOSE:      Remove all the trailing NULL characters of a string. Use in conjunction with GetUser
'
'  PARAMETERS:
'       [in]     InString (String) = String value that you want the NULL characters removes
'
'  RETURN:       String = String with NULL characters removed
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function ClipNull(InString As String) As String
    Dim intpos As Integer

    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)


        If intpos > 0 Then
            ClipNull = Left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If
    End If
End Function
