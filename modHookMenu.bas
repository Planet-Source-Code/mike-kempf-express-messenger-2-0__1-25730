Attribute VB_Name = "modHookMenu"
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10




''Variable to hold the address of the old window procedure
Public gOldProc As Long
Public Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim strMnuID, strMenuText, strRecipient, strMessage As String

       Select Case wMsg&
           Case WM_CLOSE:
               ''User has closed the window, so we should stop
               ''subclassing immediately! We do this by handing
               ''back the original window procedure.
               Call SetWindowLong(hwnd&, GWL_WNDPROC, gOldProc&)
          
           Case WM_COMMAND:
               ''WM_COMMAND is sent to the window
               ''whenever someone clicks a menu.
               ''The menu's item ID is stored in wParam.
           'Check to see if a Dynamic menu was clicked
            If wParam& >= 200 Then
                Open App.Path & "\QuickMessages.mnu" For Input As #2
                    Do While Not EOF(2)
                        Input #2, strMnuID, strMenuText, strRecipient, strMessage
                            If wParam& = CLng(strMnuID) Then
                                MsgBox "RECIPIENT: " & strRecipient & vbCrLf & vbCrLf & _
                                       "MESSAGE: " & strMessage
                            End If
                    Loop
                Close #2
            End If
       
       End Select

    ''Call original window procedure for default processing.
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)

End Function
