Attribute VB_Name = "modSystemTray"
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   MODULE NAME:        modTrayModule
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

      'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      
      Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      'Declare the constants for the API function.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      
      Global Const NIM_ADD = &H0
      Global Const NIM_MODIFY = &H1
      Global Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Global Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      
      Global Const NIF_MESSAGE = &H1
      Global Const NIF_ICON = &H2
      Global Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.
  
      'Left-click constants.
      Global Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Global Const WM_LBUTTONDOWN = &H201     'Button down
      Global Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Global Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Global Const WM_RBUTTONDOWN = &H204     'Button down
      Global Const WM_RBUTTONUP = &H205       'Button up
      
      'Declare the API function call.
      Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Global nid As NOTIFYICONDATA

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    AddToTray
'
'  AUTHOR:       Michael J. Kempf        7/30/2001 11:23:18 AM
'
'  PURPOSE:      This application add your application to the system tray
'
'  PARAMETERS:
'                TrayIcon (Variant) = Icon to represent the system tray object
'                TrayText  (String) = Text to appear on mouseover
'                TrayForm    (Form) = form to add to tray
'
'///////////////////////////////////////////////////////////////////////////////////////////
Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
    'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area.
         Shell_NotifyIcon NIM_ADD, nid
         TrayForm.Hide
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    ModifyTray
'
'  AUTHOR:       Michael J. Kempf        7/30/2001 11:23:23 AM
'
'  PURPOSE:      This procedure modifies the current system tray icon (Change icon or text )
'
'  PARAMETERS:
'                TrayIcon (Variant) = Icon to represent the system tray object
'                TrayText  (String) = Text to appear on mouseover
'                TrayForm    (Form) = form to add to tray
'
'///////////////////////////////////////////////////////////////////////////////////////////
Sub ModifyTray(TrayIcon, TrayText As String, TrayForm As Form)
    'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

    'Call the Shell_NotifyIcon function to modify the icon to the taskbar
    'status area.
         Shell_NotifyIcon NIM_MODIFY, nid
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    RemoveFromTray
'
'  AUTHOR:       Michael J. Kempf        7/30/2001 11:23:27 AM
'
'  PURPOSE:      This procedure romoved the icon from the system tray
'
'  PARAMETERS:   None.
'
'///////////////////////////////////////////////////////////////////////////////////////////
Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub



