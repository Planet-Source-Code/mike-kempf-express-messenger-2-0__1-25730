VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReadMessage 
   BackColor       =   &H8000000A&
   Caption         =   "Express Messenger"
   ClientHeight    =   5400
   ClientLeft      =   3915
   ClientTop       =   2985
   ClientWidth     =   5400
   Icon            =   "frmReadMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAutoReconnect 
      Interval        =   60000
      Left            =   1125
      Top             =   6075
   End
   Begin VB.Timer tmrNewMsg 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1575
      Top             =   5625
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1125
      Width           =   3795
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   870
      Width           =   3795
   End
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   585
      Width           =   3795
   End
   Begin VB.Timer TmrOnTop 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1125
      Top             =   5625
   End
   Begin VB.TextBox RemedyDDE 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Text            =   "RemedyDDE"
      Top             =   5625
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.ListView lstMessages 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   5175
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sender"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Recipient"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Message"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2025
      Top             =   5625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   3540
      Left            =   75
      TabIndex        =   0
      Top             =   1425
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   6244
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReadMessage.frx":0F6A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer NewMessage 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3525
      Top             =   4350
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5085
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5186
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "0/0"
            TextSave        =   "0/0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgToolBar 
      Left            =   525
      Top             =   5625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":0FE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":1F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":2ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":3E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":4DD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":578D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":6709
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":6FE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":78C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":8819
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":9771
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReadMessage.frx":9E35
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckSYS 
      Left            =   75
      Top             =   5625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   794
      ButtonWidth     =   767
      ButtonHeight    =   741
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SendMsg"
            Object.ToolTipText     =   "Send Message"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReplyMsg"
            Object.ToolTipText     =   "Reply Message"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReplyAll"
            Object.ToolTipText     =   "Reply Message to All Recipients"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FwdMsg"
            Object.ToolTipText     =   "Forward Message"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintMsg"
            Object.ToolTipText     =   "Print Message"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveMsg"
            Object.ToolTipText     =   "Save Message"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteMsg"
            Object.ToolTipText     =   "Delete Message"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LogView"
            Object.ToolTipText     =   "Message Log Viewer"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrevMsg"
            Object.ToolTipText     =   "View Previous Message"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NextMsg"
            Object.ToolTipText     =   "View Next Message"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sleep"
            Object.ToolTipText     =   "Sleep Mode - Don't popup on receipt of new message"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quick"
            Object.ToolTipText     =   "Quick Messages"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   11
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "OK..."
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Yes..."
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "No..."
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Thanks..."
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Help on Phones..."
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Not Ready"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Not Logged In"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Not Updating DPS"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Under Scheduled"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Way to Go"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5475
      X2              =   0
      Y1              =   470
      Y2              =   470
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5475
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Image imgDisconnected 
      Height          =   240
      Left            =   4575
      Picture         =   "frmReadMessage.frx":ADB1
      Top             =   5625
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgConnected 
      Height          =   240
      Left            =   4275
      Picture         =   "frmReadMessage.frx":AEFB
      Top             =   5625
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   1125
      Width           =   390
   End
   Begin VB.Label lblToFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   585
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   870
      Width           =   240
   End
   Begin VB.Image imgTrayNewMail 
      Height          =   240
      Left            =   3975
      Picture         =   "frmReadMessage.frx":B045
      Top             =   5625
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgTraySend 
      Height          =   240
      Left            =   3675
      Picture         =   "frmReadMessage.frx":B5CF
      Top             =   5625
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect to Server"
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteMsg 
         Caption         =   "&Delete Message"
      End
      Begin VB.Menu mnudeleteAll 
         Caption         =   "Delete &All Messages"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Message"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "View Previous Message"
      End
      Begin VB.Menu mnuViewNext 
         Caption         =   "View Next Message"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearMessageBuff 
         Caption         =   "Clear Last Message Buffer"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Message"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuMessage 
      Caption         =   "&Message"
      Begin VB.Menu mnuSend 
         Caption         =   "&Send Message"
      End
      Begin VB.Menu mnuReply 
         Caption         =   "&Reply"
      End
      Begin VB.Menu mnuReplyAll 
         Caption         =   "Reply to All"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
      End
      Begin VB.Menu mnuLastSentMessage 
         Caption         =   "Last Sent Messsage"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQuick 
         Caption         =   "Quick Message"
         Begin VB.Menu mnuQMOK 
            Caption         =   "OK..."
         End
         Begin VB.Menu mnuQMYes 
            Caption         =   "Yes..."
         End
         Begin VB.Menu mnuQMNo 
            Caption         =   "No..."
         End
         Begin VB.Menu mnuQMThanks 
            Caption         =   "Thanks..."
         End
         Begin VB.Menu mnuQMHelp 
            Caption         =   "Help on Phones..."
         End
         Begin VB.Menu sep4 
            Caption         =   "-"
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogViewer 
         Caption         =   "Message Log Viewer"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuhide 
         Caption         =   "Hide"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Express Messenger..."
      End
   End
   Begin VB.Menu mnuTrayClick 
      Caption         =   "TrayClick"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTrayOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuTraySend 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmReadMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmReadMessage
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Dim RegEdit As New cRegistry
Dim CurrentMsgNumber As Integer
Dim SleepButtonPushed As Boolean
Dim intTrayIcon As Integer   ' 1 - normal  2 - New Message
Dim htxt As String  'HyperLink Text

Private Sub Form_KeyPress(KeyAscii As Integer)
'If the ESC key is pressed the hide the app into the system tray
    If KeyAscii = 27 Then
        frmReadMessage.WindowState = vbMinimized
        Me.Hide
    End If
End Sub
'Sub LoadMenus()
'Dim lngMenu As Long, lngNewMenu As Long, lngNewSubMenu As Long, lngSubMenu2 As Long
'   Dim strMnuID, strMenuText, strRecipient, strMessage As String
'
'   ''Get the form's menu handle
'   lngMenu& = GetMenu(Me.hWnd)
'
'
'   lngSubMenu = GetSubMenu(lngMenu, 1)
'   lngSubMenu2 = GetSubMenu(lngSubMenu, 5)
'   newmenupos = GetMenuItemCount(lngSubMenu2)
'
'
'
'
'   Open App.Path & "\QuickMessages.mnu" For Input As #1
'    Do While Not EOF(1)
'        Input #1, strMnuID, strMenuText, strRecipient, strMessage
'        Call InsertMenu(lngSubMenu2, 1&, MF_STRING, CLng(strMnuID), CStr(strMenuText))
'    Loop
'   Close #1
'
'
'   ''Get the original window procedure, so we can call
'   ''it and we can give it back when our program is done.
'   gOldProc& = GetWindowLong(Me.hWnd, GWL_WNDPROC)
'
'   ''Now replace the old window procedure
'   Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf MenuProc)
'
'End Sub
'====================================
'   Form Load
'====================================
Private Sub Form_Load()
'Check to see if the applicaion is already running
If App.PrevInstance Then
    MsgBox "Another instance of this application is running.", vbCritical, "Express Messenger"
    Unload Me
Else
On Error Resume Next
'Set default status text and icon to disconnected
    StatusBar1.Panels(1).Text = "Not connected to server"
    StatusBar1.Panels(1).Picture = imgDisconnected.Picture
'Get Settings From Registry
    ServerIP = RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerIp")
    Serverport = RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerPort")
    Group = UCase(RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "Group"))
'Check to see if Auto NickName Lookup is enabled
    If RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoLookup") = 1 Then
       NickName = UCase(ClipNull(GetUser))
    Else
        NickName = UCase(RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "NickName"))
    End If
 'Message counter
    CurrentMsgNumber = 0
 'Disable menu and toolbar icons
    EnDisMenuTollbarItems ("Disable")
 'Connect to Server
    sckSYS.Connect ServerIP, Serverport
 'Initialize Tray Icon
    Call AddToTray(imgTraySend.Picture, Me.Caption, Me)
    intTrayIcon = 1
 'Display Username in StatusBar
    StatusBar1.Panels(2).Text = NickName
 'Check to see if the Start in sleep mode option is enabled
    If RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "StartInSleep") = 1 Then
        SleepButtonPushed = True
        Toolbar1.Buttons(11).Value = 1
    End If
 'Hide the app into the system tray
    'GetScreenPosition
    
    frmReadMessage.WindowState = vbMinimized
    frmReadMessage.Show
    Me.Hide
End If
End Sub

Sub GetScreenPosition()
On Error Resume Next
 With frmReadMessage
    .Left = RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Left")
    .Top = RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Top")
    .Width = RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Width")
    .Height = RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Height")
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim Msg As Long
    
    Msg = ScaleX(X, ScaleMode, vbPixels)
    
    Select Case Msg
        
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
                    frmReadMessage.WindowState = vbNormal
                    frmReadMessage.Show
                'Set the Window to ON TOP OF ALL SCREENS and
                'Enable Timer to ture ON TOP Off
                    AlwaysOnTop Me, True
                    TmrOnTop.Enabled = True
                'Turn off new message timer
                    tmrNewMsg = False
                    If intTrayIcon = 2 Then
                        intTrayIcon = 1
                        Call ModifyTray(imgTraySend.Picture, Me.Caption, Me)
                    End If
                    
                        If SleepButtonPushed = True Then
                            Toolbar1.Buttons(11).Value = 1
                        Else
                            Toolbar1.Buttons(11).Value = 0
                        End If
             Case WM_LBUTTONDBLCLK 'Left button double-clicked
                    frmReadMessage.WindowState = vbNormal
                    frmReadMessage.Show
                'Set the Window to ON TOP OF ALL SCREENS and
                'Enable Timer to ture ON TOP Off
                    AlwaysOnTop Me, True
                    TmrOnTop.Enabled = True
                'Turn off new message timer
                    tmrNewMsg = False
                    If intTrayIcon = 2 Then
                        intTrayIcon = 1
                        Call ModifyTray(imgTraySend.Picture, Me.Caption, Me)
                    End If
                    
                        If SleepButtonPushed = True Then
                            Toolbar1.Buttons(11).Value = 1
                        Else
                            Toolbar1.Buttons(11).Value = 0
                        End If
             Case WM_RBUTTONDOWN 'Right button pressed
                    PopupMenu mnuTrayClick
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        frmReadMessage.WindowState = vbMinimized
        Me.Hide
    Else
        UnloadAllForms
    End If
End Sub

Private Sub Form_Resize()
'Resize the form, but  do not let them resize under a certain width and height
 On Error Resume Next
If Not Me.WindowState = vbMinimized Then
    If Me.Height < 6135 Then
        Me.Height = 6135
    ElseIf Me.Width < 5550 Then
        Me.Width = 5550
    ElseIf Me.Width > (Screen.Width / 2) Then
        Me.Width = (Screen.Width / 2)
    Else
        txtMessage.Height = Me.Height - 2525
        txtMessage.Width = Me.Width - 265
        txtFrom.Width = Me.Width - 265
        txtTo.Width = Me.Width - 265
        txtTime.Width = Me.Width - 265
        Line1(1).X2 = Me.Width
        Line1(0).X1 = Me.Width
    'Save screen position and size
        With frmReadMessage
          RegEdit.SaveDword HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Left", .Left
          RegEdit.SaveDword HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Top", .Top
          RegEdit.SaveDword HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Width", .Width
          RegEdit.SaveDword HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings\Position", "Height", .Height
        End With
    End If
Else
  Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
    Set RegEdit = Nothing
'Remove icon from system tray
   Call RemoveFromTray
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuClearMessageBuff_Click()
'Clear Last Sent Message Variables
    LastSentMessageRecipients = ""
    LastSentMessage = ""
'Disable Menu Items
    mnuLastSentMessage.Enabled = False
    mnuClearMessageBuff.Enabled = False
End Sub

Private Sub mnuConnect_Click()
    sckSYS.Close
    sckSYS.Connect ServerIP, Serverport
End Sub

Private Sub mnuDeleteAll_Click()
On Error Resume Next
Dim response

response = MsgBox("CHANGES CANNOT BE UNDONE !" & Chr(13) & Chr(13) & _
            "Are you sure you would like to delete all the messages ?", vbYesNo + vbCritical, "Confirm Message Delete")
    If response = vbYes Then
    'Clear Message off screen
        txtMessage.Text = ""
        txtTime = ""
        txtFrom = ""
        txtTo = ""
    'Delete Message Queue
        lstMessages.ListItems.Clear
        CurrentMsgNumber = 0
        StatusBar1.Panels(3).Text = "0/0"
     'Hide application into system tray
        frmReadMessage.WindowState = vbMinimized
        Me.Hide
     'Disable menu and toolbar items
        EnDisMenuTollbarItems ("Disable")
     'Turn off new message timer
        tmrNewMsg.Enabled = False
    End If
End Sub

Private Sub mnuDeleteMsg_Click()
    lstMessages.ListItems.Remove CurrentMsgNumber
    txtMessage.Text = ""
    txtTime = ""
    txtFrom = ""
    txtTo = ""
 'Load next message according  to the current message position
    If CurrentMsgNumber = 1 Then
        LoadNextMessage (CurrentMsgNumber)
    Else
        LoadPrevMessage (CurrentMsgNumber - 1)
    End If
    
    If lstMessages.ListItems.Count = 0 Then
      'Disable menu and toolbar items
         EnDisMenuTollbarItems ("Disable")
      'Delete Message Queue
         CurrentMsgNumber = 0
         StatusBar1.Panels(3).Text = "0/0"
      'Turn off new message timer
        tmrNewMsg.Enabled = False
      'Hide application into system tray
         frmReadMessage.WindowState = vbMinimized
         Me.Hide
    End If
End Sub

Private Sub mnuExit_Click()
    UnloadAllForms
    End
End Sub

Private Sub mnuFile_Click()
'Check to see if there is a successful connection
If StatusBar1.Panels(1).Text = "Ready" Then
    mnuConnect.Enabled = False
Else
    mnuConnect.Enabled = True
End If
End Sub

Private Sub mnuForward_Click()
    Load frmSendMessage
    frmSendMessage.txtMessageText.TextRTF = txtFrom.Text & " Wrote:" & vbCrLf & _
    vbCrLf & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuGoodJob_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem "ITSC"
    frmSendMessage.txtMessageText = "Everyone is logged in and ready to take calls.  Way to go!!!!"
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuhide_Click()
    frmReadMessage.WindowState = vbMinimized
    Me.Hide
End Sub

Private Sub mnuLastSentMessage_Click()
Dim i As Integer
Dim vntRecipients As Variant
vntRecipients = Split(LastSentMessageRecipients, ",")

    Load frmSendMessage
        For i = 0 To UBound(vntRecipients)
            frmSendMessage.lstRecipient.AddItem vntRecipients(i)
        Next
    frmSendMessage.txtMessageText = LastSentMessage
    frmSendMessage.Show vbModal
    
End Sub

Private Sub mnuLoggedOut_Click()
    Load frmSendMessage
    frmSendMessage.txtMessageText = "You are scheduled to be logged into the phone.  Please log in or update the DPS.  Thank You"
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuLogViewer_Click()
'Change mouse pointer to hourglass
    Me.MousePointer = vbHourglass
'Load Log for the current day
    Load frmLogViewer
    frmLogViewer.Show vbModal
'Change mouse pointer to Default arrow
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuNotReady_Click()
    Load frmSendMessage
    frmSendMessage.txtMessageText = "For those who are scheduled to be on the phone and are on 'Not Ready' status, please make your self available to answer calls, if possible." & _
    "If you are dealing with an extensive problem, please let me know.  Thank You"
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuNotUpdating_Click()
    Load frmSendMessage
    frmSendMessage.txtMessageText = "Please make sure you update the DPS at least two work days in advance.  Thank You"
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuprint_Click()
    Call PrintMessage
End Sub

Private Sub mnuQMHelp_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem "ITSC"
    frmSendMessage.txtMessageText = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}{\f1\fnil\fcharset0 Arial;}} " & _
    "{\colortbl ;\red255\green0\blue0;}\viewkind4\uc1\pard\cf1\b\f0\fs32 HELP ON PHONES!\cf0\b0\f1\fs15\par }"
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuQMNo_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
    frmSendMessage.txtMessageText = "NO" & vbCrLf & ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuQMOK_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
    frmSendMessage.txtMessageText = "OK" & vbCrLf & ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuQMThanks_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
    frmSendMessage.txtMessageText = "Thanks" & vbCrLf & ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuQMYes_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
    frmSendMessage.txtMessageText = "Yes" & vbCrLf & ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuReply_Click()
    Load frmSendMessage
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
    frmSendMessage.txtMessageText.TextRTF = vbCrLf & ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuReplyAll_Click()
Dim vntRecipients As Variant
Dim i As Integer

vntRecipients = Split(txtTo.Text, ",")

    Load frmSendMessage
'Add Recipients and sender to recipients list
    frmSendMessage.lstRecipient.AddItem txtFrom.Text
        For i = 0 To UBound(vntRecipients)
            frmSendMessage.lstRecipient.AddItem vntRecipients(i)
        Next
    frmSendMessage.txtMessageText.TextRTF = ">" & txtMessage.Text
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
'Show the File save dialog box
    With frmSendMessage.CD1
        .Filter = "Text Document (*.txt)|*.txt"
        .ShowSave
    
'Save  message text to the selected location and the selected filename/type
        Open frmSendMessage.CD1.FileName For Output As #1
            Print #1, "From: " & txtFrom.Text
            Print #1, "To: " & txtTo.Text
            Print #1, "Time: " & txtTime.Text
            Print #1,
            Print #1, txtMessage.Text
        Close #1
   End With
    'txtMessage.SaveFile frmSendMessage.CD1.FileName
End Sub

Private Sub mnuSend_Click()
    Load frmSendMessage
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuTrayExit_Click()
    UnloadAllForms
End Sub

Private Sub mnuTrayOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuTrayRestore_Click()
'Show the app
     frmReadMessage.WindowState = vbNormal
     frmReadMessage.Show
'Set the Window to ON TOP OF ALL SCREENS and
'Enable Timer to ture ON TOP Off
     AlwaysOnTop Me, True
     TmrOnTop.Enabled = True
     
     If SleepButtonPushed = True Then
        Toolbar1.Buttons(11).Value = 1
     Else
        Toolbar1.Buttons(11).Value = 0
     End If
End Sub

Private Sub mnuTraySend_Click()
    frmSendMessage.Show vbModal
End Sub

Private Sub mnuUnderScheduled_Click()
    Load frmSendMessage
    frmSendMessage.txtMessageText = "You are undersheduled on the DPS.  Please adjust the DPS accordingly.  Thank You"
    frmSendMessage.Show
End Sub

Private Sub mnuViewNext_Click()
    LoadNextMessage (CurrentMsgNumber + 1)
End Sub

Private Sub mnuViewPrevious_Click()
     LoadPrevMessage (CurrentMsgNumber - 1)
End Sub

Private Sub sckSYS_Connect()
If NickName = "" Then
    MsgBox "There is NO NickName defined." & vbCrLf & vbCrLf & "Please goto View\Options and set your NickName.", vbExclamation
ElseIf Group = "" Then
    MsgBox "There is NO Group defined." & vbCrLf & vbCrLf & "Please goto View\Options and set your Group.", vbExclamation
Else
'Send user login information to the server ( NickName and Group )
    sckSYS.SendData ".UserLogin" & "||" & UCase(NickName) & "||" & " " & "||" & UCase(Group) & "||" & App.Major & "." & App.Minor & _
    "." & App.Revision
End If
End Sub

Private Sub sckSYS_DataArrival(ByVal bytesTotal As Long)
Dim pData As String   'Inital Data Received
Dim vntUserCommands As Variant 'Commands received
Dim vntUserGroup As Variant ' Username and group in packet
Dim intItems As Integer 'Counter of items in the packet
Dim i As Integer   'loop counter

'Get Data from server
    sckSYS.GetData pData
'Split up the packet into an Variant array
    vntUserCommands = Split(pData, "||")
'Get the number of items in the array
    intItems = UBound(vntUserCommands)
  

Select Case vntUserCommands(0)
     Case ".ConnectedOK": '-------------------------------------- Initial Connection Response
    'Display the current status in the status bar ( connected )
        StatusBar1.Panels(1).Text = "Ready"
        StatusBar1.Panels(1).Picture = imgConnected.Picture
     Case ".NickNameExists" '-------------------------------------- Duplicate NickName logged into server
     Dim response
        response = MsgBox("The NickName(" & NickName & ") you are using is already logged into the server !" & vbCrLf & vbCrLf & _
        "Would you like to disconnect " & Chr(34) & NickName & Chr(34) & "?" & vbCrLf & vbCrLf & vbCrLf & "NOTE - You will" & _
        " reconnect in one minute.", vbExclamation + vbYesNo, "Express Messenger")
            If response = vbYes Then
            'Send data to remote disconnect
                sckSYS.SendData ".RemoteDisconnect" & "||" & NickName
                DoEvents
            'Close Connection
                sckSYS.Close
            'Set Status to disconnected
                StatusBar1.Panels(1).Text = "Not connected to server"
                StatusBar1.Panels(1).Picture = imgDisconnected.Picture
            Else
            'Close Connection
                sckSYS.Close
            'Set Status to disconnected
                StatusBar1.Panels(1).Text = "Not connected to server"
                StatusBar1.Panels(1).Picture = imgDisconnected.Picture
            End If
            
     
     Case ".ServerDisconnect" '-------------------------------------- Disconnected from server
        sckSYS.Close
        MsgBox "You have been disconnected from the Express Messenger server!", vbExclamation, "Express Messenger"
        'If the stae of the connection to the server is closed then display the status
        'Application could have auto reconnected before user responds to the message!!
        If sckSYS.State = 0 Then
            StatusBar1.Panels(1).Text = "Disconnected from server"
            StatusBar1.Panels(1).Picture = imgDisconnected.Picture
        End If
     Case ".msg" '-------------------------------------- New Message Received
     'Display message infoemation  FROM:, TO:, TIME: and Message
          With frmReadMessage
            .txtFrom.Text = vntUserCommands(1)
            .txtTime = Now()
            .txtTo = vntUserCommands(2)
            .txtMessage.TextRTF = vntUserCommands(5)
          End With
    'Enable menu and toolbar icon on receipt of a message
        Call EnDisMenuTollbarItems("Enable")
    'Save new Message in Message Queue
        Call SaveNewMessage(CStr(vntUserCommands(1)), CStr(vntUserCommands(2)), txtTime.Text, CStr(vntUserCommands(5)))
    'Increase message counter and set current message position
        CurrentMsgNumber = lstMessages.ListItems.Count
        StatusBar1.Panels(3).Text = CurrentMsgNumber & "/" & lstMessages.ListItems.Count
    'Disable View next button since you  are on the most current message
        If CurrentMsgNumber = lstMessages.ListItems.Count Then
            Toolbar1.Buttons(10).Enabled = False
            mnuViewNext.Enabled = False
        End If
    'Disable both View next and view prev messages if only one message
        If lstMessages.ListItems.Count = 1 Then
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons(10).Enabled = False
            mnuViewNext.Enabled = False
            mnuViewPrevious.Enabled = False
        End If
    'If sleep mode is not enabled then popup message else blink in system tray
        If Not SleepButtonPushed Then
           If frmReadMessage.WindowState = vbMinimized Then
            'popup Form
              frmReadMessage.WindowState = vbNormal
              frmReadMessage.Show
            'Set the Window to "ON TOP OF ALL SCREENS" and
            'Enable Timer to turn "ON TOP OF ALL SCREENS" Off
              AlwaysOnTop Me, True
              TmrOnTop.Enabled = True
            End If
        Else
            tmrNewMsg.Enabled = True
        End If
    'Play Sound if option is enabled
        If RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "PlaySound") = 1 Then
            sndPlaySound RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SoundFile"), &H1
        End If

    Case ".UserList" '-------------------------------------- User List Received
    'Clear Tree View
        frmSendMessage.UserListTree.Nodes.Clear
    'Fill UserList with logged on users
    'Start loop at 6 because that is the peice of the packet that stores the userlist
    For i = 6 To intItems
    'seperate the user and the group
        vntUserGroup = Split(vntUserCommands(i), "_._")
        Call frmSendMessage.FillUsersList(CStr(vntUserGroup(0)), CStr(vntUserGroup(1)))
    Next i
    
    Case ".State_Check" '-------------------------------------- State Check w/ Server
        sckSYS.SendData ".Connected" & "||" & NickName
    End Select
   
End Sub
Sub EnDisMenuTollbarItems(EnableDisable As String)

Select Case EnableDisable

Case "Enable"
'Enable toolBar Icons & Menu Items
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        Toolbar1.Buttons(10).Enabled = True
        Toolbar1.Buttons(12).ButtonMenus(1).Enabled = True
        Toolbar1.Buttons(12).ButtonMenus(2).Enabled = True
        Toolbar1.Buttons(12).ButtonMenus(3).Enabled = True
        Toolbar1.Buttons(12).ButtonMenus(4).Enabled = True
        mnuDeleteAll.Enabled = True
        mnuDeleteMsg.Enabled = True
        mnuViewNext.Enabled = True
        mnuViewPrevious.Enabled = True
        mnuSave.Enabled = True
        mnuprint.Enabled = True
        mnuReply.Enabled = True
        mnuReplyAll.Enabled = True
        mnuForward.Enabled = True
        mnuQMNo.Enabled = True
        mnuQMYes.Enabled = True
        mnuQMThanks.Enabled = True
        mnuQMOK.Enabled = True
  Case "Disable"
    'Disable toolBar Icons & Menus
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(9).Enabled = False
        Toolbar1.Buttons(10).Enabled = False
        Toolbar1.Buttons(12).ButtonMenus(1).Enabled = False
        Toolbar1.Buttons(12).ButtonMenus(2).Enabled = False
        Toolbar1.Buttons(12).ButtonMenus(3).Enabled = False
        Toolbar1.Buttons(12).ButtonMenus(4).Enabled = False
        mnuDeleteAll.Enabled = False
        mnuDeleteMsg.Enabled = False
        mnuViewNext.Enabled = False
        mnuViewPrevious.Enabled = False
        mnuSave.Enabled = False
        mnuprint.Enabled = False
        mnuReply.Enabled = False
        mnuReplyAll.Enabled = False
        mnuForward.Enabled = False
        mnuQMNo.Enabled = False
        mnuQMYes.Enabled = False
        mnuQMThanks.Enabled = False
        mnuQMOK.Enabled = False
  End Select
End Sub

Sub SaveNewMessage(strSender As String, strRecipient As String, strTime As String, strMessage As String)
    'Add Message Data To Inbox
       lstMessages.ListItems.Add , , strSender
       lstMessages.ListItems(lstMessages.ListItems.Count).ListSubItems.Add , , strRecipient
       lstMessages.ListItems(lstMessages.ListItems.Count).ListSubItems.Add , , strTime
       lstMessages.ListItems(lstMessages.ListItems.Count).ListSubItems.Add , , strMessage
     'Save Message to Log file
       If RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SaveMessageLog") = 1 Then
            Call SaveMessages(strSender, strRecipient, strMessage)
       End If
End Sub

Sub SaveMessages(Sender As String, Recipient As String, Message As String)
On Error Resume Next
    Dim LogFilePath As String
'Get the Log file path from the registry
    LogFilePath = RegEdit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "MessageFilePath")
'If no log file pah exists the set to the application path
    If LogFilePath = "" Then Exit Sub
'Open Log file
    Open LogFilePath & "EM_" & Month(Date) & Day(Date) & Year(Date) & ".EMF" For Append As #1
    'Covert Text to All Richtext
        Print #1, "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}" & _
        "{\f2\fswiss Arial;}}{\colortbl\red0\green0\blue0;}\deflang1033\pard\plain\f2\fs20 " & "From: " & Sender & " \par"
        
        Print #1, "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}" & _
        "{\f2\fswiss Arial;}}{\colortbl\red0\green0\blue0;}\deflang1033\pard\plain\f2\fs20 " & "To: " & Recipient & " \par"
        
        Print #1, "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}" & _
        "{\f2\fswiss Arial;}}{\colortbl\red0\green0\blue0;}\deflang1033\pard\plain\f2\fs20 " & "Time: " & Now() & " \par \par"
       
        Print #1, Message & " \par \par"
        
    Close #1
End Sub

Private Sub sckSYS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 'Display the current status in the status bar ( Disconnected )
   ' StatusBar1.Panels(1).Picture = imgdisconnected.Picture
    StatusBar1.Panels(1).Text = "No connection to server"
    StatusBar1.Panels(1).Picture = imgDisconnected.Picture
    sckSYS.Close
End Sub

Private Sub tmrAutoReconnect_Timer()

If RegEdit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoReconnect") = 1 Then
    If sckSYS.State = 0 Then
    'Reconnect to the server
        sckSYS.Connect ServerIP, Serverport
    End If
End If

End Sub

Private Sub tmrNewMsg_Timer()
'Flash icon in the system tray to represent that a new message has arrived ( SLEEP MODE )
    If intTrayIcon = 1 Then
        Call ModifyTray(imgTrayNewMail.Picture, "New Message", Me)
        intTrayIcon = 2
    Else
        Call ModifyTray(imgTraySend.Picture, Me.Caption, Me)
        intTrayIcon = 1
    End If
End Sub

Private Sub TmrOnTop_Timer()
    AlwaysOnTop Me, False
    TmrOnTop.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Respond to button clicked
    Select Case Button.Key
        Case "SendMsg"
            Load frmSendMessage
            frmSendMessage.Show vbModal
        Case "ReplyMsg"
            mnuReply_Click
        Case "ReplyAll"
            mnuReplyAll_Click
        Case "FwdMsg"
            mnuForward_Click
        Case "PrintMsg"
            Call PrintMessage
        Case "SaveMsg"
            mnuSave_Click
        Case "DeleteMsg"
            mnuDeleteMsg_Click
        Case "LogView"
            mnuLogViewer_Click
        Case "PrevMsg"
              LoadPrevMessage (CurrentMsgNumber - 1)
        Case "NextMsg"
              LoadNextMessage (CurrentMsgNumber + 1)
        Case "sleep"
            If SleepButtonPushed = True Then
                Toolbar1.Buttons(11).Value = 0
                SleepButtonPushed = False
            Else
                Toolbar1.Buttons(11).Value = 1
                SleepButtonPushed = True
            End If
        Case "Quick"
    End Select
End Sub
Sub LoadPrevMessage(MessageNumber As Integer)
If Not MessageNumber < 1 Then
'Load Message
    txtFrom.Text = lstMessages.ListItems(MessageNumber).Text
    txtTo.Text = lstMessages.ListItems(MessageNumber).SubItems(1)
    txtTime.Text = lstMessages.ListItems(MessageNumber).SubItems(2)
    txtMessage.TextRTF = lstMessages.ListItems(MessageNumber).SubItems(3)
    
    CurrentMsgNumber = MessageNumber
    StatusBar1.Panels(3).Text = CurrentMsgNumber & "/" & lstMessages.ListItems.Count
 
    'Disable/Enable arrow buttons on toolbar
    If CurrentMsgNumber = 1 Then
        Toolbar1.Buttons(9).Enabled = False
        mnuViewPrevious.Enabled = False
    End If
    If CurrentMsgNumber < lstMessages.ListItems.Count Then
        Toolbar1.Buttons(10).Enabled = True
        mnuViewNext.Enabled = True
    End If
End If
End Sub

Sub LoadNextMessage(MessageNumber As Integer)
If Not MessageNumber > lstMessages.ListItems.Count Then
'Load Message
    txtFrom.Text = lstMessages.ListItems(MessageNumber).Text
    txtTo.Text = lstMessages.ListItems(MessageNumber).SubItems(1)
    txtTime.Text = lstMessages.ListItems(MessageNumber).SubItems(2)
    txtMessage.TextRTF = lstMessages.ListItems(MessageNumber).SubItems(3)
    
    CurrentMsgNumber = MessageNumber
    StatusBar1.Panels(3).Text = CurrentMsgNumber & "/" & lstMessages.ListItems.Count
    
'Disable/Enable arrow buttons on toolbar
    If CurrentMsgNumber > 1 Then
        Toolbar1.Buttons(9).Enabled = True
        mnuViewPrevious.Enabled = True
    End If
    If CurrentMsgNumber = lstMessages.ListItems.Count Then
        Toolbar1.Buttons(10).Enabled = False
        mnuViewNext.Enabled = False
    End If
End If
End Sub

Private Sub PrintMessage()

    On Local Error GoTo Error_Handler:
     With CD1
        .CancelError = True
        .ShowPrinter
           
                 If txtMessage.SelLength = 0 Then
                    .Flags = .Flags + cdlPDAllPages
                 Else
                    .Flags = .Flags + cdlPDSelection
                 End If
                 
            On Local Error Resume Next
                   Printer.Print ""
                   Printer.Print "From: " & txtFrom.Text
                   Printer.Print "To: " & txtTo.Text
                   Printer.Print "Time: " & txtTime.Text
                   Printer.Print ""
                   Printer.Print txtMessage.Text
                   Printer.EndDoc
          
     End With
Exit Sub

Error_Handler:
    
    If Err <> cdlCancel Then
    MsgBox " Error " & Err & "; " & Error, vbExclamation, "Express Messenger"
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.index
    Case 1
        mnuQMOK_Click
    Case 2
        mnuQMYes_Click
    Case 3
        mnuQMNo_Click
    Case 4
        mnuQMThanks_Click
    Case 5
        mnuQMHelp_Click
    Case 7
        mnuNotReady_Click
    Case 8
        mnuLoggedOut_Click
    Case 9
        mnuNotUpdating_Click
    Case 10
        mnuUnderScheduled_Click
    Case 11
        mnuGoodJob_Click
End Select

End Sub

Private Sub txtMessage_Click()
    Dim lngRet
    Dim i As Integer
    Dim strHeader As String
    Dim strTrailer As String
    
'Launch the web link executing the default browser
    If Left(htxt, 5) = "http:" Or Left(htxt, 4) = "www." Then lngRet = ShellExecute(0&, "Open", htxt, "", vbNullString, SW_SHOWNORMAL)
'Open the Specific Remedy ticket link
    If Left(htxt, 2) = "HD" Or Left(htxt, 2) = "hd" Then
        If Not Len(htxt) = 15 Then

            strHeader = Left(htxt, 2)  ' ( HD )
            strTrailer = Mid(htxt, 3, Len(htxt)) '( 00102 )
            
            'Append 0 to the Ticket number
                For i = 1 To 15 - Len(htxt)
                    strTrailer = "0" & strTrailer
                Next i
                
                htxt = strHeader & strTrailer ' ( HD0000000000102 )
        End If
        
            If Not IsItRunning("ArFrame") = 0 Then
                 Call DDEExecute("ARUSER-SERVER", "DoExecMacro", _
                 "[RunMacro (" & App.Path & ",EMTicket,TICKET=" & htxt & ")]")
            Else
                 MsgBox "Unable to open Remedy ticket!" & vbCrLf & vbCrLf & _
                 "Remedy User application is not running.", vbCritical, "Express Messenger"
            End If
       
    ElseIf Left(htxt, 1) = "#" Then
            
        If Not Len(Mid(htxt, 2, Len(htxt))) = 13 Then

                strHeader = "HD"  ' ( HD )
                strTrailer = Mid(htxt, 2, Len(htxt))
            
                'Append 0 to the Ticket number
                    For i = 1 To 13 - Len(Mid(htxt, 2, Len(htxt)))
                        strTrailer = "0" & strTrailer
                    Next i
                
                    htxt = strHeader & strTrailer ' ( HD0000000000102 )
        End If
                
                If Not IsItRunning("ArFrame") = 0 Then
                    Call DDEExecute("ARUSER-SERVER", "DoExecMacro", _
                    "[RunMacro (" & App.Path & ",EMTicket,TICKET=" & htxt & ")]")
                Else
                    MsgBox "Unable to open Remedy ticket!" & vbCrLf & vbCrLf & _
                    "Remedy User application is not running.", vbCritical, "Express Messenger"
                End If
    End If
End Sub


Private Sub txtMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    htxt = GetHyperlink(txtMessage, X, Y)
End Sub

Public Sub DDEExecute(sApplication As String, sTopic As String, sMacro As String)

On Error Resume Next

'Link Topic
    RemedyDDE.LinkTopic = sApplication & "|" & sTopic
'Set link Mode
    RemedyDDE.LinkMode = 2
'minimize application
    frmReadMessage.WindowState = vbMinimized
    Me.Hide
'Execute Command
    RemedyDDE.LinkExecute sMacro


End Sub

Private Function GetHyperlink(rch As RichTextBox, X As Single, Y As Single) As String
    'This function return any word curently under cursor and,
    'if string under cursor start with "http:","www.","HD" or "#"
    'then change mouse pointer to hand

    Dim pt As POINTAPI
    Dim pos As Integer
    Dim ch As String
    Dim txt As String
    Dim txtlen As Integer
    Dim pos_start As Integer
    Dim pos_mijloc As Integer
    Dim pos_end As Integer
    Dim i As Integer
    Dim strHeader As String
    Dim strTrailer As String

    
    ' convert mouse pos in pixels
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' position of character under cursor
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then
        Exit Function
    End If
    txt = rch.Text

    ' get start position of word under cursor
    For pos_start = pos To 1 Step -1
        ch = Mid$(txt, pos_start, 1)
        If ch = Chr(32) Or ch = vbCr Or ch = vbLf Or ch = vbNewLine Then Exit For
    Next pos_start
    pos_start = pos_start + 1

    ' get end position of word under cursor
    txtlen = Len(txt)
    For pos_end = pos To txtlen
        ch = Mid$(txt, pos_end, 1)
    If ch = Chr(32) Or ch = vbCr Then Exit For
    Next pos_end
    pos_end = pos_end - 1

    If pos_start <= pos_end Then _
        GetHyperlink = Mid$(txt, pos_start, pos_end - pos_start + 1)
        
        If Left(GetHyperlink, 5) = "http:" Or Left(GetHyperlink, 4) = "www." Or Left(GetHyperlink, 2) = "HD" _
        Or Left(GetHyperlink, 2) = "hd" Or Left(GetHyperlink, 1) = "#" Then
        
            rch.MouseIcon = LoadPicture(App.Path & "\hand.cur")
            rch.MousePointer = vbCustom
            
            If Left(GetHyperlink, 2) = "hd" Or Left(GetHyperlink, 2) = "HD" Then
                If Not Len(GetHyperlink) = 15 Then
            
                        strHeader = Left(GetHyperlink, 2)  ' ( HD )
                        strTrailer = Mid(GetHyperlink, 3, Len(GetHyperlink)) '( 00102 )
            
                'Append 0 to the Ticket number
                            For i = 1 To 15 - Len(GetHyperlink)
                                strTrailer = "0" & strTrailer
                            Next i
                
                            GetHyperlink = strHeader & strTrailer ' ( HD0000000000102 )
                End If
                
                rch.ToolTipText = "Click here to open Remedy ticket # " & GetHyperlink
        ElseIf Left(GetHyperlink, 1) = "#" Then
            
                If Not Len(Mid(GetHyperlink, 2, Len(GetHyperlink))) = 13 Then
                  
                        strHeader = "HD"  ' ( HD )
                        strTrailer = Mid(GetHyperlink, 2, Len(GetHyperlink))
            
                'Append 0 to the Ticket number
                            For i = 1 To 13 - Len(Mid(GetHyperlink, 2, Len(GetHyperlink)))
                                strTrailer = "0" & strTrailer
                            Next i
                
                            GetHyperlink = strHeader & strTrailer ' ( HD0000000000102 )
                End If
                
                rch.ToolTipText = "Click here to open Remedy ticket # " & GetHyperlink
            Else
                rch.ToolTipText = "Click here to navigate to " + GetHyperlink
            End If
        Else
            rch.ToolTipText = ""
            rch.MousePointer = 0
        End If
               
End Function

Public Sub AlwaysOnTop(frm As Form, SetOnTop As Boolean)

    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    
        SetWindowPos frm.hWnd, lFlag, _
        frm.Left / Screen.TwipsPerPixelX, _
        frm.Top / Screen.TwipsPerPixelY, _
        frm.Width / Screen.TwipsPerPixelX, _
        frm.Height / Screen.TwipsPerPixelY, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


