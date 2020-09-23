VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5250
   ClientLeft      =   4365
   ClientTop       =   3465
   ClientWidth     =   6390
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4590
      Left            =   75
      TabIndex        =   9
      Top             =   75
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   8096
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "User/Server"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Image2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkAutoLookup"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtServerPort"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtServerIP"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtServerName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtGroup"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtNickName"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Image4"
      Tab(1).Control(2)=   "Image3"
      Tab(1).Control(3)=   "Line1(7)"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Line1(4)"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Line1(5)"
      Tab(1).Control(8)=   "Line1(6)"
      Tab(1).Control(9)=   "chkSaveMessLog"
      Tab(1).Control(10)=   "txtMessLogFile"
      Tab(1).Control(11)=   "cmdFolderBrowse"
      Tab(1).Control(12)=   "chkStartInSleep"
      Tab(1).Control(13)=   "chkReconnect"
      Tab(1).Control(14)=   "txtLogPath"
      Tab(1).Control(15)=   "cmdBrowse"
      Tab(1).Control(16)=   "txtSoundFile"
      Tab(1).Control(17)=   "chkPlaySound"
      Tab(1).ControlCount=   18
      Begin VB.CheckBox chkPlaySound 
         Caption         =   "Play sound when new messages arrive"
         Height          =   315
         Left            =   -74025
         TabIndex        =   32
         Top             =   915
         Width           =   3240
      End
      Begin VB.TextBox txtSoundFile 
         Height          =   315
         Left            =   -73200
         TabIndex        =   31
         Top             =   1290
         Width           =   3465
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   315
         Left            =   -74025
         TabIndex        =   30
         Top             =   1290
         Width           =   765
      End
      Begin VB.TextBox txtLogPath 
         Height          =   315
         Left            =   -72600
         TabIndex        =   29
         Top             =   2265
         Width           =   2865
      End
      Begin VB.CheckBox chkReconnect 
         Caption         =   "Automatically reconnect if disconnected"
         Height          =   240
         Left            =   -74025
         TabIndex        =   28
         Top             =   3615
         Width           =   3240
      End
      Begin VB.CheckBox chkStartInSleep 
         Caption         =   "Start Express Messenger in sleep mode"
         Height          =   240
         Left            =   -74025
         TabIndex        =   27
         Top             =   3915
         Width           =   3240
      End
      Begin VB.CommandButton cmdFolderBrowse 
         Caption         =   "Browse"
         Height          =   315
         Left            =   -74025
         TabIndex        =   26
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox txtMessLogFile 
         Height          =   315
         Left            =   -73200
         TabIndex        =   25
         Top             =   3090
         Width           =   3465
      End
      Begin VB.CheckBox chkSaveMessLog 
         Caption         =   "Save all received messages in a log file"
         Height          =   315
         Left            =   -74025
         TabIndex        =   24
         Top             =   2715
         Width           =   3240
      End
      Begin VB.TextBox txtNickName 
         Height          =   285
         Left            =   2100
         TabIndex        =   15
         Top             =   1035
         Width           =   2115
      End
      Begin VB.TextBox txtGroup 
         Height          =   285
         Left            =   2100
         TabIndex        =   14
         Top             =   1410
         Width           =   2115
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   2100
         TabIndex        =   13
         Top             =   2760
         Width           =   2115
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   2100
         TabIndex        =   12
         Top             =   3135
         Width           =   2115
      End
      Begin VB.TextBox txtServerPort 
         Height          =   285
         Left            =   2100
         TabIndex        =   11
         Top             =   3510
         Width           =   2115
      End
      Begin VB.CheckBox chkAutoLookup 
         Caption         =   "Automatically detect NickName from Login ID"
         Height          =   315
         Left            =   975
         TabIndex        =   10
         Top             =   1710
         Width           =   3615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   6
         X1              =   -73575
         X2              =   -69110
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   5
         X1              =   -73650
         X2              =   -69110
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Sound Options"
         Height          =   240
         Left            =   -74850
         TabIndex        =   35
         Top             =   600
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -73650
         X2              =   -69110
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label9 
         Caption         =   "General Options"
         Height          =   240
         Left            =   -74850
         TabIndex        =   34
         Top             =   1815
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -73575
         X2              =   -69103
         Y1              =   1965
         Y2              =   1965
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74850
         Picture         =   "frmOptions.frx":0044
         Top             =   1065
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74850
         Picture         =   "frmOptions.frx":034E
         Top             =   2265
         Width           =   480
      End
      Begin VB.Label Label10 
         Caption         =   "Message Log Path:"
         Height          =   240
         Left            =   -74025
         TabIndex        =   33
         Top             =   2265
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "NickName:"
         Height          =   240
         Left            =   975
         TabIndex        =   23
         Top             =   1035
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Group:"
         Height          =   240
         Left            =   975
         TabIndex        =   22
         Top             =   1410
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Server Name:"
         Height          =   240
         Left            =   975
         TabIndex        =   21
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "IP Address:"
         Height          =   240
         Left            =   975
         TabIndex        =   20
         Top             =   3135
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "Server Port:"
         Height          =   240
         Left            =   975
         TabIndex        =   19
         Top             =   3510
         Width           =   915
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   150
         Picture         =   "frmOptions.frx":0C18
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "User Settings"
         Height          =   240
         Left            =   150
         TabIndex        =   18
         Top             =   600
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1275
         X2              =   5972
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   1275
         X2              =   5972
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Server Settings"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1350
         X2              =   5922
         Y1              =   2290
         Y2              =   2290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   3
         X1              =   1350
         X2              =   5922
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmOptions.frx":14E2
         Top             =   2910
         Width           =   480
      End
      Begin VB.Label Label11 
         Caption         =   "Changing these settings could affect sending and receiving of messages."
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   525
         TabIndex        =   16
         Top             =   2460
         Width           =   5190
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5175
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2925
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmOptions
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Option Explicit
Dim RegEdit2 As New cRegistry


Private Sub chkSaveMessLog_Click()
If chkSaveMessLog.Value = 0 Then
    txtMessLogFile.Enabled = False
    txtMessLogFile.BackColor = &H80000004
Else
    txtMessLogFile.Enabled = True
    txtMessLogFile.BackColor = &H80000005
End If
End Sub

Private Sub chkAutoLookup_Click()
If chkAutoLookup.Value = 1 Then
    txtNickName.Enabled = False
    txtNickName.BackColor = &H80000004
Else
    txtNickName.Enabled = True
    txtNickName.BackColor = &H80000005
End If
End Sub

Private Sub chkPlaySound_Click()
If chkPlaySound.Value = 0 Then
    txtSoundFile.Enabled = False
    txtSoundFile.BackColor = &H80000004
Else
    txtSoundFile.Enabled = True
    txtSoundFile.BackColor = &H80000005
End If
End Sub

Private Sub cmdApply_Click()
    Call SaveSettings
End Sub

Private Sub cmdBrowse_Click()
On Error Resume Next
With frmSendMessage.cd1
    .Filter = "Audio Files (*.wav)|*.wav"
    .ShowOpen
End With
    txtSoundFile = frmSendMessage.cd1.FileName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFolderBrowse_Click()
Dim ReturnValue As String

    ReturnValue = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
    
        If ReturnValue <> "" Then
          txtMessLogFile = ReturnValue & "\"
        Else
          txtMessLogFile = ""
        End If
End Sub

Private Sub cmdOK_Click()
    Call SaveSettings
    Unload Me
End Sub

Private Sub Form_Load()
Set RegEdit2 = New cRegistry
'Get all the current saved settings
    GetSettings
    
If chkPlaySound.Value = 0 Then
    txtSoundFile.Enabled = False
    txtSoundFile.BackColor = &H80000004
Else
    txtSoundFile.Enabled = True
     txtSoundFile.BackColor = &H80000005
End If

If chkSaveMessLog.Value = 0 Then
    txtMessLogFile.Enabled = False
    txtMessLogFile.BackColor = &H80000004
Else
    txtMessLogFile.Enabled = True
    txtMessLogFile.BackColor = &H80000005
End If

If chkAutoLookup.Value = 1 Then
    txtNickName.Enabled = False
    txtNickName.BackColor = &H80000004
Else
    txtNickName.Enabled = True
     txtNickName.BackColor = &H80000005
End If

End Sub

Sub GetSettings()
On Error Resume Next
    txtNickName = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "NickName")
    txtGroup = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "Group")
    txtServerName = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerName")
    txtServerIP = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerIp")
    txtServerPort = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerPort")
    chkPlaySound = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "PlaySound")
    txtSoundFile = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SoundFile")
    txtLogPath = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "LogPath")
    chkAutoLookup = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoLookup")
    chkReconnect = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoReconnect")
    chkStartInSleep = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "StartInSleep")
    chkSaveMessLog = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SaveMessageLog")
    txtMessLogFile = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "MessageFilePath")
End Sub

Sub SaveSettings()
On Error Resume Next
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "NickName", UCase(txtNickName))
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "Group", UCase(txtGroup))
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerName", txtServerName)
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerIP", txtServerIP)
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "ServerPort", txtServerPort)
    Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "PlaySound", chkPlaySound.Value)
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SoundFile", txtSoundFile)
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "LogPath", txtLogPath)
    Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoLookup", chkAutoLookup)
    Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoReconnect", chkReconnect)
    Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "StartInSleep", chkStartInSleep)
    Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "SaveMessageLog", chkSaveMessLog.Value)
    Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "MessageFilePath", txtMessLogFile)
'Set global Server and login Information
    ServerIP = txtServerIP
    Serverport = txtServerPort
    Group = txtGroup
        If RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "AutoLookup") = 1 Then
           NickName = UCase(ClipNull(GetUser))
        Else
           NickName = UCase(txtNickName)
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RegEdit2 = Nothing
End Sub

