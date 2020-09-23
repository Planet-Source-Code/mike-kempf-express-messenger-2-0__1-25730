VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5385
   ClientLeft      =   3975
   ClientTop       =   2400
   ClientWidth     =   6120
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   3750
      TabIndex        =   26
      Top             =   1125
      Width           =   390
   End
   Begin VB.CheckBox chkRemove 
      Caption         =   "Remove if not connected"
      Height          =   240
      Left            =   1425
      TabIndex        =   25
      Top             =   1425
      Width           =   3240
   End
   Begin VB.CheckBox chkVerifyConnections 
      Caption         =   "Verify all connections state every"
      Height          =   240
      Left            =   1050
      TabIndex        =   24
      ToolTipText     =   "This will check all the users connected to the server to see if there is a valid connection"
      Top             =   1125
      Width           =   2715
   End
   Begin VB.TextBox txtLogFile 
      Height          =   315
      Left            =   1950
      TabIndex        =   22
      Top             =   2175
      Width           =   3690
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   1050
      TabIndex        =   21
      Top             =   2175
      Width           =   840
   End
   Begin VB.TextBox txtMaxUsers 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   2325
      TabIndex        =   19
      Top             =   4500
      Width           =   2115
   End
   Begin VB.CheckBox chkSaveUserMsgs 
      Caption         =   "Save All Messages Sent to Users"
      Height          =   240
      Left            =   1050
      TabIndex        =   18
      Top             =   825
      Width           =   3240
   End
   Begin VB.CheckBox chkSaveGroupMsg 
      Caption         =   "Save All Messages Sent to Groups"
      Height          =   240
      Left            =   1050
      TabIndex        =   17
      Top             =   525
      Width           =   3240
   End
   Begin VB.TextBox txtServerPort 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2325
      TabIndex        =   14
      Top             =   4125
      Width           =   2115
   End
   Begin VB.TextBox txtServerIP 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2325
      TabIndex        =   12
      Top             =   3750
      Width           =   2115
   End
   Begin VB.TextBox txtServerName 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2325
      TabIndex        =   10
      Top             =   3375
      Width           =   2115
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
      Left            =   4950
      TabIndex        =   2
      Top             =   4950
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3750
      TabIndex        =   1
      Top             =   4950
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2550
      TabIndex        =   0
      Top             =   4950
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "minute(s)."
      Height          =   240
      Left            =   4200
      TabIndex        =   27
      Top             =   1125
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Log File Path:"
      Height          =   240
      Left            =   1050
      TabIndex        =   23
      Top             =   1875
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Users:"
      Height          =   240
      Left            =   1050
      TabIndex        =   20
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "frmOptions.frx":000C
      Top             =   450
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmOptions.frx":08D6
      Top             =   3525
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      X1              =   1395
      X2              =   5967
      Y1              =   2895
      Y2              =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   1395
      X2              =   5967
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label Label7 
      Caption         =   "Server Settings"
      Height          =   240
      Left            =   225
      TabIndex        =   16
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1350
      X2              =   5895
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1350
      X2              =   5895
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label Label6 
      Caption         =   "General Options"
      Height          =   240
      Left            =   150
      TabIndex        =   15
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Server Port:"
      Height          =   240
      Left            =   1050
      TabIndex        =   13
      Top             =   4125
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "IP Address:"
      Height          =   240
      Left            =   1050
      TabIndex        =   11
      Top             =   3750
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Server Name:"
      Height          =   240
      Left            =   1050
      TabIndex        =   9
      Top             =   3375
      Width           =   1065
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGE SERVER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmOptions
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Option Explicit
Dim RegEdit2 As New cRegistry

Private Sub Check1_Click()

End Sub

Private Sub chkVerifyConnections_Click()
If chkVerifyConnections.Value = 1 Then
    chkRemove.Enabled = True
    txtMinutes.Enabled = True
    txtMinutes.BackColor = vbWhite
Else
    chkRemove.Enabled = False
    txtMinutes.Enabled = False
    txtMinutes.BackColor = &H8000000F
End If
End Sub

Private Sub cmdApply_Click()
    Call SaveSettings
End Sub

Private Sub cmdBrowse_Click()
Dim ReturnValue As String

    ReturnValue = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
    
        If ReturnValue <> "" Then
          txtLogFile = ReturnValue & "\"
        Else
          txtLogFile = ""
        End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveSettings
    Unload Me
End Sub

Private Sub Form_Load()
Set RegEdit2 = New cRegistry
'Get all the current saved settings
    GetSettings
    
If chkVerifyConnections.Value = 1 Then
    chkRemove.Enabled = True
    txtMinutes.Enabled = True
    txtMinutes.BackColor = vbWhite
Else
    chkRemove.Enabled = False
    txtMinutes.Enabled = False
    txtMinutes.BackColor = &H8000000F
End If

End Sub


Sub GetSettings()
On Error Resume Next
   chkSaveUserMsgs.Value = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveUserMsgs")
   chkSaveGroupMsg.Value = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveGroupMsgs")
   chkVerifyConnections.Value = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "VerifyConn")
   chkRemove.Value = RegEdit2.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "RemoveUser")
   txtLogFile.Text = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "LogPath")
   txtMinutes.Text = RegEdit2.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "Minutes")
   txtServerIP.Text = frmServer.ServiceSocket(0).LocalIP
   txtServerName.Text = frmServer.ServiceSocket(0).LocalHostName
   txtServerPort.Text = frmServer.ServiceSocket(0).LocalPort
   txtMaxUsers.Text = "500"
End Sub

Sub SaveSettings()
On Error Resume Next

Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveUserMsgs", chkSaveUserMsgs.Value)
Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveGroupMsgs", chkSaveGroupMsg.Value)
Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "VerifyConn", chkVerifyConnections.Value)
Call RegEdit2.SaveDword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "RemoveUser", chkRemove.Value)
Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "LogPath", txtLogFile.Text)
Call RegEdit2.savestring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "Minutes", txtMinutes.Text)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RegEdit2 = Nothing
End Sub

Private Sub txtMaxUsers_Change()
If Not txtMaxUsers = "" Then
    If Not IsNumeric(txtMaxUsers) Then
        MsgBox "Maximum users must be a nemeric value !", vbExclamation
        SendKeys "{BACKSPACE}"
    End If
End If
End Sub
