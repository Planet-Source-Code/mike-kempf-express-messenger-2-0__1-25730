VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Express Messenger"
   ClientHeight    =   2385
   ClientLeft      =   5745
   ClientTop       =   4890
   ClientWidth     =   4800
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1646.169
   ScaleMode       =   0  'User
   ScaleWidth      =   4507.448
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3675
      TabIndex        =   0
      Top             =   1950
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "A rich featured instant communication application."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   975
      TabIndex        =   4
      Top             =   900
      Width           =   3840
   End
   Begin VB.Label lblCopyrighy 
      Caption         =   "Copyright © 1999 - 2001 , Kemtech Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   2025
      Width           =   3390
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version: 1.0.9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   975
      TabIndex        =   2
      Top             =   600
      Width           =   2715
   End
   Begin VB.Label lblTitle 
      Caption         =   "Express Messenger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   975
      TabIndex        =   1
      Top             =   225
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmAbout.frx":000C
      Top             =   75
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4437.019
      Y1              =   1221.685
      Y2              =   1221.685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4437.019
      Y1              =   1232.038
      Y2              =   1232.038
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmAbout
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
'Show the version of the application
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & _
    "." & App.Revision

End Sub

