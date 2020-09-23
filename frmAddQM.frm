VERSION 5.00
Begin VB.Form frmAddQM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Quick Menu"
   ClientHeight    =   1410
   ClientLeft      =   2925
   ClientTop       =   2160
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   3975
      TabIndex        =   3
      Top             =   900
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   390
      Left            =   3000
      TabIndex        =   2
      Top             =   900
      Width           =   915
   End
   Begin VB.TextBox txtQMCaption 
      Height          =   315
      Left            =   825
      TabIndex        =   1
      Top             =   450
      Width           =   4065
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   75
      Picture         =   "frmAddQM.frx":0000
      ScaleHeight     =   540
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Quick Message Caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   825
      TabIndex        =   4
      Top             =   150
      Width           =   4065
   End
End
Attribute VB_Name = "frmAddQM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
If txtQMCaption.Text = "" Then
    MsgBox "Please enter a caption for the Quick Menu you are adding.", vbExclamation, "Quick Menu"
Else
    Call AddQM(txtQMCaption.Text)
End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub AddQM(strCaption As String)


End Sub


