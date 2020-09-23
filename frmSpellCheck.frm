VERSION 5.00
Begin VB.Form frmSpellCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Check"
   ClientHeight    =   2640
   ClientLeft      =   4980
   ClientTop       =   4305
   ClientWidth     =   4185
   Icon            =   "frmSpellCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   3000
      TabIndex        =   7
      Top             =   1425
      Width           =   1110
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   420
      Left            =   3000
      TabIndex        =   6
      Top             =   900
      Width           =   1110
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   420
      Left            =   3000
      TabIndex        =   5
      Top             =   345
      Width           =   1110
   End
   Begin VB.ListBox lstWords 
      Height          =   1230
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtReplaceWith 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblReplaceWith 
      Caption         =   "Replace With"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblWord 
      Caption         =   "Word to Replace"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bCancelCheck As Boolean
Public bReplaceWord As Boolean


Private Sub cmdCancel_Click()
    bCancelCheck = True
    bReplaceWord = False
    Me.Hide
End Sub


Private Sub cmdIgnore_Click()
    bCancelCheck = False
    bReplaceWord = False
    Me.Hide
End Sub

Private Sub cmdReplace_Click()
    bCancelCheck = False
    bReplaceWord = True
    Me.Hide
End Sub


Private Sub lstWords_Click()
    txtReplaceWith.Text = lstWords.List(lstWords.ListIndex)
End Sub


Private Sub lstWords_DblClick()
    txtReplaceWith.Text = lstWords.List(lstWords.ListIndex)
    cmdReplace_Click
End Sub


