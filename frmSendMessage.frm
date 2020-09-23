VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSendMessage 
   BackColor       =   &H8000000B&
   Caption         =   " Send Message"
   ClientHeight    =   4560
   ClientLeft      =   3195
   ClientTop       =   2895
   ClientWidth     =   8775
   Icon            =   "frmSendMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Sizes 
      Height          =   315
      ItemData        =   "frmSendMessage.frx":0E42
      Left            =   8100
      List            =   "frmSendMessage.frx":0E76
      TabIndex        =   11
      Text            =   "11"
      ToolTipText     =   "Fint Size"
      Top             =   0
      Width           =   555
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   825
      ScaleHeight     =   465
      ScaleWidth      =   8415
      TabIndex        =   8
      Top             =   0
      Width           =   8415
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Object.ToolTipText     =   "Align Left"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Center"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "Align Right"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullets"
               Object.ToolTipText     =   "Bullets"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CheckSpelling"
               Object.ToolTipText     =   "Spell Check"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Color"
               Object.ToolTipText     =   "Font Color"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cboFont 
            Height          =   315
            Left            =   5025
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Font"
            Top             =   0
            Width           =   2115
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   6
      Top             =   0
      Width           =   765
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   582
         ButtonWidth     =   1349
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Send"
               Key             =   "SendMsg"
               Object.ToolTipText     =   "Send Message"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   825
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstRecipient 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      IntegralHeight  =   0   'False
      ItemData        =   "frmSendMessage.frx":0EB8
      Left            =   2700
      List            =   "frmSendMessage.frx":0EBA
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Recipient List"
      Top             =   1050
      Width           =   1515
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   225
      Top             =   4050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":0EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":1796
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":2070
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":294C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":3228
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":3B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":43E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView UserListTree 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Current Users Logged into Server"
      Top             =   1050
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   5953
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin RichTextLib.RichTextBox txtMessageText 
      Height          =   3375
      Left            =   4275
      TabIndex        =   0
      ToolTipText     =   "Message Text"
      Top             =   1050
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   5953
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmSendMessage.frx":4AA2
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
   Begin MSComctlLib.ImageList ImgToolBar 
      Left            =   1050
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":4B19
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5753
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5865
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5977
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5CAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5DBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5ED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":5FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":60F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":69CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":7821
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":7933
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSendMessage.frx":88AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      Caption         =   "  Recipients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "  Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   600
      Width           =   4365
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "  Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   600
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   465
      Left            =   0
      Top             =   525
      Width           =   8715
   End
   Begin VB.Menu mnuUserList 
      Caption         =   "UserList"
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh User List"
      End
   End
   Begin VB.Menu mnuRecpientList 
      Caption         =   "RecipientList"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete All"
      End
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmSendMessage
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////
Dim Regedit3 As cRegistry

Private Sub cmdAdd_Click()
On Error Resume Next
    If Not txtUserName.Text = "" Then
        lstRecipient.AddItem UCase(txtUserName.Text)
        txtUserName.Text = ""
    Else
        lstRecipient.AddItem UserListTree.SelectedItem.Text
    End If
End Sub

Private Sub SendMessage()
Dim i As Integer
Dim strRecipients As String

'Message Packet Format
'  0 - Message Code || 1- Sender || 2 - Recipients - Delimited by , || 3 - Group || 4 - App Version
'  || 5 - Message || 6 - UserList Delimited by _._ ==>

If lstRecipient.ListCount = 0 Then
    MsgBox "Please select a recipient.", vbInformation, "Express Messenger"
ElseIf txtMessageText.Text = "" Then
    MsgBox "Please enter a message to send.", vbInformation, "Express Messenger"
Else

'Build recipients list
    For i = 0 To lstRecipient.ListCount - 1
    'Check to see if Recipient is still logged on
      If DoesNodeExsist(lstRecipient.List(i), UserListTree) = True Then
        If strRecipients = "" Then
            strRecipients = strRecipients & lstRecipient.List(i)
        Else
            strRecipients = strRecipients & "," & lstRecipient.List(i)
        End If
      Else
        MsgBox lstRecipient.List(i) & " is not logged into the server!", vbInformation
        Exit Sub
      End If
    Next
'Convert URL links and Remedy Ticket Links
     Call convertHyperlink(txtMessageText, "http:", &HFF0000)
     Call convertHyperlink(txtMessageText, "www.", &HFF0000)
'Remedy Help Desk Software Links
     Call convertHyperlink(txtMessageText, "HD", &H8000&)
     Call convertHyperlink(txtMessageText, "#", &H8000&)
'Save Last Sent Message
     Call SaveLastSent(strRecipients, txtMessageText.TextRTF)
'Send Message
    If Not strRecipients = "" Then
       frmReadMessage.sckSYS.SendData ".msg" & "||" & UCase(NickName) & "||" & strRecipients & "||" & " " & _
        "||" & " " & "||" & txtMessageText.TextRTF
    End If
     DoEvents
     Unload Me
End If

End Sub
Function DoesNodeExsist(NodeName As String, ByVal TreeView As TreeView) As Boolean
Dim srchNode
    'This Subprocedure is used to find the n
    '     odes in the treeview
        For Each srchNode In TreeView.Nodes
            If srchNode.Text = NodeName Then
               DoesNodeExsist = True
               Exit Function
            End If
        Next srchNode
        
    DoesNodeExsist = False
End Function
Sub SaveLastSent(strRecipients As String, strMessage As String)
'Add Message Data message buffer
       LastSentMessageRecipients = strRecipients
       LastSentMessage = strMessage
'Enable menu items
       frmReadMessage.mnuLastSentMessage.Enabled = True
       frmReadMessage.mnuClearMessageBuff.Enabled = True
End Sub

Private Sub convertHyperlink(box As RichTextBox, keyWord As String, Color As OLE_COLOR)
    Dim hypStart As Integer
    Dim befor As String
    Dim after As String
    Dim cuvantAddress As String
    Dim hypEnd As Integer
      
    Dim separator1 As String
    Dim separator2 As String
  
    hypStart = box.Find(keyWord, 0, Len(box.Text))
    
  On Error Resume Next
  
    While hypStart >= 0
            separator1 = InStr(hypStart + 1, box.Text, vbCr)
            separator2 = InStr(hypStart + 1, box.Text, Chr(32))
            hypEnd = separator2
            If separator1 > separator2 Then hypEnd = separator2
            If separator2 = 0 Then hypEnd = separator1
            If separator1 = 0 And separator2 = 0 Then hypEnd = Len(box.Text) + 1
            
        cuvantAddress = Mid(box.Text, hypStart + 1, (hypEnd - hypStart))
        
        box.SelStart = hypStart
        box.Find cuvantAddress, hypStart
        box.SelUnderline = True
        box.SelColor = Color
        box.SelStart = hypStart + 1
        
        hypStart = box.Find(keyWord, hypStart + 1, Len(box.Text))
        
    Wend
End Sub

Private Sub cboFont_Click()
    txtMessageText.SelFontName = cboFont.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If the ESC key is pressed the hide the app into the system tray
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Set Regedit3 = New cRegistry
On Error Resume Next
'Disable toolbar buttons
    Toolbar2.Buttons(2).Enabled = False
    Toolbar2.Buttons(3).Enabled = False
    Toolbar2.Buttons(16).Enabled = False
'Load system Fonts
   For i = 0 To Printer.FontCount - 1  ' Determine number of fonts.
      cboFont.AddItem Printer.Fonts(i)    ' Put each font into combo box.
   Next i
'Set Default Font
   cboFont.Text = "Arial"
'Get Current userlist from server
    GetUserList
   
End Sub
  
Sub GetUserList()
On Error Resume Next
    frmReadMessage.sckSYS.SendData ".UserList"
End Sub

Public Sub FillUsersList(User As String, Group As String)

Dim ngroup As Node
Dim nperson As Node

On Error Resume Next

Set ngroup = UserListTree.Nodes.Add(, tvwParent, Group, Group, 1)
Set nperson = UserListTree.Nodes.Add(Group, tvwChild, , User, 4)
    
    
Set ngroup = Nothing
Set nperson = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Regedit3 = Nothing
End Sub

Private Sub lstRecipient_DblClick()
    lstRecipient.RemoveItem lstRecipient.ListIndex
End Sub

Private Sub lstRecipient_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then ' 46 = Delete Key
    lstRecipient.RemoveItem lstRecipient.ListIndex
End If
End Sub

Private Sub lstRecipient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not lstRecipient.ListCount = 0 Then
    If Not lstRecipient.ListIndex = -1 Then mnuDelete.Enabled = True
      If Button = 2 Then
            PopupMenu mnuRecpientList
      End If
End If

End Sub

Private Sub AlignCenter()
    txtMessageText.SelAlignment = rtfCenter
End Sub

Private Sub AlignLeft()
    txtMessageText.SelAlignment = rtfLeft
End Sub

Private Sub AlignRight()
    txtMessageText.SelAlignment = rtfRight
End Sub

Private Sub Color()
    CD1.ShowColor
    txtMessageText.SelColor = CD1.Color
End Sub

Private Sub lstRecipient_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
lstRecipient.AddItem Data.GetData(1)
End Sub

Private Sub mnuDelete_Click()
On Error Resume Next
    lstRecipient.RemoveItem lstRecipient.ListIndex
End Sub

Private Sub mnuDeleteAll_Click()
On Error Resume Next
    lstRecipient.Clear
End Sub
Private Sub SpellCheck(rtbox As RichTextBox)
On Error Resume Next
    Dim wApp As Word.Application    'Object for word application
    Dim doc As Word.Document        'Object for word document
    Dim wd As Word.Words            'object for a collection of words in the document
    Dim wSuggList As Word.SpellingSuggestions 'object for a collection of spelling suggestions (result of a method)
    Dim ss As Word.SpellingSuggestion 'object for one speeling suggestion in the above collection of spelling suggestions
    Dim bPassCheck As Boolean 'Spell Check Results
    Dim sMsg As String 'Holds text to be checked or changed
    Dim i As Integer 'counter
    
    Set wApp = New Word.Application 'Open word
    
    'Add a new document then
    'Copy the all or the selected text from the active textbox to the new document
    'I are using the InsertAfter method to add the text
    'I could have also used the InsertBefore method. In this case it does not matter
    Set doc = wApp.Documents.Add
    
        If txtMessageText.SelLength = 0 Then
            doc.Range.InsertAfter txtMessageText.Text
        Else
            doc.Range.InsertAfter txtMessageText.SelText
        End If
      
    'Create a collection of all the words on the new document
    Set wd = doc.Words
    
    'loop through all words in the list
    'Performing a spell check on each word one at a time
    i = 0
    Do
        i = i + 1
        'Perform Spell Check and store results in bPassCheck
        bPassCheck = wApp.CheckSpelling(wd(i))
        If bPassCheck = False Then 'False the spell check failed
            Set wSuggList = wApp.GetSpellingSuggestions(wd(i))  'get a list of suggestions
            Load frmSpellCheck 'load the formn that displays the list of suggestions (ignored if already loaded)
            frmSpellCheck.txtWord.Text = wd(i)  'Add the bad word
            frmSpellCheck.lstWords.Clear        'clear any existing suggestions
            If wSuggList.Count <> 0 Then        'check to see if there are any suggestions
                For Each ss In wSuggList        'Add the new suggestions from the collection
                    frmSpellCheck.lstWords.AddItem ss.name
                Next
                frmSpellCheck.lstWords.ListIndex = 0    'Select the first item in the list of suggestions
                frmSpellCheck.txtReplaceWith.Text = frmSpellCheck.lstWords.List(frmSpellCheck.lstWords.ListIndex) 'display the text also
            Else
                frmSpellCheck.txtReplaceWith.Text = "" 'No suggestions
            End If
            frmSpellCheck.Show vbModal 'display the spell check form
            'when the user selects ignore, replace, or cancel the form is hidden not unloaded
            'perform indicated action using the properties of the form
            If frmSpellCheck.bCancelCheck = True Then Exit Do
            If frmSpellCheck.bReplaceWord = True Then
                wd(i) = frmSpellCheck.txtReplaceWith.Text & " " 'Add a space as new suggestions don't have the space
            End If
        End If
    Loop Until i = wd.Count 'Loop until there are no more words in the collection
    
    'Copy the text back to the correct textbox on the activeform
        If txtMessageText.SelLength = 0 Then
            sMsg = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf) 'Word only uses CR's for hard breaks.  VB needs CR and LF
            txtMessageText.Text = sMsg
        Else
            txtMessageText.SelText = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf)
            txtMessageText.SelText = sMsg
        End If
    
    'Clean up
    wApp.Quit False  'close word application
    
    Set wApp = Nothing
    Set doc = Nothing
    Set wd = Nothing
    Set wSuggList = Nothing
    
    Unload frmSpellCheck
    
    MsgBox "Spelling Check is complete!", vbInformation, "Express Messenger"
End Sub

Private Sub mnuRefresh_Click()

'Change mouse cursor to the hourglass to show the system is working
    Me.MousePointer = vbHourglass
'Get new userlist from server
    GetUserList
'Change mouse cursor to the default arrow
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Sizes_Click()
   txtMessageText.SelFontSize = Sizes.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        Case "SendMsg"
            Call SendMessage
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        Case "Cut"
            Clipboard.SetText txtMessageText.SelRTF, vbCFRTF
            txtMessageText.SelRTF = " "
        Case "Copy"
            Clipboard.SetText txtMessageText.SelRTF, vbCFRTF
        Case "Paste"
            txtMessageText.SelRTF = Clipboard.GetText(vbCFRTF)
        Case "Bold"
            txtMessageText.SelBold = Not txtMessageText.SelBold
            Button.Value = IIf(txtMessageText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            txtMessageText.SelItalic = Not txtMessageText.SelItalic
            Button.Value = IIf(txtMessageText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            txtMessageText.SelUnderline = Not txtMessageText.SelUnderline
            Button.Value = IIf(txtMessageText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Left"
            txtMessageText.SelAlignment = rtfLeft
        Case "Center"
            txtMessageText.SelAlignment = rtfCenter
        Case "Right"
            txtMessageText.SelAlignment = rtfRight
        Case "CheckSpelling"
            Call SpellCheck(txtMessageText)
        Case "Bullets"
            txtMessageText.SelBullet = Not txtMessageText.SelBullet
            Button.Value = IIf(txtMessageText.SelBullet, tbrPressed, tbrUnpressed)
        Case "Color"
            Call Color
        Case "SaveQM"
            frmAddQM.Show
    End Select
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lstRecipient.AddItem UCase(txtUserName.Text)
    txtUserName.Text = ""
End If
End Sub

Private Sub Toolbar2_ButtonDropDown(ByVal Button As MSComctlLib.Button)
Color
End Sub



Private Sub txtMessageText_Change()
If Not txtMessageText.Text = "" Then
    Toolbar2.Buttons(2).Enabled = True
    Toolbar2.Buttons(3).Enabled = True
     If CBool(Len(Regedit3.getstring(HKEY_CLASSES_ROOT, "Word.Application\CurVer", ""))) = True Then
         Toolbar2.Buttons(16).Enabled = True
     End If
Else
    Toolbar2.Buttons(2).Enabled = False
    Toolbar2.Buttons(3).Enabled = False
    Toolbar2.Buttons(16).Enabled = False
End If
End Sub

Private Sub txtMessageText_SelChange()
    Toolbar2.Buttons(6).Value = IIf(txtMessageText.SelBold, tbrPressed, tbrUnpressed)
    Toolbar2.Buttons(7).Value = IIf(txtMessageText.SelItalic, tbrPressed, tbrUnpressed)
    Toolbar2.Buttons(10).Value = IIf(txtMessageText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    Toolbar2.Buttons(11).Value = IIf(txtMessageText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    Toolbar2.Buttons(12).Value = IIf(txtMessageText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    Toolbar2.Buttons(14).Value = IIf(txtMessageText.SelBullet, tbrPressed, tbrUnpressed)
End Sub

Private Sub UserListTree_DblClick()
On Error Resume Next
    lstRecipient.AddItem UserListTree.SelectedItem.Text

End Sub

Private Sub UserListTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserListTree.OLEDrag
If Button = 2 Then
    PopupMenu mnuUserList
End If
End Sub


