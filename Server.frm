VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Server"
   ClientHeight    =   7020
   ClientLeft      =   930
   ClientTop       =   2490
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChkStatus 
      Caption         =   "Check User's Status"
      Height          =   375
      Left            =   3900
      TabIndex        =   7
      Top             =   6150
      Width           =   1710
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Event Log"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      Top             =   6150
      Width           =   1710
   End
   Begin VB.Timer tmrRemoveDisconnected 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3075
      Top             =   6150
   End
   Begin VB.Timer tmrCheckConnections 
      Interval        =   60000
      Left            =   2625
      Top             =   6150
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   6780
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8281
            Text            =   "Server Running Time : 0 hour(s) 0 minute(s) 0 second(s)"
            TextSave        =   "Server Running Time : 0 hour(s) 0 minute(s) 0 second(s)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Users Connected "
            TextSave        =   "Users Connected "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:06 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "11/29/2001"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimeElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2175
      Top             =   6150
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   6075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Server.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Server.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmMessageServer 
      Caption         =   "Connected Users:"
      Height          =   2520
      Left            =   150
      TabIndex        =   3
      Top             =   3525
      Width           =   9075
      Begin MSComctlLib.ListView lstUsers 
         Height          =   2190
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   3863
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Connection Number"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Nick Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "IP Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Group"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Version"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date & Time"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSWinsockLib.Winsock ServiceSocket 
      Index           =   0
      Left            =   750
      Top             =   6225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3535
   End
   Begin VB.CommandButton cmdShutDown 
      Caption         =   "Shutdown Server"
      Height          =   375
      Left            =   7500
      TabIndex        =   1
      Top             =   6150
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Log:"
      Height          =   3345
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   9045
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3015
         Left            =   75
         TabIndex        =   2
         Top             =   225
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   5318
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Server.frx":1A7E
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings..."
   End
   Begin VB.Menu mnuLstRightClick 
      Caption         =   "List Right Click"
      Visible         =   0   'False
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message to User"
      End
      Begin VB.Menu mnuSendAll 
         Caption         =   "Send Message to All Users"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisConnect 
         Caption         =   "Disconnect User"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGE SERVER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmServer
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Dim Regedit As New cRegistry
Public intMax As Integer
Public MessageSent As Boolean

Dim intHoursElasped As Integer
Dim intSecondsElasped As Integer
Dim intMinutesElasped As Integer
Dim intStatusCheck As Integer
Dim IntTmrMinutes As Integer
Dim CheckStatus As Boolean

Private Function FindOpenWinsock()
Dim x As Integer
Dim i As Integer
Dim SocketExists As Boolean

On Error GoTo Find_Socket_Error

For x = 1 To ServiceSocket.UBound
    If ServiceSocket(x).State = 0 Then
    ' We found one that's state is 0, which
    ' means "closed", so let's use it
    '----------------------------------------------
    'check to make sure te index does not exist in the userlist
        For i = 1 To lstUsers.ListItems.Count
         If Not lstUsers.ListItems(i) = x Then
            SocketExists = False
         Else
            SocketExists = True
         End If
        Next i
    'The socket does not exist in the userlist so use it
            If SocketExists = False Then
                FindOpenWinsock = x
                Exit Function
            End If
    End If
Next x

    '  OK, none are open so let's make one
    Load ServiceSocket(ServiceSocket.UBound + 1)
    
    '  and then let's return it's index value
    FindOpenWinsock = ServiceSocket.UBound
    
Exit Function
Find_Socket_Error:
   SaveEventLog "Find Open Winsock Encountered error # " & Err.Number & ", " & Err.Description
   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = "Find Open Winsock Encountered error # " & Err.Number & ", " & Err.Description
   RichTextBox1.SelColor = vbBlack
End Function

Private Sub cmdChkStatus_Click()
    CheckConnections
    DoEvents
End Sub

Private Sub cmdClear_Click()
    RichTextBox1.Text = ""
    RichTextBox1.SelColor = vbBlack
End Sub

Private Sub cmdShutDown_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Set Regedit = New cRegistry
   ServiceSocket(0).Listen
'Set Application Title
   Me.Caption = "Message Server - " & App.Major & "." & App.Minor & "." & App.Revision
'Reset Timers
   intMinutesElasped = 0
   intHoursElasped = 0
   intSecondsElasped = 0
   IntTmrMinutes = 1
   
   StatusBar1.Panels(1).Text = "Server Started : " & Now()
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer

On Error Resume Next

If lstUsers.ListItems.Count = 0 Then
    Cancel = 0
    Unload Me
Else
   Dim response
    
    response = MsgBox("There are " & lstUsers.ListItems.Count & " users attached to the server." & vbCrLf & _
    vbCrLf & "Would you like to disconnect them ?", vbInformation + vbYesNo)
    
    If response = vbYes Then
    'disconnect all attached users
    On Error Resume Next
        For i = 1 To lstUsers.ListItems.Count
            ServiceSocket(lstUsers.ListItems(i)).SendData ".ServerDisconnect"
        Next i
        
        DoEvents
        
        Cancel = 0
    Else
        Cancel = 1
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Regedit = Nothing
End Sub

Private Sub lstUsers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   ' SortListView lstUsers, ColumnHeader
End Sub

Private Sub lstUsers_DblClick()
    MsgBox "State: " & ServiceSocket(Int(lstUsers.SelectedItem.Text)).State & vbCrLf & _
    "Index: " & lstUsers.SelectedItem.index & vbCrLf & _
    "State Tag: " & lstUsers.SelectedItem.Tag
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not lstUsers.ListItems.Count = 0 Then
    If Button = 2 Then
        PopupMenu mnuLstRightClick
    End If
End If
End Sub

Private Sub mnuDisConnect_Click()
On Error Resume Next
If ServiceSocket(Int(lstUsers.SelectedItem.Text)).State = 7 Then
    ServiceSocket(Int(lstUsers.SelectedItem.Text)).SendData ".ServerDisconnect"
    DoEvents
    ServiceSocket(Int(lstUsers.SelectedItem.Text)).Close
'Display and save Status in event Log
    SaveEventLog Now & ": Connection closed for " & ServiceSocket(lstUsers.SelectedItem.Text).RemoteHostIP & vbCrLf
    RichTextBox1.SelColor = &HFF&
    RichTextBox1.SelText = Now & ": Connection closed for " & ServiceSocket(lstUsers.SelectedItem.Text).RemoteHostIP & vbCrLf
    RichTextBox1.SelColor = vbBlack
'Remove item From List
   lstUsers.ListItems.Remove (lstUsers.SelectedItem.index)
Else
'Remove item From List
    lstUsers.ListItems.Remove (lstUsers.SelectedItem.index)
End If
End Sub

Private Sub mnuSendAll_Click()
Dim Message As String
Dim i As Integer
On Error Resume Next
    Message = InputBox("Please enter messages to send." & vbCrLf & vbCrLf & _
    "TO: All Users" & vbCrLf & "FROM: Message Server", "Send Message")
  'Send Message to all logged in users
    If Not Message = "" Then
        For i = 1 To lstUsers.ListItems.Count
         If ServiceSocket(i).State = 7 Then
            ServiceSocket(i).SendData ".msg" & "||" & "Message Server" & "||" & lstUsers.ListItems(i).ListSubItems(1) & "||" & " " & _
            "||" & 0 & "||" & Message
            
            DoEvents
         End If
        Next i
    End If
End Sub

Private Sub mnuSendMessage_Click()
Dim Message As String
On Error Resume Next
    Message = InputBox("Please enter messages to send." & vbCrLf & vbCrLf & _
    "TO: " & lstUsers.SelectedItem.ListSubItems(1) & vbCrLf & "FROM: Message Server", "Send Message")
  'Send Message
    If Not Message = "" Then
        If ServiceSocket(Int(lstUsers.SelectedItem.Text)).State = 7 Then
           ServiceSocket(Int(lstUsers.SelectedItem.Text)).SendData ".msg" & "||" & "Message Server" & "||" & lstUsers.SelectedItem.ListSubItems(1) & "||" & " " & _
           "||" & 0 & "||" & Message
        End If
    End If
End Sub

Private Sub mnuSettings_Click()
    frmOptions.Show vbModal
End Sub

Private Sub RichTextBox1_Change()
   RichTextBox1.SelStart = Len(RichTextBox1)
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RichTextBox1.SelStart = Len(RichTextBox1)
End Sub



Private Sub ServiceSocket_Close(index As Integer)
Dim i As Integer

On Error Resume Next

'Remove User from User List
'Let's cycle through the list, looking for their name
    For i = 1 To lstUsers.ListItems.Count
        ' Check to see if it matches
        If lstUsers.ListItems(i).Text = index Then
            ' It matches, so let's remove it form the
            ' list
            lstUsers.ListItems.Remove i
            Exit For
        End If
    Next i
'Close Socket connection
   ServiceSocket(index).Close
'Display and save Status in event Log
   SaveEventLog Now & ": Connection closed for " & ServiceSocket(index).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = &HFF&
   RichTextBox1.SelText = Now & ": Connection closed for " & ServiceSocket(index).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = vbBlack
   
End Sub

Private Sub ServiceSocket_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim Connection_Number As Integer

Connection_Number = FindOpenWinsock
'Accept the request using the created winsock
    ServiceSocket(Connection_Number).Accept requestID
'Display and save Status in event Log
        SaveEventLog Now & ": New connection request from " & ServiceSocket(Connection_Number).RemoteHostIP & vbCrLf
        RichTextBox1.SelText = Now & ": New connection request from " & ServiceSocket(Connection_Number).RemoteHostIP & vbCrLf
        RichTextBox1.SelColor = vbBlack

End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    ServiceSocket_DataArrival
'
'  AUTHOR:       Michael J. Kempf        8/1/2001 7:29:50 AM
'
'  PURPOSE:      This procedure accepts all the requests from the client application.
'
'                User Login -This request logged a user into the server
'                Message -   This request is where sending of messages is processed
'                User List - This request sends back all users and groups connected to the server
'                Connected - This request is a response from the client application that the connection is good.
'
'
'///////////////////////////////////////////////////////////////////////////////////////////
Private Sub ServiceSocket_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim vntUserCommands As Variant
Dim vntRecipients As Variant
Dim u As Integer
Dim intItems As Integer
Dim counter As Integer
Dim UserData As String
Dim MessageLogged As Boolean

On Error GoTo Data_Recv_Error

    ServiceSocket(index).GetData UserData
    
    vntUserCommands = Split(UserData, "||")

Select Case vntUserCommands(0)

Case ".UserLogin"
     Dim i As Integer
     Dim NickNameExists As Boolean
   'Check to see if the NickName trying to login already exists
     For i = 1 To lstUsers.ListItems.Count
        If vntUserCommands(1) = lstUsers.ListItems(i).SubItems(1) Then
            If ServiceSocket(lstUsers.ListItems(i)).State = 7 Then
                NickNameExists = True
                Exit For
            End If
        End If
     Next
    'NickName already exists, send response back to client
     If NickNameExists Then
        ServiceSocket(index).SendData ".NickNameExists"
     Else
    'User connected OK
        ServiceSocket(index).SendData ".ConnectedOK"
    'Add logged in user to active user list

       lstUsers.ListItems.Add , , Trim(index)
       lstUsers.ListItems(lstUsers.ListItems.Count).ListSubItems.Add , , vntUserCommands(1)
       lstUsers.ListItems(lstUsers.ListItems.Count).ListSubItems.Add , , ServiceSocket(index).RemoteHostIP
       lstUsers.ListItems(lstUsers.ListItems.Count).ListSubItems.Add , , vntUserCommands(3)
       If intItems = 4 Then
        lstUsers.ListItems(lstUsers.ListItems.Count).ListSubItems.Add , , vntUserCommands(4)
       End If
       lstUsers.ListItems(lstUsers.ListItems.Count).ListSubItems.Add , , Now()
       
       lstUsers.ListItems(lstUsers.ListItems.Count).SmallIcon = 1
   'Display ans save Status in event Log
            SaveEventLog Now & ": User: " & vntUserCommands(1) & " logged in from " & ServiceSocket(index).RemoteHostIP & vbCrLf
            RichTextBox1.SelColor = &HFF0000
            RichTextBox1.SelText = Now & ": User: " & vntUserCommands(1) & " logged in from " & ServiceSocket(index).RemoteHostIP & vbCrLf
            RichTextBox1.SelColor = vbBlack
           
    End If
Case ".msg"
    counter = 0
    vntRecipients = Split(vntUserCommands(2), ",")

    intItems = UBound(vntRecipients) + 1

Do
    'The First Socket Connecton starts at 1
           For u = 1 To lstUsers.ListItems.Count
              If lstUsers.ListItems(u).SubItems(1) = vntRecipients(counter) Then
                If ServiceSocket(lstUsers.ListItems(u)).State = 7 Then
                 'Check to see if the option is enabled to save messages
                     If Regedit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveUserMsgs") = 1 Then
                          Call SaveMessages(CStr(vntUserCommands(1)), CStr(vntUserCommands(2)), CStr(vntUserCommands(5)))
                     End If
                   'Send Message
                      ServiceSocket(lstUsers.ListItems(u)).SendData ".msg" & "||" & vntUserCommands(1) & "||" & vntUserCommands(2) & "||" & _
                      vntUserCommands(3) & "||" & vntUserCommands(4) & "||" & vntUserCommands(5)
                      MessageSent = True
                   'Display and save Status in event Log
                        SaveEventLog Now & ": User: " & vntRecipients(counter) & " received new message from " & vntUserCommands(1) & "(" & _
                        ServiceSocket(index).RemoteHostIP & ")" & vbCrLf
                        RichTextBox1.SelColor = &H8000&
                        RichTextBox1.SelText = Now & ": User: " & vntRecipients(counter) & " received new message from " & vntUserCommands(1) & "(" & _
                        ServiceSocket(index).RemoteHostIP & ")" & vbCrLf
                        RichTextBox1.SelColor = vbBlack
                      'Exit For Loop
                      Exit For
                 End If
               End If
            Next u
    'If no Message was sent to a specific user then
    'check if it needs to sent to a group
    If MessageSent = False Then
    'Reset Counter for the sockets
        u = 1
        For u = 1 To lstUsers.ListItems.Count
              If lstUsers.ListItems(u).SubItems(3) = vntRecipients(counter) Then
                If ServiceSocket(lstUsers.ListItems(u)).State = 7 Then
                   'Send Message to recipient
                      ServiceSocket(lstUsers.ListItems(u)).SendData ".msg" & "||" & vntUserCommands(1) & "||" & vntUserCommands(2) & "||" & _
                      vntUserCommands(3) & "||" & vntUserCommands(4) & "||" & vntUserCommands(5)
                   'Display and save Status in event Log
                        SaveEventLog Now & ": Group: " & vntUserCommands(2) & " User: " & lstUsers.ListItems(u).SubItems(1) & _
                        " received new message from " & vntUserCommands(1) & "(" & ServiceSocket(index).RemoteHostIP & ")" & vbCrLf
                        RichTextBox1.SelColor = &H8000&
                        RichTextBox1.SelText = Now & ": Group: " & vntUserCommands(2) & " User: " & lstUsers.ListItems(u).SubItems(1) & _
                        " received new message from " & vntUserCommands(1) & "(" & ServiceSocket(index).RemoteHostIP & ")" & vbCrLf
                        RichTextBox1.SelColor = vbBlack
                End If
              End If
         Next u
         'Check to see if the option is enabled to save messages
                     If Regedit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "SaveGroupMsgs") = 1 Then
                         If MessageLogged = False Then
                            Call SaveMessages(CStr(vntUserCommands(1)), CStr(vntUserCommands(2)), CStr(vntUserCommands(5)))
                            MessageLogged = True
                         End If
                      End If
     End If
     
     DoEvents
     
     counter = counter + 1
     MessageSent = False
   
Loop Until counter = intItems

Case ".UserList"
'Send back the userlist to the requester
    ServiceSocket(index).SendData ".UserList" & "||1||2||3||4||5" & GetUserList
    
Case ".Connected"
    RichTextBox1.SelColor = &HC000C0
    RichTextBox1.SelText = "Connection (" & index & ") Verification received from " & ServiceSocket(index).RemoteHostIP & vbCrLf
    RichTextBox1.SelColor = vbBlack
    
     For intStatusCheck = 1 To lstUsers.ListItems.Count
        If lstUsers.ListItems(intStatusCheck) = index Then
             lstUsers.ListItems(intStatusCheck).Tag = 7
             lstUsers.ListItems(intStatusCheck).SmallIcon = 1
            Exit For
        End If
        DoEvents
     Next
    
Case ".RemoteDisconnect"
    For i = 1 To lstUsers.ListItems.Count
        If lstUsers.ListItems(i).SubItems(1) = vntUserCommands(1) Then
          On Error Resume Next
           If ServiceSocket(lstUsers.ListItems(i)).State = 7 Then
                ServiceSocket(lstUsers.ListItems(i)).SendData ".ServerDisconnect"
                DoEvents
                ServiceSocket(lstUsers.ListItems(i)).Close
            'Display and save Status in event Log
                SaveEventLog Now & ": Connection closed for " & ServiceSocket(lstUsers.ListItems(i)).RemoteHostIP & vbCrLf
                RichTextBox1.SelColor = &HFF&
                RichTextBox1.SelText = Now & ": Connection closed for " & ServiceSocket(lstUsers.ListItems(i)).RemoteHostIP & vbCrLf
                RichTextBox1.SelColor = vbBlack
            'Remove item From List
                lstUsers.ListItems.Remove (lstUsers.ListItems(i).index)
            Else
            'Remove item From List
                lstUsers.ListItems.Remove (lstUsers.ListItems(i).index)
            End If
         End If
    Next
    
    
End Select
Exit Sub
Data_Recv_Error:
   SaveEventLog "ServiceSocket_DataArrival - " & Err.Number & ", " & Err.Description & vbCrLf
   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = "ServiceSocket_DataArrival - " & Err.Number & ", " & Err.Description & vbCrLf
   RichTextBox1.SelColor = vbBlack
End Sub

Private Sub ServiceSocket_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 On Error GoTo Error_Handler
 
 'Display and save Status in event Log
   SaveEventLog "User: " & lstUsers.ListItems(index).SubItems(1) & " - " & Now & ": Encountered error #" & _
   Number & ", " & Description & ", from " & ServiceSocket(index).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = "User: " & lstUsers.ListItems(index).SubItems(1) & " - " & Now & ": Encountered error #" & _
   Number & ", " & Description & ", from " & ServiceSocket(index).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = vbBlack
Exit Sub
Error_Handler:
   SaveEventLog "ServiceSocket_Error - " & Err.Number & ", " & Err.Description
   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = "ServiceSocket_Error - " & Err.Number & ", " & Err.Description
   RichTextBox1.SelColor = vbBlack
End Sub
Sub SaveEventLog(EventMessage As String)
On Error Resume Next
MkDir App.Path & "\Logs"
    Open App.Path & "\Logs\EventLog.log" For Append As #2
        Print #2, EventMessage
    Close #2
End Sub

Sub SaveMessages(Sender As String, Recipient As String, Message As String)
On Error Resume Next
    Dim LogFilePath As String
'Get the Log file path from the registry
    LogFilePath = Regedit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "LogPath")
'If no log file pah exists the set to the application path
    If LogFilePath = "" Then LogFilePath = App.Path & "\Logs\"
'Open Log file
    Open LogFilePath & Month(Date) & Day(Date) & Year(Date) & ".log" For Append As #1
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

Function GetUserList() As String
Dim i As Integer
On Error Resume Next
For i = 1 To lstUsers.ListItems.Count
'Populate the string with all the logged in users and groups
    GetUserList = GetUserList & "||" & lstUsers.ListItems(i).SubItems(1) & "_._" & lstUsers.ListItems(i).SubItems(3)
Next
End Function

Private Sub TimeElapsed_Timer()

If intSecondsElasped = 60 Then
    intMinutesElasped = intMinutesElasped + 1
    intSecondsElasped = 0
ElseIf intMinutesElasped = 60 Then
  'Clear Event Log every hour
    RichTextBox1.Text = ""
    
    intHoursElasped = intHoursElasped + 1
    intMinutesElasped = 0
Else
    intSecondsElasped = intSecondsElasped + 1
    StatusBar1.Panels(2).Text = "Users Connected : " & lstUsers.ListItems.Count
End If

StatusBar1.Panels(1).Text = "Server Running Time : " & intHoursElasped & " hour(s) " & _
intMinutesElasped & " minute(s) " & intSecondsElasped & " second(s)"
End Sub

Sub CheckConnections()
Dim i As Integer


On Error Resume Next
   RichTextBox1.SelColor = &H80FF&
   RichTextBox1.SelText = "Checking Connections......" & vbCrLf
   RichTextBox1.SelColor = vbBlack
    'Loop through all the connections and check for a connected state
    'If not connected then remove from list
     
   For intStatusCheck = 1 To lstUsers.ListItems.Count
       lstUsers.ListItems(intStatusCheck).Tag = 0
       lstUsers.ListItems(intStatusCheck).SmallIcon = 2
       ServiceSocket(lstUsers.ListItems(intStatusCheck)).SendData ".State_Check"
       Sleep Int(Regedit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "Sleep"))
   Next
   
End Sub

Sub RemoveDisconnected()
Dim i As Integer
On Error Resume Next

'Loop throug the state conection array to see who has a sent back a state response
    For i = 1 To lstUsers.ListItems.Count
        If lstUsers.ListItems(i).Tag = 0 Then
        'Display and save Status in event Log
            SaveEventLog Now & ": Auto Disconnect for " & lstUsers.ListItems(i).SubItems(1) & " - " & _
            lstUsers.ListItems(i).SubItems(2) & vbCrLf
            RichTextBox1.SelColor = &HFF&
            RichTextBox1.SelText = Now & ": Auto Disconnect for " & lstUsers.ListItems(i).SubItems(1) & " - " & _
            lstUsers.ListItems(i).SubItems(2) & vbCrLf
            RichTextBox1.SelColor = vbBlack
        'Remove User
            lstUsers.ListItems.Remove i
            DoEvents
        End If
    Next
End Sub

Private Sub tmrCheckConnections_Timer()
If Not lstUsers.ListItems.Count = 0 Then
    If IntTmrMinutes = Int(Regedit.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "Minutes")) Then
    'reset timer minutes
        IntTmrMinutes = 1
        If Regedit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "VerifyConn") = 1 Then
            'check all the connection to the server
            CheckConnections
            DoEvents
            tmrRemoveDisconnected.Enabled = True
            tmrCheckConnections.Enabled = False
        End If
    Else
        IntTmrMinutes = IntTmrMinutes + 1
    End If
End If
End Sub

Private Sub tmrRemoveDisconnected_Timer()
If Regedit.getdword(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\Message Server\Settings", "VerifyConn") = 1 Then
'Remove all disconnected users
        RemoveDisconnected
        DoEvents
        
     tmrRemoveDisconnected.Enabled = False
     tmrCheckConnections.Enabled = True
Else
     tmrRemoveDisconnected.Enabled = False
     tmrCheckConnections.Enabled = True
End If
End Sub

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub

Function GetWSAErrorString(ByVal errnum As Long) As String
    On Error Resume Next
    Select Case errnum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else: GetWSAErrorString = "Unknown Error"
    End Select
End Function
