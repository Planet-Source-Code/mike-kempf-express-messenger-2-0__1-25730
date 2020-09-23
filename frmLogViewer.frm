VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogViewer 
   Caption         =   "Express Message Log Viewer"
   ClientHeight    =   5340
   ClientLeft      =   5070
   ClientTop       =   3915
   ClientWidth     =   5130
   Icon            =   "frmLogViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4125
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTViewDate 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   150
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   59572225
      CurrentDate     =   37029
   End
   Begin VB.TextBox RemedyDDE 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "RemedyDDE"
      Top             =   5775
      Visible         =   0   'False
      Width           =   1065
   End
   Begin RichTextLib.RichTextBox rtLogView 
      Height          =   4590
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   8096
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLogViewer.frx":08CA
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmLogViewer.frx":0941
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblSelectDate 
      Caption         =   "Select Date to View:"
      Height          =   240
      Left            =   900
      TabIndex        =   2
      Top             =   220
      Width           =   1590
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuprint 
         Caption         =   "&Print Log"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   FORM NAME:          frmLogViewer
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

Option Explicit
Dim htxt As String
Dim Regedit3 As New cRegistry

Private Sub DTViewDate_CloseUp()
'View the selected log file
    ViewLog (DTViewDate.Value)
    Me.Caption = "Express Message Log Viewer - " & DTViewDate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If the ESC key is pressed the hide the app into the system tray
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Set Regedit3 = New cRegistry
ViewLog (Date)
DTViewDate.Value = Date
 Me.Caption = "Express Message Log Viewer - " & Date
End Sub

Private Sub Form_Resize()

 On Error Resume Next
If Not Me.WindowState = vbMinimized Then
    If Me.Height < 5805 Then
        Me.Height = 5805
    ElseIf Me.Width < 5250 Then
        Me.Width = 5250
    Else
        rtLogView.Height = Me.Height - 1450
        rtLogView.Width = Me.Width - 250
    End If
End If

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuprint_Click()
    Call PrintLog
End Sub

Private Sub rtLogView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 htxt = GetHyperlink(rtLogView, X, Y)
End Sub
Private Sub rtLogView_Click()
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
            Call DDEExecute("ARUSER-SERVER", "DoExecMacro", "[RunMacro (" & App.Path & ",EMTicket,TICKET=" & htxt & ")]")
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
    End If
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    DDEExecute
'
'  AUTHOR:       Kemtech Software        06/20/2001 11:47:47 AM
'
'  PURPOSE:      Execute a DDE link between the application and the Remedy Helpdesk system
'
'  PARAMETERS:
'                sApplication (String) = Application to send DDE Command to
'                sTopic       (String) = Specific task to run
'                sMacro       (String) = Macro to Launch
'
Public Sub DDEExecute(sApplication As String, sTopic As String, sMacro As String)

On Error Resume Next

'Link Topic
    RemedyDDE.LinkTopic = sApplication & "|" & sTopic
'Set link Mode
    RemedyDDE.LinkMode = 2
'minimize application
    Me.Hide
'Execute Command
    RemedyDDE.LinkExecute sMacro
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    ViewLog
'
'  AUTHOR:       Michael J. Kempf        7/30/2001 11:30:26 AM
'
'  PURPOSE:      This procedure loads a Log file for viewing
'
'  PARAMETERS:
'                dDate (Date) = Date to View
'
'///////////////////////////////////////////////////////////////////////////////////////////
Private Sub ViewLog(dDate As Date)

Dim sfile As String
Dim spath As String
Dim iyear As Integer


'Get Year Of Date
    iyear = Year(dDate)
'Get Path
    spath = Regedit3.getstring(HKEY_LOCAL_MACHINE, "Software\Kemtech Software\EM\Settings", "LogPath") & "\"
'parse date to get file name
    sfile = Format$(dDate, "m") & Format$(dDate, "d") & Format$(dDate, "yyyy") & ".log"
 On Error GoTo NotFoundError
'Load file
    rtLogView.TextRTF = ""
    rtLogView.LoadFile (spath & sfile), rtfRTF
    
Exit Sub
NotFoundError:
    Select Case Err.Number
        Case 75
            MsgBox "No Messages Logged for " & dDate & " !", vbInformation, "Express Message Log Viewer"
    End Select

End Sub

Private Function GetHyperlink(rch As RichTextBox, X As Single, Y As Single) As String
    'This function return any word curently under cursor and,
    'if string under cursor start with "http:","www.","HD",and "#"
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

    
Private Sub PrintLog()

    On Local Error GoTo Error_Handler:
     With CD1
        .CancelError = True
        .ShowPrinter
        
                 If rtLogView.SelLength = 0 Then
                    .Flags = .Flags + cdlPDAllPages
                 Else
                    .Flags = .Flags + cdlPDSelection
                 End If
                 
            On Local Error Resume Next
                   Printer.Print ""
                   Printer.Print "Daily Log for " & DTViewDate
                   Printer.Print ""
                   Printer.Print rtLogView.Text
                   Printer.EndDoc
     End With
Exit Sub

Error_Handler:
    
    If Err <> cdlCancel Then
    MsgBox " Error " & Err & "; " & Error, vbExclamation, "Express Messenger"
    End If
End Sub

