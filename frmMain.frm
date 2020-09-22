VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "P2P Pro"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   960
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3095
   End
   Begin MSWinsockLib.Winsock sckRec 
      Left            =   480
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   3095
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   960
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7802
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   360
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7802
   End
   Begin VB.CommandButton cmdDisco 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame FrameConvo 
      Caption         =   " Conversation (Not Started) "
      Height          =   3615
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   4815
      Begin MSComDlg.CommonDialog CD 
         Left            =   1800
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtSend 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   3615
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   2520
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4445
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0ECA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nickname :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5235
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status : Idle."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameHost 
      Caption         =   " Remote Host (IP Address) "
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCLS 
         Caption         =   "Clear Chat Text"
      End
      Begin VB.Menu mnuSaveConvo 
         Caption         =   "Save Conversation"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAddressBook 
         Caption         =   "Address Book"
      End
      Begin VB.Menu mnuConnectToUser 
         Caption         =   "Connect to User"
      End
      Begin VB.Menu mnuWhoIs 
         Caption         =   "WhoIs Query"
      End
      Begin VB.Menu mnuSendBugReport 
         Caption         =   "Send Bug Report!"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send a File..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About P2P Pro"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help Contents"
      End
      Begin VB.Menu mnuVisitWebpage 
         Caption         =   "Visit P2P Pro Webpage"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
If Len(Trim$(txtHost.Text)) = 0 Then
    MsgBox "Enter a host to connect to", vbCritical, "Host Required"
    txtHost.SetFocus
    Exit Sub
End If
SaveSetting sApp, "Main", "Host", EncryptHex(StrReverse$(txtHost.Text), True)
SessionInfo.RemoteHost = txtHost.Text
SessionInfo.FT_InProgress = False
SessionInfo.Connected = False
SessionInfo.Nick = txtNick.Text
SessionInfo.SessionType = SESSION_CLIENT
sckServer.Close
sckClient.Close
sckRec.Close
sckSend.Close
sckClient.Connect txtHost.Text, 7802
StatusBar.SimpleText = "Status : Connecting to User . . ."
End Sub

Private Sub cmdDisco_Click()
If Not SessionInfo.Connected Then
    MsgBox "You are not currently connected", vbCritical, "No Present Connection"
Else
    If SessionInfo.SessionType = SESSION_SERVER Then
        sckServer.Close
        SessionInfo.Connected = False
        StatusBar.SimpleText = "Status : Disconnected."
        Call AddRTFStatus("You have closed the connection.", RGB(123, 0, 0))
        sckRec.Close
        sckSend.Close
        SessionInfo.FT_InProgress = False
    ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
        sckClient.Close
        SessionInfo.Connected = False
        StatusBar.SimpleText = "Status : Disconnected."
        Call AddRTFStatus("You have closed the connection.", RGB(123, 0, 0))
        sckRec.Close
        sckSend.Close
        SessionInfo.FT_InProgress = False
    End If
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If Len(txtSend.Text) > 0 Then
    If SessionInfo.Connected = False Then
        Call AddRTFStatus("No Connection Is Present.", RGB(123, 0, 0))
        Exit Sub
    End If
    
    If SessionInfo.SessionType = SESSION_SERVER Then
        sckServer.SendData sHeader & sDelim & "M" & sDelim & txtNick.Text & sDelim & txtSend.Text
    ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
        sckClient.SendData sHeader & sDelim & "M" & sDelim & txtNick.Text & sDelim & txtSend.Text
    End If
    With txtChat
        If Len(.Text) = 0 Then
            .SelBold = True
            .SelItalic = False
            .SelUnderline = False
            .SelColor = RGB(0, 0, 123)
            .SelText = txtNick.Text
            .SelBold = False
            .SelColor = vbBlack
            .SelText = ": "
            .SelColor = RGB(0, 0, 123)
            .SelText = txtSend.Text
        Else
            .SelBold = True
            .SelItalic = False
            .SelUnderline = False
            .SelColor = RGB(0, 0, 123)
            .SelText = vbNewLine & txtNick.Text
            .SelBold = False
            .SelColor = vbBlack
            .SelText = ": "
            .SelColor = RGB(0, 0, 123)
            .SelText = txtSend.Text
        End If
    End With
    txtSend.Text = Empty
    txtSend.SetFocus
End If
       
End Sub

Private Sub Form_Load()
FileNum = FreeFile
Call MakeRecDir
Call ReadPref(False)
On Error Resume Next
txtHost.Text = StrReverse$(EncryptHex(GetSetting(sApp, "Main", "Host", ""), False))
txtNick.Text = StrReverse$(EncryptHex(GetSetting(sApp, "Main", "Nick", ""), False))

If Len(Trim$(txtHost.Text)) = 0 Then
    cmdConnect.Enabled = False
Else
    cmdConnect.Enabled = True
End If

If Len(txtNick.Text) = 0 Then
    cmdSend.Enabled = False
Else
    cmdSend.Enabled = True
End If
On Error Resume Next
sckServer.Listen
SessionInfo.SessionType = SESSION_SERVER
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Len(txtNick.Text) > 0 Then
    SaveSetting sApp, "Main", "Nick", EncryptHex(StrReverse$(txtNick.Text), True)
End If
On Error Resume Next
Unload frmAbout
Unload frmAddContact
Unload frmAddressBook
Unload frmPreferences
Unload frmRecFile
Unload frmSendFile
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Call ReadPref(False)
FrameHost.Width = Me.Width - 630
txtHost.Width = FrameHost.Width - 240
FrameConvo.Width = Me.Width - 660
txtChat.Width = FrameConvo.Width - 240
txtSend.Width = FrameConvo.Width - 1200
cmdSend.Left = txtSend.Width + 225
FrameConvo.Height = Me.Height - 2625
txtChat.Height = FrameConvo.Height - 1095
txtSend.Top = txtChat.Height + 720
cmdSend.Top = txtSend.Top
cmdDisco.Top = FrameConvo.Height + 1065
cmdConnect.Top = cmdDisco.Top
cmdConnect.Left = FrameConvo.Width - 1455
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show

End Sub

Private Sub mnuAddressBook_Click()
frmAddressBook.Show
End Sub

Private Sub mnuCLS_Click()
txtChat.Text = Empty
End Sub

Private Sub mnuConnectToUser_Click()
cmdConnect_Click
End Sub

Private Sub mnuExit_Click()
If Len(txtNick.Text) > 0 Then
    SaveSetting sApp, "Main", "Nick", EncryptHex(StrReverse$(txtNick.Text), True)
End If
On Error Resume Next
Unload frmAbout
Unload frmAddContact
Unload frmAddressBook
Unload frmPreferences
Unload frmRecFile
Unload frmSendFile
Unload Me
End Sub

Private Sub mnuHelpContents_Click()
Call OpenURL("http://www.Digital-Revolution.org/P2P/Help.html", vbNormalFocus, 0&)
End Sub

Private Sub mnuPreferences_Click()
frmPreferences.Show , Me
End Sub

Private Sub mnuSaveConvo_Click()
With CD
.DialogTitle = "Save Conversation"
.Filter = "Rich Text Files|*.rtf|Text Files|*.txt"
.ShowSave
If Len(.FileName) > 0 Then
    Dim iFile As Integer
    iFile = GetFileType(.FileName)
        If iFile = FILETYPE_RTF Then
            txtChat.SaveFile .FileName, rtfRTF
        Else
            txtChat.SaveFile .FileName, rtfText
        End If
End If
End With
End Sub



Private Sub mnuSendBugReport_Click()
Call OpenURL("http://www.Digital-Revolution.org/P2P/BugReport.html", vbNormalFocus, 0&)
End Sub

Private Sub mnuSendFile_Click()
If SessionInfo.FT_InProgress Then
    MsgBox "File transfer is already in progress; please wait for it to finish", vbCritical, "File Transfer in Progress"
    Exit Sub
Else
    frmSendFile.Show , Me
    frmSendFile.txtPath.Text = Empty: frmSendFile.lblKBPS.Caption = "0 KB/Sec": frmSendFile.lblBS.Caption = "Bytes Sent : 0": frmSendFile.StatusBar.SimpleText = "Status : Idle."
End If
End Sub

Private Sub mnuTools_Click()
If Not SessionInfo.Connected Then
mnuWhoIs.Enabled = False
mnuSendFile.Enabled = False
Else
mnuWhoIs.Enabled = True
mnuSendFile.Enabled = True
End If
End Sub

Private Sub mnuVisitWebpage_Click()
Call OpenURL("P2P Homepage : http://www.Digital-Revolution.org/P2P/Index.html", vbNormalFocus, 0&)
End Sub

Private Sub mnuWhoIs_Click()
frmWhoIs.Show
End Sub

Private Sub sckClient_Close()
On Error Resume Next
Unload frmRecFile
Unload frmSendFile
Unload frmWhoIs

SessionInfo.Connected = False
FrameConvo.Caption = " Conversation (Not Started) "
SessionInfo.SessionType = SESSION_SERVER
sckServer.Close
SessionInfo.FT_InProgress = False
sckRec.Close
sckSend.Close
sckClient.Close
sckServer.Listen
StatusBar.SimpleText = "Status : Connection to User Was Closed / Lost."
Call AddRTFStatus(sckClient.RemoteHostIP & " Has Left the Conversation.", RGB(123, 0, 0))
End Sub

Private Sub sckClient_Connect()
SessionInfo.Connected = True
StatusBar.SimpleText = "Status : Awaiting Authorization . . ."
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
Dim sData As String: sData = Empty
sckClient.GetData sData
Debug.Print sData

    Select Case Mid(sData, 19, 1)
    
    Case "S"
    Call ParseSessionReply(sData)
    Case "M"
    Call ParseMessage(sData)
    Case "F"
    Call ParseTransferRequest(sData)
    Case "R"
    Call ParseTransferReply(sData)
    Case "D"
    Call ParseReadySignal(sData)
    Case "C"
    Call ParseTransferCancel(sData)
    Case "W"
    Call ParseWhoIsRequest(sData)
    Case "X"
    Call ParseWhoIsReply(sData)
    End Select



End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SessionInfo.Connected = False
SessionInfo.SessionType = SESSION_SERVER
SessionInfo.FT_InProgress = False
sckRec.Close
sckSend.Close
sckClient.Close
sckServer.Close
On Error Resume Next
sckServer.Listen
StatusBar.SimpleText = "Status : Unable to Connect to User."
End Sub

Private Sub sckRec_Close()
Close #FileNum
TotalByteNow = 0
frmRecFile.tmrDownload.Enabled = False
SessionInfo.FT_InProgress = False

If KeepLog = True Then
    Call AddLog("File Received " & Chr$(34) & RecFileInfo.FileName & Chr$(34) & " (" & RecFileInfo.FileSize & " Bytes) at " & Now)
End If

If IsCancel = True Then
Call ResetRecForm
frmRecFile.StatusBar.SimpleText = "Status : File Transfer Canceled."
IsCancel = False
Exit Sub
Else

frmRecFile.Bar.Value = 100
frmRecFile.lblKBPS.Caption = "0 KB/Sec"
frmRecFile.lblBR.Caption = "Bytes Received : " & RecFileInfo.FileSize
frmRecFile.StatusBar.SimpleText = "Status : File Transfer Complete."
End If
End Sub

Private Sub sckRec_Connect()
Call MakeRecDir
Open App.Path & "\Received Files\" & RecFileInfo.FileName For Binary Access Write As #FileNum
SessionInfo.FT_InProgress = True
frmRecFile.Bar.Max = RecFileInfo.FileSize
frmRecFile.StatusBar.SimpleText = "Status : Receiving File . . ."
End Sub

Private Sub sckRec_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim FileData As String: FileData = Empty
frmRecFile.tmrDownload.Enabled = True
frmRecFile.Bar.Value = frmRecFile.Bar.Value + bytesTotal
frmRecFile.lblBR.Caption = "Bytes Received : " & bytesTotal
sckRec.GetData FileData
TotalByteNow = TotalByteNow + bytesTotal
Put #FileNum, , FileData

End Sub

Private Sub sckRec_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SessionInfo.FT_InProgress = False
frmRecFile.tmrDownload.Enabled = False
frmRecFile.StatusBar.SimpleText = "Status : Unable to Connect to User."
End Sub

Private Sub sckSend_Close()
sckSend.Close
frmSendFile.tmrUpload.Enabled = False
Call ResetSendForm
frmSendFile.StatusBar.SimpleText = "Status : File Transfer Canceled."

If KeepLog = True Then
    Call AddLog("File Sent " & Chr$(34) & SendFileInfo.FileName & Chr$(34) & " (" & SendFileInfo.FileSize & " Bytes) at " & Now)
End If

On Error Resume Next
sckSend.Listen

End Sub

Private Sub sckSend_ConnectionRequest(ByVal requestID As Long)
SendTotal = 0
sckSend.Close
sckSend.Accept requestID
frmSendFile.StatusBar.SimpleText = "Status : Sending File . . ."
frmSendFile.Bar.Max = SendFileInfo.FileSize
Call SendFile(SendFileInfo.FileSource)

End Sub

Private Sub sckSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SessionInfo.FT_InProgress = False
End Sub

Private Sub sckSend_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
SendTotal = SendTotal + bytesSent
On Error Resume Next
frmSendFile.Bar.Value = frmSendFile.Bar.Value + bytesSent
frmSendFile.lblBS.Caption = "Bytes Sent : " & bytesSent
If SendTotal >= SendFileInfo.FileSize Then
frmSendFile.Bar.Value = 100
frmSendFile.lblKBPS.Caption = "0 KB/Sec"
frmSendFile.lblBS.Caption = "Bytes Sent : " & SendFileInfo.FileSize
frmSendFile.StatusBar.SimpleText = "Status : File Transfer Complete."
sckSend.Close
CurByte = 0
SendTotal = 0
SessionInfo.FT_InProgress = False
End If
End Sub

Private Sub sckServer_Close()
On Error Resume Next
sckServer.Close
Unload frmRecFile
Unload frmSendFile
Unload frmWhoIs
SessionInfo.Connected = False
FrameConvo.Caption = " Conversation (Not Started) "
StatusBar.SimpleText = "Status : User Disconnected."
Call AddRTFStatus(sckServer.RemoteHostIP & " Has Left the Conversation.", RGB(123, 0, 0))
SessionInfo.FT_InProgress = False

If KeepLog = True Then
    Call AddLog(sckServer.RemoteHostIP & " disconnected at " & Now)
End If

If SessionInfo.SessionType = SESSION_SERVER Then
    sckServer.Listen
End If
End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
If sckServer.State <> sckConnected Then
    If SessionInfo.SessionType = SESSION_SERVER Then
        sckServer.Close
        sckServer.Accept requestID
        Dim iRep As Integer
        iRep = MsgBox(sckServer.RemoteHostIP & " is attempting to start a session with you. Accept ?", vbQuestion + vbYesNo, "Session Request")
            Dim sPack As String
            If iRep = vbNo Then
                sPack = sHeader & sDelim & "S" & sDelim & "Denied"
                sckServer.SendData sPack
                SessionInfo.Connected = False
            ElseIf iRep = vbYes Then
                sPack = sHeader & sDelim & "S" & sDelim & "Accepted"
                sckServer.SendData sPack
                Call AddRTFStatus(sckServer.RemoteHostIP & " Joined the Conversation.", RGB(0, 0, 123))
                SessionInfo.Connected = True
                StatusBar.SimpleText = "Status : Session Started."
                FrameConvo.Caption = " Conversation (In Progress)  "
                    If KeepLog = True Then
                        Call AddLog(sckServer.RemoteHostIP & " connected at " & Now)
                    End If
                
            End If
     End If
End If
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
Dim cData As String: cData = Empty
sckServer.GetData cData
Select Case Mid(cData, 19, 1)
    Case "M"
        Call ParseMessage(cData)
    Case "F"
        Call ParseTransferRequest(cData)
    Case "R"
        Call ParseTransferReply(cData)
    Case "D"
        Call ParseReadySignal(cData)
    Case "C"
        Call ParseTransferCancel(cData)
    Case "W"
        Call ParseWhoIsRequest(cData)
    Case "X"
        Call ParseWhoIsReply(cData)
End Select
End Sub

Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtHost_Change()
If Len(Trim$(txtHost.Text)) = 0 Then
    cmdConnect.Enabled = False
Else
    cmdConnect.Enabled = True
End If
End Sub

Private Sub txtNick_Change()
If Len(Trim$(txtNick.Text)) = 0 Then
    cmdSend.Enabled = False
    txtSend.Enabled = False
Else
    cmdSend.Enabled = True
    txtSend.Enabled = True
End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
KeyAscii = 0
cmdSend_Click
End If
End Sub
