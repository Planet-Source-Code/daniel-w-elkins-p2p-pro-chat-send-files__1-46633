Attribute VB_Name = "ModSession"
Option Explicit

Public Sub ParseSessionReply(sData As String)
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
If sBuff(2) = "Accepted" Then
frmMain.StatusBar.SimpleText = "Status : Session Started."
frmMain.FrameConvo.Caption = " Conversation (In Progress)  "

    If KeepLog = True Then
        Call AddLog("Session with " & frmMain.sckClient.RemoteHostIP & " started at " & Now)
    End If

ElseIf sBuff(2) = "Denied" Then
frmMain.StatusBar.SimpleText = "Status : User Denied Session Request."
frmMain.FrameConvo.Caption = " Conversation (Not Started) "

    If KeepLog = True Then
        Call AddLog("Session was denied from " & frmMain.sckClient.RemoteHostIP & " at " & Now)
    End If

End If
End Sub

Public Sub ParseMessage(sData As String)
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
With frmMain.txtChat
If Len(.Text) = 0 Then
    .SelBold = True
    .SelItalic = False
    .SelUnderline = False
    .SelColor = RGB(123, 0, 0)
    .SelText = sBuff(2)
    .SelBold = False
    .SelColor = vbBlack
    .SelText = ": "
    .SelColor = RGB(0, 0, 123)
    .SelText = sBuff(3)
Else
    .SelBold = True
    .SelItalic = False
    .SelUnderline = False
    .SelColor = RGB(123, 0, 0)
    .SelText = vbNewLine & sBuff(2)
    .SelBold = False
    .SelColor = vbBlack
    .SelText = ": "
    .SelColor = RGB(0, 0, 123)
    .SelText = sBuff(3)
End If
End With
End Sub
