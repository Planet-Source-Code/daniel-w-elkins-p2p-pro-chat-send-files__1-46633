Attribute VB_Name = "ModWhoIs"
Option Explicit

Public Sub ParseWhoIsRequest(sData As String)
Dim sInfo As String
Dim sDenied As String
sInfo = sHeader & sDelim & "X" & sDelim & WhoIsInfo
sDenied = sHeader & sDelim & "X" & sDelim & EncryptHex("WhoIs();->Denied", True)
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
If sBuff(2) = "WhoIs();" Then
    If WhoIs = True Then
        If SessionInfo.SessionType = SESSION_SERVER Then
            frmMain.sckServer.SendData sInfo
        ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
            frmMain.sckClient.SendData sInfo
        End If
        
        If KeepLog = True Then
            Call AddLog("WhoIs information sent at " & Now)
        End If
        
    ElseIf WhoIs = False Then
        If SessionInfo.SessionType = SESSION_SERVER Then
            frmMain.sckServer.SendData sDenied
        ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
            frmMain.sckClient.SendData sDenied
        End If
        
        If KeepLog = True Then
            Call AddLog("You denied a WhoIs query at " & Now)
        End If

    End If
End If
sInfo = Empty: sDenied = Empty
End Sub

Public Sub ParseWhoIsReply(sData As String)
On Error Resume Next
Dim sBuff() As String, sDec As String
sDec = EncryptHex(Mid(sData, 27), False)
sBuff() = Split(sDec, sDelim)
If sBuff(0) = "WhoIs();->Denied" Then
    frmWhoIs.StatusBar.SimpleText = "Status : WhoIs Not Permitted."
Else
    With frmWhoIs
    .StatusBar.SimpleText = "Status : Query Finished."
    .txtCompName = sBuff(0)
    .txtIP.Text = sBuff(1)
    .txtDate.Text = sBuff(2)
    .txtTime.Text = sBuff(3)
    End With
End If
End Sub

Public Function WhoIsInfo() As String
Dim sTmp As String, sEnc As String
sTmp = frmMain.sckClient.LocalHostName & sDelim
sTmp = sTmp & frmMain.sckClient.LocalIP & sDelim
sTmp = sTmp & Date & sDelim & Time
sEnc = EncryptHex(sTmp, True)
WhoIsInfo = sEnc
sTmp = Empty: sEnc = Empty
End Function
