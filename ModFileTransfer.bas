Attribute VB_Name = "ModFileTransfer"
Option Explicit

Public Type SendFileInfo1
FileName As String
FileSource As String
FileSize As Long
End Type

Public Type RecFileInfo1
FileName As String
FileSize As Long
End Type

Global SendFileInfo As SendFileInfo1

Global RecFileInfo As RecFileInfo1

Global CurByte As Long
Global SendTotal As Long
Global TotalBR As Long
Global UploadSpeed As Long
Global UploadSecond As Long
Global TotalSent As Long
Global KBPS As Long

Global DownloadSpeed As Long
Global DownloadSecond As Long
Global TotalByteNow As Long
Global RecKBS As Long

Global FileNum As Integer

Global IsCancel As Boolean

Public Sub ResetFTInfo()
CurByte = 0
SendTotal = 0
TotalBR = 0
UploadSpeed = 0
UploadSecond = 0
TotalSent = 0
KBPS = 0
DownloadSpeed = 0
DownloadSecond = 0
TotalByteNow = 0
RecKBS = 0
End Sub

Public Sub ParseTransferRequest(sData As String)
IsCancel = False
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
Dim iRep As Integer

If SessionInfo.SessionType = SESSION_SERVER Then

If AutoAccept And SessionInfo.SessionType = SESSION_SERVER Then
frmMain.sckServer.SendData sHeader & sDelim & "R" & sDelim & "Accepted"
RecFileInfo.FileName = sBuff(2)
RecFileInfo.FileSize = sBuff(3)
Call ResetRecForm
frmRecFile.Show
frmRecFile.lblFN.Caption = "File Name : " & sBuff(2)
frmRecFile.lblFS.Caption = "File Size (Bytes) : " & sBuff(3)
frmRecFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
Exit Sub

ElseIf AutoAccept And SessionInfo.SessionType = SESSION_CLIENT Then
frmMain.sckClient.SendData sHeader & sDelim & "R" & sDelim & "Accepted"
RecFileInfo.FileName = sBuff(2)
RecFileInfo.FileSize = sBuff(3)
Call ResetRecForm
frmRecFile.Show
frmRecFile.lblFN.Caption = "File Name : " & sBuff(2)
frmRecFile.lblFS.Caption = "File Size (Bytes) : " & sBuff(3)
frmRecFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
Exit Sub
End If


    iRep = MsgBox(frmMain.sckServer.RemoteHostIP & " would like to send you the file " & Chr$(34) & sBuff(2) & Chr$(34) & " (" & sBuff(3) & " Bytes). Accept ?", vbQuestion + vbYesNo, "P2P Pro - File Transfer Request")
        If iRep = vbNo Then
            frmMain.sckServer.SendData sHeader & sDelim & "R" & sDelim & "Denied"
        ElseIf iRep = vbYes Then
            frmMain.sckServer.SendData sHeader & sDelim & "R" & sDelim & "Accepted"
            RecFileInfo.FileName = sBuff(2)
            RecFileInfo.FileSize = sBuff(3)
            Call ResetRecForm
            frmRecFile.Show
            frmRecFile.lblFN.Caption = "File Name : " & sBuff(2)
            frmRecFile.lblFS.Caption = "File Size (Bytes) : " & sBuff(3)
            frmRecFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
        End If

ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
    iRep = MsgBox(frmMain.sckClient.RemoteHostIP & " would like to send you the file " & Chr$(34) & sBuff(2) & Chr$(34) & " (" & sBuff(3) & " Bytes). Accept ?", vbQuestion + vbYesNo, "P2P Pro - File Transfer Request")
        If iRep = vbNo Then
            frmMain.sckClient.SendData sHeader & sDelim & "R" & sDelim & "Denied"
        ElseIf iRep = vbYes Then
            frmMain.sckClient.SendData sHeader & sDelim & "R" & sDelim & "Accepted"
            TotalByteNow = 0
            RecFileInfo.FileName = sBuff(2)
            RecFileInfo.FileSize = sBuff(3)
            Call ResetRecForm
            frmRecFile.Show
            frmRecFile.lblFN.Caption = "File Name : " & sBuff(2)
            frmRecFile.lblFS.Caption = "File Size (Bytes) : " & sBuff(3)
            frmRecFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
        End If


End If
End Sub
 
Public Sub ParseTransferReply(sData As String)
On Error Resume Next
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
IsCancel = False
If SessionInfo.SessionType = SESSION_SERVER Then

    If sBuff(2) = "Denied" Then
        frmSendFile.StatusBar.SimpleText = "Status : User Denied Transfer Request."
        SessionInfo.FT_InProgress = False
    ElseIf sBuff(2) = "Accepted" Then
        frmSendFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
        SessionInfo.FT_InProgress = True
        frmMain.sckSend.Close
        frmMain.sckSend.Listen
        DoEvents
        frmMain.sckServer.SendData sHeader & sDelim & "D" & sDelim & "Ready"
    End If
    
ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
    If sBuff(2) = "Denied" Then
        frmSendFile.StatusBar.SimpleText = "Status : User Denied Transfer Request."
        SessionInfo.FT_InProgress = False
    ElseIf sBuff(2) = "Accepted" Then
        frmSendFile.StatusBar.SimpleText = "Status : Negotiating Transfer . . ."
        SessionInfo.FT_InProgress = True
        frmMain.sckSend.Close
        frmMain.sckSend.Listen
        DoEvents
        frmMain.sckClient.SendData sHeader & sDelim & "D" & sDelim & "Ready"
    End If

End If
End Sub

Public Sub ResetRecForm()
With frmRecFile
.Bar.Value = 0
.lblKBPS.Caption = "0 KB/Sec"
.lblBR.Caption = "Bytes Received : 0"
.lblFN.Caption = "File Name :"
.lblFS.Caption = "File Size (Bytes) :"
End With
End Sub

Public Sub ResetSendForm()
With frmSendFile
.Bar.Value = 0
.lblKBPS.Caption = "0 KB/Sec"
.lblBS.Caption = "Bytes Sent : 0"
.lblFN.Caption = "File Name :"
.lblFS.Caption = "File Size (Bytes) :"
End With
End Sub

Public Sub SendFile(sFilePath As String)
Dim FF As Integer: FF = FreeFile
Dim B As Long
Dim bBuffer() As Byte
frmSendFile.tmrUpload.Enabled = True
Open sFilePath For Binary Access Read As #FF
ReDim bBuffer(1 To MaxBS) As Byte
DoEvents
Do Until (SendFileInfo.FileSize - CurByte) < MaxBS
DoEvents
Get #FF, CurByte + 1, bBuffer()
CurByte = CurByte + MaxBS
On Error GoTo Err
frmMain.sckSend.SendData bBuffer
DoEvents
Loop
Dim PrevPackSize As Long
PrevPackSize = SendFileInfo.FileSize - CurByte
ReDim bBuffer(1 To PrevPackSize) As Byte
Get #FF, CurByte + 1, bBuffer()
CurByte = CurByte + PrevPackSize
frmMain.sckSend.SendData bBuffer
Close #FF
frmSendFile.tmrUpload.Enabled = False
Exit Sub
Err:
DoEvents
Debug.Print "Send Error : " & Err.Description
Exit Sub
End Sub

Public Sub ParseReadySignal(sData As String)
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
If sBuff(2) = "Ready" Then
    If SessionInfo.SessionType = SESSION_SERVER Then
        frmMain.sckRec.Close
        frmMain.sckRec.RemoteHost = frmMain.sckServer.RemoteHostIP
        frmMain.sckRec.RemotePort = 3095
        frmMain.sckRec.Connect
        frmRecFile.StatusBar.SimpleText = "Status : Connecting to User . . ."
    ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
        frmMain.sckRec.Close
        frmMain.sckRec.RemoteHost = frmMain.sckClient.RemoteHostIP
        frmMain.sckRec.RemotePort = 3095
        frmMain.sckRec.Connect
        frmRecFile.StatusBar.SimpleText = "Status : Connecting to User . . ."
    End If
End If
End Sub

Public Sub MakeRecDir()
On Error Resume Next
MkDir App.Path & "\Received Files"
End Sub

Public Sub ParseTransferCancel(sData As String)
Dim sBuff() As String: sBuff() = Split(sData, sDelim)
On Error Resume Next
If sBuff(2) = "Cancel_Transfer" Then
Call ResetFTInfo
IsCancel = True
Close #FileNum
frmSendFile.StatusBar.SimpleText = "Status : File Transfer Canceled."
frmRecFile.StatusBar.SimpleText = "Status : File Transfer Canceled."
frmMain.sckRec.Close
frmMain.sckSend.Close
frmMain.sckSend.Listen
SessionInfo.FT_InProgress = False
Call ResetRecForm
Call ResetSendForm
End If
End Sub
