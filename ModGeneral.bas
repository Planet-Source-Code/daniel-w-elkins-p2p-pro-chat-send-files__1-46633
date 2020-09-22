Attribute VB_Name = "ModGeneral"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const sHeader As String = "¬Yp 2p  "
Public Const sDelim As String = "Ã\µ‰ºž"

Public Const DEF_WIDTH As Integer = 5445
Public Const DEF_HEIGHT As Integer = 6225
Public Const SESSION_CLIENT As Integer = 1
Public Const SESSION_SERVER As Integer = 2

Public Const FILETYPE_RTF As Integer = 1
Public Const FILETYPE_TXT As Integer = 2
Public Const FILETYPE_OTHER As Integer = 3

Public Type SessionInfo1
RemoteHost As String
Nick As String
Connected As Boolean
FT_InProgress As Boolean
SessionType As Integer
End Type

Global SessionInfo As SessionInfo1

Public Sub AddRTFStatus(sStatus As String, lColor As Long)
With frmMain.txtChat
If Len(.Text) = 0 Then

.SelBold = True
.SelItalic = False
.SelUnderline = False
.SelColor = RGB(123, 0, 0)
.SelText = "- "
.SelColor = RGB(0, 123, 0)
.SelText = "Status"
.SelBold = False
.SelColor = vbBlack
.SelText = ": "
.SelColor = lColor
.SelText = sStatus

Else

.SelBold = True
.SelItalic = False
.SelUnderline = False
.SelColor = RGB(123, 0, 0)
.SelText = vbNewLine & "- "
.SelColor = RGB(0, 123, 0)
.SelText = "Status"
.SelBold = False
.SelColor = vbBlack
.SelText = ": "
.SelColor = lColor
.SelText = sStatus

End If
End With

End Sub

Public Function GetFileType(sFilePath As String) As Integer
Dim sBuff() As String: sBuff() = Split(sFilePath, ".")
Dim sTmp As String
sTmp = UCase$(sBuff(UBound(sBuff)))
Select Case sTmp
    Case "RTF"
    GetFileType = FILETYPE_RTF
    Case "TXT"
    GetFileType = FILETYPE_TXT
    Case Else
    GetFileType = FILETYPE_OTHER
End Select
End Function

Public Sub OpenURL(strURL As String, iWindowStyle As Integer, fH As Long)
Dim lSuccess As Long
lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)
End Sub

