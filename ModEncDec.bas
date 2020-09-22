Attribute VB_Name = "ModEncDec"

Public Function EncryptHex(sString As String, bEncrypt As Boolean) As String
Dim s As String
Dim sTemp As String
Dim i As Long
Dim sHex As String
Dim sNewHex As String
Dim iDec As Long
Dim sTmp As String
For i = 1 To Len(sString)
sHex = Hex$(Asc(Mid(sString, i, 1)))
If Len(sHex) = 1 Then
sHex = "0" & sHex
End If
sTemp = sTemp & sHex
DoEvents
Next
If bEncrypt Then
sTemp = Right$(sTemp, 1) & Left$(sTemp, Len(sTemp) - 1)
Else
sTemp = Mid$(sTemp, 2, Len(sTemp)) & Left$(sTemp, 1)
End If
For i = 1 To Len(sTemp) Step 2
sNewHex = Mid$(sTemp, i, 2)
iDec = Val("&H" & sNewHex)
If iDec > 0 Then
sTmp = sTmp & Chr(iDec)
If Len(sTmp) > 50 Then
s = s & sTmp
sTmp = ""
End If
End If
DoEvents
Next
s = s & sTmp
EncryptHex = s
End Function



