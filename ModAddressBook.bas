Attribute VB_Name = "ModAddressBook"
Option Explicit

Private Const ABFile As String = "\Adbk.adb"
Private Const FileDelim As String = "<&>"
Private Const P2P_LINESEP As String = "µ2ºž‰p p"

Public Sub KillAB()
On Error Resume Next
Kill App.Path & ABFile
End Sub

Public Sub AddToAB(sName As String, sIP As String, sDesc As String)
Dim FF As Integer: FF = FreeFile
Dim sTmp As String, sFileData As String
Open App.Path & ABFile For Binary Access Read As #FF
If LOF(FF) = 0 Then
sFileData = Empty
Else
sFileData = Input(LOF(FF), FF)
End If
Close #FF
sTmp = EncryptHex(sName & FileDelim & sIP & FileDelim & sDesc & P2P_LINESEP, True)
Do
DoEvents
Loop Until Len(sTmp) > 0
Call KillAB
Open App.Path & ABFile For Binary As #1: Close #1
Open App.Path & ABFile For Binary Access Write As #FF
If Len(sFileData) = 0 Then
Put #FF, , sTmp
Else
Put #FF, , sFileData & sTmp
End If
Close #FF
sTmp = Empty: sFileData = Empty
End Sub

Public Function ContactExists(sContactIP As String) As Boolean
On Error Resume Next
Dim FF As Integer: FF = FreeFile
Dim sTmp As String, sDec As String, sBuff() As String, sDat() As String
Dim A As Long
Open App.Path & ABFile For Binary Access Read As #FF

If LOF(FF) = 0 Then
    Close #FF
    ContactExists = False
    Exit Function
End If

sTmp = Input(LOF(FF), FF)
Close #FF
sDec = EncryptHex(sTmp, False)
Do
DoEvents
Loop Until Len(sDec) > 0
sBuff() = Split(sDec, P2P_LINESEP)
For A = 0 To UBound(sBuff)
    sDat() = Split(sBuff(A), FileDelim)
    If Len(sBuff(A)) > 0 Then
        If UCase$(sDat(1)) = UCase$(sContactIP) Then
            ContactExists = True
            Exit For
         End If
    End If
DoEvents
Next A
sTmp = Empty: sDec = Empty: sBuff() = Split(""): sDat() = Split("")
End Function

Public Sub ReadAB(objLV As ListView)
Dim FF As Integer: FF = FreeFile
Dim sTmp As String, sDec As String, sBuff() As String, sDat() As String
Dim A As Long
With objLV

.ListItems.Clear
Open App.Path & ABFile For Binary Access Read As #FF

If LOF(FF) = 0 Then
    Close #FF
    Exit Sub
End If

sTmp = Input(LOF(FF), FF)
Close #FF
sDec = EncryptHex(sTmp, False)
Do
DoEvents
Loop Until Len(sDec) > 0
sBuff() = Split(sDec, P2P_LINESEP)
    For A = 0 To UBound(sBuff)
        If Len(sBuff(A)) > 0 Then
            sDat() = Split(sBuff(A), FileDelim)
            .ListItems.Add , , sDat(0), , "Book"
            .ListItems(.ListItems.Count).SubItems(1) = sDat(1)
            .ListItems(.ListItems.Count).SubItems(2) = sDat(2)
        End If
    DoEvents
    Next A
End With
sTmp = Empty: sDec = Empty: sBuff() = Split(""): sDat() = Split("")
End Sub

Public Sub RemoveAB(sContactIP As String)
Dim FF As Integer: FF = FreeFile
Dim sTmp As String, sDec As String, sBuff() As String, sDat() As String
Dim sFileData As String
Dim A As Long

Open App.Path & ABFile For Binary Access Read As #FF

If LOF(FF) = 0 Then
    Close #FF
    Exit Sub
End If

sTmp = Input(LOF(FF), FF)
Close #FF
sDec = EncryptHex(sTmp, False)

Do
    DoEvents
Loop Until Len(sDec) > 0

sBuff() = Split(sDec, P2P_LINESEP)
    For A = 0 To UBound(sBuff)
        If Len(sBuff(A)) > 0 Then
            sDat() = Split(sBuff(A), FileDelim)
                If Not UCase$(sDat(1)) = UCase$(sContactIP) Then
                    sFileData = sFileData & EncryptHex(sBuff(A) & P2P_LINESEP, True)
                End If
        End If
    DoEvents
    Next A
Call KillAB
Open App.Path & ABFile For Binary As #1: Close #1
Open App.Path & ABFile For Binary Access Write As #FF
Put #FF, , sFileData
Close #FF
End Sub
