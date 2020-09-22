Attribute VB_Name = "ModLog"
Option Explicit

Public Sub KillLog()
On Error Resume Next
Kill App.Path & "\Event Log.txt"
End Sub

Public Sub AddLog(sLog As String)
Dim FF As Integer: FF = FreeFile
Open App.Path & "\Event Log.txt" For Append As #FF
Print #FF, sLog
Close #FF
End Sub
