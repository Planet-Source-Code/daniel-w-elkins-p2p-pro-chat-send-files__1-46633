Attribute VB_Name = "ModPreferences"
Option Explicit

Public Const sApp As String = "P2P"

Global StartUp As Boolean
Global AutoAccept As Boolean
Global KeepLog As Boolean
Global WhoIs As Boolean

Global MaxBS As Long

Public Sub SavePref()
SaveSetting sApp, "Pref", "StartUp", StartUp
SaveSetting sApp, "Pref", "AutoAccept", AutoAccept
SaveSetting sApp, "Pref", "KeepLog", KeepLog
SaveSetting sApp, "Pref", "WhoIs", WhoIs
SaveSetting sApp, "Pref", "MaxBS", MaxBS
End Sub

Public Sub LoadPrefDefaults()
StartUp = False
Call RemoveRegRun
AutoAccept = False
KeepLog = False
WhoIs = False
MaxBS = 4096
Call SavePref
End Sub

Public Sub ReadPref(bApply As Boolean)
On Error Resume Next
StartUp = GetSetting(sApp, "Pref", "StartUp", False)
If StartUp Then
    Call RegRun
ElseIf Not StartUp Then
    Call RemoveRegRun
End If
AutoAccept = GetSetting(sApp, "Pref", "AutoAccept", False)
KeepLog = GetSetting(sApp, "Pref", "KeepLog", False)
WhoIs = GetSetting(sApp, "Pref", "WhoIs", False)
MaxBS = GetSetting(sApp, "Pref", "MaxBS", 4096)
If KeepLog = False Then Call KillLog
If bApply Then
With frmPreferences
.chkStartUp.Value = Abs(CInt(StartUp))
.chkAutoAccept.Value = Abs(CInt(AutoAccept))
.chkLog.Value = Abs(CInt(KeepLog))
.chkWhoIs.Value = Abs(CInt(WhoIs))
.txtMaxBS = MaxBS
End With
End If
End Sub
