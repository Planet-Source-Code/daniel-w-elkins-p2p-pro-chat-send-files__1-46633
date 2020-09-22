VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmPreferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Pro - Preferences"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
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
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Load Defaults"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4815
      _cx             =   8493
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   " Advanced Settings "
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4815
      Begin VB.CheckBox chkWhoIs 
         Caption         =   "Permit WhoIs queries."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtMaxBS 
         Height          =   285
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "4096"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Max File-Send Speed (Bytes/Sec) :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " General Settings "
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   4815
      Begin VB.CheckBox chkLog 
         Caption         =   "Keep an event log."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoAccept 
         Caption         =   "Automatically accept file transfers."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chkStartUp 
         Caption         =   "Start P2P Pro when Windows starts."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
If Len(txtMaxBS.Text) = 0 Then
MsgBox "Enter a max file-send speed", vbCritical, "Input Error"
txtMaxBS.SetFocus
Exit Sub
ElseIf Val(txtMaxBS.Text) < 1024 Then
MsgBox "Max file-send speed must be at least 1024 bytes", vbCritical, "Input Error"
txtMaxBS.SetFocus: txtMaxBS.SelStart = 0: txtMaxBS.SelLength = Len(txtMaxBS.Text)
Exit Sub
End If
StartUp = chkStartUp.Value
AutoAccept = chkAutoAccept.Value
KeepLog = chkLog.Value
WhoIs = chkWhoIs.Value
MaxBS = Val(txtMaxBS.Text)
Call SavePref
Call ReadPref(False)
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDefaults_Click()
Dim iRep As Integer
iRep = MsgBox("Are you sure you want to restore the default settings ?", vbQuestion + vbYesNo, "Restore Defaults")
If iRep = vbYes Then
Call LoadPrefDefaults
Call ReadPref(True)
End If
End Sub

Private Sub Form_Load()
Flash.Movie = App.Path & "\01.swf"
Call ReadPref(True)
End Sub

Private Sub txtMaxBS_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
