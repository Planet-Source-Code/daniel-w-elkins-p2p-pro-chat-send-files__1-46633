VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSendFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Pro - Send File"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5160
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
   Icon            =   "frmSendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   3840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   " File Transfer Progress "
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   4695
      Begin ComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblBS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Sent : 0"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblKBPS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB/Sec"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   795
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4290
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status : Idle."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   " File Information "
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   4695
      Begin MSComDlg.CommonDialog CD 
         Left            =   480
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   3960
         TabIndex        =   0
         ToolTipText     =   " Browse "
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblFS 
         Alignment       =   1  'Right Justify
         Caption         =   "File Size (Bytes) :"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label lblFN 
         Alignment       =   1  'Right Justify
         Caption         =   "File Name :"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select a File :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Send files quickly and easily with this utility."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmSendFile.frx":0442
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Reset()

End Sub

Private Sub cmdBrowse_Click()
With CD
.DialogTitle = "Select a File to Send"
.Filter = "All Files (*.*)|*.*"
.Flags = cdlOFNFileMustExist
.ShowOpen
If Len(.FileName) > 0 Then
    If FileLen(.FileName) = 0 Then
        MsgBox "File is empty; please select a valid file", vbCritical, "Invalid File"
        Exit Sub
    End If
SendFileInfo.FileName = .FileTitle
SendFileInfo.FileSource = .FileName
SendFileInfo.FileSize = FileLen(.FileName)
txtPath.Text = SendFileInfo.FileSource
lblFN.Caption = "File Name : " & SendFileInfo.FileName
lblFS.Caption = "File Size (Bytes) : " & SendFileInfo.FileSize
End If
End With
End Sub

Private Sub cmdCancel_Click()
If SessionInfo.FT_InProgress Then
    If SessionInfo.SessionType = SESSION_SERVER Then
        frmMain.sckServer.SendData sHeader & sDelim & "C" & sDelim & "Cancel_Transfer"
        frmMain.sckSend.Close
        SessionInfo.FT_InProgress = False
        Call ResetRecForm
    ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
        frmMain.sckClient.SendData sHeader & sDelim & "C" & sDelim & "Cancel_Transfer"
        frmMain.sckSend.Close
        SessionInfo.FT_InProgress = False
        StatusBar.SimpleText = "Status : File Transfer Canceled."
        Call ResetRecForm
    End If
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If Not SessionInfo.Connected Then
    MsgBox "You are not connected; unable to send transfer request", vbCritical, "Not Connected"
    Exit Sub
End If
If Len(txtPath.Text) = 0 Then
    MsgBox "Please select a file to send", vbCritical, "File Required"
    cmdBrowse_Click
    Exit Sub
End If
Dim sPack As String
SendTotal = 0
TotalBR = 0
Bar.Value = 0
lblKBPS.Caption = "0 KB/Sec"
lblBS.Caption = "Bytes Sent : 0"
sPack = sHeader & sDelim & "F" & sDelim & SendFileInfo.FileName & sDelim & SendFileInfo.FileSize
If SessionInfo.SessionType = SESSION_SERVER Then
    frmMain.sckServer.SendData sPack
ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
    frmMain.sckClient.SendData sPack
End If
sPack = Empty
StatusBar.SimpleText = "Status : Awaiting User's Permission . . ."
End Sub

Private Sub tmrUpload_Timer()
UploadSpeed = SendTotal - UploadSecond
UploadSecond = SendTotal
KBPS = ((UploadSpeed / 1024) * 2)
lblKBPS.Caption = KBPS & " KB/Sec"
End Sub
