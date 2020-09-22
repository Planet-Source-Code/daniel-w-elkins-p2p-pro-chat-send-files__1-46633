VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRecFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Pro - Receive File"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5805
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
   Icon            =   "frmRecFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDownload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   3360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3810
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status : Idle."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   " File Transfer Progress "
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5535
      Begin ComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblKBPS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB/Sec"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblBR 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Received : 0"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " File Information "
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      Begin VB.Label lblFN 
         Caption         =   "File Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblFS 
         Caption         =   "File Size (Bytes) :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please do not close this window until the file tranfer is complete."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmRecFile.frx":0ECA
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmRecFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
If SessionInfo.FT_InProgress Then
    If SessionInfo.SessionType = SESSION_SERVER Then
        frmMain.sckServer.SendData sHeader & sDelim & "C" & sDelim & "Cancel_Transfer"
        frmMain.sckRec.Close
        SessionInfo.FT_InProgress = False
        Call ResetRecForm
    ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
        frmMain.sckClient.SendData sHeader & sDelim & "C" & sDelim & "Cancel_Transfer"
        frmMain.sckRec.Close
        SessionInfo.FT_InProgress = False
        StatusBar.SimpleText = "Status : File Transfer Canceled."
        Call ResetRecForm
    End If
End If
End Sub

Private Sub tmrDownload_Timer()
DownloadSpeed = TotalByteNow - DownloadSecond
DownloadSecond = TotalByteNow
RecKBS = ((DownloadSpeed / 1024) * 2)
lblKBPS.Caption = RecKBS & " KB/Sec"
End Sub
