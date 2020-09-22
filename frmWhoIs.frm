VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmWhoIs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Pro - WhoIs Query"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5085
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
   Icon            =   "frmWhoIs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   3585
      Width           =   5085
      _ExtentX        =   8969
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Results"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Results"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Query Results "
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4815
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtCompName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time :"
         Height          =   195
         Left            =   1230
         TabIndex        =   8
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         Height          =   195
         Left            =   1245
         TabIndex        =   7
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IP Address :"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Computer Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "If permitted, you canperform a WhoIs query on the remote user."
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
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmWhoIs.frx":0442
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmWhoIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
StatusBar.SimpleText = "Status : Performing Query . . ."
If SessionInfo.SessionType = SESSION_SERVER Then
    frmMain.sckServer.SendData sHeader & sDelim & "W" & sDelim & "WhoIs();"
ElseIf SessionInfo.SessionType = SESSION_CLIENT Then
    frmMain.sckClient.SendData sHeader & sDelim & "W" & sDelim & "WhoIs();"
End If
End Sub
