VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAddressBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Pro - Address Book"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
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
   Icon            =   "frmAddressBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect ->"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   " Added Contacts "
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4935
      Begin ComctlLib.ListView LVAB 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "IL"
         SmallIcons      =   "IL"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Book"
            Object.Tag             =   "Book"
            Text            =   "Name"
            Object.Width           =   2778
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "IP Address"
            Object.Width           =   2593
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   5371
         EndProperty
      End
      Begin ComctlLib.ImageList IL 
         Left            =   2280
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAddressBook.frx":0ECA
               Key             =   "Book"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "P2P Pro remembers your contacts so you don't have to!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   5415
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   120
      Picture         =   "frmAddressBook.frx":1B1C
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
frmAddContact.Show
End Sub

Private Sub cmdClear_Click()
Dim iRep As Integer
iRep = MsgBox("Are you sure you want to remove all of your contacts ?", vbQuestion + vbYesNo, "Clear Contacts")
If iRep = vbYes Then
    LVAB.ListItems.Clear
    Call KillAB
End If
End Sub

Private Sub cmdConnect_Click()
On Error Resume Next
Dim lTmp As Long, sTmp As String
If LVAB.ListItems.Count = 0 Then Exit Sub
lTmp = LVAB.SelectedItem.Index
If lTmp = 0 Then
    MsgBox "Select a contact to connect to", vbCritical, "Contact Required"
Else
    sTmp = LVAB.ListItems(lTmp).SubItems(1)
    frmMain.txtHost.Text = sTmp
    Unload Me
    On Error Resume Next
    frmMain.txtHost.SetFocus
End If
lTmp = 0: sTmp = Empty
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
Dim lTmp As Long, sTmp As String
If LVAB.ListItems.Count = 0 Then Exit Sub
lTmp = LVAB.SelectedItem.Index
If lTmp = 0 Then
    MsgBox "Select a contact to remove", vbCritical, "Contact Required"
Else
    sTmp = LVAB.ListItems(lTmp).SubItems(1)
    Call RemoveAB(sTmp)
    Call ReadAB(LVAB)
    If LVAB.ListItems.Count = 0 Then Call KillAB
End If
lTmp = 0: sTmp = Empty
End Sub

Private Sub Form_Load()
Call ReadAB(LVAB)
End Sub
