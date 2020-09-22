VERSION 5.00
Begin VB.Form frmAddContact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "P2P Pro - Add Contact"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4920
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Contact Information "
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtDesc 
         Height          =   495
         Left            =   240
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hostname / IP Address :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name (ID) :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
If Len(txtDesc.Text) = 0 Then txtDesc.Text = "< None >"

If Len(txtName.Text) = 0 Then
    MsgBox "Enter a name (ID) for the contact", vbCritical, "Name (ID) Required"
    txtName.SetFocus
    Exit Sub
ElseIf Len(Trim$(txtHost.Text)) = 0 Then
    MsgBox "Enter the contact's hostname or IP address", vbCritical, "Hostname (IP) Required"
    txtHost.SetFocus
    Exit Sub
ElseIf ContactExists(txtHost.Text) Then
    MsgBox "Contact already exists in address book (" & txtHost.Text & ")", vbCritical, "Contact Exists"
    txtHost.SetFocus: txtHost.SelStart = 0: txtHost.SelLength = Len(txtHost.Text)
    Exit Sub
End If
Call AddToAB(txtName.Text, txtHost.Text, txtDesc.Text)
Call ReadAB(frmAddressBook.LVAB)
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

