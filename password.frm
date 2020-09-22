VERSION 5.00
Begin VB.Form password 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   381.087
   ScaleMode       =   0  'User
   ScaleWidth      =   14337.7
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Message Window:"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your password to use one of the options to the left of click ""Cancel"""
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Whats Locked?"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Options"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   390
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000FFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   25
      TabIndex        =   1
      Top             =   165
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   960
   End
   Begin VB.Menu main 
      Caption         =   "Main"
      Begin VB.Menu demo 
         Caption         =   "Demo Version"
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu exit 
         Caption         =   "Exit Wiithout Password"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu wl 
         Caption         =   "Whats Locked?"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu reg 
      Caption         =   "Registration"
      Begin VB.Menu regnow 
         Caption         =   "Enter Registration Info"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
txtPassword.Text = ""
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    Form2.Enabled = True
End Sub

Private Sub cmdOK_Click()
Form2.Enabled = False
    'check for correct password
    If txtPassword = "0101" Then
    CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "0"
    End
    Else
Label1.Caption = "The password you have entered is not correct, please re-try."
txtPassword.Text = ""
txtPassword.SetFocus
SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command1_Click()
Form2.Enabled = False
    'check for correct password
    If txtPassword = "0101" Then
    Form3.Timer2.Enabled = False
    txtPassword.Text = ""
op.Visible = True
    Else
       Label1.Caption = "The password you have entered is not correct, please re-try."
        txtPassword.Text = ""
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command2_Click()
password.Enabled = False
list.Visible = True
End Sub

Private Sub demo_Click()
MsgBox "The free Demo version of Protect 2004 is being used.  To get Protect 2004 for your PC just download the free Demo version from http://www.adranix.co.uk and click the download link.", vbExclamation
End Sub

Private Sub exit_Click()
MsgBox "As the free version of Protect 2004 is being used you can remove restrictions and exit without having to enter the password, please register Protect 2004 at http://www.adranix.co.uk to disable exiting without a password.", vbExclamation
CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "0"
End
End Sub

Private Sub regnow_Click()
Form4.Visible = False
Form4.Visible = True
End Sub

Private Sub wl_Click()
list.Visible = True
End Sub

