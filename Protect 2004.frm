VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADRANIX Security"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "Protect 2004.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Error Loading Item Name..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop - Access Denied"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Protect 2004.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = False
password.Label1.Caption = "Please enter your password to use one of the options to the left of click 'Cancel'"
password.Visible = True
End Sub

Private Sub Command2_Click()
Form2.Enabled = False
password.Visible = False
password.Visible = True
End Sub

Private Sub Command3_Click()
list.Visible = True
End Sub



