VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protect 2004"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Unlock"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "UNREGISTERED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Protect 2004 is registered to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   6000
      X2              =   0
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "By Adam Ranshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.adranix.co.uk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   6000
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   6000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "This computer is protected with Protect 2004"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
password.Visible = True
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "1"
op.text1.LoadFile ("c:\windows\system32\regnamepro04.rtf")
op.text2.LoadFile ("c:\windows\system32\regcodepro04.rtf")
If op.text2.Text = "0040-0110" Then
Label6.Caption = op.text1.Text
op.text2.Enabled = False
op.text2.Locked = True
op.text2.Visible = False
op.Command5.Enabled = False
op.Command6.Enabled = True
op.Label4.Visible = False
Else
op.Check1.Enabled = False
op.Check3.Enabled = False
op.Check5.Enabled = False
op.Check7.Enabled = False
op.Check9.Enabled = False
op.Check11.Enabled = False
op.Check13.Enabled = False
op.Check15.Enabled = False
op.Check17.Enabled = False
op.Check19.Enabled = False
op.Check24.Enabled = False
op.Check27.Enabled = False
op.Check25.Enabled = False
op.Check21.Enabled = False
op.text1.Locked = False
password.demo.Visible = True
password.exit.Enabled = True
password.regnow.Enabled = True
MsgBox "The free Demo version of Protect 2004 is being used, please upgrade to the full version which will allow you to use all the options and remove this message for good.  To buy now go to http://www.adranix.co.uk", vbExclamation
End If
End Sub







Private Sub Timer1_Timer()
Form3.Visible = False
End Sub


Private Sub Timer2_Timer()
End
End Sub
