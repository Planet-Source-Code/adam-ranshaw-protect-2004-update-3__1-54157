VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form op 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protect 2004 Options"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   8760
      Top             =   120
   End
   Begin RichTextLib.RichTextBox ac 
      Height          =   375
      Left            =   120
      TabIndex        =   69
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"op.frx":0000
   End
   Begin RichTextLib.RichTextBox text2 
      Height          =   255
      Left            =   7800
      TabIndex        =   67
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   65535
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   15
      TextRTF         =   $"op.frx":0080
   End
   Begin VB.CheckBox Check22 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start Menu Settings"
      Height          =   255
      Left            =   3720
      TabIndex        =   60
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CheckBox Check27 
      BackColor       =   &H0000FFFF&
      Caption         =   "Network Connections"
      Height          =   255
      Left            =   3720
      TabIndex        =   59
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   58
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox Check26 
      BackColor       =   &H0000FFFF&
      Caption         =   "Downloads"
      Height          =   255
      Left            =   3720
      TabIndex        =   57
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9360
      Top             =   120
   End
   Begin RichTextLib.RichTextBox b1 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0102
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hide"
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   27
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox Check25 
      BackColor       =   &H0000FFFF&
      Caption         =   "Registry Editing"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check23 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command Prompt"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox Check21 
      BackColor       =   &H0000FFFF&
      Caption         =   "Network Setup Wizard"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox Check18 
      BackColor       =   &H0000FFFF&
      Caption         =   "Windows XP Tour"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sound Recorder"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H0000FFFF&
      Caption         =   "Narrator"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H0000FFFF&
      Caption         =   "Windows Messenger"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H0000FFFF&
      Caption         =   "Media Player"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H0000FFFF&
      Caption         =   "Search"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H0000FFFF&
      Caption         =   "Display Properties"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H0000FFFF&
      Caption         =   "User Accounts"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H0000FFFF&
      Caption         =   "System Tools"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H0000FFFF&
      Caption         =   "Bulit in Games"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H0000FFFF&
      Caption         =   "My Music"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H0000FFFF&
      Caption         =   "My Computer"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Printers"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shutdown"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Help"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Control Panel"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "op.frx":0184
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin RichTextLib.RichTextBox b2 
      Height          =   375
      Left            =   1200
      TabIndex        =   29
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":05C6
   End
   Begin RichTextLib.RichTextBox b3 
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0648
   End
   Begin RichTextLib.RichTextBox b4 
      Height          =   375
      Left            =   3360
      TabIndex        =   31
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":06CA
   End
   Begin RichTextLib.RichTextBox b6 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":074C
   End
   Begin RichTextLib.RichTextBox b5 
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":07CE
   End
   Begin RichTextLib.RichTextBox b7 
      Height          =   375
      Left            =   1200
      TabIndex        =   34
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0850
   End
   Begin RichTextLib.RichTextBox b8 
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":08D2
   End
   Begin RichTextLib.RichTextBox b9 
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0954
   End
   Begin RichTextLib.RichTextBox b10 
      Height          =   375
      Left            =   4440
      TabIndex        =   37
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":09D6
   End
   Begin RichTextLib.RichTextBox b11 
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0A58
   End
   Begin RichTextLib.RichTextBox b12 
      Height          =   375
      Left            =   1200
      TabIndex        =   39
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0ADA
   End
   Begin RichTextLib.RichTextBox b13 
      Height          =   375
      Left            =   2280
      TabIndex        =   40
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0B5C
   End
   Begin RichTextLib.RichTextBox b14 
      Height          =   375
      Left            =   3360
      TabIndex        =   41
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0BDE
   End
   Begin RichTextLib.RichTextBox b16 
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0C60
   End
   Begin RichTextLib.RichTextBox b15 
      Height          =   375
      Left            =   4440
      TabIndex        =   43
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0CE2
   End
   Begin RichTextLib.RichTextBox b17 
      Height          =   375
      Left            =   1200
      TabIndex        =   44
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0D64
   End
   Begin RichTextLib.RichTextBox b18 
      Height          =   375
      Left            =   2280
      TabIndex        =   45
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0DE6
   End
   Begin RichTextLib.RichTextBox b19 
      Height          =   375
      Left            =   3360
      TabIndex        =   46
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0E68
   End
   Begin RichTextLib.RichTextBox b20 
      Height          =   375
      Left            =   4440
      TabIndex        =   47
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0EEA
   End
   Begin RichTextLib.RichTextBox b21 
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0F6C
   End
   Begin RichTextLib.RichTextBox b22 
      Height          =   375
      Left            =   1200
      TabIndex        =   49
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":0FEE
   End
   Begin RichTextLib.RichTextBox b23 
      Height          =   375
      Left            =   2280
      TabIndex        =   50
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":1070
   End
   Begin RichTextLib.RichTextBox b24 
      Height          =   375
      Left            =   3360
      TabIndex        =   51
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":10F2
   End
   Begin RichTextLib.RichTextBox b26 
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":1174
   End
   Begin RichTextLib.RichTextBox b25 
      Height          =   375
      Left            =   4440
      TabIndex        =   53
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":11F6
   End
   Begin RichTextLib.RichTextBox b27 
      Height          =   375
      Left            =   1200
      TabIndex        =   54
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":1278
   End
   Begin RichTextLib.RichTextBox b28 
      Height          =   375
      Left            =   2280
      TabIndex        =   55
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":12FA
   End
   Begin RichTextLib.RichTextBox b29 
      Height          =   375
      Left            =   3360
      TabIndex        =   56
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"op.frx":137C
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Protect 2004 Access Options"
      Height          =   3615
      Left            =   6240
      TabIndex        =   61
      Top             =   840
      Width           =   3495
      Begin VB.CommandButton Command6 
         Caption         =   "&Unregister"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Register"
         Height          =   375
         Left            =   2040
         TabIndex        =   68
         Top             =   3120
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox text1 
         Height          =   255
         Left            =   1560
         TabIndex        =   66
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   65535
         BorderStyle     =   0
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   16
         TextRTF         =   $"op.frx":13FE
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Do not allow use of Protect 2004"
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Allow use of Protect 2004"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Registration Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Your Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3480
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "What To Protect:"
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      Begin VB.CheckBox Check24 
         BackColor       =   &H0000FFFF&
         Caption         =   "Windows Explorer"
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H0000FFFF&
         Caption         =   "Remote Desktop"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H0000FFFF&
         Caption         =   "Accsessibillity Wizard"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0000FFFF&
         Caption         =   "My Pictures"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Free Demo Version is in use - Please Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   71
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Protect 2004 Options - Version 1.50"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   26
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "op"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
b1.Text = "1"
b1.SaveFile ("c:\windows\system32\value1.rtf")
Else
b1.Text = "0"
b1.SaveFile ("c:\windows\system32\value1.rtf")
End If
End Sub


Private Sub Check10_Click()
On Error Resume Next
If Check10.Value = 1 Then
b10.Text = "1"
b10.SaveFile ("c:\windows\system32\value10.rtf")
Else
b10.Text = "0"
b10.SaveFile ("c:\windows\system32\value10.rtf")
End If
End Sub

Private Sub Check11_Click()
On Error Resume Next
If Check11.Value = 1 Then
b11.Text = "1"
b11.SaveFile ("c:\windows\system32\value11.rtf")
Else
b11.Text = "0"
b11.SaveFile ("c:\windows\system32\value11.rtf")
End If
End Sub

Private Sub Check12_Click()
On Error Resume Next
If Check12.Value = 1 Then
b12.Text = "1"
b12.SaveFile ("c:\windows\system32\value12.rtf")
Else
b12.Text = "0"
b12.SaveFile ("c:\windows\system32\value12.rtf")
End If
End Sub

Private Sub Check13_Click()
On Error Resume Next
If Check13.Value = 1 Then
b13.Text = "1"
b13.SaveFile ("c:\windows\system32\value13.rtf")
Else
b13.Text = "0"
b13.SaveFile ("c:\windows\system32\value13.rtf")
End If
End Sub

Private Sub Check14_Click()
On Error Resume Next
If Check14.Value = 1 Then
b14.Text = "1"
b14.SaveFile ("c:\windows\system32\value14.rtf")
Else
b14.Text = "0"
b14.SaveFile ("c:\windows\system32\value14.rtf")
End If
End Sub

Private Sub Check15_Click()
On Error Resume Next
If Check15.Value = 1 Then
b15.Text = "1"
b15.SaveFile ("c:\windows\system32\value15.rtf")
Else
b15.Text = "0"
b15.SaveFile ("c:\windows\system32\value15.rtf")
End If
End Sub

Private Sub Check16_Click()
On Error Resume Next
If Check16.Value = 1 Then
b16.Text = "1"
b16.SaveFile ("c:\windows\system32\value16.rtf")
Else
b16.Text = "0"
b16.SaveFile ("c:\windows\system32\value16.rtf")
End If
End Sub

Private Sub Check17_Click()
On Error Resume Next
If Check17.Value = 1 Then
b17.Text = "1"
b17.SaveFile ("c:\windows\system32\value17.rtf")
Else
b17.Text = "0"
b17.SaveFile ("c:\windows\system32\value17.rtf")
End If
End Sub

Private Sub Check18_Click()
On Error Resume Next
If Check18.Value = 1 Then
b18.Text = "1"
b18.SaveFile ("c:\windows\system32\value18.rtf")
Else
b18.Text = "0"
b18.SaveFile ("c:\windows\system32\value18.rtf")
End If
End Sub

Private Sub Check19_Click()
On Error Resume Next
If Check19.Value = 1 Then
b19.Text = "1"
b19.SaveFile ("c:\windows\system32\value19.rtf")
Else
b19.Text = "0"
b19.SaveFile ("c:\windows\system32\value19.rtf")
End If
End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
b2.Text = "1"
b2.SaveFile ("c:\windows\system32\value2.rtf")
Else
b2.Text = "0"
b2.SaveFile ("c:\windows\system32\value2.rtf")
End If
End Sub


Private Sub Check20_Click()
On Error Resume Next
If Check20.Value = 1 Then
b20.Text = "1"
b20.SaveFile ("c:\windows\system32\value20.rtf")
Else
b20.Text = "0"
b20.SaveFile ("c:\windows\system32\value20.rtf")
End If
End Sub

Private Sub Check21_Click()
On Error Resume Next
If Check21.Value = 1 Then
b21.Text = "1"
b21.SaveFile ("c:\windows\system32\value21.rtf")
Else
b21.Text = "0"
b21.SaveFile ("c:\windows\system32\value21.rtf")
End If
End Sub

Private Sub Check22_Click()
On Error Resume Next
If Check22.Value = 1 Then
b22.Text = "1"
b22.SaveFile ("c:\windows\system32\value22.rtf")
Else
b22.Text = "0"
b22.SaveFile ("c:\windows\system32\value22.rtf")
End If
End Sub

Private Sub Check23_Click()
On Error Resume Next
If Check23.Value = 1 Then
b23.Text = "1"
b23.SaveFile ("c:\windows\system32\value23.rtf")
Else
b23.Text = "0"
b23.SaveFile ("c:\windows\system32\value23.rtf")
End If
End Sub

Private Sub Check24_Click()
On Error Resume Next
If Check24.Value = 1 Then
b24.Text = "1"
b24.SaveFile ("c:\windows\system32\value24.rtf")
Else
b24.Text = "0"
b24.SaveFile ("c:\windows\system32\value24.rtf")
End If
End Sub

Private Sub Check25_Click()
On Error Resume Next
If Check25.Value = 1 Then
On Error Resume Next
Set b = CreateObject("wscript.shell")
s = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
b.RegWrite s, 1, "REG_DWORD"
Else
On Error Resume Next
Set b = CreateObject("wscript.shell")
s = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
b.RegDelete s
End If
End Sub

Private Sub Check26_Click()
On Error Resume Next
If Check26.Value = 1 Then
b26.Text = "1"
b26.SaveFile ("c:\windows\system32\value26.rtf")
Else
b26.Text = "0"
b26.SaveFile ("c:\windows\system32\value26.rtf")
End If
End Sub

Private Sub Check27_Click()
On Error Resume Next
If Check27.Value = 1 Then
b27.Text = "1"
b27.SaveFile ("c:\windows\system32\value27.rtf")
Else
b27.Text = "0"
b27.SaveFile ("c:\windows\system32\value27.rtf")
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 1 Then
b3.Text = "1"
b3.SaveFile ("c:\windows\system32\value3.rtf")
Else
b3.Text = "0"
b3.SaveFile ("c:\windows\system32\value3.rtf")
End If
End Sub

Private Sub Check4_Click()
On Error Resume Next
If Check4.Value = 1 Then
b4.Text = "1"
b4.SaveFile ("c:\windows\system32\value4.rtf")
Else
b4.Text = "0"
b4.SaveFile ("c:\windows\system32\value4.rtf")
End If
End Sub

Private Sub Check5_Click()
On Error Resume Next
If Check5.Value = 1 Then
b5.Text = "1"
b5.SaveFile ("c:\windows\system32\value5.rtf")
Else
b5.Text = "0"
b5.SaveFile ("c:\windows\system32\value5.rtf")
End If
End Sub

Private Sub Check6_Click()
On Error Resume Next
If Check6.Value = 1 Then
b6.Text = "1"
b6.SaveFile ("c:\windows\system32\value6.rtf")
Else
b6.Text = "0"
b6.SaveFile ("c:\windows\system32\value6.rtf")
End If
End Sub

Private Sub Check7_Click()
On Error Resume Next
If Check7.Value = 1 Then
b7.Text = "1"
b7.SaveFile ("c:\windows\system32\value7.rtf")
Else
b7.Text = "0"
b7.SaveFile ("c:\windows\system32\value7.rtf")
End If
End Sub

Private Sub Check8_Click()
On Error Resume Next
If Check8.Value = 1 Then
b8.Text = "1"
b8.SaveFile ("c:\windows\system32\value8.rtf")
Else
b8.Text = "0"
b8.SaveFile ("c:\windows\system32\value8.rtf")
End If
End Sub

Private Sub Check9_Click()
On Error Resume Next
If Check9.Value = 1 Then
b9.Text = "1"
b9.SaveFile ("c:\windows\system32\value9.rtf")
Else
b9.Text = "0"
b9.SaveFile ("c:\windows\system32\value9.rtf")
End If
End Sub

Private Sub Command1_Click()
MsgBox "Warning: Any changes made will only take effect after you restart Protect 2004.", vbExclamation, "Warning!"
op.Visible = False
End Sub

Private Sub Command2_Click()
CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "0"
If op.text2.Text = "0040-0110" Then
End
Else
MsgBox "Please register Protect 2004 to unlock all features and remove this message.", vbExclamation
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
ac.Text = "y"
ac.SaveFile ("c:\windows\system32\pro04access.rft")
MsgBox "Access to Protect 2004 has now been made allowed.", vbInformation
End Sub

Private Sub Command4_Click()
On Error Resume Next
ac.Text = "n"
ac.SaveFile ("c:\windows\system32\pro04access.rft")
MsgBox "Access to Protect 2004 has now been blocked.", vbInformation
End Sub

Private Sub Command5_Click()
On Error Resume Next
If text2.Text = "0040-0110" Then
text1.SaveFile ("c:\windows\system32\regnamepro04.rtf")
text2.SaveFile ("c:\windows\system32\regcodepro04.rtf")
MsgBox "Thank you for registering Protect 2004 with ADRANIX", vbInformation, "Thanks!"
End
Else
MsgBox "The Registration Code given is not correct, please try again.", vbCritical
End If
End Sub




Private Sub Command6_Click()
On Error Resume Next
text2.Text = ""
text2.SaveFile ("c:\windows\system32\regcodepro04.rtf")
MsgBox "You are now Unregistered. To Register again enter your registration code again.", vbExclamation
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim pin As Integer
Randomize
ac.LoadFile ("c:\windows\system32\pro04access.rft")
If ac.Text = "n" Then
MsgBox "Access to Protect 2004 is denied.  To continue you must click OK and enter the overide code otherwise click OK and click Cancel.", vbCritical
pin1 = InputBox("Please Enter Overide Code to continue:")
If pin1 = "23450" Then
op.Visible = True
Else
MsgBox "Overide code is not correct.", vbCritical
End
End If
End If
End Sub


Private Sub Timer2_Timer()
If text2.Text = "0040-0110" Then
Timer2.Enabled = False
Else
op.Check1.Value = 0
op.Check3.Value = 0
op.Check5.Value = 0
op.Check7.Value = 0
op.Check9.Value = 0
op.Check11.Value = 0
op.Check13.Value = 0
op.Check15.Value = 0
op.Check17.Value = 0
op.Check19.Value = 0
op.Check21.Value = 0
op.Check24.Value = 0
op.Check25.Value = 0
op.Check27.Value = 0
End If
End Sub
