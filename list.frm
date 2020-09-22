VERSION 5.00
Begin VB.Form list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Locked Items List"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6240
      TabIndex        =   28
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label34 
      Caption         =   "http://www.adranix.co.uk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label33 
      Caption         =   "By Adam Ranshaw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   3720
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Label Label32 
      Caption         =   "Plus More"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   7080
      Width           =   5175
   End
   Begin VB.Label Label31 
      Caption         =   "Internet Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   31
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label30 
      Caption         =   "Almost all Folders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   6120
      Width           =   5175
   End
   Begin VB.Label Label29 
      Caption         =   "Recycle Bin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   29
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   600
      Y2              =   9240
   End
   Begin VB.Label Label28 
      Caption         =   "Synchrize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label27 
      Caption         =   "Local Security Policy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Label Label26 
      Caption         =   "Windows Messenger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   25
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Label Label25 
      Caption         =   "Remote Desktop Connection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Label Label24 
      Caption         =   "Windows Media Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label Label23 
      Caption         =   "Microsoft Hearts Network"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label Label22 
      Caption         =   "Spider Solitare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label21 
      Caption         =   "Solitare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label20 
      Caption         =   "3D Pinball"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label19 
      Caption         =   "Display Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label18 
      Caption         =   "Scheduled Tasks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   8520
      Width           =   5175
   End
   Begin VB.Label Label17 
      Caption         =   "Sound Recorder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   8040
      Width           =   5175
   End
   Begin VB.Label Label16 
      Caption         =   "Windows Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7560
      Width           =   5175
   End
   Begin VB.Label Label15 
      Caption         =   "Network Setup Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Width           =   5175
   End
   Begin VB.Label Label14 
      Caption         =   "Narrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label13 
      Caption         =   "Command Prompt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6120
      Width           =   5175
   End
   Begin VB.Label Label12 
      Caption         =   "Backup or Restore Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label11 
      Caption         =   "Windows XP Tour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label10 
      Caption         =   "My Pictures"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Label Label9 
      Caption         =   "My Music"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Label Label8 
      Caption         =   "Help and Support"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "Printers and Faxes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "My Network Places"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "My Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Items that are Locked are shown below:"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
list.Visible = False
password.Enabled = True
End Sub
