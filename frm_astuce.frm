VERSION 5.00
Begin VB.Form frm_astuce 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Astuce ..."
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Change Name and/or size of personal icon -> Left Click"
      Height          =   420
      Index           =   6
      Left            =   765
      TabIndex        =   16
      Top             =   4665
      Width           =   6585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Personalise buton icon -> Drag/drop a .ico file (from windows explorer for example) to button program."
      Height          =   750
      Index           =   5
      Left            =   765
      TabIndex        =   15
      Top             =   4125
      Width           =   6585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Win 8 Metro  V 1.0.3 |  M_simohamed@hotmail.com (with all my compliments)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   14
      Top             =   7710
      Width           =   5535
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edition Mode "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   4
      Left            =   165
      TabIndex        =   13
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Personnal interface, buton style , colors, save/restore configuration, activate edition mode ...etc)"
      Height          =   510
      Index           =   3
      Left            =   855
      TabIndex        =   12
      Top             =   1395
      Width           =   5670
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Use search zone (on buttom on the left) to search file/folder existant on desktop or start menu."
      Height          =   600
      Index           =   2
      Left            =   765
      TabIndex        =   11
      Top             =   5670
      Width           =   6630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   465
      TabIndex        =   10
      Top             =   5265
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   465
      TabIndex        =   9
      Top             =   1980
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Mode "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   1
      Left            =   165
      TabIndex        =   8
      Top             =   45
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Show/Hide tool Box-> Right click on empty space"
      Height          =   210
      Index           =   1
      Left            =   765
      TabIndex        =   7
      Top             =   1110
      Width           =   4200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   465
      TabIndex        =   6
      Top             =   435
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "- Find shortcut destination -> CTRL + Shift + left Click (in case of shurtcut only)"
      Height          =   540
      Left            =   765
      TabIndex        =   5
      Top             =   3015
      Width           =   6060
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- open parent folder -> Shift + Left Click."
      Height          =   210
      Left            =   765
      TabIndex        =   4
      Top             =   2685
      Width           =   3390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Launch Program -> simple Click."
      Height          =   210
      Left            =   765
      TabIndex        =   3
      Top             =   2370
      Width           =   2715
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Open Folder inside Form (without using windows Explorer) -> CTRL + Click (in case of folder only)."
      Height          =   795
      Left            =   765
      TabIndex        =   2
      Top             =   3570
      Width           =   6780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "- draw a rectangle with mouse to select many buton. Use {CTRL} ou {SHIFT} to add/delete a buton to selection."
      Height          =   510
      Left            =   765
      TabIndex        =   1
      Top             =   6750
      Width           =   6900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Change Wallpaper - > Dbl Click on empty space."
      Height          =   210
      Index           =   0
      Left            =   765
      TabIndex        =   0
      Top             =   780
      Width           =   4065
   End
End
Attribute VB_Name = "frm_astuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

