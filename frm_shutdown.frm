VERSION 5.00
Begin VB.Form frm_shutdown 
   BorderStyle     =   0  'None
   Caption         =   "Ordinateur"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_shutdown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic_action_ordinateur 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   0
      Picture         =   "frm_shutdown.frx":0CCA
      ScaleHeight     =   1830
      ScaleWidth      =   3660
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3660
      Begin VB.PictureBox pic_close 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3360
         Picture         =   "frm_shutdown.frx":169E4
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   465
         Left            =   2760
         TabIndex        =   3
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   465
         Left            =   1590
         TabIndex        =   2
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   465
         Left            =   375
         TabIndex        =   1
         Top             =   645
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_shutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 118 Then 'v - verrouiller
    Label1_Click
ElseIf KeyAscii = 116 Then 't - Arrêter
    Label2_Click
ElseIf KeyAscii = 114 Then 'r - redémarrer
    Label3_Click
End If

End Sub

Private Sub Form_Load()
frmmain.Enabled = False
Me.Width = Me.pic_action_ordinateur.Width
Me.Height = Me.pic_action_ordinateur.Height




''Shutdown Computer
''Shutdown.exe -s -t 00
''Restart Computer
''Shutdown.exe -r -t 00
''Lock Workstation
''rundll32.exe User32.dll, LockWorkStation
''Hibernate Computer
''rundll32.exe powrprof.dll, SetSuspendState
''Sleep Computer
''rundll32.exe powrprof.dll,SetSuspendState 0,1,0
''verrouiller
''Arrêter
''Redémarrer

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call frmmain.no_transparence
frmmain.Enabled = True
frmmain.Refresh
End Sub

Private Sub Label1_Click()
Shell "rundll32.exe User32.dll, LockWorkStation"
Unload Me
End Sub

Private Sub Label2_Click()
''Shutdown Computer
Shell "Shutdown.exe -s -t 00"
Unload Me
End Sub

Private Sub Label3_Click()
''Restart Computer
Shell "Shutdown.exe -r -t 00"
Unload Me
End Sub


Private Sub pic_action_ordinateur_Click()
'Unload Me
End Sub


Private Sub pic_close_Click()
Unload Me
End Sub


