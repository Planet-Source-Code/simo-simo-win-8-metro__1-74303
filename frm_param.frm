VERSION 5.00
Begin VB.Form frm_param 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Config"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   240
      Picture         =   "frm_param.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   390
      TabIndex        =   5
      Top             =   1050
      Width           =   390
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminate personal icon and restore the original one"
      Height          =   1215
      Left            =   2205
      TabIndex        =   3
      Top             =   1095
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Name"
      Height          =   390
      Left            =   2835
      TabIndex        =   2
      Top             =   195
      Width           =   1155
   End
   Begin VB.ListBox List_taille_ico 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   1590
      ItemData        =   "frm_param.frx":09A2
      Left            =   750
      List            =   "frm_param.frx":09BB
      TabIndex        =   1
      Top             =   1005
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   225
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   195
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1860
      X2              =   1860
      Y1              =   1020
      Y2              =   2325
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D8E9EC&
      BackStyle       =   0  'Transparent
      Caption         =   "Icon dimension :"
      Height          =   285
      Left            =   225
      TabIndex        =   4
      Top             =   750
      Width           =   1545
   End
End
Attribute VB_Name = "frm_param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Trim$(text1) <> "" Then
        frmmain.pic_prog(frmmain.indexencours).Tag = frmmain.mtag(frmmain.indexencours) & ";" & Trim$(text1)
        frmmain.actualise_ico_caption frmmain.indexencours, frmmain.mcaption(frmmain.indexencours)
        'Unload Me
    Else
        text1.SetFocus
    End If


End Sub

Private Sub Command2_Click()

        If MsgBox("Restaurer l'icone du Programme?", vbYesNo + vbQuestion, "Restauration d'icone") = vbYes Then
            Call WriteToINI("config", "progico" & frmmain.indexencours, "", App.path & "\config.ini")
            Call WriteToINI("config", "progicotaille" & frmmain.indexencours, "", App.path & "\config.ini")
            frmmain.actualise_ico_caption frmmain.indexencours
            Me.Height = 1095
            List_taille_ico.Visible = False
            Command2.Visible = False
            If text1.Enabled Then text1.SetFocus
            'List_taille_ico.Visible = False
        End If


End Sub

Private Sub Form_Activate()
'On Error Resume Next
If text1.Enabled Then text1.SelStart = 0: text1.SelLength = Len(text1)
List_taille_ico.ListIndex = 0
End Sub

Private Sub List_taille_ico_Click()

Dim taille As String * 3
Dim valeur As String
If indexencours <> -1 Then

    valeur = Trim$(List_taille_ico)
    If valeur = "" Then Exit Sub
    
'    If valeur = "Aucune" Then 'plus de icon perso
'    Else
    'If LCase(Right$(Tempo, 4)) = ".ico" Then 'ico perso
        taille = Mid$(valeur, 1, 3)
        taille = Replace(taille, "x", "")
        Call WriteToINI("config", "progicotaille" & frmmain.indexencours, taille, App.path & "\config.ini")
        frmmain.actualise_ico_caption frmmain.indexencours
        Exit Sub
    'End If
'    End If

End If


End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    'Text_name.Visible = False
ElseIf KeyAscii = 13 Then
    'Text_name.Visible = False
    't = InputBox("Veuillez spécifier le Nom du Programme", "Nom du Programme", mcaption(Index))
    'pic_prog(index).Enabled = True
    Command1_Click
End If


End Sub


