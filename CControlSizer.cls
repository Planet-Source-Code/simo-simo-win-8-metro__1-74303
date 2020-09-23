VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControlSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=======================================
'Downloaded from Visual Basic Thunder
'www.vbthunder.com
'Created on: 9/12/2002 6:59:25 PM
'=======================================
'Last modified 9/15/2002
' - Fixed resizing error when the size rect's right edge
'   or bottom edge was less than the left or top
' - Made GridSize and HandleSize properties, and fixed
'   the problem with controls not snapping correctly
'   to the form grid if the control placement in VB's
'   design mode used a different grid size than the
'   CControlSizer grid size.
' - Set handles' TabStop to False
' - The Load/SaveControls methods now use App.Title
'   and CurFrm.Name in saving positions to the registry,
'   so that we can get rid of the ControlKey property
'   and still have a unique save location for each
'   application/form.
' - Added an error trap for controls that have one or
'   more read-only size-related properties.
'=======================================

Public GridSize As Long
Public HandleSize As Long

'Eight pictureboxes to use as resize
'handles.
Dim WithEvents pTL As PictureBox 'top left
Attribute pTL.VB_VarHelpID = -1
Dim WithEvents pT As PictureBox 'top
Attribute pT.VB_VarHelpID = -1
Dim WithEvents pTR As PictureBox 'top right
Attribute pTR.VB_VarHelpID = -1
Dim WithEvents pL As PictureBox 'left
Attribute pL.VB_VarHelpID = -1
Dim WithEvents pR As PictureBox 'right
Attribute pR.VB_VarHelpID = -1
Dim WithEvents pBL As PictureBox 'bottom left
Attribute pBL.VB_VarHelpID = -1
Dim WithEvents pB As PictureBox 'bottom
Attribute pB.VB_VarHelpID = -1
Dim WithEvents pBR As PictureBox 'bottom right
Attribute pBR.VB_VarHelpID = -1
Dim WithEvents pM As PictureBox 'mover
Attribute pM.VB_VarHelpID = -1

'The form and control clients that are
'utilizing this class
Dim CurFrm As VB.Form
Dim CurCtl As VB.Control

'Flags for moving/sizing
Dim bSizing As Boolean
Dim lStartX As Long
Dim lStartY As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim rc As RECT
Dim sizerc As RECT
Dim bErase As Boolean



Private OldX As Long
Private OldY As Long
Private MoveIt As Boolean
Dim touche_ctrl As Boolean

Public Sub AttachForm(Frm As VB.Form)

On Error Resume Next

Set CurFrm = Frm

If Frm.Controls("pTL") Is Nothing Then
    'Add the handles and size them to HandleSize by HandleSize pixels
    Set pTL = Frm.Controls.Add("VB.PictureBox", "pTL")
    pTL.MousePointer = vbSizeNWSE
    SizeHandle pTL
    Set pT = Frm.Controls.Add("VB.PictureBox", "pT")
    pT.MousePointer = vbSizeNS
    SizeHandle pT
    Set pTR = Frm.Controls.Add("VB.PictureBox", "pTR")
    pTR.MousePointer = vbSizeNESW
    SizeHandle pTR

    Set pL = Frm.Controls.Add("VB.PictureBox", "pL")
    pL.MousePointer = vbSizeWE
    SizeHandle pL
    Set pR = Frm.Controls.Add("VB.PictureBox", "pR")
    pR.MousePointer = vbSizeWE
    SizeHandle pR

    Set pBL = Frm.Controls.Add("VB.PictureBox", "pBL")
    pBL.MousePointer = vbSizeNESW
    SizeHandle pBL
    Set pB = Frm.Controls.Add("VB.PictureBox", "pB")
    pB.MousePointer = vbSizeNS
    SizeHandle pB
    Set pBR = Frm.Controls.Add("VB.PictureBox", "pBR")
    pBR.MousePointer = vbSizeNWSE
    SizeHandle pBR

    'Add the mover handle
    Set pM = Frm.Controls.Add("VB.PictureBox", "pM")
    pM.MousePointer = vbSizeAll
    SizeHandle pM
    pM.Width = CurFrm.ScaleX(HandleSize * 2, vbPixels, CurFrm.ScaleMode)
    pM.Height = CurFrm.ScaleY(HandleSize * 2, vbPixels, CurFrm.ScaleMode)

End If

End Sub

Public Sub AttachControl(ctl As VB.Control)
Dim i As Integer
Dim lOfs As Long
lOfs = CurFrm.ScaleX(3, vbPixels, CurFrm.ScaleMode)

'Save the current control
Set CurCtl = ctl


pTL.Move CurCtl.Left - pTL.Width - lOfs, CurCtl.Top - pTL.Height - lOfs
'top
pT.Move CurCtl.Left + ((CurCtl.Width - pT.Width) \ 2), _
    CurCtl.Top - pT.Height - lOfs
'top right
pTR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top - pTR.Height - lOfs

'left
pL.Move CurCtl.Left - pL.Width - lOfs, _
    CurCtl.Top + ((CurCtl.Height - pL.Height) \ 2)
'right
pR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top + ((CurCtl.Height - pL.Height) \ 2)

'bottom left
pBL.Move CurCtl.Left - pTL.Width - lOfs, _
    CurCtl.Top + CurCtl.Height + lOfs
'bottom
pB.Move CurCtl.Left + ((CurCtl.Width - pT.Width) \ 2), _
    CurCtl.Top + CurCtl.Height + lOfs
'bottom right
pBR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top + CurCtl.Height + lOfs


pM.Width = CurCtl.Width '- 40
pM.Height = CurCtl.Height ' - 40

pM.Move CurCtl.Left, CurCtl.Top

pM.Picture = CurCtl.Image


pM.OLEDropMode = 1
'pM.BorderStyle = 0

'pM.AutoRedraw = True
pM.Refresh


''Show them!
Set pTL.Container = ctl.Parent: pTL.Visible = True: pTL.ZOrder 0
Set pT.Container = ctl.Parent: pT.Visible = True: pT.ZOrder 0
Set pTR.Container = ctl.Parent: pTR.Visible = True: pTR.ZOrder 0

Set pL.Container = ctl.Parent: pL.Visible = True: pL.ZOrder 0
Set pR.Container = ctl.Parent: pR.Visible = True: pR.ZOrder 0

Set pBL.Container = ctl.Parent: pBL.Visible = True: pBL.ZOrder 0
Set pB.Container = ctl.Parent: pB.Visible = True: pB.ZOrder 0
Set pBR.Container = ctl.Parent: pBR.Visible = True: pBR.ZOrder 0

Set pM.Container = ctl.Parent: pM.Visible = True: pM.ZOrder 0
'
End Sub

Public Sub HideHandles()

pTL.Visible = False
pT.Visible = False
pTR.Visible = False

pL.Visible = False
pR.Visible = False

pBL.Visible = False
pB.Visible = False
pBR.Visible = False

pM.Visible = False

End Sub

Function restauration_needed() As Boolean
Dim i As Integer

restauration_needed = False



For i = 0 To frmmain.pic_prog.UBound
    If frmmain.pic_prog(i).WhatsThisHelpID = 1 Then 'And CurCtl.Index <> i Then  'Or i = index Then 'déjà sel
        If frmmain.pic_prog(i).Left < 15 Then
            restauration_needed = True
        ElseIf frmmain.pic_prog(i).Top < 15 Then
            restauration_needed = True
        ElseIf frmmain.pic_prog(i).Left + frmmain.pic_prog(i).Width > frmmain.Width Then
            restauration_needed = True
        ElseIf frmmain.pic_prog(i).Top + frmmain.pic_prog(i).Height > frmmain.Height Then
            restauration_needed = True
        End If
        
        If restauration_needed = True Then Exit Function
    End If
Next i




End Function

Private Sub SizeHandle(Pic As PictureBox)

'Initializes the basic properties that are
'common to all the drag handles

'Set the Tag - this class checks the Tag
'during control loading/saving, so that it
'doesn't inadvertently save the position of
'the drag handles.
Pic.TabStop = False
Pic.Tag = "VBTDragHandle"
Pic.BackColor = vbHighlight
Pic.BorderStyle = 0
Pic.ScaleMode = vbPixels
Pic.Width = CurFrm.ScaleX(HandleSize, vbPixels, CurFrm.ScaleMode)
Pic.Height = CurFrm.ScaleY(HandleSize, vbPixels, CurFrm.ScaleMode)
'Bring to the front
Pic.ZOrder 0

End Sub

Private Sub DrawHandle(Pic As PictureBox)

'Draws a white border along the handle
Pic.Line (0, 0)-(Pic.ScaleWidth - 1, Pic.ScaleHeight - 1), vbHighlightText, B

End Sub

Private Sub Class_Initialize()

GridSize = 8
HandleSize = 8

End Sub

Private Sub pB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
        
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Bottom = ((rc.Bottom - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If
''
End Sub

Private Sub pB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y



End Sub

Private Sub pB_Paint()

DrawHandle pB

End Sub

Private Sub pBL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pBL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Left = ((rc.Left - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    sizerc.Bottom = ((rc.Bottom - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If
''

End Sub

Private Sub pBL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pBL_Paint()

DrawHandle pBL

End Sub

Private Sub pBR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pBR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect


Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    'sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - x), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'sizerc.Right = sizerc.Left + CurCtl.Width
    'sizerc.Bottom = sizerc.Top + CurCtl.Height
    sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'''    sizerc.Left = ((rc.Left - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
'''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
'''    sizerc.Top = ((rc.Top - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
'''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
'''    sizerc.Right = sizerc.Left + CurCtl.Width
'''    sizerc.Bottom = sizerc.Top + CurCtl.Height
''    sizerc.Right = ((rc.Right - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    sizerc.Bottom = ((rc.Bottom - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If


End Sub

Private Sub pBR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pBR_Paint()

DrawHandle pBR

End Sub

Private Sub pL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'
'DrawSizeRect

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Left = ((rc.Left - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If


End Sub

Private Sub pL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pL_Paint()

DrawHandle pL

End Sub

Private Sub pM_Click()

        'juste ou cas ou
        If frmmain.indexencours = -1 Then Exit Sub

        If touche_ctrl = True And frmmain.pic_prog(frmmain.indexencours).WhatsThisHelpID = 1 Then
            frmmain.pic_prog(frmmain.indexencours).WhatsThisHelpID = 0
            frmmain.actualise_ico_caption frmmain.indexencours
            'pic_prog(i).ZOrder 0
            'ctl.AttachControl pic_prog(i)
            HideHandles
            touche_ctrl = False
        End If

End Sub

Private Sub pM_DragDrop(Source As Control, X As Single, Y As Single)

frmmain.pic_prog_DragDrop frmmain.indexencours, Source, X, Y

End Sub

Private Sub pM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

''BeginSize x, y
''
''If Button = 2 Then
''    'modifier le caption en mode edition
''    frmmain.pic_prog_MouseDown frmmain.indexencours, 2, 0, 0, 0
''ElseIf Button = 1 Then
''
''    frmmain.pic_prog_MouseDown frmmain.indexencours, 1, Shift, x, y
''
''End If
    
    If Shift = 2 Then touche_ctrl = True

    OldX = X
    OldY = Y
    MoveIt = True
    
    
    save_selection 'en cas de move en dehors de la form, j'utilise le restore_selection_position
    
End Sub

Sub save_selection()

Dim i As Integer
            
    'CurCtl.Move valx, valy
    Open App.path & "\sel_tmp.ini" For Output As #2
    
    For i = 0 To frmmain.pic_prog.UBound
        If frmmain.pic_prog(i).WhatsThisHelpID = 1 Then 'And CurCtl.Index <> i Then  'Or i = index Then 'déjà sel
            'frmmain.pic_prog(i).Move frmmain.pic_prog(i).Left - pasx, frmmain.pic_prog(i).Top - pasy
            Call WriteToINI("config", "progleft" & i, frmmain.pic_prog(i).Left, App.path & "\sel_tmp.ini")
            Call WriteToINI("config", "progtop" & i, frmmain.pic_prog(i).Top, App.path & "\sel_tmp.ini")
        End If
    Next i

    Close 2

End Sub


Sub restaure_selection_position()

Dim i As Integer
            
    'CurCtl.Move valx, valy
    Open App.path & "\sel_tmp.ini" For Input As #2
    
    For i = 0 To frmmain.pic_prog.UBound
        If frmmain.pic_prog(i).WhatsThisHelpID = 1 Then 'And CurCtl.Index <> i Then  'Or i = index Then 'déjà sel
            frmmain.pic_prog(i).Left = GetFromINI("config", "progleft" & i, App.path & "\sel_tmp.ini")
            frmmain.pic_prog(i).Top = GetFromINI("config", "progtop" & i, App.path & "\sel_tmp.ini")
            frmmain.pic_prog(i).Refresh
        End If
    Next i

    Close 2
    AttachControl CurCtl

End Sub
Private Sub pM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer
Dim gx As Single
Dim gy As Single
Dim valx As Integer
Dim valy As Integer

Dim pasx As Integer
Dim pasy As Integer

'If bSizing = False Then
'    Exit Sub
'End If
'
'''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
'    DrawSizeRect_form True
'    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'    sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - x), _
'        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'    sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - y), _
'        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'    sizerc.Right = sizerc.Left + CurCtl.Width
'    sizerc.Bottom = sizerc.Top + CurCtl.Height
'    'Commented out: By using the following lines instead of the
'    'above uncommented lines, the Bottom and Right will also
'    'snap to the grid during the move, which could cause some
'    'resizing if the grid size changes at all.
'    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
'    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
'    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'    DrawSizeRect_form
'    SizeControl
'
'
'
'''Else
'''    For i = 0 To frmmain.pic_conteneur.uBound
'''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
'''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
'''            Exit For
'''        End If
'''    Next i
'''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'''    sizerc.Left = ((rc.Left - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
'''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
'''    sizerc.Top = ((rc.Top - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
'''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
'''    sizerc.Right = sizerc.Left + CurCtl.Width
'''    sizerc.Bottom = sizerc.Top + CurCtl.Height
'''    DrawSizeRect_control frmmain.pic_conteneur(i)
'''
'''End If

    If MoveIt Then
        
        If frmmain.check_move_libre.Value = 0 Then 'move libre
            valx = CurCtl.Left + X - OldX
            valy = CurCtl.Top + Y - OldY
            'If valx < 15 Then valx = 15
            'If CurCtl.Left + CurCtl.Width > frmmain.Width Then valx = frmmain.Width - CurCtl.Width  ' - CurCtl.Left
            'If valy < 15 Then valy = 15
            'If CurCtl.Top + CurCtl.Height > frmmain.Height Then valy = frmmain.Height - CurCtl.Height  ' - CurCtl.Left
            
            
            pasx = CurCtl.Left - valx
            pasy = CurCtl.Top - valy
            
            CurCtl.Move valx, valy
            
            'appliquer sur tous les objets sélectionnés
            For i = 0 To frmmain.pic_prog.UBound
                If frmmain.pic_prog(i).WhatsThisHelpID = 1 And CurCtl.index <> i Then  'Or i = index Then 'déjà sel
'                    valx = frmmain.pic_prog(i).Left + x - OldX
'                    valy = frmmain.pic_prog(i).Top + y - OldY
'                    'If valx < frmmain.Left Then valx = frmmain.Left
'                    'If valy > frmmain.Left + frmmain.Width + frmmain.pic_prog(i).Width Then valy = frmmain.Left + frmmain.Width + frmmain.pic_prog(i).Width
'
'                    If valx < 15 Then valx = 15
'                    If frmmain.pic_prog(i).Left + frmmain.pic_prog(i).Width > frmmain.Width Then valx = frmmain.Width - frmmain.pic_prog(i).Width  ' - CurCtl.Left
'                    If valy < 15 Then valy = 15
'                    If frmmain.pic_prog(i).Top + frmmain.pic_prog(i).Height > frmmain.Height Then valy = frmmain.Height - frmmain.pic_prog(i).Height  ' - CurCtl.Left
                    
                    frmmain.pic_prog(i).Move frmmain.pic_prog(i).Left - pasx, frmmain.pic_prog(i).Top - pasy
                End If
            Next i
            
        Else 'respecter la grille
            gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
            gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
            'sizerc.Left =
            valx = ((CurCtl.Left - CurFrm.ScaleX((OldX - X), _
                vbPixels, CurFrm.ScaleMode)) \ gx) * gx
            'sizerc.Top =
            valy = ((CurCtl.Top - CurFrm.ScaleY((OldY - Y), _
                vbPixels, CurFrm.ScaleMode)) \ gy) * gy
            'valx = sizerc.Left '+ x - OldX
            'valy = sizerc.Top '+ y - OldY
            'If valx < frmmain.Left Then valx = frmmain.Left
            'If valy > frmmain.Left + frmmain.Width + CurCtl.Width Then valy = frmmain.Left + frmmain.Width + CurCtl.Width
            
            pasx = CurCtl.Left - valx
            pasy = CurCtl.Top - valy
            
            CurCtl.Move valx, valy
            
            'appliquer sur tous les objets sélectionnés
            For i = 0 To frmmain.pic_prog.UBound
                If frmmain.pic_prog(i).WhatsThisHelpID = 1 And CurCtl.index <> i Then 'Or i = index Then 'déjà sel
'                    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'                    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'                    valx = ((frmmain.pic_prog(i).Left - CurFrm.ScaleX((OldX - x), _
'                        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'                    valy = ((frmmain.pic_prog(i).Top - CurFrm.ScaleY((OldY - y), _
'                        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'                    'If valx < frmmain.Left Then valx = frmmain.Left
'                    'If valy > frmmain.Left + frmmain.Width + frmmain.pic_prog(i).Width Then valy = frmmain.Left + frmmain.Width + frmmain.pic_prog(i).Width
                    
'                    frmmain.pic_prog(i).Move valx, valy
                    frmmain.pic_prog(i).Move frmmain.pic_prog(i).Left - pasx, frmmain.pic_prog(i).Top - pasy
                End If
            Next i
        End If
        
'            CurCtl.Move valx, valy
            'CurCtl.left = CurCtl.Left + x - OldX
            'CurCtl.Top = valy 'CurCtl.Top + y - OldY
        
        
        
'        Form2.Top = Me.Top - Form2.Height
'        Form2.Left = Me.Left
'
'        Form3.Top = Me.Top
'        Form3.Left = Me.Left + Me.Width
'
'        Form4.Top = Me.Top + Me.Height + 80
'        Form4.Left = (Me.Width * 0.5) + Me.Left - (Form4.Width * 0.5)
'
'        Form5.Top = Me.Top
'        Form5.Left = Me.Left - Form5.Width '(Me.Width * 0.5) + Me.Left - (Form2.Width * 0.5)
        
        AttachControl CurCtl
        frmmain.Refresh
    End If


End Sub

Private Sub pM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer

''EndSize

''    'If MoveIt = False Then 'un clique pour déselctionner
''        If CurCtl.WhatsThisHelpID = 1 Then
''            CurCtl.WhatsThisHelpID = 0
''            frmmain.actualise_ico_caption frmmain.indexencours
''            'pic_prog(i).ZOrder 0
''            'ctl.AttachControl pic_prog(i)
''            HideHandles
''        End If
''    'End If


''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    'Clear the resizing rectangle
    DrawSizeRect_form True
    'Clear the sizing flag
    If bSizing Then
        bSizing = False
        'Resize the control to the size that was requested
        SizeControl
        'Reposition the drag handles
        AttachControl CurCtl
        frmmain.Cls
        frmmain.DrawGrid
    End If
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            'Clear the resizing rectangle
''            DrawSizeRect_control frmmain.pic_conteneur(i), True
''            'Clear the sizing flag
''            bSizing = False
''            'Resize the control to the size that was requested
''            SizeControl
''            'Reposition the drag handles
''            AttachControl CurCtl
''            frmmain.pic_conteneur(i).Cls
''            Exit For
''        End If
''    Next i
''End If

    MoveIt = False
    'qd je déplace plusieurs pic_prog et que la position final dépasse les bords de la form,
    'je les remet a leur place avant déplacement
    If restauration_needed Then restaure_selection_position
End Sub

Private Sub pM_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

frmmain.pic_prog_OLEDragDrop frmmain.indexencours, Data, Effect, Button, Shift, X, Y
AttachControl frmmain.pic_prog(frmmain.indexencours)

End Sub

Private Sub pM_Paint()

DrawHandle pM

End Sub

Private Sub pR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'
'DrawSizeRect

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Right = ((rc.Right - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If


End Sub

Private Sub pR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pR_Paint()

DrawHandle pR

End Sub

Private Sub pT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect_form True
'
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect_form

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Top = ((rc.Top - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If
''


End Sub

Private Sub pT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pT_Paint()

DrawHandle pT

End Sub

Private Sub pTL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pTL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect

Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Left = ((rc.Left - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    sizerc.Top = ((rc.Top - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If



End Sub

Private Sub pTL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pTL_Paint()

DrawHandle pTL

End Sub

Private Sub pTR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pTR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim gx As Single
'Dim gy As Single
'
'If bSizing = False Then
'    Exit Sub
'End If
'
'DrawSizeRect True
'
'gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
'gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - x), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
'
'DrawSizeRect


Dim i As Integer
Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

''If frmmain.mparent = "frmmain" Or frmmain.mparent = "" Then
    DrawSizeRect_form True
    gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
    gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
    sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
        vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
        vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    'Commented out: By using the following lines instead of the
    'above uncommented lines, the Bottom and Right will also
    'snap to the grid during the move, which could cause some
    'resizing if the grid size changes at all.
    'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    '    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
    'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    '    vbPixels, CurFrm.ScaleMode)) \ gy) * gy
    DrawSizeRect_form
''Else
''    For i = 0 To frmmain.pic_conteneur.uBound
''        If frmmain.pic_conteneur(i).Name & "(" & i & ")" = frmmain.mparent Then
''            DrawSizeRect_control frmmain.pic_conteneur(i), True '', PictureBox)  'leconteneur_control
''            Exit For
''        End If
''    Next i
''    gx = frmmain.pic_conteneur(i).ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
''    gy = frmmain.pic_conteneur(i).ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
''    sizerc.Right = ((rc.Right - frmmain.pic_conteneur(i).ScaleX((lStartX - x), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gx) * gx
''    sizerc.Top = ((rc.Top - frmmain.pic_conteneur(i).ScaleY((lStartY - y), _
''        vbPixels, frmmain.pic_conteneur(i).ScaleMode)) \ gy) * gy
''    DrawSizeRect_control frmmain.pic_conteneur(i)
''
''End If


End Sub

Private Sub pTR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pM_MouseUp Button, Shift, X, Y

End Sub

Private Sub pTR_Paint()

DrawHandle pTR

End Sub


Private Sub DrawSizeRect_form(Optional ByVal bEraseOnly As Boolean = False)

'Draw the sizing rectangle during a size or move operation

Dim tW As Long
Dim tS As DrawModeConstants


tW = CurFrm.DrawWidth
tS = CurFrm.DrawMode
CurFrm.DrawWidth = 2
CurFrm.DrawMode = vbInvert

If bErase = True Then
    If bSizing = False Then
        CurFrm.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbRed, B
    Else
        CurFrm.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbRed, B
        'CurCtl.Move sizerc.Left / Screen.TwipsPerPixelX, sizerc.Top / Screen.TwipsPerPixelY
    End If
End If

If bEraseOnly = True Then
    bErase = False
    GoTo Cleanup
End If

rc.Left = CurCtl.Left
rc.Top = CurCtl.Top
rc.Right = rc.Left + CurCtl.Width
rc.Bottom = rc.Top + CurCtl.Height

If bSizing = False Then
    CurFrm.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbRed, B
Else
    CurFrm.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbRed, B
End If

bErase = True

Cleanup:
CurFrm.DrawWidth = tW
CurFrm.DrawMode = tS

        
End Sub
Private Sub DrawSizeRect_control(ctrl As Control, Optional ByVal bEraseOnly As Boolean = False)

'Draw the sizing rectangle during a size or move operation

Dim tW As Long
Dim tS As DrawModeConstants

ctrl.Cls

tW = ctrl.DrawWidth
tS = ctrl.DrawMode
ctrl.DrawWidth = 2

If bErase = True Then
    If bSizing = False Then
        ctrl.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbWhite, B
    Else
        ctrl.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbWhite, B
    End If
End If

If bEraseOnly = True Then
    bErase = False
    GoTo Cleanup
End If

rc.Left = CurCtl.Left
rc.Top = CurCtl.Top
rc.Right = rc.Left + CurCtl.Width
rc.Bottom = rc.Top + CurCtl.Height

If bSizing = False Then
    ctrl.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbWhite, B
Else
    ctrl.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbWhite, B
End If

bErase = True

Cleanup:
ctrl.DrawWidth = tW
ctrl.DrawMode = tS

        
        
End Sub


Private Sub SizeControl()

Dim lTemp As Long

'Check for the rect being inverted and fix it
If sizerc.Bottom < sizerc.Top Then
    lTemp = sizerc.Top
    sizerc.Top = sizerc.Bottom
    sizerc.Bottom = lTemp
End If
If sizerc.Right < sizerc.Left Then
    lTemp = sizerc.Left
    sizerc.Left = sizerc.Right
    sizerc.Right = lTemp
End If
'Make sure we aren't trying to resize the control
'to 0 width or height
If sizerc.Right - sizerc.Left < 1 Then
    sizerc.Right = sizerc.Left + 1
End If
If sizerc.Bottom - sizerc.Top < 1 Then
    sizerc.Bottom = sizerc.Top + 1
End If

On Error Resume Next
'Error trap catches "Height property is read-only"
'for certain controls.

'See the EndSize subroutine
CurCtl.Left = sizerc.Left
CurCtl.Top = sizerc.Top
CurCtl.Width = sizerc.Right - sizerc.Left
CurCtl.Height = sizerc.Bottom - sizerc.Top

End Sub

Private Sub EndSize()

'Halt the sizing operation - called from all drag handles.
'Clear the resizing rectangle
DrawSizeRect_form True
'Clear the sizing flag
bSizing = False
'Resize the control to the size that was requested
SizeControl
'Reposition the drag handles
AttachControl CurCtl

End Sub

Private Sub BeginSize(ByVal X As Long, ByVal Y As Long)
Dim i As Integer
'Begin the sizing operation - called from all drag handles.
HideHandles

    DrawSizeRect_form

bSizing = True
lStartX = X
lStartY = Y
sizerc = rc

End Sub

Public Sub SaveControls()

''Dim i As Long
''
''For i = 0 To CurFrm.Controls.Count - 1
''    If InStr(1, CurFrm.Controls(i).Tag, "VBTDragHandle", vbBinaryCompare) = 0 Then
''        'Not a CControlSizer handle, so save the position
''        SaveSetting App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Left", _
''            CurFrm.Controls(i).Left
''        SaveSetting App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Top", _
''            CurFrm.Controls(i).Top
''        SaveSetting App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Width", _
''            CurFrm.Controls(i).Width
''        SaveSetting App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Height", _
''            CurFrm.Controls(i).Height
''
''    End If
''Next

End Sub

Public Sub LoadControls()

''On Error Resume Next
'''Error trap catches "Height property is read-only"
'''for certain controls.
''
''Dim i As Long
''
''For i = 0 To CurFrm.Controls.Count - 1
''    If InStr(1, CurFrm.Controls(i).Tag, "VBTDragHandle", vbBinaryCompare) = 0 Then
''        'Not a CControlSizer handle, so load the position
''        CurFrm.Controls(i).Left = GetSetting(App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Left", CurFrm.Controls(i).Left)
''        CurFrm.Controls(i).Top = GetSetting(App.Title, CurFrm.Name, _
''            CurFrm.Controls(i).Name & "Top", CurFrm.Controls(i).Top)
''        CurFrm.Controls(i).Width = GetSetting(App.Title, CurFrm.Name, _
''                CurFrm.Controls(i).Name & "Width", CurFrm.Controls(i).Width)
''        CurFrm.Controls(i).Height = GetSetting(App.Title, CurFrm.Name, _
''                CurFrm.Controls(i).Name & "Height", CurFrm.Controls(i).Height)
''
''    End If
''Next

End Sub
