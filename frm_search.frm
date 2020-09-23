VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_search 
   BackColor       =   &H002D3234&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8295
   ClientLeft      =   -45
   ClientTop       =   -90
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frm_search.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_search 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   90
      Picture         =   "frm_search.frx":6DC2
      ScaleHeight     =   390
      ScaleWidth      =   4830
      TabIndex        =   15
      ToolTipText     =   "Rechercher  ... (Bureau & Menu démarrer)"
      Top             =   255
      Width           =   4830
      Begin VB.TextBox text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   90
         Width           =   4395
      End
   End
   Begin VB.PictureBox pic_tri 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   105
      Picture         =   "frm_search.frx":D054
      ScaleHeight     =   180
      ScaleWidth      =   315
      TabIndex        =   14
      ToolTipText     =   "Trier (Asc/Desc)"
      Top             =   675
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4650
      Picture         =   "frm_search.frx":D396
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   12
      ToolTipText     =   "Afficher/cacher Information Chemin"
      Top             =   660
      Width           =   225
   End
   Begin VB.PictureBox pic_close 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   4320
      Picture         =   "frm_search.frx":D648
      ScaleHeight     =   165
      ScaleWidth      =   600
      TabIndex        =   11
      Top             =   45
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C07C2D&
      Caption         =   "Afficher colonne Chemin"
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   225
      TabIndex        =   9
      Top             =   1125
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   420
      Hidden          =   -1  'True
      Left            =   5790
      System          =   -1  'True
      TabIndex        =   8
      Top             =   750
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Top             =   1740
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   7365
      TabIndex        =   6
      Top             =   2100
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox pic16 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   8895
      ScaleHeight     =   390
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.FileListBox File2 
      Appearance      =   0  'Flat
      Height          =   420
      Hidden          =   -1  'True
      Left            =   5790
      System          =   -1  'True
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtFolder 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6285
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "..."
      Height          =   255
      Left            =   9645
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView lstFind 
      Height          =   6615
      Left            =   90
      TabIndex        =   0
      Top             =   900
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   11668
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   5910
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   30
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00606060&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   75
      TabIndex        =   10
      Top             =   7575
      Width           =   4815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher dans :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6450
      TabIndex        =   3
      Top             =   3330
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   15
      Picture         =   "frm_search.frx":DBB2
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2415
   End
End
Attribute VB_Name = "frm_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function LockWindowUpdate _
Lib "user32" (ByVal hwndLock As Long) _
As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SMALL_ICON        As Integer = 16
'Private Const Speed             As Integer = 60
'Private Expand                  As Boolean
'Private Frame                   As Integer
'Private DoResize                As Boolean
''Private WithEvents cFind        As clsSearch
'Private mbDate                  As Boolean
'Private mbType                  As Boolean
'Private mbSize                  As Boolean


'pour le recherche
Private ItemCount As Integer 'Long
'pour savoir le chemin des dossiers system (desktop,favoris, program, ...)
Private Enum SystemFolderIDs
    fldDESKTOP = &H0
    fldPROGRAMS = &H2
    fldPERSONAL = &H5
    fldFAVORITES = &H6
    fldSTARTUP = &H7
    fldRECENT = &H8
    fldSENDTO = &H9
    fldSTARTMENU = &HB
    fldDESKTOPDIRECTORY = &H10
    fldNETHOOD = &H13
    fldFONTS = &H14
    fldTEMPLATES = &H15
    fldCOMMON_STARTMENU = &H16
    fldCOMMON_PROGRAMS = &H17
    fldCOMMON_STARTUP = &H18
    fldCOMMON_DESKTOPDIRECTORY = &H19
    fldAPPDATA = &H1A
    fldPRINTHOOD = &H1B
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As SystemFolderIDs, ByRef pidl As Long) As Long
'pour detecter la touche appuyer (ex: shift, ctrl ...)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Const NOERROR = 0




'pour l'icone
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Private Const SHGFI_LARGEICON = &H0        ' Large icon
Private Const SHGFI_SMALLICON = &H1        ' Small icon
Private Const ILD_TRANSPARENT = &H1        ' Display transparent
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

Private ShInfo As SHFILEINFO


Dim rech_en_cours As Boolean
Dim rech_cancel As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
    lstFind.ColumnHeaders(2).Width = 6000 '.Add , , "Path", 0 ' 6000 'File Path Column
    Me.Width = 9965
Else
    lstFind.ColumnHeaders(2).Width = 0
    Me.Width = 4965
End If
End Sub

Private Sub cmd_tri_Click()

End Sub

Private Sub cmdFolder_Click()
'    Dim sTmp                    As String
'    sTmp = IIf(Len(Dir(txtFolder.Text, vbDirectory)) > 0, txtFolder.Text, "") 'If the text in the Folder text box is not a directory, we cannot use it as the start up for the Folder Open API or it will crash.
'    sTmp = BrowseForFolder(sTmp, "Select Folder To Search...") 'Show the Open Folder dialog.
'    If Len(sTmp) > 0 Then txtFolder.Text = sTmp
End Sub

Private Sub Command1_Click()



End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
Dir1.Refresh
File1.Refresh

End Sub

Private Sub Form_Activate()
text1.SetFocus
If text1 <> "" Then text1.SelStart = Len(text1): pic_search_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If rech_en_cours = False Then Call pic_close_Click
End If
End Sub


Private Sub Form_Load()
    'Set cFind = New clsSearch
    Dim i                       As Integer
'    For I = 0 To MenuHeader.Count - 1 'Setup the XP Menu frames.
'        If FrameXPMenu(I).Height = imgUp.Height Then
'            MenuHeader(I).Picture = imgDown.Picture
'        Else
'            MenuHeader(I).Picture = imgUp.Picture
'        End If
'        MenuHeader(I).Height = 375
'        MenuHeader(I).Width = FrameXPMenu(I).Width
'    Next
'    DoResize = False 'To stop it resizing the form while we make adjustments.
    
    pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX 'Set the Temp Picture Box properties.
    pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY 'Set the Temp Picture Box properties.

    lstFind.View = 3 'IIf(IsNumeric(sTmp), CInt(sTmp), 3)

'OpDate(0).Enabled = False 'Disable all the check box's while they are not visible.
'    OpDate(1).Enabled = False
'    OpLast(0).Enabled = False
'    OpLast(1).Enabled = False
'    txtDays.Enabled = False
'    txtMonths.Enabled = False
'    UpDownDays.Enabled = False
'    UpDownMonths.Enabled = False
'    cmbSize.Enabled = False
'    txtSize.Enabled = False
'    UpDownSize.Enabled = False
'    cmbSize.ListIndex = 0
    With lstFind 'Add column headers.
        .ColumnHeaders.Add , , "Nom", 4200 '600 'File Name Column
        .ColumnHeaders.Add , , "Path", 0 ' 6000 'File Path Column
        '.ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_SIZE), 1000 'File Size Column
        '.ColumnHeaders.Add , , "Byte Size", 0 'Raw Size Column for sorting only !
        '.ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_TYPE), 2000 'Type Column
        '.ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_MODIFIED), 3000 'Type Column
        '.ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_CREATED), 3000  'Type Column
        .HideColumnHeaders = True
    End With
    
    'Label1.BackStyle = 0
    Label1.Caption = "Prêt"
    
    'Image1.Visible = False
    Me.BackColor = 2961972 'vbBlack
    
End Sub
Private Sub Form_Resize()
Image1.Width = Me.Width - 60
'Shape1.Width = Me.Width - 30
pic_close.Left = Me.Width - pic_close.Width - 80
lstFind.Width = Me.Width - 180
Picture1.Left = Me.Width - Picture1.Width - 120 '60
Label1.Width = lstFind.Width
'    If Me.WindowState = vbMinimized Then Exit Sub 'It will error if its minimized !
'    If Me.Height <= 7305 Then Me.Height = 7305 'Don't let it go too small !
'    If Me.Width <= 10185 Then Me.Width = 10185
'    ' Now calculate the size's for resizing.
'    Picture3.Height = Me.Height - 1000
'    lstFind.Height = Picture3.Top + Picture3.Height - lstFind.Top
'    lstFind.Width = Me.Width - 195 - lstFind.Left
'    StatusBar1.Panels(1).Width = Me.Width - 200 '- StatusBar1.Panels(2).Width - 400
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    Dim vValue                  As Integer
'    'Save some data to the registry for next load time.
'    'vValue = chkSubFolders.Value
'    Call SaveSetting("Search Sub Folders", CStr(vValue))
'    vValue = chkHidden.Value
'    Call SaveSetting("Search Hidden Folders", CStr(vValue))
'    vValue = chkSystem.Value
'    Call SaveSetting("Search System Folders", CStr(vValue))
'    vValue = chkZips.Value
'    Call SaveSetting("Search Zip Folders", CStr(vValue))
'    Call SaveSetting("Last Location", txtFolder.Text)
'    vValue = lstFind.View
'    Call SaveSetting("View Type", CStr(vValue))
    'Set cFind = Nothing 'Unset the Search Object
'    End
    'Alot of people say don't use End, this is BULLSHIT.
    'End stops the process in its tracks, and unloads all variables from memory.
    'One thing you must do before using this, is to unset all the objects you may have loaded, otherwise memory leaks will occur.
End Sub
Private Sub lstFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'This event is fired when a user clicks on one of the column headers.
    If ColumnHeader.index = 3 Then
        'If they click on the Size column, we need to set the SortKey to the raw size column.
        'Example:
        
        '1      MB
        '100    KB
        '2      MB
        '200    Bytes
        
        'I think thats right anyway, but thats how it would be sorted,
        'But our hidden column contains...
        
        '000002097152 ( 2 MB )
        '000001048576 ( 1 MB )
        '000000102400 ( 100 KB )
        '000000000200 ( 200 Bytes )
        'Now do you see the method to my madness ? =D
        
        lstFind.SortKey = 3
    Else
        'Otherwise set the SortKey to the column clicked.
        lstFind.SortKey = ColumnHeader.index - 1
    End If
    lstFind.Sorted = True 'Sort it.
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending) 'SortOrder = Not SortOrder
End Sub
Private Sub lstFind_DblClick()
Dim chemin As String
If lstFind.ListItems.Count = 0 Then Exit Sub
'Dim fichier As String
chemin = Mid$(lstFind.SelectedItem.ListSubItems(1).Text, 1, Len(lstFind.SelectedItem.ListSubItems(1).Text) - Len(lstFind.SelectedItem.Text) - 1)

    If GetKeyState(vbKeyShift) < 0 Then 'shift appuyer
        If GetKeyState(vbKeyControl) < 0 Then 'touche control
                If LCase(Right$(lstFind.SelectedItem.Text, 4)) = ".lnk" Then chemin = GetTarget(chemin & "\" & lstFind.SelectedItem.Text)
                If File_folder_Exists(chemin) = True Then ', vbArchive + vbHidden + vbNormal + vbNormal + vbSystem
                    Shell "explorer.exe /select, " & chemin, vbNormalFocus
                'ElseIf File_folder_Exists(chemin) = True Then ', vbDirectory + vbSystem + vbHidden + vbSystem + vbHidden
                '    Shell "explorer.exe /select, " & chemin, vbNormalFocus
                End If
        Else
            'Call ShellExecute(0, "", lstFind.SelectedItem.ListSubItems(1).Text, vbNullString, vbNullString, 1)  'Shell Execute the selected Item.
            Shell "explorer.exe /select, " & chemin & "\" & lstFind.SelectedItem.Text, vbNormalFocus
        End If
    Else
        Call ShellExecute(0, "", lstFind.SelectedItem.ListSubItems(1).Text, vbNullString, vbNullString, 1)  'Shell Execute the selected Item.
    End If
End Sub
Private Sub lstFind_ItemClick(ByVal item As MSComctlLib.ListItem)
Exit Sub
'    StatusBar1.Panels(1).Text = "In Folder " & item.ListSubItems(1) & "; Type " & item.ListSubItems(4) & "; Date Modified " & item.Tag & "; Size " & item.ListSubItems(2)
    'Set the Status Bar Panel to hold the info of the current clicked item.
End Sub

Private Sub lstFind_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then lstFind_DblClick

End Sub


Private Sub lstFind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    On Error Resume Next 'We need this to know if an item is selected.
'    If Button = 2 Then
'        mnProperties.Visible = False 'Set both of the Item independant menu items to not visibile.
'        mnOpenFolder.Visible = False
'        If Not Len(lstFind.SelectedItem.Text) = 0 Then 'If no item is selected, this will error, but we are resuming so this should be False
'            'But if by some miracle, the above statement ends up true, with no item selected, it will have errored anyway.
'            'So, if no error, then we have an item.
'            If Err.Number = 0 Then
'                'So show the menu items needed !
'                mnProperties.Visible = True
'                mnOpenFolder.Visible = True
'            End If
'        Else
'            'Otherwise reset them to false anyway.
'            mnProperties.Visible = False
'            mnOpenFolder.Visible = False
'        End If
'        PopupMenu mnRClick 'Show the popupmenu.
'    End If
End Sub

Private Sub pic_tri_Click()
    lstFind.SortKey = 0 'ColumnHeader.Index - 1
    lstFind.Sorted = True 'Sort it.
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending) 'SortOrder = Not SortOrder

End Sub

Private Sub Picture1_Click()
Check1.Value = IIf(Check1.Value = 0, 1, 0)
End Sub

Private Sub pic_close_Click()
Unload Me
End Sub

Private Sub pic_search_Click()
           
           If Trim$(text1) = "" Then Beep: Exit Sub
           
           text1 = Replace(text1, "*", "")
           
           text1.SetFocus
           
           ItemCount = 0
           
            Screen.MousePointer = 13
            rech_en_cours = True
            rech_cancel = False
            'LockWindowUpdate (lstFind.hwnd)

            
            lstFind.ListItems.Clear
            lstFind.SmallIcons = Nothing
            'lstFind.Icons = imgNull
            'iml32.ListImages.Clear 'Now that they aren't in use, clear all the Icons from both Image Lists
            iml16.ListImages.Clear
            iml16.ListImages.Add , "a0", Me.Icon
            lstFind.SmallIcons = iml16
            lstFind.Sorted = False
            
            If txtFolder = "" Then
                If GetKeyState(vbKeyShift) < 0 Then 'shift appuyer
                    Call init_recherche(False)
                Else
                    Call init_recherche(True)
                End If
            Else
                ItemCount = 0
                List1.Clear
                Me.lstFind.ListItems.Clear ': List2.Clear
                If GetKeyState(vbKeyShift) < 0 Then 'shift appuyer
                    Call mon_search(txtFolder, True, "", False)
                Else
                    Call mon_search(txtFolder, True, "", True)
                End If
            End If
            'LockWindowUpdate (0&)
            
            
            
            Screen.MousePointer = 0
            If rech_cancel = False Then
                Label1.Caption = "Recherche Terminé : " & lstFind.ListItems.Count & " élément(s) trouvé(s)"
            Else
                Label1.Caption = "Recherche Annulé : " & lstFind.ListItems.Count & " élément(s) trouvé(s)"
            End If
            
            rech_en_cours = False: rech_cancel = False
            
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        pic_search_Click
    End If
    
        
            
End Sub


Sub DrawIcon(path) ', Index, dossier As Boolean, selection As Boolean, Optional blt = True)

Dim hImgLarge&
  
    '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
'        'utiliser le cache :)
'    pb rencontrer : la couleur est sauver avec l'icone donc au changement du skin, badabom
'    c la mouvaise couleur, c normal et puis ya pas vraiement un gain au niveau de la vitesse
'    au contraire, mm le save des icons fait clignoter le témoin du disque dur
'    donc je désactive le cache :)
    '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
'Dim CurX As Integer

hImgLarge& = SHGetFileInfo(path, 0&, ShInfo, Len(ShInfo), _
BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

  pic16.Picture = Nothing
    ImageList_Draw hImgLarge&, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT
'              '2_l'ajouter dans une imagelist temporaire (pour garder la transparence)
    'iml16.ListImages.Add iml16.ListImages.Count + 1, "a" & iml16.ListImages.Count + 1, pic16.Image
     
    iml16.ListImages.Add , , pic16.Image

    'ImageList_Draw hImgLarge&, ShInfo.iIcon, picItem(Index).hDc, 4, 3, ILD_TRANSPARENT
            
            
'si raccourci donc ajouter une petite flèche pour destinguer les racc des autres fichiers
'    If dossier = False And UCase(Right$(path, 3)) = "LNK" Then
''        'DrawIcon_from_imagelist 6, Index & 0, False
'        If selection And frm_lecteurs.anim_ico Then
'            Call PaintStandardPicture(picItem(Index).hDc, frm_search_AppBar.ma_liste_ico.ListImages(6).ExtractIcon, 3, 2, 16, 16)
'        Else
'            Call PaintStandardPicture(picItem(Index).hDc, frm_search_AppBar.ma_liste_ico.ListImages(6).ExtractIcon, 4, 3, 16, 16)
'        End If
'    End If
  
End Sub


Sub init_recherche(rech_ds_sous_dossier As Boolean)
    
    Dim i As Integer

    Dim sPath As String
    Dim IDL As Long
    Dim strPath As String
    Dim lngPos As Long
    'Dim recherche_ds_bureau_perso As Boolean
    
    'pour lancer la recherche ds un dossier qui existe ailleurs mais qui a son raccourcis sur le bureau
    Dim bureau_chemin As String
    Dim dossier_cible As String
    
    
    ItemCount = 0
    List1.Clear
    
    Me.lstFind.ListItems.Clear ': List2.Clear
    
    
    'GetCursorPos pt
    'posx = pt.X * Screen.TwipsPerPixelX
    'posy = pt.Y * Screen.TwipsPerPixelY
    'show_menu_principal posx, posy ', dossier_principal

    'Set frm_mn2.frm_appelant = Me
    'Set frm_mn2.imagelistico = frm_search.ImageList1
    
    
    
    
''        'If ItemCount > 0 Then
'            ItemCount = ItemCount + 1
'            frm_mn2.addmenu frm_mn2, ItemCount - 1, "", "", 0, True
''        'End If

    
    
        'user
        If SHGetSpecialFolderLocation(0, fldSTARTMENU, IDL) = NOERROR Then
            sPath = String$(255, 0)
            SHGetPathFromIDListA IDL, sPath
            lngPos = InStr(sPath, Chr(0))
            If lngPos > 0 Then
                strPath = Left$(sPath, lngPos - 1)
            End If
        End If
        'Call mon_search(strPath, False)
        Call mon_search(strPath, True, "Menu Démarrer (USER)", rech_ds_sous_dossier)
        
        
        
        '2-all users
        If SHGetSpecialFolderLocation(0, fldCOMMON_STARTMENU, IDL) = NOERROR Then
            sPath = String$(255, 0)
            SHGetPathFromIDListA IDL, sPath
            lngPos = InStr(sPath, Chr(0))
            If lngPos > 0 Then
                strPath = Left$(sPath, lngPos - 1)
            End If
        End If
        Call mon_search(strPath, True, "Menu Démarrer (All USERS)", rech_ds_sous_dossier)
        
       
        
        '3-desktop
        If SHGetSpecialFolderLocation(0, fldDESKTOPDIRECTORY, IDL) = NOERROR Then
            sPath = String$(255, 0)
            SHGetPathFromIDListA IDL, sPath
            lngPos = InStr(sPath, Chr(0))
            If lngPos > 0 Then
                strPath = Left$(sPath, lngPos - 1)
            End If
        End If
        
        
        bureau_chemin = strPath
        Call mon_search(strPath, True, "Bureau (USER)", rech_ds_sous_dossier)
        
        'If Trim$(dossier_principal) <> Trim$(strPath) Then recherche_ds_bureau_perso = True
        
        
        '4-desktop All Users
        If SHGetSpecialFolderLocation(0, fldCOMMON_DESKTOPDIRECTORY, IDL) = NOERROR Then
            sPath = String$(255, 0)
            SHGetPathFromIDListA IDL, sPath
            lngPos = InStr(sPath, Chr(0))
            If lngPos > 0 Then
                strPath = Left$(sPath, lngPos - 1)
            End If
        End If
        
        
        Call mon_search(strPath, True, "Bureau (ALL USERS)", rech_ds_sous_dossier)
        
        
        
        'spécial traitement ////////////////////
            'si sur le bureau user y a des raccourcis vers un dossier
            'lancer la recherche ds ces dossiers aussi
            File2.path = bureau_chemin
            File2.Pattern = "*.lnk"
            File2.Refresh
            If File2.ListCount <> 0 Then
                For i = 0 To File2.ListCount - 1
                    dossier_cible = GetTarget(bureau_chemin & "\" & File2.List(i))
                    If folder_exists(dossier_cible) Then
                        Call mon_search(dossier_cible, True, "Depuis lien (Raccourcis) sur le bureau", rech_ds_sous_dossier)
                    End If
                Next i
            End If
        '///////////////////////////////////////
        
        
    If ItemCount = 0 Then
        For i = 1 To 5
            Beep
        Next i
    End If

    rech_en_cours = False
    
End Sub


Private Function folder_exists(sFullPath As String) As Boolean
    'Dim oFile As New Scripting.FileSystemObject
    'FileExists = oFile.FileExists(sFullPath)
    
    
'Dim fso As FileSystemObject
'Dim sFilePath As String
'Set fso = New FileSystemObject
''sFilePath = "c:\test.txt"
'If fso.FileExists(sFullPath) Then
'    FileExists = True
'Else
'    FileExists = False
'End If


    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("Scripting.FileSystemObject") 'CreateObject("WScript.FileSystemObject")
    'Set wshLink = wshShell.FileExists(strPath)
    'File_folder_Exists = wshShell.FileExists(sFullPath)
    
    'vérifie si c un dossier
    'If File_folder_Exists = False Then
        folder_exists = wshShell.FolderExists(sFullPath)
    'End If
    'GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
End Function

Private Function GetTarget(strPath As String) As String
    'Gets target path from a shortcut file
    On Error GoTo Error_Loading
    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(strPath)
    'récupérer le chemin du fichier source depuis le raccourcis
    GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
    'vérifie si le fichier source exist
    If File_folder_Exists(GetTarget) = False Then GetTarget = "" 'source raccourcis introuvable

    
    Exit Function

Error_Loading:
    'GetTarget = "Erreur : " & Err.Description
    'End
    'MsgBox Err.Number & Err.Description
    If err.Number = -2147352567 Then 'Le chemin du raccourci doit finir par .lnk ou .url.
        'c pas un raccourcis donc envoyer le mm chemin
        GetTarget = strPath
    Else
        GetTarget = ""
    End If
    
    Exit Function
End Function

Private Function File_folder_Exists(sFullPath As String) As Boolean
    'Dim oFile As New Scripting.FileSystemObject
    'FileExists = oFile.FileExists(sFullPath)
    
    
'Dim fso As FileSystemObject
'Dim sFilePath As String
'Set fso = New FileSystemObject
''sFilePath = "c:\test.txt"
'If fso.FileExists(sFullPath) Then
'    FileExists = True
'Else
'    FileExists = False
'End If


    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("Scripting.FileSystemObject") 'CreateObject("WScript.FileSystemObject")
    'Set wshLink = wshShell.FileExists(strPath)
    File_folder_Exists = wshShell.FileExists(sFullPath)
    
    'vérifie si c un dossier
    If File_folder_Exists = False Then
        File_folder_Exists = wshShell.FolderExists(sFullPath)
    End If
    'GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
End Function

Sub mon_search(rechercherdans As String, ajout_seperateur_avant As Boolean, separateur_menu_txt As String, rech_ds_sous_dossier As Boolean)

'On Error GoTo err_system

    'StatusBar1.Refresh

    'On Error GoTo bere
    Dim num As Long
    Dim a As Integer ', i As Integer
    Dim lv As ListItem
    Dim lvSI As ListSubItem
    Dim path As String
    'Me.tag = "Start"
'*************************************
    Dim i As Long
    Dim separateur_ajouter As Boolean
    
    'initialisation
    ''ItemCount = 0
    'List2.Clear

    If folder_exists(rechercherdans) = False Then Exit Sub
    'If Dir$(rechercherdans, vbDirectory Or vbSystem Or vbHidden) = "" Then Exit Sub
    Dir1.path = rechercherdans
    
    '== RESET PAST RESULTS =======
    'ListView1.ListItems.Clear
    List1.Clear
    'lstFind.ListItems.Clear
    '=============================



    '== ADD CURRENT DIRECTORY TO DIRECOTRY TO SEARCH =====
    List1.AddItem Dir1.path
    List1.ListIndex = 0
    '==============================================


ere:
    
    
    'If Me.tag = "Stop" Then GoTo bere
    'Me.Caption = "Search " & ListView1.ListItems.Count & " Results"

    '== GO TO THE NEXT PATH TO SEARCH
    If Dir$(List1.Text, vbDirectory Or vbHidden Or vbSystem) <> "" Then
    
        Label1.Caption = "Recherche en cours : - > " & List1.Text
        Dir1.path = List1.Text
        '===========================
        For a = 0 To Dir1.ListCount - 1
            
            'touche escape
            If GetKeyState(vbKeyEscape) < 0 Then rech_cancel = True: Exit Sub 'shift appuyer

            List1.AddItem Dir1.List(a)
            ''ajouter dossier
            If InStr(1, LCase(ExtractFileName(Dir1.List(a))), LCase(text1.Text)) > 0 Then
                
                
                ItemCount = ItemCount + 1
                'ajout_element Dir1.path, ExtractFileName(Dir1.List(a)), ItemCount     'i ', 3 'chemin_menu_user & "\" &
                'ItemCount = ItemCount + 1
                'frm_mn2.addmenu frm_mn2, ItemCount - 1, Dir1.List(a), "shell", 0, False
                DrawIcon Dir1.List(a) ', Index, False, False
                'lstFind.ListItems.Add , "k" & ItemCount, Dir1.List(a), , iml16.ListImages.Count   ', "shell", 0, False
                lstFind.ListItems.Add ItemCount, , ExtractFileName(Dir1.List(a)), , iml16.ListImages.Count  ', "shell", 0, False
                lstFind.ListItems(ItemCount).SubItems(1) = Dir1.List(a)
                
                If LCase(Right$(lstFind.ListItems(ItemCount).Text, 4)) = ".lnk" Then
                    lstFind.ListItems(ItemCount).ForeColor = vbRed
                    lstFind.ListItems(ItemCount).ListSubItems(1).ForeColor = vbRed
                End If
                
            End If
            DoEvents
        Next a
        '============================================================
    
        '== SERACH ALL FILES FOR SERACH STRING IF FOUND ADD TO LISTVIEW
        For i = 0 To File1.ListCount - 1
            
            'touche escape
            If GetKeyState(vbKeyEscape) < 0 Then rech_cancel = True: Exit Sub 'shift appuyer
            
            
            If InStr(1, LCase(File1.List(i)), LCase(text1.Text)) > 0 Then
                'Set lv = ListView1.ListItems.Add(, , File1.List(i), 2, 2)
                'lv.ListSubItems.Add , , Dir1.path
                'ch = LCase(ListView1.ListItems(i).ListSubItems(1).Text & "\" & ListView1.ListItems(i))  '.Text)
                'If InStr(ch, Text3) <> 0 Then
                    '-------------------
                    
                    
                    ItemCount = ItemCount + 1
                    'ajout_element Dir1.path, File1.List(i), ItemCount     'i ', 3 'chemin_menu_user & "\" &
                    'frm_mn2.addmenu frm_mn2, ItemCount - 1, Dir1.Path & "\" & File1.List(i), "shell", 0, False
                    If Right$(Dir1.path, 1) = "\" Then
                        path = Dir1.path & File1.List(i)
                    Else
                        path = Dir1.path & "\" & File1.List(i)
                    End If
                    DrawIcon path
                    lstFind.ListItems.Add ItemCount, , ExtractFileName(path), , iml16.ListImages.Count  ', "shell", 0, False
                    lstFind.ListItems(ItemCount).SubItems(1) = path 'Dir1.path & "\" & File1.List(i)
                    'ListView1.ListItems.item(ItemCount).SubItems(2) = Trim$(Mid$(txt, Pos + 1, 100))
                    
                    If LCase(Right$(lstFind.ListItems(ItemCount).Text, 4)) = ".lnk" Then
                        lstFind.ListItems(ItemCount).ForeColor = vbRed
                        lstFind.ListItems(ItemCount).ListSubItems(1).ForeColor = vbRed
'                        For intIndex = 1 To lstFind.ColumnHeaders.Count - 1
'                            Set lvSI = lstFind.ListItems(ItemCount).ListSubItems(intIndex)
'
'                            lvSI.ForeColor = vbRed 'RowColor
'                        Next
                    End If
                    
                    '-------------------
                'End If
                'DoEvents
            End If
            DoEvents
        Next i
        '=====================================

    End If

    '== TO SET NEXT PATH  TO CHANGE TO
    'si pas de rech ds sous dossier alors quitter
    If rech_ds_sous_dossier = False Then Exit Sub
    
    If List1.ListIndex + 1 >= List1.ListCount Then
        GoTo bere
    Else
        List1.ListIndex = List1.ListIndex + 1
    End If
    GoTo ere
    '================================


    '==SEARCH IS NOW COMPLETE
bere:
    'MsgBox "COMPLETE"
    '==============================

    'old_nbre_element_menu = List2.ListCount
Exit Sub

err_system:

MsgBox "Erreur :" & err.Number & " - " & err.Description, vbCritical, App.FileDescription

End Sub



Private Function ExtractFileName(ByVal strPath As String) As String

Dim Pos As Integer

On Error GoTo err1
  ' StrReverse is only working in VB6
    Pos = InStr(1, strPath, "\")
    If Pos <> 0 Then
        strPath = StrReverse(strPath)
        strPath = Left(strPath, InStr(strPath, "\") - 1)
        ExtractFileName = StrReverse(strPath)
    Else
        ExtractFileName = Trim$(strPath)
    End If
    
Exit Function
err1:
  ExtractFileName = Trim$(strPath)
  Exit Function
  
End Function


