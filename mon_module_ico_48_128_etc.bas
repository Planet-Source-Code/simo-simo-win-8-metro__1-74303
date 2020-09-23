Attribute VB_Name = "mon_module_ico_48_128_etc"
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Public Const DI_MASK = &H1&
Public Const DI_IMAGE = &H2&
Public Const DI_NORMAL = &H3&
Public Const DI_COMPAT = &H4&
Public Const DI_DEFAULTSIZE = &H8&

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Const LR_LOADMAP3DCOLORS = &H1000
    Public Const LR_LOADFROMFILE = &H10
    Public Const LR_LOADTRANSPARENT = &H20
'Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
'Public Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
'Public Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
'Public Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Public Const ILD_TRANSPARENT = 1&
Public Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage  As Long
    xExt    As Long
    yExt    As Long
End Type
Public Type Guid
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
    ' This is all magic if you ask me:
    Dim NewPic As Picture, PicConv As PictDesc, IGuid As Guid
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect PicConv, IGuid, True, NewPic
    Set IconToPicture = NewPic
End Function


