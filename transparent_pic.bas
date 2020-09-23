Attribute VB_Name = "Transparent"
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Sub TransparentBlt(dsthDC As Long, srchDC As Long, X As Integer, Y As Integer, Width As Integer, Height As Integer, TransColor As Long)
    Dim maskDC As Long      'DC for the mask
    Dim tempDC As Long      'DC for temporary data
    Dim hMaskBmp As Long    'Bitmap for mask
    Dim hTempBmp As Long    'Bitmap for temporary data
    'First create some DC's. These are our gateways to assosiated bitmaps in RAM
    maskDC = CreateCompatibleDC(dsthDC)
    tempDC = CreateCompatibleDC(dsthDC)
    'Then we need the bitmaps. Note that we create a monochrome bitmap here!
    'this is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dsthDC, Width, Height)
    '..then we can assign the bitmaps to the DCs
    hMaskBmp = SelectObject(maskDC, hMaskBmp)
    hTempBmp = SelectObject(tempDC, hTempBmp)
    'Now we can create a mask..First we set the background color to the
    'transparent color then we copy the image into the monochrome bitmap.
    'When we are done, we reset the background color of the original source.
    TransColor = SetBkColor(srchDC, TransColor)
    BitBlt maskDC, 0, 0, Width, Height, srchDC, 0, 0, vbSrcCopy
    TransColor = SetBkColor(srchDC, TransColor)
    'The first we do with the mask is to MergePaint it into the destination.
    'this will punch a WHITE hole in the background exactly were we want the
    'graphics to be painted in.
    BitBlt tempDC, 0, 0, Width, Height, maskDC, 0, 0, vbSrcCopy
    BitBlt dsthDC, X, Y, Width, Height, tempDC, 0, 0, vbMergePaint
    'Now we delete the transparent part of our source image. To do this
    'we must invert the mask and MergePaint it into the source image. the
    'transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, Width, Height, maskDC, 0, 0, vbNotSrcCopy
    BitBlt tempDC, 0, 0, Width, Height, srchDC, 0, 0, vbSrcCopy
    BitBlt tempDC, 0, 0, Width, Height, maskDC, 0, 0, vbMergePaint
    'Both target and source are clean, all we have to do is to AND them together!
    BitBlt dsthDC, X, Y, Width, Height, tempDC, 0, 0, vbSrcAnd
    'Now all we have to do is to clean up after us and free system resources..
    DeleteObject (hMaskBmp)
    DeleteObject (hTempBmp)
    DeleteDC (maskDC)
    DeleteDC (tempDC)
End Sub


