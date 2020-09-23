Attribute VB_Name = "ApiDibs"
'ApiDibs.bas

'Unless otherwise defined:
Option Base 1  ' Arrays base 1
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings


'------------------------------------------------------------
' API used to check patching

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'------------------------------------------------------------
' API used fir shifting paint point

Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long


Public Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long

'res = GetCursorPos(LP)
'res = SetCursorPos(LP.x,LP.y)

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public LP As POINTAPI

' -----------------------------------------------------------
' API to Fill background, creating masks

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1
' --------------------------------------------------------------

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' --------------------------------------------------------------
' Function & constants to make Window stay on top

'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
'ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
'ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

'Public Const hWndInsertAfter = -1
'Public Const wFlags = &H40 Or &H20

' -----------------------------------------------------------
'  This required instead of Screen.Height & Width for resizing

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen

' -----------------------------------------------------------
Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function SetPixelV Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

' -----------------------------------------------------------
' APIs for getting DIB bits to PicMem

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long

'------------------------------------------------------------------------------
'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

'------------------------------------------------------------------------------
' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring aDrawing in an array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'wUsage is one of:-
Public Const DIB_PAL_COLORS = 1 '  uses system....
Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'dwRop is vbSrcCopy

'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' Publics

Public bred As Byte, bgreen As Byte, bblue As Byte

' Sized :-
Public bP1M() As Byte '4, PICW, PICH)
Public bP2M() As Byte '4, PICW, PICH)
Public bPPM() As Byte '4, PICW, PICH)

' Picture details
Public PICW, PICH
Public ImageBytes
Public ScanLineBytes

' Paths
Public PathSpec$, CurrPath$
' Loaed file
Public FileSpec$

' Operation parameters
Public Threshold, ClumpIndex
Public FilterParam

' Booleans
Public aMagON As Boolean
Public aSelectShapeON As Boolean
Public aSelectionMade As Boolean
Public aDrawing As Boolean
Public aCulPickerON As Boolean
Public aDeSpotON As Boolean
Public aDeSpotterStarted As Boolean
Public aPaintON As Boolean
Public aPatchON As Boolean
Public aPatchFilled As Boolean
Public aClickON As Boolean
Public aColorFormON As Boolean
'---------------------------------

' Misc
Public PickedCul     ' From CulPicker
Public PaintSize     ' 1 pix or 2x2 pix
Public SLCul         ' Color of selected shape lines
Public SLWidth       ' Select line thickness 1 or 2 pixels
Public NSLines       ' Number of shape lines
Public ixTL, iyTL, ixBR, iyBR ' picPatchMask mask bounding rect
Public ixSL0, iySL0           ' Starting coords of shape lines


'Public bP1M() As Byte  ' Holds source picture
'Public bP2M() As Byte  ' Holds dest picture
'Public bPPM() AS Byte  ' Hold masks

Public Sub FillBMPStruc(ByVal mwidth, ByVal mheight)
  
  With bm.bmiH
   .biSize = 40
   .biwidth = mwidth
   .biheight = mheight
   .biPlanes = 1
   .biBitCount = 32     ' BGRA
   .biCompression = 0
   
   ' Ensure expansion to 4B boundary
   ScanLineBytes = (((mwidth * .biBitCount) + 31) \ 32) * 4

   .biSizeImage = ScanLineBytes * Abs(.biheight)
   .biXPelsPerMeter = 0
   .biYPelsPerMeter = 0
   .biClrUsed = 0
   .biClrImportant = 0
 End With

End Sub

Public Sub GETDIBS(ByVal PICIM As Long, picnum)

' PICIM = pic1.Image or pic2.Image or picPatchMask.Image
' picnum = 1 for pic1, 2 for pic2, 3 for picPatchMask

On Error GoTo DIBError

'Get info on picture loaded into PIC
GetObjectAPI PICIM, Len(bmp), bmp
' which fills:-
'Public Type BITMAP
'   bmType As Long              ' Type of bitmap
'   bmWidth As Long             ' Pixel width
'   bmHeight As Long            ' Pixel height
'   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
'   bmPlanes As Integer         ' Color depth of bitmap
'   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
'   bmBits As Long              ' This is the pointer to the bitmap data  !!!
'End Type
'Public bmp As BITMAP


NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

FillBMPStruc bmp.bmWidth, bmp.bmHeight
' which fills:-
' Structures for StretchDIBits
'Public Type BITMAPINFOHEADER ' 40 bytes
'   biSize As Long
'   biwidth As Long
'   biheight As Long
'   biPlanes As Integer
'   biBitCount As Integer
'   biCompression As Long
'   biSizeImage As Long
'   biXPelsPerMeter As Long
'   biYPelsPerMeter As Long
'   biClrUsed As Long
'   biClrImportant As Long
'End Type

' Set pic mems to receive color bytes or indexes or bits
PICW = bmp.bmWidth
PICH = bmp.bmHeight
If picnum = 1 Then
   ReDim bP1M(4, PICW, PICH)   ' pic1
   ' Load color bytes to bP1M
   ret = GetDIBits(NewDC, PICIM, 0, PICH, bP1M(1, 1, 1), bm, 1)
ElseIf picnum = 2 Then
   ReDim bP2M(4, PICW, PICH)   ' pic2
   ' Load color bytes to bP2M
   ret = GetDIBits(NewDC, PICIM, 0, PICH, bP2M(1, 1, 1), bm, 1)
Else
   ReDim bPPM(4, PICW, PICH)   ' picPatchMask
   ' Load color bytes to bPPM
   ret = GetDIBits(NewDC, PICIM, 0, PICH, bPPM(1, 1, 1), bm, 1)
End If

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

'PtrP1M = VarPtr(bP1M(1, 1, 1))
ImageBytes = ScanLineBytes * PICH

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS"
  DoEvents
  Unload Form1
  End
End Sub

Public Sub LngToRGB(LCul)
'Convert Long Colors() to RGB components
bred = (LCul And &HFF&)
bgreen = (LCul And &HFF00&) / &H100&
bblue = (LCul And &HFF0000) / &H10000
End Sub

