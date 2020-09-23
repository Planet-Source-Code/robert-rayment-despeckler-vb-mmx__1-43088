VERSION 5.00
Begin VB.Form MagForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "RRMag"
   ClientHeight    =   3240
   ClientLeft      =   105
   ClientTop       =   135
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H00C0C0FF&
      Caption         =   "X"
      Height          =   210
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Hide"
      Top             =   2970
      Width           =   270
   End
   Begin VB.PictureBox PicCul 
      Height          =   195
      Left            =   1140
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   4
      Top             =   2970
      Width           =   585
   End
   Begin VB.OptionButton OptMag 
      BackColor       =   &H00FFC0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2985
      Width           =   345
   End
   Begin VB.OptionButton OptMag 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2985
      Width           =   360
   End
   Begin VB.OptionButton OptMag 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2985
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   60
      ScaleHeight     =   192
      ScaleMode       =   0  'User
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   45
      Width           =   2880
   End
   Begin VB.Label LabRGB 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabRGB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1740
      TabIndex        =   5
      ToolTipText     =   "R G B"
      Top             =   2970
      Width           =   930
   End
End
Attribute VB_Name = "MagForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MagForm.frm  by  Robert Rayment

' Self contained magnifier form

' Magnification shown around cursor
' Buttons for magnification 2,4 or 8
' Move Mouse Down on picture box to move headerless form

' Color at Cursor Hot X,Y

' The API StretchBits is simpler, than the method here, but will
' not do the edge of the screen.

Option Base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings

               
' ## API's & STRUCTURES #######################################

'  Windows API to make application stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H20

'------------------------------------------------------------------------------
' Windows API's to get color from anywhere ( see AlphaSpy by nitrix on PSC)

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'------------------------------------------------------------------------------
' Structures for StretchDIBits
Private Type BITMAPINFOHEADER ' 40 bytes
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

Private Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'bmiC As RGBTRIPLE            'NB Palette NOT NEEDED for 24 & 32-bit
End Type
Private bm As BITMAPINFO

' For transferring an array to Form or Picture Box
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal SrcW As Long, ByVal SrcH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

' Constants for StretchDIBits
Private Const DIB_PAL_COLORS = 1
'Const SRCCOPY = &HCC0020
'Const SRCINVERT = &H660046
'-----------------------------------------------------------------

Dim PicSize       ' Picture1 size
Dim LongSize      ' LongSurf() size
Dim LongSurf()    ' Display surface
Dim Mag           ' Magnification 2,4 or 8

Dim NewPos As POINTAPI  ' Mouse screen position
Dim OldPos As POINTAPI  ' To check if mouse has moved

Dim redb As Byte, greenb As Byte, blueb As Byte 'RGB byte components
Dim Topy, LeftX      ' To move headerless form
Dim MagDone As Boolean



Private Sub Form_Load()

If App.PrevInstance Then End  ' Only allow it to run once.

'------------------------------------------------
' Size and position Picture1
PicSize = 192 '128   ' Can try other sizes BUT must be a multiple of 8
   
   With Picture1
      .Top = 4
      .Left = 4
      .Height = PicSize
      .Width = PicSize
      .AutoRedraw = True
   End With


'------------------------------------------------

' Size & Make application stay on top
fTop = 164
fLeft = 80
fWidth = PicSize + 8 ' 192+8=200 (3000)for a bit of border
fHeight = PicSize + 24 ' 192+24+22=238(3570) for buttons and a bit of border
ret& = SetWindowPos(MagForm.hwnd, hWndInsertAfter, fLeft, fTop, fWidth, fHeight, wFlags)
'------------------------------------------------


Show


' Set Initial Mag to 2 and size LongSurf()
OptMag_MouseDown 0, 1, 0, 0, 0

'---------------------------------------------------

'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biSize = 40
bm.bmiH.biwidth = LongSize
bm.bmiH.biheight = LongSize
bm.bmiH.biPlanes = 1
bm.bmiH.biBitCount = 32 '24 '8
bm.bmiH.biCompression = 0
bm.bmiH.biSizeImage = 0 ' Not needed
bm.bmiH.biXPelsPerMeter = 0
bm.bmiH.biYPelsPerMeter = 0
bm.bmiH.biClrUsed = 0
bm.bmiH.biClrImportant = 0
'---------------------------------------------------

' Dummy old mouse coords
OldPos.X = 0
OldPos.Y = 0

MagDone = False

' =========   ACTION LOOP ======================

Do

   Call GetCursorPos(NewPos)
   
   
   ' Only re-scan if mouse moved
   If NewPos.X <> OldPos.X Or NewPos.Y <> OldPos.Y Then
    
      OldPos = NewPos      ' Save new coords
      
      ' Get rectangle to scan
      ixyoff = LongSize \ 2 - 1
      ix0 = NewPos.X - ixyoff
      iy0 = NewPos.Y - ixyoff
      ix1 = ix0 + LongSize - 1
      iy1 = iy0 + LongSize - 1
 
      rDC = GetDC(0&)   ' Get Device Context to whole screen
 
      ' Fill LongSurf(kx,ky) with colors
      ky = 1
      For iy = iy1 To iy0 Step -1   ' Need to switch y for LongSurf()
      If iy = NewPos.Y Then iym = ky
      kx = 1
         For ix = ix0 To ix1
            If ix = NewPos.X Then ixm = kx
            LongCul = GetPixel(rDC, ix, iy)
            CulToRGB LongCul, redb, greenb, blueb  ' Get RGB components
            LongSurf(kx, ky) = RGB(blueb, greenb, redb)  ' Need BGR for LongSurf()
            kx = kx + 1
         Next ix
         ky = ky + 1
      Next iy
 
      LongCul = GetPixel(rDC, NewPos.X, NewPos.Y)
      
      CulToRGB LongCul, redb, greenb, blueb  ' Get RGB components
      
      ReleaseDC 0&, rDC    ' Important to avoid using up resources
   
      LongSurf(ixm, iym) = RGB(0, 0, 0)
      
      ' Display LongSurf()
      ShowLongSurfONCE
      
      
      If LongCul >= 0 Then PicCul.BackColor = LongCul  '<<<<<
      
      LabRGB.Caption = Format(Str$(redb), " ##0") & Format(Str$(greenb), " ##0") & Format(Str$(blueb), " ##0")
   
   End If

   DoEvents    ' To allow other controls

Loop Until MagDone

' ==============================================

End Sub

' ########  Stretch LongSurf() to Picture1 ###############

Public Sub ShowLongSurfONCE()

Picture1.Cls

ptLS = VarPtr(LongSurf(1, 1)) 'Pointer to long surface

   succ& = StretchDIBits(Picture1.hdc, _
   0, 0, _
   PicSize, PicSize, _
   0, 0, _
   LongSize, LongSize, _
   ByVal ptLS, bm, _
   DIB_PAL_COLORS, vbSrcCopy)

End Sub

' ####### CHANGE MAGNIFICATION & LongSurf() ################

Private Sub OptMag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

OptMag(Index) = True
Select Case Index
Case 0: Mag = 2
Case 1: Mag = 4
Case 2: Mag = 8
End Select

LongSize = PicSize \ Mag

' Set varying parameters for StretchDIBits
bm.bmiH.biwidth = LongSize
bm.bmiH.biheight = LongSize

' Resize LongSurf() to pick-up screen colors
ReDim LongSurf(LongSize, LongSize)

End Sub


' #######  GET RGB COMPONENTS #########

Private Sub CulToRGB(LongCul&, re As Byte, gr As Byte, bl As Byte)
' Input LongCul&:  Output: R G B components

re = LongCul& And &HFF&
gr = (LongCul& And &HFF00&) / &H100&
bl = (LongCul& And &HFF0000) / &H10000

End Sub

' ########  MOVING HEADERLESS FORM #########################

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Topy = Y
LeftX = X

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 0 Then
   Me.Top = Me.Top + Y - Topy
   Me.Left = Me.Left + X - LeftX
End If

End Sub

' ###### EXIT ##############

Private Sub cmdHide_Click()
   Picture1.SetFocus
   MagForm.Hide
   aMagON = False
End Sub

'Private Sub cmdExit_Click()
'Form_Unload 0

'MagDone = True
'DoEvents
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'aMagON = True
'Unload Me
'End Sub

