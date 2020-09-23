VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Basic Despeckler by Robert Rayment"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11625
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   Begin VB.Frame fraBrightness 
      BackColor       =   &H00008000&
      Height          =   945
      Left            =   4920
      TabIndex        =   42
      Top             =   7275
      Width           =   2985
      Begin VB.CommandButton cmdBrightnessExit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
         Height          =   300
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   510
         Width           =   330
      End
      Begin VB.PictureBox picLevelBar 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   225
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   43
         Top             =   495
         Width           =   2000
      End
      Begin VB.Label LabLevel 
         BackColor       =   &H00808000&
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   270
         TabIndex        =   45
         Top             =   270
         Width           =   1545
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   660
         Left            =   120
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.PictureBox picPatchMask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   10440
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   41
      Top             =   1065
      Width           =   480
   End
   Begin VB.PictureBox picPatch 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   10440
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   300
      Width           =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Height          =   435
      Left            =   7185
      TabIndex        =   31
      Top             =   75
      Width           =   1710
      Begin VB.Label LabTim 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ms"
         Height          =   255
         Left            =   735
         TabIndex        =   33
         ToolTipText     =   "Time millisec "
         Top             =   135
         Width           =   900
      End
      Begin VB.Label LabActivity 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   75
         TabIndex        =   32
         ToolTipText     =   "Activity "
         Top             =   135
         Width           =   630
      End
   End
   Begin VB.Frame fraPaint 
      BackColor       =   &H8000000C&
      Height          =   2280
      Left            =   105
      TabIndex        =   23
      Top             =   3240
      Width           =   390
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   3
         Left            =   60
         Picture         =   "Form1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1050
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   4
         Left            =   60
         Picture         =   "Form1.frx":04CC
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1350
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   2
         Left            =   60
         Picture         =   "Form1.frx":0556
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   750
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   6
         Left            =   60
         Picture         =   "Form1.frx":05E0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1950
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   5
         Left            =   60
         Picture         =   "Form1.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1650
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "Form1.frx":06F4
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   150
         Width           =   300
      End
      Begin VB.CommandButton cmdPaint 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   1
         Left            =   60
         Picture         =   "Form1.frx":077E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   450
         Width           =   300
      End
   End
   Begin VB.Frame fraSelRect 
      BackColor       =   &H8000000C&
      Height          =   765
      Left            =   105
      TabIndex        =   21
      Top             =   2475
      Width           =   390
      Begin VB.CommandButton cmdSelShape 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   450
         Width           =   300
      End
      Begin VB.CommandButton cmdSelShape 
         BackColor       =   &H80000009&
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "Form1.frx":0808
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   150
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdMagnifier 
      BackColor       =   &H80000009&
      Height          =   330
      Left            =   135
      Picture         =   "Form1.frx":0892
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2145
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9705
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   270
      Width           =   435
   End
   Begin VB.PictureBox picThreshold 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   9705
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   15
      Top             =   570
      Width           =   435
   End
   Begin VB.CommandButton cmdCopyPic 
      BackColor       =   &H80000009&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Copy Picture 1 to Picture 2 "
      Top             =   225
      Width           =   450
   End
   Begin VB.Frame fraDespecklers 
      BackColor       =   &H8000000C&
      Height          =   1980
      Left            =   105
      TabIndex        =   10
      Top             =   90
      Width           =   390
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   5
         Left            =   60
         Picture         =   "Form1.frx":09DC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1650
         Width           =   300
      End
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   4
         Left            =   60
         Picture         =   "Form1.frx":0A66
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1350
         Width           =   300
      End
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   3
         Left            =   60
         Picture         =   "Form1.frx":0AF0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1050
         Width           =   300
      End
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   2
         Left            =   60
         Picture         =   "Form1.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   750
         Width           =   300
      End
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   1
         Left            =   60
         Picture         =   "Form1.frx":0C04
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   450
         Width           =   300
      End
      Begin VB.OptionButton optClump 
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "Form1.frx":0C8E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdCopyPic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copy Picture 2 to Picture 1 "
      Top             =   225
      Width           =   450
   End
   Begin VB.PictureBox picC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   6195
      Index           =   1
      Left            =   5460
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   4
      Top             =   690
      Width           =   4170
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   15
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   190
         TabIndex        =   5
         Top             =   0
         Width           =   2850
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   1170
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6960
      Width           =   3330
   End
   Begin VB.VScrollBar VS 
      Height          =   2715
      Left            =   570
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   255
   End
   Begin VB.PictureBox picC 
      BackColor       =   &H00008000&
      Height          =   6210
      Index           =   0
      Left            =   900
      ScaleHeight     =   410
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   0
      Top             =   555
      Width           =   4485
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   3885
         Left            =   30
         Picture         =   "Form1.frx":0D18
         ScaleHeight     =   259
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   368
         TabIndex        =   3
         Top             =   150
         Width           =   5520
         Begin VB.Line SL 
            BorderWidth     =   2
            Index           =   0
            X1              =   55
            X2              =   73
            Y1              =   36
            Y2              =   26
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   435
      Left            =   1800
      TabIndex        =   28
      Top             =   75
      Width           =   1950
      Begin VB.Label LabPaint 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   555
         TabIndex        =   34
         ToolTipText     =   "Selected Paint Color "
         Top             =   135
         Width           =   480
      End
      Begin VB.Label LabPicCul 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   30
         ToolTipText     =   "Picture color "
         Top             =   135
         Width           =   465
      End
      Begin VB.Label LabXY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X,Y"
         Height          =   255
         Left            =   1050
         TabIndex        =   29
         ToolTipText     =   "Picture X,Y coords & Bar level "
         Top             =   135
         Width           =   840
      End
   End
   Begin VB.Label LabShadeBar 
      BackColor       =   &H00008000&
      Caption         =   "Shade   bar"
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   9690
      TabIndex        =   39
      Top             =   4575
      Width           =   480
   End
   Begin VB.Label LabSDP 
      BackColor       =   &H00008000&
      Caption         =   "Destination     Picture 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   6195
      TabIndex        =   9
      Top             =   90
      Width           =   900
   End
   Begin VB.Label LabSDP 
      BackColor       =   &H00008000&
      Caption         =   "Source       Picture 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Top             =   105
      Width           =   795
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   45
      Left            =   60
      TabIndex        =   6
      Top             =   5550
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load bmp, jpg, gif"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Picture 2, 24bpp bmp"
      End
      Begin VB.Menu zbrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPNG 
         Caption         =   "Load &png"
      End
      Begin VB.Menu mnuSaveJPGGIFPNG 
         Caption         =   "Save &jpg, gif, png"
      End
      Begin VB.Menu zbrk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReLoad 
         Caption         =   "Re-load"
      End
      Begin VB.Menu zbrk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCorS 
         Caption         =   "&Click or Down && Slide on bars"
         Begin VB.Menu mnuClickorSlide 
            Caption         =   "&Click"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuClickorSlide 
            Caption         =   "&Slide"
            Index           =   1
         End
      End
      Begin VB.Menu zbrk5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelLineThickness 
         Caption         =   "&Shape Tool  line thickness"
         Begin VB.Menu mnuSelPixThick 
            Caption         =   "&1 pixel"
            Index           =   0
         End
         Begin VB.Menu mnuSelPixThick 
            Caption         =   "&2 pixel"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu zbrk56 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetAnyColor 
         Caption         =   "C&olor selector"
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "F&ilters"
      Begin VB.Menu mnuSaltPepper 
         Caption         =   "&Add Salt 'n Pepper Noise"
      End
      Begin VB.Menu zbrk6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSharpSoftness 
         Caption         =   "&Sharp-Softness bar"
      End
      Begin VB.Menu mnuDarkBrightness 
         Caption         =   "&Dark-Brightness bar"
      End
   End
   Begin VB.Menu mnuRESIZE 
      Caption         =   "&RESIZE"
   End
   Begin VB.Menu mnuVBASM 
      Caption         =   "&VB->ASM"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuFileName 
      Caption         =   "FileName"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1.frm

' Update  13/2/03
'         Re-load added
'         Add Salt 'n Pepper scaled to (PICW+PICH)\2

' Basic Despeckler by  Robert Rayment  Feb 2003 V1.0

' Thanks to:

' VBAccelerator.com for OSDialog.cls (Modified)
' VBSpeed <www.xbeat.net/vbspeed/index.htm> for CTimingPC.cls
' Ulli (PSC  CodeId=42051) for CToolTip.cls
' Carles P V for inspiration and debugging, but remaining
'  errors are all mine!
' Manuel Santos (PSC CodeId=26303) filters reference
'  & original Bart 'n God image.


' Option for loading & saving non-standard VB pictures
' Using FREEWARE i_view32.exe by Irfan Skiljan
' <www.irfanview.com>


Option Base 1  ' Arrays base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings

Dim fraTop
Dim fraLeft

Dim tmHowLong As New CTimingPC
Dim CommonDialog1 As New OSDialog
Dim Tooltips As New Collection
Dim Tooltip As cToolTip

Private Sub Form_Load()

'Public TempBmp$ ' used by i_view32.rxr
TempBmp$ = "~!@temp.bmp"   ' NB UNIQUENESS NOT CHECKED!

' Boolean switches OFF
SwitchOffBooleans
aSelectionMade = False

aMagON = False
ASMON = False

aClickON = True   ' Click or slide on shade bar

aColorFormON = False

SLWidth = 2
SL(0).BorderWidth = SLWidth
SL(0).Visible = False      ' Shape Lines

fraBrightness.Visible = False

picPatchMask.Visible = False
picPatch.Visible = False

FileSpec$ = ""
mnuReLoad.Enabled = False

RESIZE
''''''''''''''''''''''''''''
' Threshold bar
For iy = 0 To 255
   If iy Mod 3 = 0 Then Cul = RGB(0, 160, 0) Else Cul = RGB(iy, iy, iy)
   picThreshold.Line (0, iy)-(picThreshold.Width, iy), Cul
Next iy
picThreshold.Refresh

' Sharp >---> Brightness bar
For ix = 0 To 255 Step 2
   Cul = RGB(ix, ix, ix)
   If ix = 128 Then Cul = RGB(0, 250, 250)
   picLevelBar.Line (ix \ 2, 0)-(ix \ 2, picLevelBar.Height), Cul
Next ix
picLevelBar.Refresh

'Public PathSpec$, CurrPath$

PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

CheckForIncludedFiles

LabTim = "millisec"

' Initial Clump size
ClumpIndex = 0
optClump(ClumpIndex).Value = True

'------------------------------------------------
GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M
ReDim bP2M(4, PICW, PICH)   ' Picture 2 mem
ReDim bPPM(4, PICW, PICH)   ' picPatchMask mem
' Make picPatchMask white for full despeckling
picPatchMask.BackColor = vbBlack
GETDIBS picPatchMask.Image, 3   '6
' Shape bounding rect
ixTL = 0: iyTL = 0
ixBR = PICW - 1: iyBR = PICH - 1

'------------------------------------------------

' No activity
LabActivity.BackColor = vbGreen

' Initial paintcolor
PickedCul = vbRed

' Initial threshold
Threshold = 64
Text1.Text = "64"

' The high byte = the alphabyte = bPPM(4, ix, iy) = 0

'#### MAKE ULLI's TOOLTIPS ############################################

UlliToolTips

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Load Machine code from bin file
Loadmcode PathSpec$ & "Despec.bin", bDespecMC()
ptrMC = VarPtr(bDespecMC(1))
ptrStruc = VarPtr(ASMStruc.PICH)
'BB = ExtractMC(1)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'To Load Machine code frm Res file
'Put Despec.bin in res file first
' Public bDespecMC() As Byte  'Array to hold machine code
' Public ptrMC, ptrStruc      ' Ptrs to Machine Code & Structure

' bDespecMC = LoadResData("DESPEC", "ASM")
' ptrMC = VarPtr(bDespecMC(0))
' ptrStruc = VarPtr(ASMStruc.PICH)
'BB = ExtractMC(0)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Private Sub CheckForIncludedFiles()

'Public PathSpec$, CurrPath$

If Dir(PathSpec$ & "DesPicPath.txt") = "" Then
   ' Create despec.ini
   CurrPath$ = PathSpec$
   Open PathSpec$ & "DesPicPath.txt" For Output As #1
   Print #1, PathSpec$
   Close
Else  ' Read despec.ini
   Open PathSpec$ & "DesPicPath.txt" For Input As #1
   Line Input #1, CurrPath$
   Close
   ' Check CurrPath$ exists
   If Dir(CurrPath$, vbDirectory) = "" Then
      ' Create new despec.ini
      Open PathSpec$ & "DesPicPath.txt" For Output As #1
      Print #1, PathSpec$
      Close
      CurrPath$ = PathSpec$
   Else
   End If
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' Test if i_view32.exe & Help.txt in App folder

If Dir(PathSpec$ & "i_view32.exe") = "" Then
   mnuLoadPNG.Enabled = False
   mnuSaveJPGGIFPNG.Enabled = False
End If

If Dir(PathSpec$ & "Help.txt") = "" Then
   mnuHelp.Enabled = False
End If
End Sub

Private Sub RESIZE()

'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Public Const SM_CXSCREEN = 0 'X Size of screen
'Public Const SM_CYSCREEN = 1 'Y Size of Screen

TPX = Screen.TwipsPerPixelX
TPY = Screen.TwipsPerPixelY
FW = GetSystemMetrics(SM_CXSCREEN)
FH = GetSystemMetrics(SM_CYSCREEN)

If FW = 640 Then
   Form1.Width = FW * TPX * 0.98
Else
   Form1.Width = FW * TPX * 0.976
End If
Form1.Height = FH * TPY * 0.9

Form1.Left = (FW * TPX - Form1.Width) \ 2
Form1.Top = 250

Refresh

' picC(0/1) = Containers
' pic1, pic2 = Pictures
If FW = 640 Then
   picC(0).Width = Int(0.028 * Form1.Width + 0.5)
   picC(0).Height = Int(0.0484 * Form1.Height + 0.5)
Else
   picC(0).Width = Int(0.0286 * Form1.Width + 0.5)
   picC(0).Height = Int(0.0511 * Form1.Height + 0.5)
End If

picC(1).Width = picC(0).Width
picC(1).Height = picC(0).Height
picC(1).Top = picC(0).Top
picC(1).Left = picC(0).Left + picC(0).Width + 3

pic1.Top = 0
pic1.Left = 0
With pic2
   .Width = pic1.Width
   .Height = pic1.Height
   .Top = 0
   .Left = 0
End With

With picPatchMask
   .Width = pic1.Width
   .Height = pic1.Height
   .Top = 300
   .Left = 50
End With

With picPatch
   .Width = pic1.Width
   .Height = pic1.Height
   .Top = 300
   .Left = 300
End With


FixScrollbars picC(0), pic1, HS, VS

cmdCopyPic(1).Left = picC(0).Left + picC(0).Width - cmdCopyPic(0).Width - 12
cmdCopyPic(0).Left = picC(1).Left + 12

Text1.Left = picC(1).Left + picC(1).Width + 1
picThreshold.Left = picC(1).Left + picC(1).Width + 1


LabSDP(0).Left = picC(0).Left + 2
LabSDP(1).Left = cmdCopyPic(0).Left + cmdCopyPic(0).Width + 22
Frame2.Left = LabSDP(1).Left + LabSDP(1).Width + 12

LabShadeBar.Left = picThreshold.Left + 1

fraBrightness.Top = picC(0).Top
fraBrightness.Left = picC(0).Left

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Form As Form
Me.MousePointer = vbDefault
Erase bP1M(), bP2M(), bPPM()

' Save CurrPath$ to despec.ini
Open PathSpec$ & "DesPicPath.txt" For Output As #1
Print #1, CurrPath$
Close

' Make sure all forms cleared
For Each Form In Forms
   Unload Form
   Set Form = Nothing
Next Form
End
End Sub



'#### OPTION MENUS #####################################

Private Sub mnuClickorSlide_Click(Index As Integer)

aClickON = Not aClickON
mnuClickorSlide(0).Checked = Not mnuClickorSlide(0).Checked
mnuClickorSlide(1).Checked = Not mnuClickorSlide(1).Checked
End Sub

Private Sub mnuReLoad_Click()

If Len(FileSpec$) = 0 Then
   Exit Sub
End If

pdot = InStrRev(FileSpec$, ".")
If pdot = 0 Then Exit Sub
Extension$ = LCase$(Mid$(FileSpec$, pdot + 1))

pnum = InStrRev(FileSpec$, "\")
FileSpecPath$ = Left$(FileSpec$, pnum)

If Extension$ <> "png" Then

   pic1.Picture = LoadPicture(FileSpec$)

Else

   LoadPNG FileSpec$
   ' temp bmp TempBmp$ will be in FileSpec$ folder
   
   ' Load temp bmp
   pic1.Picture = LoadPicture(FileSpecPath$ & TempBmp$)
   ' Kill temp bmp
   Kill FileSpecPath$ & TempBmp$
   
End If

GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M

DoEvents

End Sub

Private Sub mnuSelPixThick_Click(Index As Integer)

SLWidth = Index + 1
mnuSelPixThick(0).Checked = Not mnuSelPixThick(0).Checked
mnuSelPixThick(1).Checked = Not mnuSelPixThick(1).Checked
End Sub

Private Sub mnuGetAnyColor_Click()

If aColorFormON = False Then
   aColorFormON = True
   frmCulPick.Show
Else
   frmCulPick.Hide
   aColorFormON = False
End If
End Sub

'########################################################

Private Sub mnuExit_Click()
Form_Unload 0
End Sub

Private Sub mnuHelp_Click()
frmHelp2.Show vbModal, Me
End Sub

Private Sub mnuRESIZE_Click()
RESIZE
End Sub

'#### COPY PICTURES ########################################

Private Sub cmdCopyPic_Click(Index As Integer)

If Index = 0 Then
   ' pic1 <<--<<-- pic2
   BitBlt pic1.hdc, 0, 0, PICW, PICH, pic2.hdc, 0, 0, vbSrcCopy
   '------------------------------------------------
   GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M
   '------------------------------------------------
   pic1.Refresh
Else
   ' pic1 >>-->>-- pic2
   BitBlt pic2.hdc, 0, 0, PICW, PICH, pic1.hdc, 0, 0, vbSrcCopy
   '------------------------------------------------
   GETDIBS pic2.Image, 2  ' pic1 32bpp DIBS to bP1M
   '------------------------------------------------
   pic2.Refresh
End If
End Sub

'#### FILTER MENUS ###############################################

Private Sub mnuFilters_Click()

fraBrightness.Visible = False
End Sub

Private Sub mnuSaltPepper_Click()

LabTim = ""
LabActivity.BackColor = vbRed: LabActivity.Refresh
tmHowLong.Reset

SaltPepper

ShowPicture 2   ' bP2M to pic2
LabTim = Format(Int(tmHowLong.Elapsed), "@@@@") & " ms"
LabActivity.BackColor = vbGreen: LabActivity.Refresh
End Sub


'#### SHARPNESS - SOFTNESS, DARKNESS - BRIGHTNESS ############

Private Sub mnuSharpSoftness_Click()
fraBrightness.Visible = Not fraBrightness.Visible
LabLevel.Caption = "Sharp-Softness bar"
End Sub

Private Sub mnuDarkBrightness_Click()
fraBrightness.Visible = Not fraBrightness.Visible
LabLevel.Caption = "Dark-Brightness bar"
End Sub

Private Sub picLevelBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case LabLevel.Caption

Case "Sharp-Softness bar"

   LabTim = ""
   LabActivity.BackColor = vbRed
   LabActivity.Refresh
   tmHowLong.Reset
   
   If FilterParam < 5 Then
   
      Sharpness

   Else
   
      FilterParam = FilterParam - 4
      Softness

   End If
   
   ShowPicture 2   ' bP2M to pic2
   LabTim = Format(Int(tmHowLong.Elapsed), "@@@@") & " ms"
   LabActivity.BackColor = vbGreen: LabActivity.Refresh

Case "Dark-Brightness bar"
   
   If FilterParam <> 0 Then
      LabTim = ""
      LabActivity.BackColor = vbRed
      LabActivity.Refresh
      tmHowLong.Reset
      
      Brightness
      
      ShowPicture 2   ' bP2M to pic2
      LabTim = Format(Int(tmHowLong.Elapsed), "@@@@") & " ms"
      LabActivity.BackColor = vbGreen: LabActivity.Refresh
   End If

End Select
End Sub

Private Sub picLevelBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X < 0 Then X = 0
If X > 128 Then X = 128

Select Case LabLevel.Caption
Case "Sharp-Softness bar"
   FilterParam = (8 * X) \ 128 + 1  ' 1-4 Sharp 5-7 Soft
   If FilterParam > 7 Then FilterParam = 7
   If FilterParam < 5 Then
      num$ = LTrim$(Str$(FilterParam - 5))
   Else
      num$ = "+" & LTrim$(Str$(FilterParam - 4))
   End If
   
Case "Dark-Brightness bar"
   
   FilterParam = (8 * X) \ 128 + 1
   If FilterParam > 8 Then FilterParam = 8
   If FilterParam < 5 Then
      num$ = LTrim$(Str$(FilterParam - 5))
   Else
      num$ = "+" & LTrim$(Str$(FilterParam - 4))
   End If

End Select

LabXY.Caption = num$
LabXY.Refresh

If Not aClickON And Button = vbLeftButton Then
   picLevelBar_MouseDown Button, Shift, X, Y
End If
End Sub

Private Sub cmdBrightnessExit_Click()

fraBrightness.Visible = False
End Sub

Private Sub fraBrightness_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

fraTop = Y: fraLeft = X
End Sub

Private Sub fraBrightness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 0 Then
   fraBrightness.Top = fraBrightness.Top + (Y - fraTop) \ Screen.TwipsPerPixelY
   fraBrightness.Left = fraBrightness.Left + (X - fraLeft) \ Screen.TwipsPerPixelX
End If
End Sub

Private Sub LabLevel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraTop = Y: fraLeft = X

End Sub

Private Sub LabLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 0 Then
   fraBrightness.Top = fraBrightness.Top + (Y - fraTop) \ Screen.TwipsPerPixelY
   fraBrightness.Left = fraBrightness.Left + (X - fraLeft) \ Screen.TwipsPerPixelX
End If

End Sub


'### ASM - VB SWITCH ########################################

Private Sub mnuVBASM_Click()

ASMON = Not ASMON
If ASMON Then
   mnuVBASM.Caption = "&ASM->VB"
Else
   mnuVBASM.Caption = "&VB->ASM"
End If
End Sub

'### SELECT DESPECKLER OPTIONS ###############################

Private Sub optClump_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

ClumpIndex = Index
End Sub


'#### Picture 1 COLOR PICKER & SHAPE SELECTION ########################

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ixSL0 = CLng(X)
iySL0 = CLng(Y)

If aCulPickerON Then
   PickedCul = Abs(GetPixel(pic1.hdc, ixSL0, iySL0))
   LabPaint.BackColor = PickedCul
   LabPaint.Refresh
   Exit Sub
End If

'----------------------------------------------------------------------
If aSelectShapeON Then
   aDrawing = True
   
   If NSLines > 1 Then
      For i = 2 To NSLines: Unload SL(i - 1): Next
   End If
   NSLines = 1

   With SL(0)
      .X1 = ixSL0: .Y1 = iySL0
      .X2 = ixSL0: .Y2 = iySL0
      .Visible = True
      .BorderWidth = SLWidth
   End With
   
   ' Start bounding rect coords
   ixTL = ixSL0: iyTL = iySL0
   ixBR = ixSL0: iyBR = iySL0

   picPatchMask.BackColor = 0
   GETDIBS picPatchMask.Image, 3   '1

   picPatchMask.PSet (ixSL0, iySL0), vbWhite
   
   aSelectionMade = True

End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ix = CLng(X)
iy = CLng(Y)

If ix <= 0 Then ix = 1   ' Leave a gap
If ix >= PICW - 1 Then ix = PICW - 2
If iy <= 0 Then iy = 1
If iy >= PICH - 1 Then iy = PICH - 2

LabXY.Caption = LTrim$(Str$(ix + 1)) & "," & Str$(PICH - iy)
' PICH
' |
' 1->PICW

LabPicCul.BackColor = Abs(GetPixel(pic1.hdc, ix, iy))
' 0->PICW-1
' |
' PICH-1
''''''''''''''''''''''''''''''''''''''''''''''''''''''
If aSelectShapeON And aDrawing Then
   ' Show B/W lines
   If Rnd < 0.5 Then SLCul = vbBlack Else SLCul = vbWhite
   'If SLCul = vbWhite Then SLCul = vbBlack Else SLCul = vbWhite
   
   NSLines = NSLines + 1
   Load SL(NSLines - 1)
   With SL(NSLines - 1)
      .X1 = SL(NSLines - 2).X2: .Y1 = SL(NSLines - 2).Y2
      .X2 = ix: .Y2 = iy
      .BorderColor = SLCul
      .Visible = True
   End With
   
   ' Find bounding rect coords of selection
   If ix < ixTL Then ixTL = ix
   If ix > ixBR Then ixBR = ix
   If iy < iyTL Then iyTL = iy
   If iy > iyBR Then iyBR = iy
   
   picPatchMask.Line -(ix, iy), vbWhite
End If
End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If aSelectShapeON And aDrawing Then
'''''''''''''''''''''''''''''''''''''''''''
   NSLines = NSLines + 1
   Load SL(NSLines - 1)
   With SL(NSLines - 1)
      .X1 = SL(NSLines - 2).X2: .Y1 = SL(NSLines - 2).Y2
      .X2 = SL(0).X1: .Y2 = SL(0).Y1
      .BorderColor = SLCul
      .Visible = True
      '.Refresh
   End With


   picPatchMask.Line -(ixSL0, iySL0), vbWhite
   'pic1.Refresh
   picPatchMask.Refresh
   
   ' Switch off aDrawing
   aDrawing = False
   
   ' Make B/W mask

   ' Have a white outline of select image
   ' Make mask: change outer region black to white
   picPatchMask.DrawStyle = vbSolid
   picPatchMask.DrawMode = 13
   picPatchMask.DrawWidth = 1
   picPatchMask.FillColor = vbWhite
   picPatchMask.FillStyle = vbFSSolid
   px = 0: py = 0    'FloodFill point
   Cul = picPatchMask.Point(px, py)
   If Cul <> vbBlack Then
      For i = 1 To 10
         px = px + 1: py = py + 1
         Cul = picPatchMask.Point(px, py)
         If Cul = vbBlack Then Exit For
      Next
      If i = 11 Then ' Black not found
         picPatchMask.FillStyle = vbFSTransparent  'Default (Transparent)
         MsgBox "Can't make mask.  Try aDrawing again", , "Despeckler"
         Exit Sub
      End If
   End If
   FillPtcul& = vbBlack  'picPatchMask.Point(pX,pY)
   'FLOODFILLSURFACE = 1
   'Fills with FillColor so long as point surrounded by FillPtcul&
   rs = ExtFloodFill(picPatchMask.hdc, px, py, FillPtcul&, FLOODFILLSURFACE)
   picPatchMask.Refresh
   picPatchMask.FillStyle = vbFSTransparent  'Default (Transparent)
   
   ' Make picPatchMask mem for testing black underlay
   GETDIBS picPatchMask.Image, 3  '2
   
   '------------------------------------
   ' Form picPatch, Black outside selected picture area
   picPatch.BackColor = vbWhite
   BitBlt picPatch.hdc, 0, 0, PICW, PICH, _
      picPatchMask.hdc, 0, 0, vbSrcAnd
   BitBlt picPatch.hdc, 0, 0, PICW, PICH, _
      pic1.hdc, 0, 0, vbSrcPaint
      
End If
End Sub


'#### Picture 2 COLOR PICKER, DE-SPOTTING, PAINTING & PATCHING ####

Private Sub pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ix = CLng(X)
iy = CLng(Y)

'------------------------------------------------------------
If aCulPickerON Then
   PickedCul = Abs(GetPixel(pic2.hdc, ix, iy))
   LabPaint.BackColor = PickedCul
   LabPaint.Refresh
   Exit Sub
End If

'------------------------------------------------------------
If aDeSpotON Then
   aDeSpotterStarted = True
   De_Spotter ix, iy
End If

'------------------------------------------------------------
If aPaintON Then
   If Button = 1 Then
      
      If PaintSize = 1 Then   ' 1x1 pixel
         SetPixelV pic2.hdc, ix, iy, PickedCul
      Else  ' 2x2 pixels
         SetPixelV pic2.hdc, ix, iy, PickedCul
         If ix < PICW - 1 Then SetPixelV pic2.hdc, ix + 1, iy, PickedCul
         If iy < PICH - 1 Then SetPixelV pic2.hdc, ix, iy + 1, PickedCul
         If ix < PICW - 1 And iy < PICH - 1 Then SetPixelV pic2.hdc, ix + 1, iy + 1, PickedCul
      End If
   
   Else  ' RC undo
      ' Undo 1x1
      bblue = bP2M(1, ix + 1, PICH - iy)
      bgreen = bP2M(2, ix + 1, PICH - iy)
      bred = bP2M(3, ix + 1, PICH - iy)
      SetPixelV pic2.hdc, ix, iy, RGB(bred, bgreen, bblue)
   
      If PaintSize = 2 Then   ' Undo 2x2
         If ix < PICW - 1 Then
            bblue = bP2M(1, ix + 2, PICH - iy)
            bgreen = bP2M(2, ix + 2, PICH - iy)
            bred = bP2M(3, ix + 2, PICH - iy)
            SetPixelV pic2.hdc, ix + 1, iy, RGB(bred, bgreen, bblue)
         End If
         If iy < PICH - 1 Then
            bblue = bP2M(1, ix + 1, PICH - iy - 1)
            bgreen = bP2M(2, ix + 1, PICH - iy - 1)
            bred = bP2M(3, ix + 1, PICH - iy - 1)
            SetPixelV pic2.hdc, ix, iy + 1, RGB(bred, bgreen, bblue)
            If ix < PICW Then
               bblue = bP2M(1, ix + 2, PICH - iy - 1)
               bgreen = bP2M(2, ix + 2, PICH - iy - 1)
               bred = bP2M(3, ix + 2, PICH - iy - 1)
               SetPixelV pic2.hdc, ix + 1, iy + 1, RGB(bred, bgreen, bblue)
            End If
         End If
      
      End If
   End If
   
   pic2.Refresh

   res = GetCursorPos(LP)
   If PaintSize = 1 Then
      res = SetCursorPos(LP.X + 1, LP.Y + 1)
   Else
      res = SetCursorPos(LP.X + 2, LP.Y + 2)
   End If
   DoEvents

End If

'------------------------------------------------------------
If aPatchON And aSelectionMade Then
  
   picPatchMask.Refresh
   aPatchFilled = True

   ixSL0 = ix
   iySL0 = iy
      
   PatchWidth = ixBR - ixTL '+ 1
   PatchHeight = iyBR - iyTL + 1

   ShowPicture 2  ' Clear all including any unfixed pixels
   
   ' Show selected patch
   BitBlt pic2.hdc, ix, iy, PatchWidth, PatchHeight, _
      picPatchMask.hdc, ixTL, iyTL, vbMergePaint
   BitBlt pic2.hdc, ix, iy, PatchWidth, PatchHeight, _
      picPatch.hdc, ixTL, iyTL, vbSrcAnd

   pic2.Refresh
   
End If
End Sub

Private Sub pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim R As RECT, RP As RECT


ix = CLng(X)
iy = CLng(Y)

LabXY.Caption = LTrim$(Str$(ix + 1)) & "," & Str$(PICH - iy)
' PICH
' |
' 1->PICW

LabPicCul.BackColor = Abs(GetPixel(pic2.hdc, ix, iy))
' 0->PICW-1
' |
' PICH-1

'LngToRGB

'------------------------------------------------------------
If aPatchON And aPatchFilled Then

   PatchWidth = ixBR - ixTL + 1
   PatchHeight = iyBR - iyTL + 1
   
   ' Clear previous patch position
   ' Check patch in bounds
   ' Set the rectangle's values
   ' Suggestion from Carles P V to stop flickering
   ' & Blit error on some computers (but not on mine??)
   
   SetRect R, 0, 0, PICW - 1, PICH - 1
   SetRect RP, ixSL0, iySL0, ixSL0 + PatchWidth - 1, iySL0 + PatchHeight - 1
   IntersectRect R, R, RP
   
   If (IsRectEmpty(R) = 0) Then
      ' Inside, clear last patch
   
      iyR = PICH - (iySL0 + PatchHeight)
      
      If StretchDIBits(pic2.hdc, _
         ixSL0, iySL0, PatchWidth, PatchHeight, _
         ixSL0, iyR, PatchWidth, PatchHeight, _
         bP2M(1, 1, 1), bm, _
         DIB_RGB_COLORS, vbSrcCopy) = 0 Then
            MsgBox "Blit Error, pic2_MouseMove", , "Despeckler"
            Unload Form1
      End If
   End If

   ' Show new selection position of patch
   BitBlt pic2.hdc, ix, iy, PatchWidth, PatchHeight, _
      picPatchMask.hdc, ixTL, iyTL, vbMergePaint
   BitBlt pic2.hdc, ix, iy, PatchWidth, PatchHeight, _
      picPatch.hdc, ixTL, iyTL, vbSrcAnd
   
   ixSL0 = ix
   iySL0 = iy
   pic2.Refresh

End If

'------------------------------------------------------------
If aDeSpotON And aDeSpotterStarted Then
   De_Spotter ix, iy
End If
End Sub

Private Sub pic2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'------------------------------------------------------------

If aDeSpotON Then aDeSpotterStarted = False

'------------------------------------------------------------
If aPatchON And aPatchFilled Then
   
   aPatchFilled = False

End If
End Sub

Private Sub De_Spotter(ByVal ix, ByVal iy)

If ix > 0 And ix < PICW - 1 And iy > 0 And iy < PICH - 1 Then
   LngToRGB GetPixel(pic2.hdc, ix - 1, iy - 1)
   red1 = bred: green1 = bgreen: blue1 = bblue
   LngToRGB GetPixel(pic2.hdc, ix + 1, iy + 1)
   culR = (1& * red1 + bred) \ 2
   culG = (1& * green1 + bgreen) \ 2
   culB = (1& * blue1 + bblue) \ 2
   
   LngToRGB GetPixel(pic2.hdc, ix + 1, iy - 1)
   red1 = bred: green1 = bgreen: blue1 = bblue
   LngToRGB GetPixel(pic2.hdc, ix - 1, iy + 1)
   culR = (culR + (1& * red1 + bred) \ 2) \ 2
   culG = (culG + (1& * green1 + bgreen) \ 2) \ 2
   culB = (culB + (1& * blue1 + bblue) \ 2) \ 2
   
   SetPixelV pic2.hdc, ix, iy, RGB(culR, culG, culB)

ElseIf (ix = 0 Or ix = PICW - 1) And (iy <> 0 And iy <> PICH - 1) Then
   ' Vert edge
   LngToRGB GetPixel(pic2.hdc, ix, iy - 1)
   red1 = bred: green1 = bgreen: blue1 = bblue
   LngToRGB GetPixel(pic2.hdc, ix, iy + 1)
   culR = (1& * red1 + bred) \ 2
   culG = (1& * green1 + bgreen) \ 2
   culB = (1& * blue1 + bblue) \ 2
   
   SetPixelV pic2.hdc, ix, iy, RGB(culR, culG, culB)

ElseIf (iy = 0 Or iy = PICH - 1) And (ix <> 0 And ix <> PICW - 1) Then
   ' Horz edge
   LngToRGB GetPixel(pic2.hdc, ix - 1, iy)
   red1 = bred: green1 = bgreen: blue1 = bblue
   LngToRGB GetPixel(pic2.hdc, ix + 1, iy)
   culR = (1& * red1 + bred) \ 2
   culG = (1& * green1 + bgreen) \ 2
   culB = (1& * blue1 + bblue) \ 2
   
   SetPixelV pic2.hdc, ix, iy, RGB(culR, culG, culB)

End If

pic2.Refresh
End Sub


'#### CALL DESPECKLERS #####################################

Private Sub picThreshold_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbHourglass

Threshold = CLng(Y)

If Threshold < 0 Then Threshold = 0
If Threshold > 255 Then Threshold = 255

Text1.Text = Str$(Threshold)
Text1.Refresh

LabTim = ""
LabActivity.BackColor = vbRed
LabActivity.Refresh
tmHowLong.Reset

Select Case ClumpIndex
Case 0: Despeckle1   ' 1x1 Clump
Case 1: Despeckle2   ' 2x2 Clump
Case 2: Despeckle3   ' 3x3 Clump
Case 3: Despeckle4   ' 1 pix Scratch
Case 4: Despeckle5   ' 2 pix Scratch
Case 5: Despeckle6   ' 3 pix Scratch
End Select

ShowPicture 2   ' bP2M to pic2
LabTim = Format(Int(tmHowLong.Elapsed), "@@@@") & " ms"
LabActivity.BackColor = vbGreen: LabActivity.Refresh

' Restore bP1M  leave bP2M the same
'------------------------------------------------
GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M
'------------------------------------------------

Screen.MousePointer = vbDefault
End Sub

Private Sub picThreshold_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Y < 0 Then Y = 0
If Y > 255 Then Y = 255

Threshold = CLng(Y)
Text1.Text = Str$(Y)
Text1.Refresh
DoEvents

'Exit Sub
'If ASMON And ImageBytes < 500000 Then

If Not aClickON And Button = vbLeftButton Then

   Screen.MousePointer = vbHourglass

   LabTim = ""
   LabActivity.BackColor = vbRed
   LabActivity.Refresh
   tmHowLong.Reset

   Select Case ClumpIndex
   Case 0: Despeckle1   ' 1x1 Clump
   Case 1: Despeckle2   ' 2x2 Clump
   Case 2: Despeckle3   ' 3x3 Clump
   Case 3: Despeckle4   ' 1 pix Scratch
   Case 4: Despeckle5   ' 2 pix Scratch
   Case 5: Despeckle6   ' 3 pix Scratch
   End Select

   ShowPicture 2   ' bP2M to pic2
   LabTim = Format(Int(tmHowLong.Elapsed), "@@@@") & " ms"
   LabActivity.BackColor = vbGreen: LabActivity.Refresh

   ' Restore bP1M  leave bP2M the same
   '------------------------------------------------
   GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M
   '------------------------------------------------

   Screen.MousePointer = vbDefault

End If
End Sub

Private Sub ShowPicture(n)

'Public Const DIB_PAL_COLORS = 1 '  uses system colors
'Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD colors
Select Case n
Case 1
   If StretchDIBits(pic1.hdc, _
      0, 0, PICW, PICH, _
      0, 0, PICW, PICH, _
      bP1M(1, 1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Despeckler"
         Form_Unload 0
   End If
   pic1.Refresh
Case 2
   If StretchDIBits(pic2.hdc, _
      0, 0, PICW, PICH, _
      0, 0, PICW, PICH, _
      bP2M(1, 1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Despeckler"
         Form_Unload 0
   End If
   pic2.Refresh
Case 3
   If StretchDIBits(picPatchMask.hdc, _
      0, 0, PICW, PICH, _
      0, 0, PICW, PICH, _
      bPPM(1, 1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Despeckler"
         Form_Unload 0
   End If
   picPatchMask.Refresh
End Select
End Sub


'#### MAGNIFIER #######################################

Private Sub cmdMagnifier_Click()

If aMagON = False Then
   aMagON = True
   MagForm.Show
Else
   MagForm.Hide
   aMagON = False
End If
End Sub


'#### SELECT SHAPE TOGGLER OR DESELECT ############################

Private Sub cmdSelShape_Click(Index As Integer)

If Index = 0 Then

   If aSelectShapeON = False Then   ' Select shape
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(101, vbResCursor)
      aSelectShapeON = True
      aCulPickerON = False
      aDeSpotON = False
      aPaintON = False
      PaintSize = 0
      ''''''''''''''''''''''''''''''''''''''
      If NSLines > 1 Then
         For i = 2 To NSLines: Unload SL(i - 1): Next
      End If
      SL(0).BorderWidth = SLWidth
      NSLines = 1
      picPatchMask.BackColor = vbBlack
      picPatchMask.Refresh
      GETDIBS picPatchMask.Image, 3   '3
      ''''''''''''''''''''''''''''''''''''''
   Else
      Me.MousePointer = vbDefault
      SwitchOffBooleans
      aSelectionMade = True
   End If

Else  ' Deselect
      Me.MousePointer = vbDefault
      SwitchOffBooleans
      ''''''''''''''''''''''''''''''''''''''
      If NSLines > 1 Then
         For i = 2 To NSLines: Unload SL(i - 1): Next
      End If
      NSLines = 1
      SL(0).BorderWidth = SLWidth
      SL(0).Visible = False
      
      picPatchMask.BackColor = vbBlack
      picPatchMask.Refresh
      GETDIBS picPatchMask.Image, 3  '4
      ' Shape bounding rect
      ixTL = 0: iyTL = 0
      ixBR = PICW - 1: iyBR = PICH - 1
      ''''''''''''''''''''''''''''''''''''''
      aSelectionMade = False
End If
End Sub

'#### Select COLOR PICKER, PIXEL PAINTER, DE-SPOTTER,#############
'#### PATCHER, FIX OR UNDO ########################################

Private Sub cmdPaint_Click(Index As Integer)

Select Case Index

Case 0   ' Cul Picker ON
   aCulPickerON = Not aCulPickerON
   If aCulPickerON Then
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(102, vbResCursor)
      
      aDeSpotON = False
      aSelectShapeON = False
      aPaintON = False
      aPatchON = False
      PaintSize = 0
   Else
      Me.MousePointer = vbDefault
      SwitchOffBooleans
   End If

Case 1   ' Pixel painter ON
   If aPaintON = False Then
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(103, vbResCursor)
      PaintSize = 1
      aPaintON = True
      aSelectShapeON = False
      aDeSpotON = False
      aCulPickerON = False
      aPatchON = False
   Else
      Me.MousePointer = vbDefault
      SwitchOffBooleans
   End If

Case 2   ' 2x2 Pixel painter ON
   If aPaintON = False Then
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(104, vbResCursor)
      PaintSize = 2
      aPaintON = True
      aSelectShapeON = False
      aDeSpotON = False
      aCulPickerON = False
      aPatchON = False
   Else
      Me.MousePointer = vbDefault
      SwitchOffBooleans
   End If
Case 3   ' De-spotter
   If aDeSpotON = False Then
      aDeSpotON = True
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(106, vbResCursor)
      aSelectShapeON = False
      aCulPickerON = False
      aPaintON = False
      PaintSize = 0
      aPatchON = False
   Else
      Me.MousePointer = vbDefault
      SwitchOffBooleans
   End If

Case 4   ' Patch
   If aPatchON = False Then
      If aSelectionMade = False Then
         MsgBox " No Selection made", vbExclamation, "Despeckler"
         Exit Sub
      End If
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(105, vbResCursor)
      aPatchON = True
      aCulPickerON = False
      aDeSpotON = False
      aPaintON = False
      PaintSize = 0
      
      For i = 0 To 3
         If i < 2 Then cmdSelShape(i).Enabled = False
         cmdPaint(i).Enabled = False
      Next i
   
   Else
      
      aPatchON = False
      Me.MousePointer = vbDefault
      For i = 0 To 3
         If i < 2 Then cmdSelShape(i).Enabled = True
         cmdPaint(i).Enabled = True
      Next i
   
   End If

Case 5   ' Default & fix colors on Picture 2
   Me.MousePointer = vbDefault
   SwitchOffBooleans
   GETDIBS pic2.Image, 2  ' pic2 32bpp DIBS to bP2M

   For i = 0 To 3
      If i < 2 Then cmdSelShape(i).Enabled = True
      cmdPaint(i).Enabled = True
   Next i

Case 6   ' Undo all painting
   Me.MousePointer = vbDefault
   SwitchOffBooleans
   ShowPicture 2   ' bP2M to pic2

   For i = 0 To 3
      If i < 2 Then cmdSelShape(i).Enabled = True
      cmdPaint(i).Enabled = True
   Next i

End Select
End Sub


'#### SCROLL BARS #########################################

Private Sub HS_Change()

pic1.Left = -HS.Value
pic2.Left = -HS.Value
End Sub

Private Sub HS_Scroll()

pic1.Left = -HS.Value
pic2.Left = -HS.Value
End Sub

Private Sub VS_Change()

pic1.Top = -VS.Value
pic2.Top = -VS.Value
End Sub

Private Sub VS_Scroll()

pic1.Top = -VS.Value
pic2.Top = -VS.Value
End Sub

Private Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)

' picC = Container
' picP = Picture
HS.Max = picP.Width - picC.Width
VS.Max = picP.Height - picC.Height
HS.LargeChange = picC.Width \ 10
HS.SmallChange = 1
VS.LargeChange = picC.Height \ 10
VS.SmallChange = 1
HS.Top = picC.Top + picC.Height + 1
HS.Left = picC.Left
HS.Width = picC.Width
If picP.Width < picC.Width Then
   HS.Visible = False
   'HS.Enabled = False
Else
   HS.Visible = True
   'HS.Enabled = True
End If
VS.Top = picC.Top
VS.Left = picC.Left - VS.Width - 1
VS.Height = picC.Height
If picP.Height < picC.Height Then
   VS.Visible = False
   'VS.Enabled = False
Else
   VS.Visible = True
   'VS.Enabled = True
End If
End Sub


'#### LOAD PICTURE ############################################

Private Sub mnuLoad_Click()

' LOAD STANDARD VB PICTURES

If aMagON = True Then
   MagForm.Hide
   aMagON = False
End If

If aColorFormON = True Then
   frmCulPick.Hide
   aColorFormON = False
End If

MousePointer = vbDefault

Title$ = "Load a picture file"
Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
InDir$ = CurrPath$

CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

pnum = InStrRev(FileSpec$, "\")
FileSpecPath$ = Left$(FileSpec$, pnum)
mnuFileName.Caption = "File = " & Mid$(FileSpec$, pnum + 1)

CurrPath$ = FileSpecPath$

pic1.Picture = LoadPicture
pic1.Picture = LoadPicture(FileSpec$)
DoEvents

ConditionLoadedPicture

End Sub

Private Sub mnuLoadPNG_Click()

' LOAD PNG  using i_view32.exe

If aMagON = True Then
   MagForm.Hide
   aMagON = False
End If

If aColorFormON = True Then
   frmCulPick.Hide
   aColorFormON = False
End If

MousePointer = vbDefault

Title$ = "Load a PNG picture file"
Filt$ = "png|*.png"
InDir$ = CurrPath$

CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

LoadPNG FileSpec$
' temp bmp TempBmp$ will be in FileSpec$ folder

' InstrRev VB6 only
pnum = InStrRev(FileSpec$, "\")
FileSpecPath$ = Left$(FileSpec$, pnum)
mnuFileName.Caption = "File = " & Mid$(FileSpec$, pnum + 1)

CurrPath$ = FileSpecPath$

pic1.Picture = LoadPicture

' Load temp bmp
pic1.Picture = LoadPicture(FileSpecPath$ & TempBmp$)
' Kill temp bmp
Kill FileSpecPath$ & TempBmp$

DoEvents

ConditionLoadedPicture
End Sub

Private Sub ConditionLoadedPicture()

For i = 0 To 3
   If i < 2 Then cmdSelShape(i).Enabled = True
   cmdPaint(i).Enabled = True
Next i

PICW = pic1.Width
PICH = pic1.Height

If PICW < 16 Or PICH < 16 Then
   MsgBox "Can't have width or height < 16 pixels", vbExclamation, "Despeckler"
   If PICW < 16 Then PICW = 16
   If PICH < 16 Then PICH = 16
   pic1.Width = PICW
   pic1.Height = PICH
End If

pic2.Picture = LoadPicture

pic2.Width = pic1.Width
pic2.Height = pic1.Height

picPatchMask.Width = pic1.Width
picPatchMask.Height = pic1.Height

picPatch.Width = pic1.Width
picPatch.Height = pic1.Height

FixScrollbars picC(0), pic1, HS, VS

'------------------------------------------------
GETDIBS pic1.Image, 1  ' pic1 32bpp DIBS to bP1M

ReDim bP2M(4, PICW, PICH)   ' Picture 2 mem
'pic2.BackColor = vbBlack
'GETDIBS pic2.Image, 2  ' pic2 32bpp DIBS to bP2M

ReDim bPPM(4, PICW, PICH)   ' picPatchMask mem
' Make picPatchMask white for full despeckling
picPatchMask.BackColor = vbBlack
GETDIBS picPatchMask.Image, 3  '5
picPatchMask.Visible = False   ' picPatchMask
'------------------------------------------------

''''''''''''''''''''''''''''''''''''''
If NSLines > 1 Then
   For i = 2 To NSLines: Unload SL(i - 1): Next
End If
NSLines = 1
SL(0).Visible = False
''''''''''''''''''''''''''''''''''''''
' Boolean switches OFF
SwitchOffBooleans

'For Re-load
mnuReLoad.Enabled = True

End Sub

'#### SAVE PICTURE ############################################

Private Sub mnuSave_Click()

'  SAVE 24bpp BMP

If aMagON = True Then
   MagForm.Hide
   aMagON = False
End If

If aColorFormON = True Then
   frmCulPick.Hide
   aColorFormON = False
End If

Title$ = "Save 24 bpp bmp"
Filt$ = "Save bmp|*.bmp"
InDir$ = CurrPath$

CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

CurrPath$ = FileSpec$

SavePicture pic2.Image, FileSpec$
End Sub

Private Sub mnuSaveJPGGIFPNG_Click()

' SAVE JPG GIG PNG

If aMagON = True Then
   MagForm.Hide
   aMagON = False
End If

If aColorFormON = True Then
   frmCulPick.Hide
   aColorFormON = False
End If

Title$ = "Save As JPG, GIF or PNG"
Filt$ = "jpg|*.jpg|gif|*.gif|png|*.png"
InDir$ = CurrPath$
CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

If Len(FileSpec$) = 0 Then
   Exit Sub
End If

' InstrRev VB6 only
FileSpecPath$ = Left$(FileSpec$, InStrRev(FileSpec$, "\"))
CurrPath$ = FileSpecPath$

' Save temp bmp to save folder for i_view32 to convert
SavePicture pic2.Image, FileSpecPath$ & TempBmp$

DoEvents

SaveJPGGIFPNG FileSpec$

End Sub

'#### ULLI TOOLTIPS ######################################

Private Sub UlliToolTips()

Cul = &HC0FFFF

Set Tooltip = New cToolTip
Tooltip.Create Text1, "Shade bar|to activate", TTBalloonIfActive, False, , _
   "Threshold value", vbBlack, Cul, 150, 20000
Tooltips.Add Tooltip, Text1.Name

For i = 0 To 5
   Set Tooltip = New cToolTip
   CollKey$ = "optClump" & "(" & "." & Trim$(Str$(i)) & ")"
   First$ = "Clump" & Str$(i + 1)
   If i > 2 Then First$ = "Scratch" & Str$(i - 2)
   Tooltip.Create optClump(i), "Shade bar|to activate", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$
Next i

' Toggle magnifier
   Set Tooltip = New cToolTip
   CollKey$ = "cmdMagnifier"
   First$ = "Toggle Magnifier"
   Tooltip.Create cmdMagnifier, "-", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' Select shape
   Set Tooltip = New cToolTip
   CollKey$ = "cmdSelShape(0)"
   First$ = "Toggle Shape Tool"
   Tooltip.Create cmdSelShape(0), "select by mouse down, move|& up on Picture 1", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$
   
' DeSelect shape
   Set Tooltip = New cToolTip
   CollKey$ = "cmdSelShape(1)"
   First$ = "Deselect Shape"
   Tooltip.Create cmdSelShape(1), "-", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' Color Picker ON
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(0) '" & "(" & "." & Trim$(Str$(i)) & ")"
   First$ = "Toggle Color Picker"
   Tooltip.Create cmdPaint(0), "Click on Picture", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' Pixel painter ON
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(1)"
   First$ = "Toggle Pixel Painter"
   Tooltip.Create cmdPaint(1), "L Click to paint|R Click to undo|on Picture 2", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' 2x2 Pixel painter ON
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(2)"
   First$ = "Toggle 2x2 Pixel Painter"
   Tooltip.Create cmdPaint(2), "L Click to paint|R Click to undo|on Picture 2", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

'  Pixel Despotter
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(3)"
   First$ = "Toggle Pixel De-spotter"
   Tooltip.Create cmdPaint(3), "mouse down or |mouse down & move |over spots on Picture 2", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' Patch
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(4)"
   First$ = "Patch Selected Shape"
   Tooltip.Create cmdPaint(4), "by mouse down, move|& up on Picture 2", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' Pixel painter OFF & FIX
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(5)"
   First$ = "Paint, Patch & De-spotter Off"
   Tooltip.Create cmdPaint(5), "and FIX", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

' UNDO all
   Set Tooltip = New cToolTip
   CollKey$ = "cmdPaint(6)"
   First$ = "Undo"
   Tooltip.Create cmdPaint(6), "all Paint, Patch & De-spotter", TTBalloonIfActive, False, , _
      First$, vbBlack, Cul, 150, 20000
   Tooltips.Add Tooltip, CollKey$

End Sub

Private Sub SwitchOffBooleans()

aSelectShapeON = False
'aSelectionMade = False    ' NO
aDeSpotON = False
aCulPickerON = False
aPaintON = False
PaintSize = 0
aPatchON = False
aPatchFilled = False
End Sub
