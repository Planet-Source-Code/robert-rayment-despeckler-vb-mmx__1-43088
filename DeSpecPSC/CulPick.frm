VERSION 5.00
Begin VB.Form frmCulPick 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Picker"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1305
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   87
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCul 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   690
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdCul 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   2400
      Width           =   645
   End
   Begin VB.VScrollBar VScr 
      Height          =   1485
      Index           =   2
      LargeChange     =   4
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   30
      Width           =   270
   End
   Begin VB.VScrollBar VScr 
      Height          =   1485
      Index           =   1
      LargeChange     =   4
      Left            =   510
      Max             =   255
      TabIndex        =   1
      Top             =   30
      Width           =   285
   End
   Begin VB.VScrollBar VScr 
      Height          =   1485
      Index           =   0
      LargeChange     =   4
      Left            =   180
      Max             =   255
      TabIndex        =   0
      Top             =   30
      Width           =   285
   End
   Begin VB.Label LabRGB 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   810
      TabIndex        =   6
      Top             =   1530
      Width           =   300
   End
   Begin VB.Label LabRGB 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1530
      Width           =   300
   End
   Begin VB.Label LabRGB 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   1530
      Width           =   300
   End
   Begin VB.Label LabCul 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   90
      TabIndex        =   3
      Top             =   1830
      Width           =   1020
   End
End
Attribute VB_Name = "frmCulPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CulPick.frm

Option Base 1
DefBool A
DefByte B
DefLng C-W
DefSng X-Y

'  Windows API to make application stay on top
' -----------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H20

'------------------------------------------------------------------------------

' Public PickedCul As Long
' Public aColorFormON As Boolean
' & Form1.LabPaint.BackColor

Dim brr, bgg, bbb

Private Sub cmdCul_Click(Index As Integer)
If Index = 0 Then
   PickedCul = RGB(brr, bgg, bbb)
   Form1.LabPaint.BackColor = PickedCul
Else
   frmCulPick.Hide
   aColorFormON = False
   'Unload Me
End If
End Sub

Private Sub Form_Load()

If App.PrevInstance Then End  ' Only allow it to run once.

fLeft = 384
fTop = 120
fWidth = 94
fHeight = 201
ret = SetWindowPos(frmCulPick.hwnd, hWndInsertAfter, _
   fLeft, fTop, fWidth, fHeight, wFlags)
'------------------------------------------------


'Show

PickedCul = RGB(120, 130, 140)   ' TEST

LngToRGB PickedCul

VScr(0).Value = brr
VScr(1).Value = bgg
VScr(2).Value = bbb
LabCul.BackColor = RGB(brr, bgg, bbb)
End Sub

Private Sub LngToRGB(LCul)
' Dim brr, bgg, bbb bytes
'Convert Long color LCul to RGB components
brr = (LCul And &HFF&)
bgg = (LCul And &HFF00&) / &H100&
bbb = (LCul And &HFF0000) / &H10000
End Sub
Private Sub VScr_Change(Index As Integer)
Select Case Index
Case 0: brr = VScr(0).Value: LabRGB(0) = Trim$(Str$(brr))
Case 1: bgg = VScr(1).Value: LabRGB(1) = Trim$(Str$(bgg))
Case 2: bbb = VScr(2).Value: LabRGB(2) = Trim$(Str$(bbb))
End Select
LabCul.BackColor = RGB(brr, bgg, bbb)
LabCul.Refresh
End Sub

Private Sub VScr_Scroll(Index As Integer)
Select Case Index
Case 0: brr = VScr(0).Value: LabRGB(0) = Trim$(Str$(brr))
Case 1: bgg = VScr(1).Value: LabRGB(1) = Trim$(Str$(bgg))
Case 2: bbb = VScr(2).Value: LabRGB(2) = Trim$(Str$(bbb))
End Select
LabCul.BackColor = RGB(brr, bgg, bbb)
LabCul.Refresh
End Sub

