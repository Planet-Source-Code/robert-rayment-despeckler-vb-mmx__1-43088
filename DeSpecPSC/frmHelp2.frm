VERSION 5.00
Begin VB.Form frmHelp2 
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Despeckler Help & Acknowledgements"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
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
      Height          =   2745
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmHelp2.frx":0000
      Top             =   45
      Width           =   4350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   2865
      Width           =   4365
   End
End
Attribute VB_Name = "frmHelp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmHelp2.frm

'Unless otherwise defined:
Option Base 1  ' Arrays base 1
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings


Private Sub Form_Load()
Open PathSpec$ & "Help.txt" For Binary As #1
A$ = Space$(LOF(1))
Get #1, , A$
Text1.Text = A$
Close
Text1.Refresh
A$ = ""
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

