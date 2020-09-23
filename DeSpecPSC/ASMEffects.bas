Attribute VB_Name = "ASMEffects"
' ASMEffects.bas  by  Robert Rayment

Option Base 1  ' Arrays base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings

'-----------------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long
'-----------------------------------------------------------------

Type ASMVars
   PICH As Long
   PICW As Long
   ptrP1M As Long
   ptrP2M As Long
   ptrPPM As Long
   Threshold As Long
   aSelectShapeON As Long   ' 0 False
   FilterParam As Long
End Type
Public ASMStruc As ASMVars

Public ptrMC, ptrStruc        ' Ptrs to MCode & Structure
Public bDespecMC() As Byte    ' Byte array to hold MCode

Public Const Despec1 As Long = 0
Public Const Despec2 As Long = 1
Public Const Despec3 As Long = 2
Public Const Despec4 As Long = 3
Public Const Despec5 As Long = 4
Public Const Despec6 As Long = 5
Public Const Sharp As Long = 6
Public Const Soft As Long = 7
Public Const Bright As Long = 8

Public ASMON As Boolean

Public Sub FillASMStruc()
With ASMStruc
   .PICH = PICH
   .PICW = PICW
   .ptrP1M = VarPtr(bP1M(1, 1, 1))
   .ptrP2M = VarPtr(bP2M(1, 1, 1))
   .ptrPPM = VarPtr(bPPM(1, 1, 1))
   .Threshold = Threshold
   .aSelectShapeON = CLng(aSelectShapeON)   ' 0 False
   .FilterParam = FilterParam
End With
End Sub

Public Sub ASM_Caller(OpCode&)
A = FilterParam
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, OpCode)

hexres$ = Hex$(res)

numr$ = Hex$(bP1M(3, 1, 1))
numg$ = Hex$(bP1M(2, 1, 1))
numb$ = Hex$(bP1M(1, 1, 1))

End Sub

Public Sub Loadmcode(InFile$, MCCode() As Byte)
'Load machine code into MCCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox InFile$ & " missing", vbCritical, "Despeckler"
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox InFile$ & " zero size", vbCritical, "Despeckler"
   Close
   Kill InFile$
   DoEvents
   Unload Form1
   End
End If

ReDim MCCode(MCSize&)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

