Attribute VB_Name = "Effects"
' Effects.bas by  Robert Rayment

Option Base 1  ' Arrays base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings


'ReDim bP1M(4, PICW, PICH)  ' bytes  ' Picture 1
'ReDim bP2M(4, PICW, PICH)  ' bytes  ' Picture 2
'ReDim bPPM(4, PICW, PICH)  ' bytes  ' picPatchMask

' Copies modified pixels in bP1M to bP2M

Public Sub Despeckle1()

' 1x1 Clump

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes


If ASMON Then
   FillASMStruc
   ASM_Caller Despec1
   Exit Sub
End If

'Init Threshold = 64    ' Critical, on gradient bar

For iy = 2 To PICH - 1
For ix = 2 To PICW - 1
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------

      ' Sum 8 surrounding pixels
      gBAV = 0: gGAV = 0: gRAV = 0
      For iiy = -1 To 1
      For iix = -1 To 1
         If iix <> 0 Or iiy <> 0 Then
            gBAV = gBAV + bP1M(1, ix + iix, iy + iiy)
            gGAV = gGAV + bP1M(2, ix + iix, iy + iiy)
            gRAV = gRAV + bP1M(3, ix + iix, iy + iiy)
         End If
      Next iix
      Next iiy

      gBAV = gBAV \ 8: gGAV = gGAV \ 8: gRAV = gRAV \ 8
      
      DDiffB = Abs(gBAV - bP1M(1, ix, iy))
      DDiffG = Abs(gGAV - bP1M(2, ix, iy))
      DDiffR = Abs(gRAV - bP1M(3, ix, iy))
      
      If DDiffB > Threshold Or DDiffG > Threshold Or DDiffR > Threshold Then
         bP2M(1, ix, iy) = gBAV
         bP2M(2, ix, iy) = gGAV
         bP2M(3, ix, iy) = gRAV
         ' Copy back to source memory
         bP1M(1, ix, iy) = gBAV
         bP1M(2, ix, iy) = gGAV
         bP1M(3, ix, iy) = gRAV
      End If

   '---------------------------------------------------------
   End If
Next ix
Next iy

If Not aSelectShapeON Then
   ' Do Edges
   HorzEdges 1
   HorzEdges PICH

   VertEdges 1
   VertEdges PICW
End If

End Sub

Public Sub Despeckle2()

' 2x2 clump

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes
   
If ASMON Then
   FillASMStruc
   ASM_Caller Despec2
   Exit Sub
End If

For iy = 2 To PICH - 2 Step 2
For ix = 2 To PICW - 2 Step 2
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
   
      ' Sum 12 pixels around block of 4
      gBAV = 0: gGAV = 0: gRAV = 0
      For iiy = -1 To 2
      For iix = -1 To 2
         If Not ((iix = 0 Or iix = 1) And _
                 (iiy = 0 Or iiy = 1)) Then
            gBAV = gBAV + bP1M(1, ix + iix, iy + iiy)
            gGAV = gGAV + bP1M(2, ix + iix, iy + iiy)
            gRAV = gRAV + bP1M(3, ix + iix, iy + iiy)
         End If
      Next iix
      Next iiy
      
      gBAV = gBAV \ 12: gGAV = gGAV \ 12: gRAV = gRAV \ 12
      
      For iiy = 0 To 1
      For iix = 0 To 1
         DDiffB = Abs(gBAV - bP1M(1, ix + iix, iy + iiy))
         DDiffG = Abs(gGAV - bP1M(2, ix + iix, iy + iiy))
         DDiffR = Abs(gRAV - bP1M(3, ix + iix, iy + iiy))
         If DDiffB > Threshold Or DDiffG > Threshold Or DDiffR > Threshold Then
            bP2M(1, ix + iix, iy + iiy) = gBAV
            bP2M(2, ix + iix, iy + iiy) = gGAV
            bP2M(3, ix + iix, iy + iiy) = gRAV
   
            ' Copy back to source memory
            bP1M(1, ix + iix, iy + iiy) = gBAV
            bP1M(2, ix + iix, iy + iiy) = gGAV
            bP1M(3, ix + iix, iy + iiy) = gRAV
         End If
      Next iix
      Next iiy
   
   '---------------------------------------------------------
   End If

Next ix
Next iy

If Not aSelectShapeON Then
   ' 2x2 Edges
   HorzEdges 1
   HorzEdges 2
   HorzEdges PICH
   HorzEdges PICH - 1
   
   VertEdges 1
   VertEdges PICW - 1
   VertEdges PICW
End If

End Sub

Public Sub Despeckle3()

' 3x3 clump

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes


If ASMON Then
   FillASMStruc
   ASM_Caller Despec3
   Exit Sub
End If

For iy = 3 To PICH - 2 Step 3
For ix = 3 To PICW - 2 Step 3
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
      
      ' Sum 12 pixels around block of 4
      gBAV = 0: gGAV = 0: gRAV = 0
      For iiy = -2 To 2
      For iix = -2 To 2
         If Not ((iix = -1 Or iix = 0 Or iix = 1) And _
                 (iiy = -1 Or iiy = 0 Or iiy = 1)) Then
            gBAV = gBAV + bP1M(1, ix + iix, iy + iiy)
            gGAV = gGAV + bP1M(2, ix + iix, iy + iiy)
            gRAV = gRAV + bP1M(3, ix + iix, iy + iiy)
         End If
      Next iix
      Next iiy
      
      gBAV = gBAV \ 16: gGAV = gGAV \ 16: gRAV = gRAV \ 16
      
      For iiy = -1 To 1
      For iix = -1 To 1
         DDiffB = Abs(gBAV - bP1M(1, ix + iix, iy + iiy))
         DDiffG = Abs(gGAV - bP1M(2, ix + iix, iy + iiy))
         DDiffR = Abs(gRAV - bP1M(3, ix + iix, iy + iiy))
         If DDiffB > Threshold Or DDiffG > Threshold Or DDiffR > Threshold Then
            bP2M(1, ix + iix, iy + iiy) = gBAV
            bP2M(2, ix + iix, iy + iiy) = gGAV
            bP2M(3, ix + iix, iy + iiy) = gRAV
            
            ' Copy back to source memory
            bP1M(1, ix + iix, iy + iiy) = gBAV
            bP1M(2, ix + iix, iy + iiy) = gGAV
            bP1M(3, ix + iix, iy + iiy) = gRAV
         End If
      Next iix
      Next iiy

   '---------------------------------------------------------
   End If
Next ix
Next iy

If Not aSelectShapeON Then
   ' 3x3 edges
   HorzEdges 1
   HorzEdges 2
   HorzEdges PICH - 1
   HorzEdges PICH

   VertEdges 1
   VertEdges 2
   VertEdges PICW - 1
   VertEdges PICW
End If

End Sub

Public Sub Despeckle4()

' 1 pixel wide Horizontal & Vertical scratches

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes

If ASMON Then
   FillASMStruc
   ASM_Caller Despec4
   Exit Sub
End If

' Ymemlim  PICH to 1
For iy = PICH To 1 Step -1
   HorzEdges iy
Next iy
' Xmemlim     1 to PICW
For ix = 1 To PICW
   VertEdges ix
Next ix

End Sub

Public Sub Despeckle5()

' 2 pixel wide Horizontal & Vertical scratches

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes

If ASMON Then
   FillASMStruc
   ASM_Caller Despec5
   Exit Sub
End If

For iy = 2 To PICH - 2
For ix = 2 To PICW - 2
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------

      ' Average of above & below
      gBAV = (1& * bP1M(1, ix, iy - 1) + bP1M(1, ix, iy + 2)) \ 2
      gGAV = (1& * bP1M(2, ix, iy - 1) + bP1M(2, ix, iy + 2)) \ 2
      gRAV = (1& * bP1M(3, ix, iy - 1) + bP1M(3, ix, iy + 2)) \ 2
   
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
      
      If BDIFF > Threshold Or GBIFF > Threshold Or RDIFF > Threshold Then
         For iiy = 0 To 1
            bP2M(1, ix, iy + iiy) = gBAV
            bP2M(2, ix, iy + iiy) = gGAV
            bP2M(3, ix, iy + iiy) = gRAV
   
            ' Copy back to source memory
            bP1M(1, ix, iy + iiy) = gBAV
            bP1M(2, ix, iy + iiy) = gGAV
            bP1M(3, ix, iy + iiy) = gRAV
         Next iiy
      End If
      
      ' Average of left & right
      gBAV = (1& * bP1M(1, ix - 1, iy) + bP1M(1, ix + 2, iy)) \ 2
      gGAV = (1& * bP1M(2, ix - 1, iy) + bP1M(2, ix + 2, iy)) \ 2
      gRAV = (1& * bP1M(3, ix - 1, iy) + bP1M(3, ix + 2, iy)) \ 2
   
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
      
         If BIFF > Threshold Or GDIFF > Threshold Or RDIFF > Threshold Then
            For iix = 0 To 1
               bP2M(1, ix + iix, iy) = gBAV
               bP2M(2, ix + iix, iy) = gGAV
               bP2M(3, ix + iix, iy) = gRAV
               
               ' Copy back to source memory
               bP1M(1, ix + iix, iy) = gBAV
               bP1M(2, ix + iix, iy) = gGAV
               bP1M(3, ix + iix, iy) = gRAV
            Next iix
         End If
   '---------------------------------------------------------
   
   End If
Next ix
Next iy

'If Not aSelectShapeON Then
'   HorzEdges 1
'   HorzEdges 2
'   HorzEdges PICH - 1
'   HorzEdges PICH
'
'   VertEdges 1
'   VertEdges 2
'   VertEdges PICW - 1
'   VertEdges PICW
'End If

End Sub

Public Sub Despeckle6()

' 3 pixel wide Horizontal & Vertical scratches

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes


If ASMON Then
   FillASMStruc
   ASM_Caller Despec6
   Exit Sub
End If

For iy = 3 To PICH - 2
For ix = 3 To PICW - 2
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------

      ' Average of above & below
      gBAV = (1& * bP1M(1, ix, iy - 2) + bP1M(1, ix, iy + 2)) \ 2
      gGAV = (1& * bP1M(2, ix, iy - 2) + bP1M(2, ix, iy + 2)) \ 2
      gRAV = (1& * bP1M(3, ix, iy - 2) + bP1M(3, ix, iy + 2)) \ 2
      
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
      
         If BDIFF > Threshold Or GDIFF > Threshold Or RDIFF > Threshold Then
            For iiy = -1 To 1
               bP2M(1, ix, iy + iiy) = gBAV
               bP2M(2, ix, iy + iiy) = gGAV
               bP2M(3, ix, iy + iiy) = gRAV
            
               ' Copy back to source memory
               bP1M(1, ix, iy + iiy) = gBAV
               bP1M(2, ix, iy + iiy) = gGAV
               bP1M(3, ix, iy + iiy) = gRAV
            Next iiy
         End If
      
      ' Average of left & right
      gBAV = (1& * bP1M(1, ix - 2, iy) + bP1M(1, ix + 2, iy)) \ 2
      gGAV = (1& * bP1M(2, ix - 2, iy) + bP1M(2, ix + 2, iy)) \ 2
      gRAV = (1& * bP1M(3, ix - 2, iy) + bP1M(3, ix + 2, iy)) \ 2
   
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
      
         If BDIFF > Threshold Or GDIFF > Threshold Or RDIFF > Threshold Then
            For iix = -1 To 1
               bP2M(1, ix + iix, iy) = gBAV
               bP2M(2, ix + iix, iy) = gGAV
               bP2M(3, ix + iix, iy) = gRAV
            
               ' Copy back to source memory
               bP1M(1, ix + iix, iy) = gBAV
               bP1M(2, ix + iix, iy) = gGAV
               bP1M(3, ix + iix, iy) = gRAV
            Next iix
         End If
   '---------------------------------------------------------
   End If
   
Next ix
Next iy

'If Not aSelectShapeON Then
'   HorzEdges 1
'   HorzEdges 2
'   HorzEdges 3
'   HorzEdges PICH - 2
'   HorzEdges PICH - 1
'   HorzEdges PICH
'
'   VertEdges 1
'   VertEdges 2
'   VertEdges 2
'   VertEdges PICW - 2
'   VertEdges PICW - 1
'   VertEdges PICW
'End If

End Sub

Public Sub HorzEdges(iy)

   For ix = 2 To PICW - 1
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
      gBAV = (1& * bP1M(1, ix - 1, iy) + bP1M(1, ix + 1, iy)) \ 2
      gGAV = (1& * bP1M(2, ix - 1, iy) + bP1M(2, ix + 1, iy)) \ 2
      gRAV = (1& * bP1M(3, ix - 1, iy) + bP1M(3, ix + 1, iy)) \ 2
      
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
         
      If BDIFF > Threshold Or GDIFF > Threshold Or RDIFF > Threshold Then
         bP2M(1, ix, iy) = gBAV
         bP2M(2, ix, iy) = gGAV
         bP2M(3, ix, iy) = gRAV
            
         ' Copy back to source memory
         bP1M(1, ix, iy) = gBAV
         bP1M(2, ix, iy) = gGAV
         bP1M(3, ix, iy) = gRAV
      End If
   End If
Next ix

End Sub

Public Sub VertEdges(ix)

For iy = 2 To PICH - 1
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
      
      gBAV = (1& * bP1M(1, ix, iy - 1) + bP1M(1, ix, iy + 1)) \ 2
      gGAV = (1& * bP1M(2, ix, iy - 1) + bP1M(2, ix, iy + 1)) \ 2
      gRAV = (1& * bP1M(3, ix, iy - 1) + bP1M(3, ix, iy + 1)) \ 2
      
      BDIFF = Abs(gBAV - bP1M(1, ix, iy))
      GDIFF = Abs(gGAV - bP1M(2, ix, iy))
      RDIFF = Abs(gRAV - bP1M(3, ix, iy))
      
      If BDIFF > Threshold Or GDIFF > Threshold Or RDIFF > Threshold Then
         bP2M(1, ix, iy) = gBAV
         bP2M(2, ix, iy) = gGAV
         bP2M(3, ix, iy) = gRAV
            
         ' Copy back to source memory
         bP1M(1, ix, iy) = gBAV
         bP1M(2, ix, iy) = gGAV
         bP1M(3, ix, iy) = gRAV
      End If
   End If
Next iy

End Sub

Public Sub SaltPepper()

' ADD NOISE

' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes

Increment = 80
For i = 1 To 20 * Increment
   ix = PICW * Rnd + 1: If ix > PICW Then ix = PICW
   iy = PICH * Rnd + 1: If iy > PICH Then iy = PICH
   If Rnd < 0.5 Then
      bP2M(1, ix, iy) = 0   ' B
      bP2M(2, ix, iy) = 0   ' G
      bP2M(3, ix, iy) = 0   ' R
   Else
      bP2M(1, ix, iy) = 255 ' B
      bP2M(2, ix, iy) = 255 ' G
      bP2M(3, ix, iy) = 255 ' R
   End If
Next i

End Sub


Public Sub Sharpness()
  
' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes
  
If ASMON Then
   FillASMStruc
   ASM_Caller Sharp
   Exit Sub
End If
'                            S     SS
'mf = 20 '20 '22 '24       '26    '48
'dF = mf - 16 '4  '6  '8    '10    '32

' FilterParam= 1,2 3 or 4
dF = 2 ^ FilterParam ' 2,4 8 or 16 ie shift right in asm
FP = dF + 8

For iy = 2 To PICH - 1
For ix = 2 To PICW - 1
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
  
      ' Sum 8 surrounding pixels
      gBAV = 0: gGAV = 0: gRAV = 0
      For iiy = -1 To 1
      For iix = -1 To 1
         If iix <> 0 Or iiy <> 0 Then
            gBAV = gBAV + bP1M(1, ix + iix, iy + iiy)
            gGAV = gGAV + bP1M(2, ix + iix, iy + iiy)
            gRAV = gRAV + bP1M(3, ix + iix, iy + iiy)
         End If
      Next iix
      Next iiy
      
      gBAV = (FP * bP1M(1, ix, iy) - gBAV) \ dF
      gGAV = (FP * bP1M(2, ix, iy) - gGAV) \ dF
      gRAV = (FP * bP1M(3, ix, iy) - gRAV) \ dF
      'gBAV = (FP * bP1M(1, ix, iy) - 2 * gBAV) \ dF
      'gGAV = (FP * bP1M(2, ix, iy) - 2 * gGAV) \ dF
      'gRAV = (FP * bP1M(3, ix, iy) - 2 * gRAV) \ dF
      If gBAV > 255 Then gBAV = 255
      If gBAV < 0 Then gBAV = 0
      If gGAV > 255 Then gGAV = 255
      If gGAV < 0 Then gGAV = 0
      If gRAV > 255 Then gRAV = 255
      If gRAV < 0 Then gRAV = 0
      bP2M(1, ix, iy) = gBAV
      bP2M(2, ix, iy) = gGAV
      bP2M(3, ix, iy) = gRAV
   '---------------------------------------------------------
   End If
Next ix
Next iy

End Sub

Public Sub Softness()
  
' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes
  
If ASMON Then
   FillASMStruc
   ASM_Caller Soft
   Exit Sub
End If

'mf = FilterParam * 8
'dF = FilterParam + 1

For iy = 2 To PICH - 1
For ix = 2 To PICW - 1
   
   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
   
      ' Sum 8 surrounding pixels
      gBAV = 0: gGAV = 0: gRAV = 0
      For iiy = -1 To 1
      For iix = -1 To 1
         If iix <> 0 Or iiy <> 0 Then
            gBAV = gBAV + bP1M(1, ix + iix, iy + iiy)
            gGAV = gGAV + bP1M(2, ix + iix, iy + iiy)
            gRAV = gRAV + bP1M(3, ix + iix, iy + iiy)
         End If
      Next iix
      Next iiy
      
      gBAV = (bP1M(1, ix, iy) + gBAV \ 8) \ 2
      gGAV = (bP1M(2, ix, iy) + gGAV \ 8) \ 2
      gRAV = (bP1M(3, ix, iy) + gRAV \ 8) \ 2
      
      'If gBAV > 255 Then gBAV = 255
      'If gBAV < 0 Then gBAV = 0
      'If gGAV > 255 Then gGAV = 255
      'If gGAV < 0 Then gGAV = 0
      'If gRAV > 255 Then gRAV = 255
      'If gRAV < 0 Then gRAV = 0
      bP2M(1, ix, iy) = gBAV
      bP2M(2, ix, iy) = gGAV
      bP2M(3, ix, iy) = gRAV

   '---------------------------------------------------------
   
   End If
Next ix
Next iy


If FilterParam > 1 Then  ' re-sample from bP2M()

   For nn = 2 To FilterParam
   
      For iy = 2 To PICH - 1
      For ix = 2 To PICW - 1
         
         ' Check if underlying area is black ie selected area
         If bPPM(1, ix, iy) = 0 And _
            bPPM(2, ix, iy) = 0 And _
            bPPM(3, ix, iy) = 0 Then
         '---------------------------------------------------------
         
            ' Sum 8 surrounding pixels
            gBAV = 0: gGAV = 0: gRAV = 0
            For iiy = -1 To 1
            For iix = -1 To 1
               If iix <> 0 Or iiy <> 0 Then
                  gBAV = gBAV + bP2M(1, ix + iix, iy + iiy)
                  gGAV = gGAV + bP2M(2, ix + iix, iy + iiy)
                  gRAV = gRAV + bP2M(3, ix + iix, iy + iiy)
               End If
            Next iix
            Next iiy
            
            gBAV = (bP2M(1, ix, iy) + gBAV \ 8) \ 2
            gGAV = (bP2M(2, ix, iy) + gGAV \ 8) \ 2
            gRAV = (bP2M(3, ix, iy) + gRAV \ 8) \ 2
            
            'If gBAV > 255 Then gBAV = 255
            'If gBAV < 0 Then gBAV = 0
            'If gGAV > 255 Then gGAV = 255
            'If gGAV < 0 Then gGAV = 0
            'If gRAV > 255 Then gRAV = 255
            'If gRAV < 0 Then gRAV = 0
            bP2M(1, ix, iy) = gBAV
            bP2M(2, ix, iy) = gGAV
            bP2M(3, ix, iy) = gRAV
      
         '---------------------------------------------------------
         
         End If
      Next ix
      Next iy
   
   Next nn

End If

End Sub


Public Sub Brightness()
  
' Brighter mdiv =  30
' Darker   mdiv = -30
  
' bP2M <<--<<-- bP1M
CopyMemory bP2M(1, 1, 1), bP1M(1, 1, 1), ImageBytes  'PICH * ScanLineBytes
  
If ASMON Then
   FillASMStruc
   If FilterParam > 0 Then
     ASM_Caller Bright
   End If
   Exit Sub
End If

' FilterParam = 1,2,3,4,5,6,7,8

   If FilterParam < 5 Then
      FP = 2 ^ (FilterParam + 1)
      FP = -FP
   Else
      FP = 8 - FilterParam
      FP = 2 ^ (FP + 1)
   End If

For iy = PICH To 1 Step -1
For ix = PICW To 1 Step -1

   ' Check if underlying area is black ie selected area
   If bPPM(1, ix, iy) = 0 And _
      bPPM(2, ix, iy) = 0 And _
      bPPM(3, ix, iy) = 0 Then
   '---------------------------------------------------------
         gBAV = bP1M(1, ix, iy) + (bP1M(1, ix, iy) \ FP)
         gGAV = bP1M(2, ix, iy) + (bP1M(2, ix, iy) \ FP)
         gRAV = bP1M(3, ix, iy) + (bP1M(3, ix, iy) \ FP)
         If gBAV > 255 Then gBAV = 255
         If gBAV < 0 Then gBAV = 0
         If gGAV > 255 Then gGAV = 255
         If gGAV < 0 Then gGAV = 0
         If gRAV > 255 Then gRAV = 255
         If gRAV < 0 Then gRAV = 0
         bP2M(1, ix, iy) = gBAV
         bP2M(2, ix, iy) = gGAV
         bP2M(3, ix, iy) = gRAV
   '---------------------------------------------------------
   
   End If
Next ix
Next iy

End Sub

