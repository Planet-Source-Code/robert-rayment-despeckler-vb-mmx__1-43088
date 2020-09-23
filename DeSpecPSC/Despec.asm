;Despec.bin  NASM  by  Robert Rayment Feb 2003
; NB Assumes MMX present. cpuid can be used to
; to test for MMX if wanted.

; Contents          Approx line numbers
; Despec1  1 pix            Line 176
; Despec2  2x2 pix          Line 281
; Despec3  3x3 pix          Line 457
; Despec4  1 pix scratch    Line 680
; Despec5  2 pix scratch    Line 710
; Despec6  3 pix scratch    Line 854
; Sharpness                 Line 1008
; Softness                  Line 1098
; Brightness                Line 1268
; HorzEdges                 Line 1358
; VertEdges                 Line 1448
; Sum8              Line 1535
; Sum12             Line 1570
; Sum16             Line 1632
; TestThreshold     Line 1710
; Jump table        Line 1757

; VB
;Type ASMVars
;   PICH As Long
;   PICW As Long
;   ptrP1M As Long
;   ptrP2M As Long
;   ptrPPM As Long
;   Threshold As Long
;   aSelectShapeON As Long   ' 0 False
;   FilterParam As Long
;End Type
;Public ASMStruc As ASMVars
;
;Public ptrMC, ptrStruc        ' Ptrs to MCode & Structure
;Public bDespecMC() As Byte    ' Byte array to hold MCode
;
;Public Const Despec1 As Long = 0
;Public Const Despec2 As Long = 1
;Public Const Despec3 As Long = 2
;Public Const Despec4 As Long = 3
;Public Const Despec5 As Long = 4
;Public Const Despec6 As Long = 5
;Public Const Sharp As Long = 6
;Public Const Soft As Long = 7
;Public Const Bright As Long = 8
;
;Public ASMON As Boolean
;
;Public Sub FillASMStruc()
;With ASMStruc
;   .PICH = PICH
;   .PICW = PICW
;   .ptrP1M = VarPtr(bP1M(1, 1, 1))
;   .ptrP2M = VarPtr(bP2M(1, 1, 1))
;   .ptrPPM = VarPtr(bPPM(1, 1, 1))
;   .Threshold = Threshold
;   .aSelectShapeON = CLng(aSelectShapeON)   ' 0 False
;   .FilterParam = FilterParam
;End With
;End Sub
;
;res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, OpCode)
;                             8         12  16  20

%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%define PICH          [ebp-4]
%define PICW          [ebp-8]
%define ptrP1M        [ebp-12]
%define ptrP2M        [ebp-16]
%define ptrPPM        [ebp-20]
%define Threshold     [ebp-24]
%define SelectShapeON [ebp-28]
%define ZeroFilterParam    [ebp-32]

; Some variables
%define FilterParam        [ebp-36]
%define Zero2    [ebp-40]
%define ZeroFP   [ebp-44]
%define FP       [ebp-48]    
%define dF       [ebp-52]    

%define kx       [ebp-56]    
%define ky       [ebp-60]

%define ptmc     [ebp-64]   ; ptr mcode buffer

%define pPPM     [ebp-68]
%define PICW4    [ebp-72]
%define pPPMorg  [ebp-76]


[bits 32]

    ; Get absolute address to mcode buffer in VB
    Call BufferADDR     ; 5B
    sub eax,5           ; -5B for CallBufferADDR

    push ebp
    mov ebp,esp
    sub esp,76
    push edi
    push esi
    push ebx
    
    mov ptmc,eax


    ;Fill structure
    mov ebx,[ebp+8]
    movab PICH,           [ebx]
    movab PICW,           [ebx+4]
    movab ptrP1M,         [ebx+8]
    movab ptrP2M,         [ebx+12]
    movab ptrPPM,         [ebx+16]
    movab Threshold,      [ebx+20]
    movab SelectShapeON,  [ebx+24]
    movab FilterParam,    [ebx+28]
;----------------------------

    mov ebx,PICW
    shl ebx,2
    mov PICW4,ebx

    ; Make mm6 = 4 Threshold words TH TH TH TH
    mov eax,Threshold
    mov edx,eax
    shl edx,16
    add eax,edx
    movd mm6,eax
    movq mm7,mm6
    punpckldq mm6,mm7   ;mm6 = TH TH TH TH

    ; Zero dwords above for psrlw mm# FilterPram & FP
    xor eax,eax
    mov ZeroFilterParam,eax
    mov ZeroFP,eax

    ; Using jump table
    mov ebx,Sub0
    mov eax,[ebp+20]    ;OpCode
    cmp eax,9
    jge GETOUT           ; Error
    shl eax,2           ; x4
    add ebx,eax
    add ebx,ptmc
    mov eax,[ebx]
    add eax,ptmc
    Call eax

GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

;========================================
; Get mcode buffer absolute start address
BufferADDR:
    pop eax
    push eax
RET
;######################################################################

Despec1:

    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx

    mov ebx, PICW4

    mov ecx,1
iyDS1:
    push ecx
    mov ecx,1
ixDS1:
    push ecx

    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixDS1   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum8       ; mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    psrlw mm0,3     ; w\8 mm0=|AvA|AvR|AvG|AvR|
    
    movq mm4,mm0    ; mm4=|AvA|AvR|AvG|AvR|
    
    movd mm1,[edi+ebx+4]    ;P1M(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    movq mm2,mm0    ; mm2=|AvA|AvR|AvG|AvR|
    psubusw mm0,mm1
    psubusw mm1,mm2
    por mm0,mm1     ; mm0=ABS |AvA-A|AvR-R|AvG-G|AvB-B|
    
    pcmpgtw mm0,mm6     ; is data in mm0 > TH
    packsswb mm0,mm0    ; pack result into mm0
    movd eax,mm0
    cmp eax,0
    je NexixDS1         ; false all data <= TF

DS1Replace:

    packuswb mm4,mm7    ; mm4 = | |ARGB|
    movd[esi+ebx+4],mm4
    movd[edi+ebx+4],mm4

;;;;;;;;;;;;;;;;;;
NexixDS1:
    add edi,4
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    cmp ecx,eax     ;ix-(PICW-1)
    jl near ixDS1
NexiyDS1:
    add edi,8
    add esi,8
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    cmp ecx,eax     ;iy-(PICH-1)
    jl near iyDS1

    emms

; EDGES
        mov eax,SelectShapeON
        cmp eax,0
        jne DS1ret  ; Not False Select shape On ie no edges

        mov eax,1       ; ky=1
        mov ky,eax
        Call HorzEdges
        mov eax,PICH
        mov ky,eax      ; ky=PICH
        Call HorzEdges

        mov eax,1
        mov kx,eax      ; ix=1
        Call VertEdges
        mov eax,PICW
        mov kx,eax      ; ix=PICW
        Call VertEdges

DS1ret:
    emms

RET
;=============================================================

Despec2:

    ; num\3 = (num * .3333 * 64)/64 == (num * 21) shr 6

    mov eax,21
    mov edx,eax
    shl edx,16
    add eax,edx
    movd mm5,eax
    movq mm4,mm5
    punpckldq mm5,mm4   ; Multiplier mm5 = 21 21 21 21



    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx
    mov pPPMorg,edx

    mov ebx, PICW4

    mov ecx,1
iyDS2:
    push ecx
    push edi
    push esi
    mov edx,pPPM
    mov pPPMorg,edx
   
    mov ecx,1
ixDS2:
    push ecx
    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixDS2   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum12      ; mm0=|SA|SumR|SumG|SumB|
    ; max Sum = 12*255 = 3060
    ; max multiplier for words = 65536\3060 = 21
    ; \12
    pmullw mm0,mm5      ; mm1=|21*A|21*R|21*G|21*B|
    psrlw mm0,8         ; mm1= w\(64*4) == mm0 w\12
    movq mm4,mm0        ; also mm4=RGBAv

    push edi
    push esi

DS2DoIXIY:
    add edi,ebx
    add edi,4           ; edi->P1M(1,ix,iy)
    add esi,ebx
    add esi,4           ; esi->P2M(1,ix,iy)
    
    CALL TestThreshold
;-----------------

DS2DoIXp1IY:
    add edi,4       ; ix+1,iy
    add esi,4

    CALL TestThreshold
;-----------------

DS2DoIXIYp1:
    pop esi
    pop edi
    push edi
    push esi

    add edi,ebx
    add edi,ebx
    add edi,4       ; ix,iy+1
    add esi,ebx
    add esi,ebx
    add esi,4

    CALL TestThreshold
;-----------------

DS2DoIXp1IYp1:
    add edi,4       ; ix+1,iy+1
    add esi,4

    CALL TestThreshold
;-----------------

DS2pop:
    pop esi
    pop edi

;;;;;;;;;;;;;;;;;;
NexixDS2:
    add edi,8
    add esi,8           ; ix+2
    
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    inc ecx
    mov eax,PICW
    dec eax             ; PICW-1
    dec eax             ; PICW-2
    cmp ecx,eax
    jl near ixDS2

NexiyDS2:
    pop esi
    pop edi
    
    add edi,ebx
    add edi,ebx     ; iy+2
    add esi,ebx
    add esi,ebx
    
    mov edx,pPPMorg
    add edx,ebx
    add edx,ebx
    mov pPPM,edx
    mov pPPMorg,edx

    pop ecx
    inc ecx
    inc ecx
    mov eax,PICH
    dec eax             ; PICH-1
    dec eax             ; PICH-2
    cmp ecx,eax
    jl near iyDS2
    emms

; EDGES
        mov eax,SelectShapeON
        cmp eax,0
        jne DS2ret  ; Not False Select shape On ie no edges

        mov eax,1
        mov ky,eax          ; iy=1
        Call HorzEdges
        
        mov eax,2
        mov ky,eax          ; iy=2
        Call HorzEdges
        
        mov eax,PICH
        mov ky,eax          ; iy=PICH
        Call HorzEdges

        mov eax,PICH
        dec eax
        mov ky,eax          ; iy=PICH-1
        Call HorzEdges

        mov eax,1
        mov kx,eax          ; ix=1
        Call VertEdges
        
        mov eax,PICW
        dec eax             ; ix=PICW-1
        mov kx,eax
        Call VertEdges
        
        mov eax,PICW        ; ix=PICW
        mov kx,eax
        Call VertEdges

DS2ret:

    emms

RET
;=============================================================

Despec3:

    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx
    mov pPPMorg,edx

    mov ebx, PICW4

    mov ecx,1
iyDS3:
    push ecx
    push edi
    push esi
    mov edx,pPPM
    mov pPPMorg,edx

    mov ecx,1
ixDS3:
    push ecx

    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixDS3   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum16      ; mm0=|SA|SumR|SumG|SumB|
    psrlw mm0,4     ; w\16 mm0=|AvA|AvR|AvG|AvR|
    movq mm4,mm0        ; also mm4=RGBAv


    push edi
    push esi

DS3DoIXIY:  ; ROW 1
    add edi,ebx
    add edi,4           ; edi->P1M(1,ix,iy)
    add esi,ebx
    add esi,4           ; esi->P2M(1,ix,iy)
    
    CALL TestThreshold
;-----------------

DS3DoIXp1IY:
    add edi,4       ; ix+1,iy
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXp2IY:
    add edi,4       ; ix+2,iy
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXIYp1:    ; ROW 2
    pop esi
    pop edi
    push edi
    push esi

    add edi,ebx
    add edi,ebx
    add edi,4       ; ix,iy+1
    add esi,ebx
    add esi,ebx
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXp1IYp1:
    add edi,4       ; ix+1,iy+1
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXp2IYp1:
    add edi,4       ; ix+2,iy+1
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXIYp2:    ; ROW 3
    pop esi
    pop edi
    push edi
    push esi

    add edi,ebx
    add edi,ebx
    add edi,ebx
    add edi,4       ; ix,iy+2
    add esi,ebx
    add esi,ebx
    add esi,ebx
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXp1IYp2:
    add edi,4       ; ix+1,iy+2
    add esi,4

    CALL TestThreshold
;-----------------

DS3DoIXp2IYp2:
    add edi,4       ; ix+2,iy+2
    add esi,4

    CALL TestThreshold
;-----------------

DS3pop:
    pop esi
    pop edi

;;;;;;;;;;;;;;;;;;
NexixDS3:
    add edi,8
    add esi,8           ; ix+2
    
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    inc ecx
    mov eax,PICW
    dec eax             ; PICW-1
    dec eax             ; PICW-2
    dec eax             ; PICW-3
    cmp ecx,eax
    jl near ixDS3

NexiyDS3:
    pop esi
    pop edi
    
    add edi,ebx
    add edi,ebx     ; iy+2
    add esi,ebx
    add esi,ebx
    
    mov edx,pPPMorg
    add edx,ebx
    add edx,ebx
    mov pPPM,edx
    mov pPPMorg,edx

    pop ecx
    inc ecx
    inc ecx
    mov eax,PICH
    dec eax             ; PICH-1
    dec eax             ; PICH-2
    dec eax             ; PICH-3
    cmp ecx,eax
    jl near iyDS3
    emms

; EDGES
        mov eax,SelectShapeON
        cmp eax,0
        jne DS3ret  ; Not False Select shape On ie no edges

        mov eax,1           ; ky= 1
        mov ky,eax
        Call HorzEdges
        
        mov eax,2           ; ky=2
        mov ky,eax
        Call HorzEdges
        
        mov eax,PICH
        dec eax             ; ky=PICH-2
        dec eax
        mov ky,eax
        Call HorzEdges


        mov eax,PICH
        dec eax             ; PICH-1
        mov ky,eax
        Call HorzEdges

        mov eax,PICH        ; PICH
        mov ky,eax
        Call HorzEdges

        mov eax,1           ; ix=1
        mov kx,eax
        Call VertEdges

        mov eax,2           ; ix=2
        mov kx,eax
        Call VertEdges

        mov eax,PICW
        dec eax             ; ix=PICW-1
        mov kx,eax
        Call VertEdges

        mov eax,PICW
        mov kx,eax          ; ix=PICW
        Call VertEdges

DS3ret:

    emms

RET
;=============================================================

Despec4:    ;1 pixel scratches

    ; Horz
    mov ecx,PICH
ForHDS4:
    push ecx
    mov ky,ecx
    Call HorzEdges
NexHDS4:
    pop ecx
    dec ecx
    jnz ForHDS4

    ; Vert
    mov ecx,1
ForVDS4:
    push ecx
    mov kx,ecx
    Call VertEdges
NexVDS4:
    pop ecx
    inc ecx
    cmp ecx,PICW
    jle ForVDS4

    emms    

RET
;=============================================================

Despec5:    ; 2 pixel scratches

    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx
    mov pPPMorg,edx

    mov ebx, PICW4
    pxor mm7,mm7

    mov ecx,1
iyDS5:
    push ecx

    push edi
    push esi
    mov edx,pPPM
    mov pPPMorg,edx

    mov ecx,1
ixDS5:
    push ecx

    push edi
    push esi

    mov edx,pPPM
    mov eax,[edx+ebx+4] ; ix+1,iy+1
    cmp eax,0
    jne near NexixDS5   ; Not black
    ;;;;;;;;;;;;;;;;;;

    ; ABOVE & BELOW - HORIZONTAL SCRATCGES
    
;    push edi
;    push esi
    
    movd mm0,[edi]      ; y
    add edi,ebx
    add edi,ebx
    add edi,ebx
    movd mm1,[edi]      ; y+3
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1     ; w\2 mm0=|AvA|AvR|AvG|AvR|
    
    movq mm4,mm0    ; mm4=|AvA|AvR|AvG|AvR|

    pop esi
    pop edi         ; y
    push edi
    push esi
    
    add edi,ebx     ; y+1
    add esi,ebx
    
    Call TestThreshold

    add edi,ebx     ; y+2
    add esi,ebx
    
    Call TestThreshold

    pop esi
    pop edi     ; x,y
    push edi
    push esi

    ; LEFT & RIGHT - VERTICAL SCRATCGES

    movd mm0,[edi]      ; y
    add edi,12          ; x+3
    movd mm1,[edi] 

    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1     ; w\2 mm0=|AvA|AvR|AvG|AvR|
    
    movq mm4,mm0    ; mm4=|AvA|AvR|AvG|AvR|

    pop esi
    pop edi
    push edi
    push esi


    add esi,4       ; x+1
    add edi,4

    Call TestThreshold

    add esi,4       ; x+2,y
    add edi,4

    Call TestThreshold

;;;;;;;;;;;;;;;;;;
NexixDS5:
    pop esi
    pop edi

    add edi,4       ; x+1,y
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    dec eax
    dec eax
    cmp ecx,eax     ;ix-(PICW-3)
    jl near ixDS5
NexiyDS5:

    pop esi
    pop edi
    
    add edi,ebx     ; iy+1
    add esi,ebx
    
    mov edx,pPPMorg
    add edx,ebx
    mov pPPM,edx
    mov pPPMorg,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    dec eax
    dec eax
    cmp ecx,eax     ;iy-(PICH-3)
    jl near iyDS5

    emms
RET
;=============================================================

Despec6:    ; 3 pixel scratches


    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx
    mov pPPMorg,edx

    mov ebx, PICW4
    pxor mm7,mm7

    mov ecx,1
iyDS6:
    push ecx
    push edi
    push esi
    mov edx,pPPM
    mov pPPMorg,edx

    mov ecx,1
ixDS6:
    push ecx

    push edi
    push esi

    mov edx,pPPM
    mov eax,[edx+ebx+4] ; ix+1,iy+1
    cmp eax,0
    jne near NexixDS6   ; Not black
    ;;;;;;;;;;;;;;;;;;

    ; ABOVE & BELOW - HORIZONTAL SCRATCGES
    
;    push edi
;    push esi
    
    movd mm0,[edi]      ; y
    add edi,ebx
    add edi,ebx
    add edi,ebx
    add edi,ebx
    movd mm1,[edi]      ; y+4
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1     ; w\2 mm0=|AvA|AvR|AvG|AvR|
    
    movq mm4,mm0    ; mm4=|AvA|AvR|AvG|AvR|

    pop esi
    pop edi         ; y
    push edi
    push esi
    
    add edi,ebx     ; y+1
    add esi,ebx
    
    Call TestThreshold

    add edi,ebx     ; y+2
    add esi,ebx
    
    Call TestThreshold

    add edi,ebx     ; y+3
    add esi,ebx
    
    Call TestThreshold

    pop esi
    pop edi     ; x,y
    push edi
    push esi

    ; LEFT & RIGHT - VERTICAL SCRATCHES

    movd mm0,[edi]      ; y
    add edi,16          ; x+4
    movd mm1,[edi] 

    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1     ; w\2 mm0=|AvA|AvR|AvG|AvR|
    
    movq mm4,mm0    ; mm4=|AvA|AvR|AvG|AvR|

    pop esi
    pop edi
    push edi
    push esi


    add esi,4       ; x+1
    add edi,4

    Call TestThreshold

    add esi,4       ; x+2,y
    add edi,4

    Call TestThreshold

    add esi,4       ; x+3,y
    add edi,4

    Call TestThreshold

;;;;;;;;;;;;;;;;;;
NexixDS6:
    pop esi
    pop edi
    add edi,4       ; x+1,y
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    dec eax
    dec eax
    cmp ecx,eax     ;ix-(PICW-3)
    jl near ixDS6
NexiyDS6:

    pop esi
    pop edi
    
    add edi,ebx     ; iy+1
    add esi,ebx
    
    mov edx,pPPMorg
    add edx,ebx
    mov pPPM,edx
    mov pPPMorg,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    dec eax
    dec eax
    cmp ecx,eax     ;iy-(PICH-3)
    jl near iyDS6

    emms
RET
;=============================================================

Sharpness:

    mov ecx,FilterParam ; 1,2,3,4
    mov eax,1
    shl eax,CL          ; \dF = 2^FilterParam = 2,4,8,16
    add eax,8
    mov edx,eax
    shl edx,16
    add eax,edx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5   ; Multiplier mm4 = FP FP FP FP

    mov eax,0FFFFFFh
    movd mm5,eax        ; mm5 to mask out Alpha
    
    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx

    mov ebx, PICW4

    mov ecx,1
iySharp:
    push ecx
    mov ecx,1
ixSharp:
    push ecx

    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixSharp   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum8       ; mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    movd mm1,[edi+ebx+4]    ;P1M(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    ; mm4 = FP FP FP FP
    ; mm5 = mask to GET RGB

    pmullw mm1,mm4          ; mm1=|FP*A|FP*R|FP*G|FP*B|
    psubusw mm1,mm0         ; mm1= FPARGB - SumARGB
    
    psrlw mm1,FilterParam   ; \dF   NB ZeroFilterParam must = 0
    
    packuswb mm1,mm7        ; | |ARGB|          
    pand mm1,mm5            ; Mask out alpha

    movd[esi+ebx+4],mm1     ; Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

;;;;;;;;;;;;;;;;;;
NexixSharp:
    add edi,4
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    cmp ecx,eax     ;ix-(PICW-1)
    jl near ixSharp
NexiySharp:
    add edi,8
    add esi,8
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    cmp ecx,eax     ;iy-(PICH-1)
    jl near iySharp

    emms

RET
;=============================================================

Softness:

    mov eax,0FFFFFFh
    movd mm5,eax        ; mm5 to mask out Alpha
    
    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx

    mov ebx, PICW4

    mov ecx,1
iySoft:
    push ecx
    mov ecx,1
ixSoft:
    push ecx

    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixSoft   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum8       ; mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    psrlw mm0,3     ; \8     mm0 = AVA AvR AvG AvB

    movd mm1,[edi+ebx+4]    ;P1M(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    paddw mm1,mm0           ; mm1= FPARGB + AvARGB
    
    psrlw mm1,1             ; \2
    
    packuswb mm1,mm7        ; | |ARGB|          

    ; mm5 = mask to GET RGB

    pand mm1,mm5            ; Mask out alpha

    movd[esi+ebx+4],mm1     ; Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

;;;;;;;;;;;;;;;;;;
NexixSoft:
    add edi,4
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    cmp ecx,eax     ;ix-(PICW-1)
    jl near ixSoft
NexiySoft:
    add edi,8
    add esi,8
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    cmp ecx,eax     ;iy-(PICH-1)
    jl near iySoft


    mov eax,FilterParam
    cmp eax,1
    je near SoftRET
    ;====================
    mov ecx,eax
ForNN:
    push ecx
    ;====================

    ;mov edi,ptrP1M
    
    mov edi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx

    mov ebx, PICW4

    mov ecx,1
iySoftNN:
    push ecx
    mov ecx,1
ixSoftNN:
    push ecx

    mov edx,pPPM
    mov eax,[edx+ebx+4]
    cmp eax,0
    jne near NexixSoftNN   ; Not black
    ;;;;;;;;;;;;;;;;;;

    Call Sum8       ; mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    psrlw mm0,3     ; \8     mm0 = AVA AvR AvG AvB

    movd mm1,[edi+ebx+4]    ;P1M(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    paddw mm1,mm0           ; mm1= FPARGB + AvARGB
    
    psrlw mm1,1             ; \2
    
    packuswb mm1,mm7        ; | |ARGB|          

    ; mm5 = mask to GET RGB

    pand mm1,mm5            ; Mask out alpha

    movd[edi+ebx+4],mm1     ; Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

;;;;;;;;;;;;;;;;;;
NexixSoftNN:
    add edi,4
    ;add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICW
    dec eax
    cmp ecx,eax     ;ix-(PICW-1)
    jl near ixSoftNN
NexiySoftNN:
    add edi,8
    ;add esi,8
    mov edx,pPPM
    add edx,8
    mov pPPM,edx

    pop ecx
    inc ecx
    mov eax,PICH
    dec eax
    cmp ecx,eax     ;iy-(PICH-1)
    jl near iySoftNN

    ;====================

NexNN:
    pop ecx
    dec ecx
    jnz near ForNN
    ;====================

SoftRET:
    emms

RET
;=============================================================

Brightness:
    
    xor eax,eax
    mov FP,eax

    mov eax,dword FilterParam
    cmp eax,5
    jge GE5
    
    inc eax
    jmp StoreFP
GE5:
    mov eax,10
    sub eax,FilterParam
StoreFP:
    mov dword FP,eax        ; shr = 2,3,4,5 or 5,4,3,2

    mov eax,0FFFFFFh
    movd mm5,eax        ; mm5 to mask out Alpha
    
    mov edi,ptrP1M
    mov esi,ptrP2M
    mov edx,ptrPPM
    mov pPPM,edx

    xor edx,edx
    mov eax,PICH
    mov ebx,PICW
    mul ebx
    mov ecx,eax

    mov ebx, PICW4
    pxor mm7,mm7


ForBright:
    push ecx

    mov edx,pPPM
    mov eax,[edx]
    cmp eax,0
    jne near NexBright   ; Not black
    ;;;;;;;;;;;;;;;;;;

    movd mm1,[edi]          ;P1M(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|

    movq mm2,mm1            ; mm2=|A|R|G|B|

    psrlw mm2,FP        ; shr 5,4,3,2 \32,16,8,4
                        ; NB ZeroFP must be = 0

shrdone:
    mov eax,FilterParam
    cmp eax,5
    jge GE52
    
    psubusw mm1,mm2         ; ARGB - ARGB\FilterParam
    jmp NowPack
GE52:
    paddusw mm1,mm2         ; ARGB + ARGB\FilterParam

NowPack:
    
    packuswb mm1,mm7        ; | |ARGB|          

    ; mm5 = mask to GET RGB

    pand mm1,mm5            ; Mask out alpha

    movd[esi],mm1       ; Stalling sequence??
                        ; would need to unroll loop
                        ; to take 2 pixel in more mms
;;;;;;;;;;;;;;;;;;
NexBright:
    add edi,4
    add esi,4
    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    pop ecx
    dec ecx
    jnz near ForBright

    emms

RET
;=============================================================

HorzEdges:  ; In ky=iy

    xor edx,edx
    mov eax,ky
    dec eax
    mov ebx,PICW4
    mul ebx         ; Line Offset eax = PICW4*(ky-1)

    mov edi,ptrP1M
    add edi,eax     ; edi-> bP1M(1,1,iy)
    mov esi,ptrP2M
    add esi,eax     ; esi-> bP2M(1,1,iy)
    mov ebx, PICW4

    mov edx,ptrPPM
    add edx,eax
    mov pPPM,edx


    mov ecx,2

ForH:

    mov edx,pPPM
    mov eax,[edx+4]
    cmp eax,0
    jne near NexixH   ; Not black
    ;;;;;;;;;;;;;;;;;;

    ; RGBAV
    pxor mm7,mm7
    movd mm0,[edi]
    movd mm1,[edi+8]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1         ; mm0 = RGB AV
    movq mm4,mm0        ; mm4 = RGB AV

    ; RGBDiff           ; RGBIFF <= Threshold
    movd mm2,[edi+4]    ; P1M(1,ix,iy)
    punpcklbw mm2,mm7       ; A R G B
    movq mm3,mm0            ; mm0 = RGB AV
    psubusw mm0,mm2         ; RGBAV-RGB
    psubusw mm2,mm3         ; RGB-RGBAV
    por mm0,mm2             ; mm0 = ABS(ARGBAV-ARGB)
    
    movd eax,mm0    ; AvG-G|AvB-B
    and eax,0FFFFh  ; AvB-B
    cmp eax,Threshold
    jg HReplace

    movd eax,mm0
    shr eax,16      ; AvG-G
    cmp eax,Threshold
    jg HReplace

    psrlq mm0,32
    movd eax,mm0    ; AvA-A|AvR-R
    and eax,0FFFFh  ; AvR-R
    cmp eax,Threshold
    jle near NexixH

HReplace:

    packuswb mm4,mm7    ; mm4 = | |ARGB|
    movd[esi+4],mm4
    movd[edi+4],mm4

;;;;;;;;;;;;;;;;;;
NexixH:

    add edi,4
    add esi,4

    mov edx,pPPM
    add edx,4
    mov pPPM,edx

    inc ecx
    mov eax,PICW
    dec eax             ; PICW-1
    cmp ecx,eax
    jl near ForH
    
    emms

RET
;=============================================================
    
VertEdges:  ;In kx=ix = 1,2    PICW-1, PICW
    
    mov eax,kx
    dec eax
    shl eax,2       ; 4*(kx-1)

    mov edi,ptrP1M
    add edi,eax     ; edi-> bP1M(1,kx,1)
    mov esi,ptrP2M
    add esi,eax     ; esi-> bP2M(1,kx,1)
    mov ebx, PICW4

    mov edx,ptrPPM
    add edx,eax
    mov pPPM,edx

    mov ecx,2

ForV:

    mov edx,pPPM
    mov eax,[edx+ebx]
    cmp eax,0
    jne near NexiyV   ; Not black
    ;;;;;;;;;;;;;;;;;;

    ; RGBAV
    pxor mm7,mm7
    movd mm0,[edi]
    movd mm1,[edi+2*ebx]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    paddusw mm0,mm1
    psrlw mm0,1         ; mm0 = RGB AV
    movq mm4,mm0        ; mm4 = RGB AV

    ; RGBDiff           ; RGBIFF <= Threshold
    movd mm2,[edi+ebx]  ; P1M(1,ix,iy)
    punpcklbw mm2,mm7       ; A R G B
    movq mm3,mm0            ; mm0 = RGB AV
    psubusw mm0,mm2         ; RGBAV-RGB
    psubusw mm2,mm3         ; RGB-RGBAV
    por mm0,mm2             ; mm0 = ABS(ARGBAV-ARGB)
    
    movd eax,mm0    ; AvG-G|AvB-B
    and eax,0FFFFh  ; AvB-B
    cmp eax,Threshold
    jg VReplace

    movd eax,mm0
    shr eax,16      ; AvG-G
    cmp eax,Threshold
    jg VReplace

    psrlq mm0,32
    movd eax,mm0    ; AvA-A|AvR-R
    and eax,0FFFFh  ; AvR-R
    cmp eax,Threshold
    jle near NexiyV

VReplace:

    packuswb mm4,mm7    ; mm4 = | |ARGB|
    movd[esi+ebx],mm4
    movd[edi+ebx],mm4

;;;;;;;;;;;;;;;;;;
NexiyV:
    add edi,ebx
    add esi,ebx

    mov edx,pPPM
    add edx,ebx
    mov pPPM,edx


    inc ecx
    mov eax,PICH
    dec eax             ; PICH-1
    cmp ecx,eax
    jl near ForV
    
    emms

RET
;====================================================

Sum8:   ; In edi->BotLeft, ebx PICW*4
        ; Out: mm0 = A% R% G% B% lob
    pxor mm7,mm7

    ; Row 1
    movd mm0,[edi]
    movd mm1,[edi+4]
    movd mm2,[edi+8]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    
    ; Row 2
    movd mm1,[edi+ebx]
    movd mm2,[edi+ebx+8]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    
    
    ; Row 3
    movd mm1,[edi+2*ebx]
    movd mm2,[edi+2*ebx+4]
    movd mm3,[edi+2*ebx+8]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    paddusw mm0,mm3 ; mm0 = A% R% G% B% lob
RET
;====================================================

Sum12:  ; In edi->BotLeft, ebx PICW*4

    push edi

    pxor mm7,mm7

    ; Row 1
    movd mm0,[edi]
    movd mm1,[edi+4]
    movd mm2,[edi+8]
    movd mm3,[edi+12]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    add edi,ebx

    paddusw mm0,mm3

    ; Row 2
    movd mm1,[edi]
    movd mm2,[edi+12]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    add edi,ebx

    paddusw mm0,mm2
    
    
    ; Row 3
    movd mm1,[edi]
    movd mm2,[edi+12]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    add edi,ebx

    paddusw mm0,mm2
  
    ; Row 4
    movd mm1,[edi]
    movd mm2,[edi+4]
    movd mm3,[edi+8]
    movd mm4,[edi+12]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    punpcklbw mm4,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    paddusw mm0,mm3
    paddusw mm0,mm4 ; A% R% G% B% lob

    pop edi

RET
;====================================================

Sum16:  ; In edi->BotLeft, ebx PICW*4
    
    push edi

    pxor mm7,mm7

    ; Row 1
    movd mm0,[edi]
    movd mm1,[edi+4]
    movd mm2,[edi+8]
    movd mm3,[edi+12]
    movd mm4,[edi+16]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    punpcklbw mm4,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    paddusw mm0,mm3
    add edi,ebx

    paddusw mm0,mm4

    ; Row 2
    movd mm1,[edi]
    movd mm2,[edi+16]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    add edi,ebx

    paddusw mm0,mm2
    
    
    ; Row 3
    movd mm1,[edi]
    movd mm2,[edi+16]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    add edi,ebx

    paddusw mm0,mm2

    
    ; Row 4
    movd mm1,[edi]
    movd mm2,[edi+16]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    add edi,ebx

    paddusw mm0,mm2

    ; Row 5
    movd mm1,[edi]
    movd mm2,[edi+4]
    movd mm3,[edi+8]
    movd mm4,[edi+12]
    movd mm5,[edi+16]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    punpcklbw mm4,mm7
    punpcklbw mm5,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    paddusw mm0,mm3
    paddusw mm0,mm4
    paddusw mm0,mm5 ; A% R% G% B% lob


    pop edi
RET
;====================================================

TestThreshold:   ; TEST THRESHOLD In:  mm0=mm4=|AvA|AvR|AvG|AvB|
                 ;          &  edi,esi -> P1M(ix,iy), P2M(ix,iy)
                 ;          &  mm6 = TH TH TH TH

    movd mm1,[edi]
    punpcklbw mm1,mm7
    movq mm2,mm0
    psubusw mm0,mm1
    psubusw mm1,mm2
    por mm0,mm1         ; mm0=ABS |AvA-A|AvR-R|AvG-G|AvB-B|
    
    pcmpgtw mm0,mm6     ; is data in mm0 > TH in mm6
    packsswb mm0,mm0    ; pack result into mm0
    movd eax,mm0
    and eax,0FFFFFFh
    cmp eax,0
    je TTRET            ; false all data <= TH

    movq mm0,mm4
    
    packuswb mm4,mm7    ; mm4 = | |ARGB|
    movd[esi],mm4
    movd[edi],mm4
    
    movq mm4,mm0
TTRET:
    movq mm0,mm4
RET

;====================================================
;
; NOT USED
;GetAddr:    ; IN: edi, kx,ky  OUT: edi-> pmem(1,kx,ky)
;            ; Addr=edi= 4 * [PICW * (ky-1) + (kx-1)]
;    mov eax,ky
;    dec eax
;    mov ebx,PICW
;    mul ebx
;    mov ebx,kx
;    dec ebx
;    add eax,ebx
;    shl eax,2   ;*4
;    add edi,eax
;RET
;
;=============================================================

; Jump table
[SECTION .data]
Sub0 dd Despec1
Sub1 dd Despec2
Sub2 dd Despec3
Sub3 dd Despec4
Sub4 dd Despec5
Sub5 dd Despec6
Sub6 dd Sharpness
Sub7 dd Softness
Sub8 dd Brightness
;Trebor Tnemyar