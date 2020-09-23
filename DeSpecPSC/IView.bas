Attribute VB_Name = "IView"
' IView.bas by Robert Rayment

' Using FREEWARE i_view32.exe by Irfan Skiljan
' <http://www.irfanview.com>

Option Base 1

'Unless otherwise defined:
DefBool A      ' Boolean
DefByte B      ' Byte
DefLng C-W     ' Long
DefSng X-Z     ' Singles
               ' $ for strings

Public TempBmp$  ' Temporary bmp file name


Public Sub SaveJPGGIFPNG(SaveFileSpec$)

' CONVERT TempBmp$ to SaveFileSpec$

' Assumes that the file TempBmp$ is saved in the App folder
' AND that i_view32.exe is in the same folder
' The SaveFileSpec$ extension determines what conversion is done.
' It follows that the extension must by a valid format that i_view32
' can deal with.

AppPath$ = App.Path
If Right$(AppPath$, 1) <> "\" Then AppPath$ = AppPath$ & "\"

IViewPath$ = AppPath$ & "i_view32.exe "     ' NB space after .exe

' InStrRev for VB6 only
SavePath$ = Left$(SaveFileSpec$, InStrRev(SaveFileSpec$, "\"))

aString$ = IViewPath$ & SavePath$ & TempBmp$
bString$ = " /convert=" & SaveFileSpec$   ' NB space before /convert
cString$ = aString$ & bString$
res& = Shell(cString$, vbNormalFocus)

Kill SavePath$ & TempBmp$

End Sub

Public Sub LoadPNG(OpenFileSpec$)

' CONVERT OpenFileSpec$ to TempBmp$

' The OpenFileSpec$ extension shows i_view32 what type of file
' to convert to the temporary file TempBmp$.
' The TempBmp$ file will be saved in the OpenFileSpec$ folder.
' It follows that the extension MUST be a true representation
' of the file's format.

IViewPath$ = App.Path
If Right$(IViewPath$, 1) <> "\" Then IViewPath$ = IViewPath$ & "\"
IViewPath$ = IViewPath$ & "i_view32.exe " ' NB space after .exe

' InStrRev for VB6 only
SavePath$ = Left$(OpenFileSpec$, InStrRev(OpenFileSpec$, "\"))

aString$ = IViewPath$ & OpenFileSpec$
bString$ = " /convert=" & SavePath$ & TempBmp$   ' NB space before /convert
cString$ = aString$ & bString$
res& = Shell(cString$, vbNormalFocus)

End Sub

