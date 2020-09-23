VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimingPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CTimingPC.cls

' VBSpeed:-
' <www.xbeat.net/vbspeed/index.htm>

' Use eg

' At Form1

' Dim tmHowLong As New CTimingPC

' RESET
'  tmHowLong.Reset

' SHOW msec
'  Text1 = Format(tmHowLong.Elapsed, "0000") & " msec"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' CTimingPC - (c) don 19990921, donald@xbeat.net
'           - update: 20000908
' ******************************************************************************
' Performance Timer with ± n MicroSecond-Resolution (dep. on CPU)
' NOTE: uses API QueryPerformanceCounter which some CPUs do NOT support!
' ******************************************************************************
'
' use:
' Module declare: Private tmr as CTimingPC
' Make Instance:  Set tmr = New CTimingPC
' start timer:    tmr.Reset
' stop/get time:  dTime=tmr.Elapsed (returns Double millisec with microsec as fraction)
'                 sTime=tmr.sElapsed (returns String millisec with microsec as fraction and " msec")
' wait:           wait for n milliseconds (total suspend, no DoEvents)
' note:
' tmr.Elapsed actually returns Currency but it can be mapped on a double

Option Explicit

Private seqFreq As Currency
Private seqStart As Currency
Private fAvailable As Boolean

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency) As Long

Private Sub Class_Initialize()
' determine if system has Performance Counter and get its frequency
' returns ticks/sec (1.193.180 on my machine PII 350)
' note: the Currency var displays this as "119.318", gotta multiply
'       by 10,000 to get the actual integer value
  If QueryPerformanceFrequency(seqFreq) = 0& Then
    MsgBox "PerfCounter not available", vbExclamation
  Else
    fAvailable = True
  End If
End Sub

Friend Sub Reset()
' establishes a "start time" reference
  QueryPerformanceCounter seqStart
End Sub

Friend Function Elapsed() As Double 'more handy than Currency
' return elapsed time in milliseconds
  Dim seqNow As Currency
  QueryPerformanceCounter seqNow
  If fAvailable Then
    Elapsed = 1000 * (seqNow - seqStart) / seqFreq
  End If
End Function

