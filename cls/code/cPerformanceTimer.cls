VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPerformanceTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Type LARGE_INTEGER
'    LowPart As Long
'    HighPart As Long
'End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (OutCounter As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (OutFrequency As Currency) As Long

Private mcFrequency As Currency

Private mcStarted   As Currency
Private mcStopped   As Currency

Private Sub Class_Initialize()
    QueryPerformanceFrequency mcFrequency
End Sub

Public Function TimerStart() As Boolean
    TimerStart = QueryPerformanceCounter(mcStarted) <> 0
    mcStopped = 0
End Function

Public Function TimerElapsed() As Double
    If mcStopped = 0 Then
        Dim lcTemp As Currency
        QueryPerformanceCounter lcTemp
        lcTemp = lcTemp
        TimerElapsed = (lcTemp - mcStarted) / mcFrequency
    Else
        TimerElapsed = (mcStopped - mcStarted) / mcFrequency
    End If
End Function

Public Sub TimerStop()
    QueryPerformanceCounter mcStopped
End Sub
