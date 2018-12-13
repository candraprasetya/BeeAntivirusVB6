Attribute VB_Name = "basVarPub"
Option Explicit

Public Detik            As Long
Public Menit            As Long
Public Jam              As Integer

Public xDrive()         As String
Public qPath(10)        As String
Public cPath            As String

Public xNumChecksum()   As String
Public xNamChecksum()   As String
Public xJumChecksum     As String
Public xPENam()         As String
Public xPENum()         As String
Public xPEJum           As String
Public IsItFirst        As Boolean

Public FileToScan       As String
Public FoldToScan       As String
Public FileSpeed        As String
Public FileScan         As String
Public FileIgnore       As String
Public FileRemain       As String
Public WithBuffer       As Boolean
Public StopScan         As Boolean
Public IdxScan          As Integer
Public xScanPath        As String
Public xMaxFile         As String
Public xPercentage      As String
Public isPause          As Boolean

Public XinjectVir       As String
Public XnamVirT         As String
Public XnamScrT         As String

Public IsPE32EXE        As Boolean

Public RTPpath          As String
Public ScanRTPmod       As Boolean

Public LastFlashVolume  As Long
Public PathCustomScan   As String
Public PathContextScan() As String

Public isFromContext    As Boolean

Public xSectionAkhir    As String
Public xNamaSectionAkhir As String

Public xSectionAkhir2    As String
Public xNamaSectionAkhir2 As String
Public xSectionJum      As String

Public DbDefinition     As String
Public VirusStatus      As Boolean
Public ProcessScan      As Boolean
