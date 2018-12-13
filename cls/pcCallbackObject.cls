VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcCallbackObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcCallbackObject.cls      10/2/05
'
'           PURPOSE:
'               Set a short timer and make a method call in its notification.
'
'==================================================================================================

Option Explicit

Implements iTimer

Private mvArgs            As Variant
Private moMe              As pcCallbackObject
Private moCallbackObject  As Object
Private msCallbackProc    As String
Private miCallbackType    As VbCallType

Friend Sub fInit(ByVal oObject As Object, ByVal sProcName As String, ByVal iType As VbCallType, ByRef vArgs As Variant)
    Set moMe = Me
    Set moCallbackObject = oObject
    msCallbackProc = sProcName
    miCallbackType = iType
    If VarType(vArgs) = vbObject Then Set mvArgs = vArgs Else mvArgs = vArgs
    Timer_Install Me, ZeroL, OneL
End Sub

Private Sub iTimer_Proc(ByVal iId As Long, ByVal iElapsed As Long)
    Timer_Remove Me, ZeroL
    On Error Resume Next
    CallByName moCallbackObject, msCallbackProc, miCallbackType, mvArgs
    On Error GoTo 0
    Set moMe = Nothing
End Sub

Private Sub Class_Terminate()
    Timer_Remove Me, ZeroL
End Sub
