VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cDebug.cls                      1/23/05
'
'           PURPOSE:
'               Track menu, string, memory, gdi and other handles for debugging purposes.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit

#If bDebug Then

Private moCollKeys         As Collection
Private moCollCollections  As Collection
    
Public Sub Add(ByRef sType As String, ByVal H As Long)
        
    On Error Resume Next
        
    Dim loColl      As Collection
    Set loColl = moCollCollections.Item(sType)
        
    If loColl Is Nothing Then
        Set loColl = New Collection
        moCollCollections.Add loColl, sType
        moCollKeys.Add sType, sType
    End If
        
    Err.Clear
    loColl.Add CStr(H), CStr(H)
    If Err.Number Then pErr "Add"
        
    On Error GoTo 0
        
    End Sub
    
Public Sub Remove(ByRef sType As String, ByVal H As Long)
    On Error GoTo handler
    moCollCollections.Item(sType).Remove CStr(H)
    On Error GoTo 0
    Exit Sub
handler:
    pErr "Remove"
    On Error GoTo 0
    End Sub
    
Public Function Exists(ByRef sType As String, ByVal H As Long) As Boolean
    On Error GoTo handler
    Call moCollCollections.Item(sType).Item(CStr(H))
    Exists = True
handler:
    On Error GoTo 0
    End Function
    
Public Function Count(ByRef sType As String) As Long
    On Error GoTo handler
    Count moCollCollections.Item(sType).Count
handler:
    On Error GoTo 0
    End Function
    
Private Sub pErr(ByRef sProc As String)
    Debug.Print "DEBUG " & sProc & " Error: " & Err.Number & vbNewLine & Err.Description
    'debug.assert False
    End Sub
    
Private Sub Class_Initialize()
    Set moCollCollections = New Collection
    Set moCollKeys = New Collection
    End Sub
    
Private Sub Class_Terminate()
    On Error GoTo handler
        
    Dim lvKey      As Variant
    Dim lsKey      As String
        
    Dim loColl As Collection
        
    For Each lvKey In moCollKeys
        lsKey = lvKey
        Set loColl = moCollCollections.Item(lsKey)
        If loColl.Count Then
            Debug.Print "LEAKED " & loColl.Count & " " & lsKey & "(s)!!!"
            MsgBox "LEAKED " & loColl.Count & " " & lsKey & "(s)!!!"
        End If
    Next
        
    On Error GoTo 0
    Exit Sub
        
handler:
    pErr "Terminate"
    Resume Next
        
End Sub

#End If
