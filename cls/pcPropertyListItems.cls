VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcPropertyListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2070
'==================================================================================================
'cPropertyListItems.cls                 7/5/04
'
'           PURPOSE:
'               Creates and maintains a collection of cookie/name pairs for use with iPerPropertyBrowsing.
'
'==================================================================================================
Option Explicit

Private mColl     As Collection
Private miCookie  As Long

Private Sub Class_Initialize()
    Set mColl = New Collection
End Sub
Public Property Get NewEnum() As stdole.IUnknown
    Set NewEnum = mColl.[_NewEnum]
End Property

Public Property Get Exists(ByVal Cookie As Long) As Boolean
    On Error Resume Next
    IsObject mColl.Item("C" & Cookie)
    Exists = (Err.Number = 0&)
    On Error GoTo 0
End Property

Public Property Get Item(ByVal Cookie As Long) As pcPropertyListItem
    Set Item = mColl.Item("C" & Cookie)
End Property

Public Property Get Count() As Long
    Count = mColl.Count
End Property

Public Function Add(ByVal DisplayName As String) As pcPropertyListItem
    Dim loItem      As pcPropertyListItem
    
    For Each loItem In mColl
        If StrComp(DisplayName, loItem.DisplayName) < 0& Then Exit For
    Next
        
    miCookie = miCookie + 1&
        
    Dim oItem      As pcPropertyListItem
    Set oItem = New pcPropertyListItem
    oItem.Cookie = miCookie
    oItem.DisplayName = DisplayName
        
    If loItem Is Nothing Then
        mColl.Add oItem, "C" & miCookie
    Else
        mColl.Add oItem, "C" & miCookie, "C" & loItem.Cookie
    End If
    Set Add = oItem
End Function

