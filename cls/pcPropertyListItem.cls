VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcPropertyListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2040
'==================================================================================================
'cPropertyListItem.cls                  7/5/04
'
'           PURPOSE:
'               R/W access to name and cookie value for Properties through IPerPropertyBrowsing
'
'==================================================================================================
Option Explicit

Private msString  As String
Private miCookie  As Long

Public Property Get Cookie() As Long
    Cookie = miCookie
End Property
Friend Property Let Cookie(ByVal newVal As Long)
    miCookie = newVal
End Property
Public Property Get DisplayName() As String
    DisplayName = msString
End Property
Friend Property Let DisplayName(ByVal newVal As String)
    msString = newVal
End Property

Friend Function lpDisplayName() As Long
    Dim liLen      As Long
    liLen = LenB(msString) + 2
    lpDisplayName = CoTaskMemAlloc(liLen)
    #If bDebug Then
    DEBUG_Remove DEBUG_hMemCoTask, lpDisplayName
    #End If
    CopyMemory ByVal lpDisplayName, ByVal StrPtr(msString & vbNullChar), liLen
End Function
