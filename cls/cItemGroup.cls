VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItemGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cItemGroup.cls                    12/15/04
'
'           PURPOSE:
'               Represent an item group in a listview control.
'
'               May only be used with a link to comctl32 version 6
'
'==================================================================================================

Option Explicit

Private moOwner  As ucListView
Private miIndex  As Long
Private miId     As Long

Friend Sub fInit(ByVal oOwner As ucListView, ByVal iId As Long, ByVal iIndex As Long)
    Set moOwner = oOwner
    miIndex = iIndex
    miId = iId
End Sub

Friend Function fIsOwner(ByVal oOwner As ucListView) As Boolean
    fIsOwner = (oOwner Is moOwner)
End Function

Friend Property Get fId() As Long
    fId = miId
End Property

Public Property Get Text() As String
    Text = moOwner.fItemGroup_Text(miIndex, miId)
End Property
'Public Property Let Text(ByVal sNew As String)
'    moOwner.fItemGroup_Text(miIndex, miId) = sNew
'End Property

Public Property Get Key() As String
    Key = moOwner.fItemGroup_Key(miIndex, miId)
End Property
Public Property Let Key(ByVal sNew As String)
    moOwner.fItemGroup_Key(miIndex, miId) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fItemGroup_Index(miIndex, miId)
End Property
