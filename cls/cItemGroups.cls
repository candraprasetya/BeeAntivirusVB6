VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItemGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cItemGroups.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of groups in a listview control.
'
'               May only be used with a link to comctl32 version 6
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner  As ucListView

Friend Sub fInit(ByVal oOwner As ucListView)
    Set moOwner = oOwner
End Sub

Public Function Add( _
Optional ByVal sKey As String, _
Optional ByVal sText As String, _
Optional ByVal vGroupBefore As Variant) _
As cItemGroup
    Set Add = moOwner.fItemGroups_Add(sKey, sText, vGroupBefore)
End Function
   
Public Property Get Count() As Long
    Count = moOwner.fItemGroups_Count()
End Property
Public Sub Clear()
    moOwner.fItemGroups_Clear
End Sub
Public Sub Remove(ByVal vGroup As Variant)
    moOwner.fItemGroups_Remove vGroup
End Sub
Public Property Get Enabled() As Boolean
    Enabled = moOwner.fItemGroups_Enabled
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
    moOwner.fItemGroups_Enabled = bNew
End Property

Public Property Get Exists(ByVal vGroup As Variant) As Boolean
    Exists = moOwner.fItemGroups_Exists(vGroup)
End Property

Public Property Get Item(ByVal vGroup As Variant) As cItemGroup
    Set Item = moOwner.fItemGroups_Item(vGroup)
End Property

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fItemGroups_NextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = NegOneL
    tEnum.iControl = moOwner.fItemGroups_Control()
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub

Public Property Get NewEnum() As stdole.IUnknown
    Dim loEnum      As pcEnumeration
    Set loEnum = New pcEnumeration
    
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fItemGroups_Control())
End Property
