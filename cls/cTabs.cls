VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cTabs.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of tabs in a tab control.
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner  As ucTabStrip

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fTabs_Enum_NextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iControl = moOwner.fTabs_Enum_Control()
    tEnum.iIndex = NegOneL
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    moOwner.fTabs_Enum_Skip tEnum, iSkipCount, bSkippedAll
End Sub

Public Property Get NewEnum() As IUnknown
    Dim loEnum      As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fTabs_Enum_Control())
End Property


Friend Sub fInit(ByVal oOwner As ucTabStrip)
    Set moOwner = oOwner
End Sub

Public Function Add(ByRef sText As String, Optional ByVal iIconIndex As Long = NegOneL, Optional ByRef sKey As String, Optional ByVal vTabInsertBefore As Variant) As cTab
    Set Add = moOwner.fTabs_Add(sText, iIconIndex, sKey, vTabInsertBefore)
End Function

Public Sub Remove(ByVal vTab As Variant)
    moOwner.fTabs_Remove vTab
End Sub

Public Property Get Exists(ByVal vTab As Variant) As Boolean
    Exists = moOwner.fTabs_Exists(vTab)
End Property

Public Property Get Item(ByVal vTab As Variant) As cTab
    Set Item = moOwner.fTabs_Item(vTab)
End Property

Public Sub Clear()
    moOwner.fTabs_Clear
End Sub

Public Property Get Count() As Long
    Count = moOwner.fTabs_Count()
End Property
