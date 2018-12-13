VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcSupportFontPropPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcSupportFontPropPage.cls                      2/3/05
'
'           PURPOSE:
'               Simplify the implementation of ppFont.pag for the many controls that
'               use it.  It would be better if this could be done using the 'Implements'
'               feature, but that only works if the controls are not part of a control array.
'
'==================================================================================================

Option Explicit

Event AddFonts(ByVal o As ppFont)
Event GetAmbientFont(ByRef o As StdFont)

Friend Sub AddFonts(ByVal o As ppFont)
    RaiseEvent AddFonts(o)
End Sub
Friend Sub GetAmbientFont(ByRef o As StdFont)
    RaiseEvent GetAmbientFont(o)
End Sub
