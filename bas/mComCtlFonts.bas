Attribute VB_Name = "mComCtlFonts"
'==================================================================================================
'mComCtlFonts.bas                      2/13/05
'
'           PURPOSE:
'               Provide read/write access to a PropertyBag object with cFont.cls
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Public Function Font_CreateDefault(ByVal oAmbient As StdFont, Optional ByVal iDefaultSource As eFontSource = fntSourceAmbient) As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Create a font and set the source and the default ambient font.
    '---------------------------------------------------------------------------------------
    Set Font_CreateDefault = New cFont
    Font_CreateDefault.Source = iDefaultSource
    Font_CreateDefault.OnAmbientFontChanged oAmbient
End Function

Public Function Font_Read(ByVal oPropBag As PropertyBag, ByRef sPropName As String, ByVal oAmbient As StdFont, Optional ByVal iDefaultSource As eFontSource = fntSourceAmbient) As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Read a font object from the property bag.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim loDefault      As cFont
    Set loDefault = New cFont
    loDefault.Source = iDefaultSource
    
    Set Font_Read = oPropBag.ReadProperty(sPropName, loDefault)
    If Not oAmbient Is Nothing Then Font_Read.OnAmbientFontChanged oAmbient
    Exit Function
handler:
    Set Font_Read = pIdeFix(oPropBag, sPropName)
    If Not oAmbient Is Nothing Then Font_Read.OnAmbientFontChanged oAmbient
End Function

Public Sub Font_Write(ByVal oFont As cFont, ByVal oPropBag As PropertyBag, ByRef sPropName As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Write a font object to the property bag.
    '---------------------------------------------------------------------------------------
    '    oPropBag.WriteProperty sPropName, oFont, New cFont
End Sub

Private Function pIdeFix(ByVal oPropBag As PropertyBag, ByRef sPropName As String) As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Account for vb's confusions when using a project group.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    
    Set pIdeFix = New cFont
    
    Dim o      As Object
    Set o = oPropBag.ReadProperty(sPropName)
    
    With o
        If .Source = fntSourceCustom Then
            pIdeFix.Charset = .Charset
            pIdeFix.ClipPrecision = .ClipPrecision
            pIdeFix.Escapement = .Escapement
            pIdeFix.FaceName = .FaceName
            pIdeFix.Height = .Height
            pIdeFix.Italic = .Italic
            pIdeFix.Orientation = .Orientation
            pIdeFix.OutPrecision = .OutPrecision
            pIdeFix.PitchAndFamily = .PitchAndFamily
            pIdeFix.Quality = .Quality
            pIdeFix.Strikeout = .Strikeout
            pIdeFix.Underline = .Underline
            pIdeFix.Weight = .Weight
            pIdeFix.Width = .Width
        ElseIf .Source = fntSourceAmbient Or (.Source <= fntSourceSysSmallCaption And .Source >= fntSourceSysMenu) Then
            pIdeFix.Source = .Source
        Else
            'debug.assert False
        End If
    End With
    
    Exit Function
handler:
    'we'll get here if an incompatible version has been compiled since the font was saved.
    ''debug.assert False
End Function
