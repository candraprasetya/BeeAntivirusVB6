VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cDeviceNames.cls                    12/15/04
'
'           PURPOSE:
'               Wrap a DEVNAMES structure.
'
'==================================================================================================

Option Explicit

Event Changed()

Private mbDefault  As Boolean
Private msDriver   As String
Private msDevice   As String
Private msOutput   As String

Private Const DEF_Default As Boolean = False
Private Const DEF_Driver As String = vbNullString
Private Const DEF_Device As String = vbNullString
Private Const DEF_Output As String = vbNullString

Private Const PROP_Default As String = "Def"
Private Const PROP_Driver As String = "Drv"
Private Const PROP_Device As String = "Dev"
Private Const PROP_Output As String = "Out"
 
Friend Sub fInit(ByVal bDefault As Boolean, ByRef sDriver As String, ByRef sDevice As String, ByRef sOutput As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Initialize all properties to the given values.
    '---------------------------------------------------------------------------------------
    If ((bDefault Xor mbDefault) Or StrComp(sDriver, msDriver) Or StrComp(sDevice, msDevice) Or StrComp(sOutput, msOutput)) Then
        mbDefault = bDefault
        msDriver = sDriver
        msDevice = sDevice
        msOutput = sOutput
        RaiseEvent Changed
    End If
End Sub

Public Property Get default() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the default indicator stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    default = mbDefault
End Property
Public Property Let default(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the default indicator stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    If bNew Xor mbDefault Then
        mbDefault = True
        RaiseEvent Changed
    End If
End Property

Public Property Get DeviceName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the device name stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    DeviceName = msDevice
End Property
Public Property Let DeviceName(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the device name stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    If StrComp(sNew, msDevice) Then
        msDevice = sNew
        RaiseEvent Changed
    End If
End Property

Public Property Get DriverName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the driver name stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    DriverName = msDriver
End Property
Public Property Let DriverName(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the driver name stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    If StrComp(sNew, msDriver) Then
        msDriver = sNew
        RaiseEvent Changed
    End If
End Property

Public Property Get OutputPort() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the output port stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    OutputPort = msOutput
End Property
Public Property Let OutputPort(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the output port stored by the DEVNAMES structure.
    '---------------------------------------------------------------------------------------
    If StrComp(sNew, msOutput) Then
        msOutput = sNew
        RaiseEvent Changed
    End If
End Property


Private Sub Class_InitProperties()
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Initalize properties to their defaults.
    '---------------------------------------------------------------------------------------
    mbDefault = DEF_Default
    msDriver = DEF_Driver
    msDevice = DEF_Device
    msOutput = DEF_Output
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Read property values from a previously saved instance.
    '---------------------------------------------------------------------------------------
    With PropBag
        mbDefault = .ReadProperty(PROP_Default, DEF_Default)
        msDriver = .ReadProperty(PROP_Driver, DEF_Driver)
        msDevice = .ReadProperty(PROP_Device, DEF_Device)
        msOutput = .ReadProperty(PROP_Output, DEF_Output)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Store property values for future use.
    '---------------------------------------------------------------------------------------
    With PropBag
        .WriteProperty PROP_Default, mbDefault, DEF_Default
        .WriteProperty PROP_Driver, msDriver, DEF_Driver
        .WriteProperty PROP_Device, msDevice, DEF_Device
        .WriteProperty PROP_Output, msOutput, DEF_Output
    End With
End Sub
