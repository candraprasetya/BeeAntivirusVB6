VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Constants used to detect keyboard state.
Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

' Subclassing Pointers.
Public frmParent As Form
Private Const GWL_WNDPROC = -4

' Type Declarations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Mouse Events.
Public Event MouseDown(Button As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, x As Single, y As Single)
Public Event MouseWheelUp(x As Single, y As Single)
Public Event MouseWheelDown(x As Single, y As Single)

Public Sub HookMouseEvents(frmPassed As Form)
    
    ' Start intercepting mouse events, and process them in the routine "MouseHook".

    If modHooking.cMouseEvents Is Nothing Then
        Set frmParent = frmPassed
        Set modHooking.cMouseEvents = Me
        modHooking.hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHook, App.hInstance, App.ThreadID)
    End If

End Sub

Public Sub UnhookMouseEvents()
   
    ' Terminate the interception of mouse events and returns control to it's previous state.
    
    If Not (modHooking.cMouseEvents Is Nothing) Then
        Call UnHookWindowsHookEx(modHooking.hHook)
        Set modHooking.cMouseEvents = Nothing
    End If

End Sub

Friend Sub MouseButton(bPressed As Boolean, Button As Integer)
On Error Resume Next

    Dim x As Single, y As Single, Mouse As POINTAPI
    
    ' Find the position of the mouse pointer.
    Call GetCursorPos(Mouse)

    ' Raise MouseUp and MouseDown Events.
    If bPressed Then
        ' A Mouse Buttons is pressed.
        RaiseEvent MouseDown(CInt(Button), CSng(Mouse.x), CSng(Mouse.y))
    Else
        ' No Mouse Buttons are pressed.
        RaiseEvent MouseUp(CInt(Button), CSng(Mouse.x), CSng(Mouse.y))
    End If

End Sub

Public Sub MouseWheelUsed(bWheelUp As Boolean)
On Error Resume Next
    
    ' Notify the parent application of a MouseWheel event.
    
    Dim Mouse As POINTAPI
    
    ' Find the current mouse pointer position (to be passed on with the event).
    Call GetCursorPos(Mouse)
    
    ' Now raise the appropriate MouseWheel Event.
    If bWheelUp Then
        RaiseEvent MouseWheelUp(CSng(Mouse.x), CSng(Mouse.y))
    Else
        RaiseEvent MouseWheelDown(CSng(Mouse.x), CSng(Mouse.y))
    End If

End Sub

