VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DownloadStarted()
Public Event DownloadComplete()
Public Event DownloadAborted(ByVal ErrNumber As Long, ErrDescription As String)
Public Event DownloadProgress(ByVal ReceivedBytes As Long, ByVal TotalBytes As Long)
Public Event Authenticate(ByRef Username As String, ByRef Password As String)

Implements IAuthenticate
Implements IBindStatusCallback

Private mobjBinding     As IBinding
Private mlngParenthWnd  As Long

Private mstrUsername    As String
Private mstrPassword    As String

Private Sub IAuthenticate_Authenticate(phwnd As Long, pszUsername As Long, pszPassword As Long)
    If Len(mstrUsername) = 0 And Len(mstrPassword) = 0 Then
        RaiseEvent Authenticate(mstrUsername, mstrPassword)
    End If
    pszUsername = pvStr2Ptr(mstrUsername)
    pszPassword = pvStr2Ptr(mstrPassword)
End Sub

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As olelib.BINDF, pbindinfo As olelib.BINDINFO)
'
End Sub

Private Function IBindStatusCallback_GetPriority() As Long
'
End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As olelib.BSCF, ByVal dwSize As Long, pformatetc As olelib.FORMATETC, pStgmed As olelib.STGMEDIUM)
'
End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)
'
End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As olelib.UUID, ByVal pUnk As stdole.IUnknown)
'
End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As olelib.BINDSTATUS, ByVal szStatusText As Long)
    If ulProgressMax > 0 Then
        RaiseEvent DownloadProgress(ulProgress, ulProgressMax)
    End If
End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As olelib.IBinding)
    Set mobjBinding = pib
    RaiseEvent DownloadStarted
End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)
    If hresult = 1 Then
        RaiseEvent DownloadComplete
    Else
        RaiseEvent DownloadAborted(hresult, ErrorDescription(hresult))
    End If
    Set mobjBinding = Nothing
End Sub

Private Function pvStr2Ptr(ByVal pstrString As String) As Long
   If Len(pstrString) > 0 Then
      pvStr2Ptr = CoTaskMemAlloc(LenB(pstrString) + 2)
      MoveMemory ByVal pvStr2Ptr, ByVal StrPtr(pstrString), LenB(pstrString) + 2
   End If
End Function

Public Sub StartDownload(ByVal Source As String, ByVal Dest As String, Optional ByVal Username As String, Optional ByVal Password As String)
    If mobjBinding Is Nothing Then
        mstrUsername = Username
        mstrPassword = Password
        URLDownloadToFileW Me, Source, Dest, 0, Me
        mstrUsername = vbNullString
        mstrPassword = vbNullString
    End If
End Sub

Public Sub StopDownload()
    If Not (mobjBinding Is Nothing) Then
        mobjBinding.Abort
        Set mobjBinding = Nothing
    End If
End Sub

Private Function ErrorDescription(ByVal plngErrNum As Long) As String
    Select Case plngErrNum
        Case INET_E_AUTHENTICATION_REQUIRED
            ErrorDescription = "Authentication Failure."
        Case INET_E_CANNOT_CONNECT
            ErrorDescription = "Cannot Connect"
        Case INET_E_CANNOT_INSTANTIATE_OBJECT
            ErrorDescription = "Cannot Instantiate Object."
        Case INET_E_CANNOT_LOAD_DATA
            ErrorDescription = "Cannot Load Data."
        Case INET_E_CANNOT_LOCK_REQUEST
            ErrorDescription = "Cannot Lock Request."
        Case INET_E_CANNOT_REPLACE_SFP_FILE
            ErrorDescription = "Cannot Replace SFP File."
        Case INET_E_CODE_DOWNLOAD_DECLINED
            ErrorDescription = "Code Download Declined."
        Case INET_E_CONNECTION_TIMEOUT
            ErrorDescription = "Connection Timeout."
        Case INET_E_DATA_NOT_AVAILABLE
            ErrorDescription = "Data Not Available."
        Case INET_E_DEFAULT_ACTION
            ErrorDescription = "Default Action."
        Case INET_E_DOWNLOAD_FAILURE
            ErrorDescription = "Download Failure."
        Case INET_E_INVALID_REQUEST
            ErrorDescription = "Invalid Request."
        Case INET_E_INVALID_URL
            ErrorDescription = "Invalid URL."
        Case INET_E_NO_SESSION
            ErrorDescription = "No Session."
        Case INET_E_NO_VALID_MEDIA
            ErrorDescription = "No Valid Media."
        Case INET_E_OBJECT_NOT_FOUND
            ErrorDescription = "File Not Found."
        Case INET_E_QUERYOPTION_UNKNOWN
            ErrorDescription = "QueryOption Unknown."
        Case INET_E_REDIRECT_FAILED
            ErrorDescription = "Redirect Failed."
        Case INET_E_REDIRECT_TO_DIR
            ErrorDescription = "Redirect To Dir."
        Case INET_E_REDIRECTING
            ErrorDescription = "Redirecting."
        Case INET_E_RESOURCE_NOT_FOUND
            ErrorDescription = "Resource Not Found."
        Case INET_E_RESULT_DISPATCHED
            ErrorDescription = "Result Dispatched."
        Case INET_E_SECURITY_PROBLEM
            ErrorDescription = "Security Problem."
        Case INET_E_UNKNOWN_PROTOCOL
            ErrorDescription = "Unknown Protocol."
        Case Else
            ErrorDescription = "Unknown Error."
   End Select
End Function


