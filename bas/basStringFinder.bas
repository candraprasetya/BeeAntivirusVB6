Attribute VB_Name = "basStringFinder"
Private strHeader(5) As Long        ' Header for the StringArray Map
Private patHeader(5) As Long        ' Header for the PatternArray Map
Private Const FILE_SHARE_READ = &H1
'Private Const FILE_SHARE_WRITE = &H2
'Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_BEGIN = 0
Private Const CREATE_NEW = 1
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READWRITE = &HC0000000
Private Const PAGE_READWRITE = &H4
Private Const FILE_MAP_WRITE = &H2
Private Const FILE_MAP_READ = &H4
Private Const FILE_MAP_READWRITE = &H6
'Private Const FADF_FIXEDSIZE = &H10
Private Const PAGE_READONLY = &H2
Private Const INVALIDHANDLE = -1
Private Const CLOSEDHANDLE = 0
'MEMORY MAPPED FILE VARIABLES-------------------------------------------------
Private hFile As Long       'HANDLE TO FILE
Private hFileMap As Long    'HANDLE TO FILE MAPPING
Private hMapView As Long    'HANDLE TO MAP VIEW
Private mBaseAddr As Long   'FILE POINTER = HANDLE TO MAP VIEW
Private mFileSize As Long   'SIZE OF THE MEMORY MAPPED FILE
'=============================================================================


Private Function SearchBMH(ByVal sPath As String, ByVal StartAt As Long) As Long
Dim i As Long
Dim j As Long
Dim k As Long
    hFile = CreateFile(sPath, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    hFileMap = CreateFileMapping(hFile, 0, PAGE_READONLY, 0, 0, vbNullString)
    hMapView = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
    strHeader(3) = mBaseAddr
    strHeader(4) = mFileSize
   SearchBMH = 0&
   i = CLng(StartAt + m_PatLength - 2)
   
   If m_PatLength > 1 Then
   
        Do While i < mFileSize
           k = m_PatLength - 1
           j = i
           Do While strArrayB(j) = patArrayB(k)
              If k = 0 Then
                 SearchBMH = j + 1&
                 Exit Function
              End If
              k = k - 1&
              j = j - 1&
           Loop
           i = i + m_Skip32Table(strArrayB(i))
        Loop
        
   ElseIf m_PatLength = 1 Then
        
        k = m_PatLength - 1
        Do While i < mFileSize
           'k = m_PatLength - 1
           j = i
           If strArrayB(j) = patArrayB(k) Then
                 SearchBMH = j + 1&
                 Exit Function
           End If
           i = i + 1 'm_Skip32Table(strArrayB(i))
        Loop
   
   End If
   
End Function
