VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classFindString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CRC Checksum Class
'------------------------------------
'
'A very fast solution to calculate the
'CRC Checksum (at the moment CRC16 and
'CRC32 values) with the help of some
'pre-compiled assembler code
'
'(c) 2000, Fredrik Qvarfort
'

Option Explicit
Public StopSearch As Boolean
Public Enum CRCAlgorithms
  CRC16
  CRC32
End Enum
Private m_Algorithm As Boolean

Private m_CRC16 As Long
Private m_CRC16Asm() As Byte
Private m_CRC16Init As Boolean
Private m_CRC16Table(0 To 255) As Long

Private m_CRC32 As Long
Private m_CRC32Asm() As Byte
Private m_CRC32Init As Boolean
Private m_CRC32Table(0 To 255) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'=================================================================================
'CRC Checksum & String Search Functions
'=================================================================================
'
'A very fast solution to calculate the CRC Checksum (CRC16 and CRC32 values)
'and String Searching with the help of some pre-compiled assembler code
'
'(c) 2008, by D.Senthilathiban
'E-mail:samram20@hotmail.com
'This code is for Personal usage only! and not to be used for commercial purpose.
'===============================MEMORY MAPPING ===============================
'
'MEMORY MAPPING APIs----------------------------------------------------------
'
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFile As Long, _
                                                           ByVal dwDesiredAccess As Long, _
                                                           ByVal dwFileOffsetHigh As Long, _
                                                           ByVal dwFileOffsetLow As Long, _
                                                           ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, _
                                                           ByVal lpAttributes As Long, ByVal flProtect As Long, _
                                                           ByVal dwMaximumSizeHigh As Long, _
                                                           ByVal dwMaximumSizeLow As Long, _
                                                           ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Boolean
'Private Declare Function ReadFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
                                                    ByVal nNumberOfBytesToRead As Long, _
                                                    lpOverlapped As OVERLAPPED, _
                                                    ByVal lpCompletionRoutine As Long) As Long
'Private Declare Function FileIOCompletionRoutine Lib "kernel32" (dwErrorCode As Long, _
                                                                 dwNumberOfBytesTransfered As Long, _
                                                                 lpOverlapped As OVERLAPPED)
'Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                        ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
                                                                        ByVal lpSecurityAttributes As Long, _
                                                                        ByVal dwCreationDisposition As Long, _
                                                                        ByVal dwFlagsAndAttributes As Long, _
                                                                        ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'USER TYPE -------------------------------------------------------------------
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

'CONSTANTS--------------------------------------------------------------------
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

'===============================STRING MAPPING ===============================
' Thanks to Chris Lucas for the String mapping Concepts.

Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long

Private strHeader(5) As Long        ' Header for the StringArray Map
Private patHeader(5) As Long        ' Header for the PatternArray Map

'Character string array = 2 Bytes per element for UNICODE Char
Private strArrayW() As Integer      ' Maps onto the Text's string
Private patArrayW() As Integer      ' Maps onto the Pattern's string

'Private UB_strArrayW  As Long       ' UBound size of strArrayW()
'Private UB_patArrayW  As Long       ' UBound size of patArrayW()

'Character string array = 1 Byte per element for ANSI Char
Private strArrayB() As Byte         ' Maps onto the Text's string
Private patArrayB() As Byte         ' Maps onto the Pattern's string

'Private UB_strArrayB  As Long       ' UBound size of strArrayB()
'Private UB_patArrayB  As Long       ' UBound size of patArrayB()

Private m_StrBuffer As String      ' variable to hold string to search in
'Private m_PatBuffer As String      ' variable to hold Pattern string to search for
'=============================================================================

'===============================STRING SEARCHING==============================
' (c) 2008, by D.Senthilathiban, E-mail:samram20@hotmail.com
' ******************************* WARNING ************************************
' The code is tested for binary comparison with ANSI code input only.UNICODE
' character input and strings have ligatures are not tested. It is upto the
' user to test the code with UNICODE/special character input and validate.
' It is assumed that the pattern text length always small in size hence the
' pattern string is converted into byte array using VB in-built functions.
' ****************************************************************************

Public Enum eSearchAlgorithms
        Asm_BMHA 'Boyer-Moore horspool string searching Algorithm-Assembly code
        Asm_STRC 'C language StrStr String searching Algorithm   -Assembly code
        Asm_BYTE 'Byte by byte string searching using Assembly string instruction
        Vb_BMHA  'Boyer-Moore horspool string searching Algorithm-Vb Native code
        Vb_InStr 'Visual basic Instr function
        'Vb_InStrMAP 'Visual basic Instr function
End Enum

Public Enum eCharCode
        ANSISTD_CHAR 'ASCII character set (1 byte per char 0-255)
        UNICODE_CHAR 'UniCode character set (2 byte per char 0 to 65,535)
        'DEFAULT_CHAR 'check the sting and set ASCII OR UniCode character set
        'DBCSSTD_CHAR 'double byte char set
End Enum
Private m_SearchAlgorithm As eSearchAlgorithms
Private m_MyFindAlgorithm As eSearchAlgorithms
Private m_BMHA32Asm() As Byte
Private m_BYTE32Asm() As Byte
Private m_STRC32Asm() As Byte

Private m_BMHA32Init As Boolean
Private m_BYTE32Init As Boolean
Private m_STRC32Init As Boolean

Private m_Skip32Table(0 To 255) As Long
Private m_PatData(0 To 2) As Long
Private m_PatArray() As Byte
Private m_PatString As String
Private m_PatLength As Long
Private m_PatFormat As eCharCode
Private m_PatChrLen As Long
'
'
'====================================================================================

Private Function CloseMapFile()
    
        UnmapViewOfFile hMapView
        CloseHandle hFileMap
        CloseHandle hFile
        hMapView = CLOSEDHANDLE
        hFileMap = CLOSEDHANDLE
        hFile = INVALIDHANDLE
        mBaseAddr = CLOSEDHANDLE
        mFileSize = CLOSEDHANDLE
        
End Function

Sub FileMapSearchCRC(ByVal FileName As String, ByVal SearchText As String, ByRef strCRC As String, ByRef strFound As Boolean)

Dim bOpenMap As Boolean

strCRC = vbNullString
strFound = False
Clear 'clear old CRC value stored in the variable m_CRC16 or m_CRC32
mFileSize = FileLen(FileName) 'Find the length of the target file.

If mFileSize > 0 Then
    '===========================================================================
    '*************************FILE MAPPING**************************************
    '===========================================================================
    'JUST MAP THE FILE AND GET THE FILE BASE ADDRESS IN THE VIRTUAL MEMORY SPACE.
    'SAY GOOD BYE FOR BYTE ARRAY MEMORY ALLOCATION & MOVING CHUNK INTO THE ARRAY.
    'SO OBVIOUSLY WE HAVE SAVED THE PROCESSOR TIME,HUGE MEMORY SPACE.
    hFile = CreateFile(FileName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    hFileMap = CreateFileMapping(hFile, 0, PAGE_READONLY, 0, 0, vbNullString)
    hMapView = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
    'hFile = CreateFile(FileName, GENERIC_READWRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    'hFileMap = CreateFileMapping(hFile, 0, PAGE_READWRITE, 0, 0, 0) ' vbNullString)
    'hMapView = MapViewOfFile(hFileMap, FILE_MAP_READWRITE, 0, 0, 0)
    '===========================================================================

    If hMapView <> CLOSEDHANDLE Then
        
       mBaseAddr = hMapView
       bOpenMap = True 'SET TRUE IF THE FILE IS MAPPED
       'Set error trapping incase any error
       On Local Error GoTo NoData
      
      'Run the pre-compiled assembler code
      'for the current selected algorithm
      Select Case m_Algorithm
      
        Case CRC16
        
        Call CallWindowProc(VarPtr(m_CRC16Asm(0)), VarPtr(m_CRC16), _
                            ByVal mBaseAddr, VarPtr(m_CRC16Table(0)), mFileSize)
      
        Case CRC32
        
        Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), _
                            ByVal mBaseAddr, VarPtr(m_CRC32Table(0)), mFileSize)
      
      End Select
    
    End If

Else
    
    strCRC = "00000000"
    strFound = False
    'If InstrFIND(SearchText) > 0 Then
    '    strFound = False
    '    strCRC = Right$(strCRC & Hex$(Value), 8)
    'End If
    Exit Sub
    
End If

NoData:

  'Return the current CRC value
  strCRC = "00000000"
  strFound = False
  If bOpenMap = True Then
        If InstrFIND(SearchText) > 0 Then
            strFound = True
            strCRC = Right$(strCRC & Hex$(Value), 8)
        End If
  End If
  CloseMapFile 'close the handle
  
End Sub

Public Function FileMapCRC(ByVal FileName As String) As String

Dim bOpenMap As Boolean


Clear 'clear old CRC value stored in the variable m_CRC16 or m_CRC32
mFileSize = FileLen(FileName) 'Find the length of the target file.

If mFileSize > 0 Then
    '===========================================================================
    '*************************FILE MAPPING**************************************
    '===========================================================================
    'JUST MAP THE FILE AND GET THE FILE BASE ADDRESS IN THE VIRTUAL MEMORY SPACE.
    'SAY GOOD BYE FOR BYTE ARRAY MEMORY ALLOCATION & MOVING CHUNK INTO THE ARRAY.
    'SO OBVIOUSLY WE HAVE SAVED THE PROCESSOR TIME,HUGE MEMORY SPACE.
    hFile = CreateFile(FileName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    hFileMap = CreateFileMapping(hFile, 0, PAGE_READONLY, 0, 0, vbNullString)
    hMapView = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
    '===========================================================================
    
    If hMapView <> CLOSEDHANDLE Then
       mBaseAddr = hMapView
       bOpenMap = True 'SET TRUE IF THE FILE IS MAPPED
       'Set error trapping incase any error
       On Local Error GoTo NoData
      
      'Run the pre-compiled assembler code
      'for the current selected algorithm
      Select Case m_Algorithm
      
        Case CRC16
        
        Call CallWindowProc(VarPtr(m_CRC16Asm(0)), VarPtr(m_CRC16), _
                            ByVal mBaseAddr, VarPtr(m_CRC16Table(0)), mFileSize)
      
        Case CRC32
        
        Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), _
                            ByVal mBaseAddr, VarPtr(m_CRC32Table(0)), mFileSize)
      
      End Select
    
    End If

Else
    
    FileMapCRC = "00000000"
    Exit Function
    
End If

NoData:
  'Return the current CRC value
  FileMapCRC = Right$("00000000" & Hex$(Value), 8)
  CloseMapFile 'close the handle
  
End Function

Public Function FileMapSearch(ByVal FileName As String, ByVal SearchText As String, ByVal sStart As Long) As Long

FileMapSearch = 0
mFileSize = FileLen(FileName) 'Find the length of the target file.

If mFileSize > 0 Then
    '===========================================================================
    '*************************FILE MAPPING**************************************
    '===========================================================================
    'JUST MAP THE FILE AND GET THE FILE BASE ADDRESS IN THE VIRTUAL MEMORY SPACE.
    'SAY GOOD BYE FOR BYTE ARRAY MEMORY ALLOCATION & MOVING CHUNK INTO THE ARRAY.
    'SO OBVIOUSLY WE HAVE SAVED THE PROCESSOR TIME,HUGE MEMORY SPACE.
    hFile = CreateFile(FileName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    hFileMap = CreateFileMapping(hFile, 0, PAGE_READONLY, 0, 0, vbNullString)
    hMapView = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
    '===========================================================================
    'Set error trapping incase any error
    On Local Error GoTo NoData
    
    If hMapView <> CLOSEDHANDLE Then
       mBaseAddr = hMapView
       FileMapSearch = InstrFIND(SearchText, sStart, ANSISTD_CHAR)
    End If

Else
    
    FileMapSearch = 0
    Exit Function
    
End If

NoData:
  CloseMapFile 'close the handle
  
End Function

Private Function InstrFIND(ByVal msText As String, Optional ByVal StartAt As Long = 1, Optional chrFormat As eCharCode = ANSISTD_CHAR) As Long

Dim offset As Long

'On Error Resume Next

    If Len(msText) = 0 Then
        InstrFIND = 1
        Exit Function
    End If
    
    If mFileSize = 0 Then
        InstrFIND = 0
        Exit Function
    End If
    'm_SearchAlgorithm = Asm_BMHA
    'verify any change is happed in pattern text or in text format or in Search Algorithm
    If m_PatString <> msText Or m_PatFormat <> chrFormat Or m_SearchAlgorithm <> m_MyFindAlgorithm Then
        m_PatString = msText
        m_PatFormat = chrFormat
        m_PatChrLen = chrFormat + 1
        m_MyFindAlgorithm = m_SearchAlgorithm
        'Initialize the pattern
        InitializePattern chrFormat
        'Boyer-Moore horspool Algorithm skip table
        'InitializeSkip32  'chrFormat
    End If
    
    'reduce 1 from startAt as the byteArray is assumed to start at index 0
    If StartAt < 1 Then StartAt = 1
    offset = StartAt - 1
    'Run the pre-compiled assembler code
    
    Select Case m_SearchAlgorithm
      
        Case Asm_BMHA
          
          'Boyer-Moore horspool string searching Algorithm' for patlength > 2
          'Assembly string instruction used for the string searching upto patlength<=2
            Call CallWindowProc(VarPtr(m_BMHA32Asm(0)), VarPtr(m_PatData(0)), ByVal (mBaseAddr + offset), _
                                VarPtr(m_Skip32Table(0)), (mFileSize - offset))
                        
        Case Asm_STRC
          'C language StrStr String searching Algorithm' for patlength > 1
          'Assembly string instruction for the string searching upto patlength=1
            Call CallWindowProc(VarPtr(m_STRC32Asm(0)), VarPtr(m_PatData(0)), ByVal (mBaseAddr + offset), _
                                ByVal 0&, (mFileSize - offset))
        
        Case Asm_BYTE
          'Byte by byte string searching using Assembly string instruction for any patlength
            Call CallWindowProc(VarPtr(m_BYTE32Asm(0)), VarPtr(m_PatData(0)), ByVal (mBaseAddr + offset), _
                                ByVal 0&, (mFileSize - offset))
                        
        Case Vb_BMHA
        
            strHeader(3) = mBaseAddr  'pointer to MMF '= StrPtr(m_StrBuffer)' Pointer to StrArray
            strHeader(4) = mFileSize  'MMF Size '=  String Array size (max= &H7FFFFFFF)
            m_PatData(2) = SearchBMH(StartAt)
        
        Case Vb_InStr
          'Vb Native function
            strHeader(3) = mBaseAddr '= StrPtr(m_StrBuffer)' Pointer to StrArray
            strHeader(4) = mFileSize  ' String Array size (max= &H7FFFFFFF)
            '===========================================================================
            '**************WARNING*************WARNING**************WARNING*************
            '===========================================================================
            'THE VB IDE MAY CRASH AT THIS LOCATION IF THE SIZE OF MEMORY MAPPED FILE(MMF)
            'IS VERY LARGE. ALSO WE LOST THE PORPOSE OF USING MEMORY MAPPED FILE POINTER
            'BECAUSE THE 'StrConv' FUNCTION CONVERTS THE WHOLE BYTE ARRAY BACK TO STRING
            'BY ALLOCATING REQD MEMOREY SIZE & MOVES ALL THE CHARS INTO STRING VARIABLE.
            'SO OBVIOUSLY WE HAVE WASTED THE PROCESSOR TIME,HUGE MEMORY SPACE.
            m_StrBuffer = StrConv(strArrayB, vbUnicode) 'convert byte array to string
            '===========================================================================
            m_PatData(2) = InStr(StartAt, m_StrBuffer, m_PatString)
            
      End Select
    
    If StopSearch = True Then 'flag to triminate the process
            InstrFIND = 0
            '=====================================================================
            'incase user teriminates the loop this must execute else IDE will crash
            TerminateHeaders
            '=====================================================================
            Exit Function
    End If
    
    InstrFIND = m_PatData(2)
    
End Function


Private Function ValidateContainedText(ByVal msText As String) As Boolean
    
    If Len(msText) = 0 Then
        ValidateContainedText = True
        Exit Function
    End If
    
    If mFileSize = 0 Then
        ValidateContainedText = False
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim lbFirstTime As Boolean
    Dim lsWide As String
    Dim myMainArray() As Byte
    Dim chunkSize As Long
    Dim readSize As Long
    Dim leftSize As Long
    Dim offset As Long
    Dim miDoubleTextLen As Long
    
    offset = 0
    chunkSize = 65536 '65536=64kb'104857600 '100 * 1024 * 1024
    miDoubleTextLen = Len(StrConv(msText, vbUnicode))  ' Len(msText) * 2
    If chunkSize < miDoubleTextLen Then chunkSize = miDoubleTextLen
    lbFirstTime = True
    
    If mFileSize > chunkSize Then
        readSize = chunkSize
        leftSize = mFileSize - readSize
    Else
        readSize = mFileSize
        leftSize = 0
    End If
    
    Do While (readSize > 0 Or leftSize > 0)
                If lbFirstTime Then
                    lbFirstTime = False
                    ReDim myMainArray(0 To readSize - 1)
                    CopyMemory myMainArray(0), ByVal (mBaseAddr + offset), readSize
                    offset = offset + readSize - miDoubleTextLen
                    If leftSize = 0 Then readSize = 0
                Else
                    If (leftSize + miDoubleTextLen) > chunkSize Then
                        ''''pending''''''''''''''''''''''''
                        readSize = chunkSize
                        ReDim myMainArray(0 To readSize - 1) 'clear past data
                        CopyMemory myMainArray(0), ByVal (mBaseAddr + offset), readSize
                        leftSize = (leftSize + miDoubleTextLen) - chunkSize
                        offset = offset + readSize - miDoubleTextLen
                    Else
                        ''''''pending''''''''''''''''''''''''
                        readSize = (leftSize + miDoubleTextLen)
                        ReDim myMainArray(0 To readSize - 1) 'clear past data
                        CopyMemory myMainArray(0), ByVal (mBaseAddr + offset), readSize
                        leftSize = 0: readSize = 0 'flag to exit loop
                    End If
                End If
                
                lsWide = StrConv(myMainArray, vbUnicode)
                If InStr(1, myMainArray, msText, vbTextCompare) > 0 Then
                        ValidateContainedText = True
                        Exit Do
                End If
                 If InStr(1, lsWide, msText, vbTextCompare) > 0 Then
                        ValidateContainedText = True
                        Exit Do
                End If
    Loop
        
End Function


Public Function ByteArrayToString(ByteArray() As Byte) As String
    Dim strResult As String
    Dim lngPos As Long
    
    strResult = StrConv(ByteArray, vbUnicode)
    lngPos = InStr(strResult, Chr(0))
    If lngPos > 0 Then strResult = Left(strResult, lngPos - 1)
    
    ByteArrayToString = strResult
 
 End Function

Private Sub InitializeHeaders()

    ' Set TextArrayHeader for search in strings
    strHeader(0) = 1              ' Number of dimensions
    strHeader(1) = m_PatChrLen    ' Bytes per element (= 2 for integer)/(= 1 for byte)
    'strHeader(2) = '
    'strHeader(3) = mBaseAddr'= StrPtr(m_StrBuffer)' Pointer to StrArray
    'strHeader(4) = mFileSize  ' String Array size (max= &H7FFFFFFF)
    ' set strArray to use strHeader as its own header
    CopyMemory ByVal ArrPtr(strArrayW), VarPtr(strHeader(0)), 4 ' Word Array
    CopyMemory ByVal ArrPtr(strArrayB), VarPtr(strHeader(0)), 4 ' Byte Array
    
    ' Set PatArrayHeader to look at the search text
    patHeader(0) = 1              ' Number of dimensions
    patHeader(1) = m_PatChrLen    ' Bytes per element (= 2 for integer)/(= 1 for byte)
    'patHeader(2) = '
    patHeader(3) = m_PatData(0)   ' Pointer to pattern array 'StrPtr(m_PatBuffer)
    patHeader(4) = m_PatLength    ' Pattern Array size (max= &H7FFFFFFF)
    ' Set patArray to use patHeader as its own header
    CopyMemory ByVal ArrPtr(patArrayW), VarPtr(patHeader(0)), 4 ' Word Array
    CopyMemory ByVal ArrPtr(patArrayB), VarPtr(patHeader(0)), 4 ' Byte Array
    'convert byte array to string
    'm_PatBuffer = StrConv(m_PatArray, vbUnicode)
    'm_PatBuffer = m_PatString
    m_StrBuffer = vbNullString             ' dummy string to hold Text string to search in
    'm_PatBuffer = vbNullString             ' dummy string to hold Pattern string to search for
    
End Sub

Private Sub TerminateHeaders()
    ' Set back strArray and patArray to use their own headers
    
    ' If this code doesn't run the IDE will crash
    CopyMemory ByVal ArrPtr(strArrayW), 0&, 4
    CopyMemory ByVal ArrPtr(patArrayW), 0&, 4
    
    CopyMemory ByVal ArrPtr(strArrayB), 0&, 4
    CopyMemory ByVal ArrPtr(patArrayB), 0&, 4
    
End Sub

Private Sub InitializePattern(Optional charFormat As eCharCode = ANSISTD_CHAR)

   m_PatLength = Len(m_PatString)
   If m_PatLength > 0 Then
   'If m_PatString <> "" Then
      
      If charFormat = ANSISTD_CHAR Then
      'METHOD-1
            'm_PatLength = Len(m_PatString)
            ReDim m_PatArray(m_PatLength - 1)
            'If Not CaseSensitive Then m_PatString = StrConv(m_PatString, vbLowerCase)
            '==================================
            m_PatArray = StrConv(m_PatString, vbFromUnicode)
            '==================================
      'METHOD-2
            'm_PatArray = StrConv(m_PatString, vbFromUnicode)
            'm_PatLength = UBound(m_PatArray) + 1
      ElseIf charFormat = UNICODE_CHAR Then
      'METHOD-1
            m_PatArray = m_PatString 'set String pointer to m_PatArray
            m_PatLength = UBound(m_PatArray) + 1 'OR = m_PatLength * 2
      'METHOD-2
            'm_PatLength = LenB(m_PatString)'OR = m_PatLength * 2
            'ReDim m_PatArray(m_PatLength - 1)
            'CopyMemory m_PatArray(0), ByVal StrPtr(m_PatString), m_PatLength
      End If
      
      'set the input/output variables
      m_PatData(0) = VarPtr(m_PatArray(0))  'ptrPatArray(0)
      m_PatData(1) = CLng(m_PatLength - 1)  'PatLength-1
      m_PatData(2) = CLng(0)                'Default Match Location = 0 (no Match)
      
      If m_SearchAlgorithm = Asm_BMHA Then
            'set skip table for Boyer-Moore horspool String searching Algorithm
            InitializeSkip32
      ElseIf m_SearchAlgorithm = Vb_BMHA Then
      'set skip table for Boyer-Moore horspool String searching Algorithm
            InitializeSkip32
            InitializeHeaders
      ElseIf m_SearchAlgorithm = Vb_InStr Then
            'set safe array headers to map the String in byte or Integer array
            InitializeHeaders
      End If
      
   End If

End Sub

Private Sub InitializeSkip32()

Dim k As Long

   If m_PatLength > 0 Then
            'set skip table for Boyer-Moore horspool String searching Algorithm
            For k = 0 To 255
               m_Skip32Table(k) = CLng(m_PatLength - 1)
            Next k
            For k = 0 To m_PatLength - 2
               m_Skip32Table(m_PatArray(k)) = CLng(m_PatLength - k - 1)
            Next k
   End If

End Sub

Public Function IsUnicode(msText As String) As Boolean
   Dim i As Long
   Dim lngByteSize As Long
   Dim MapStringArray() As Byte
   'msText = msText & ChrW(&H6B22)
   If LenB(msText) Then
      MapStringArray = msText 'set pointer
      'CHECK EACH WIDECHAR HI BYTE VALUE
      lngByteSize = UBound(MapStringArray)
      For i = 1 To lngByteSize Step 2
         If (MapStringArray(i) > 0) Then
            IsUnicode = True
            Exit Function
         End If
      Next
   End If
End Function

Private Function SearchBMH(ByVal StartAt As Long) As Long
Dim i As Long
Dim j As Long
Dim k As Long

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

Public Property Let SearchAlgorithm(New_Value As eSearchAlgorithms)

  'Set the new algorithm
  m_SearchAlgorithm = New_Value

  'Make sure we have initialized the
  'current selected algorithm
  Select Case m_SearchAlgorithm
    Case Asm_BMHA
        If (Not m_BMHA32Init) Then Call InitializeBMHA32
    Case Asm_STRC
        If (Not m_STRC32Init) Then Call InitializeSTRC32
    Case Asm_BYTE
        If (Not m_BYTE32Init) Then Call InitializeBYTE32
  End Select

  'Make sure we reset the data of the
  'current selected algorithm
  Call ClearPatData
  
End Property

Public Property Get SearchAlgorithm() As eSearchAlgorithms

  SearchAlgorithm = m_SearchAlgorithm
  
End Property

Public Function CalculateFileCRC(ByVal strFilePath As String) As String
Dim bArrayFile() As Byte
Dim lngCRC32 As Long
Dim lngChunkSize As Long
Dim lngSize As Long
Dim lngReadPos As Long
Dim FileNo
On Local Error GoTo NoData

lngSize = FileLen(strFilePath)
lngChunkSize = 4096 '=4kb '65536=64kb'

If lngSize <> 0 Then
    ' Read byte array from file
    FileNo = FreeFile
    Open strFilePath For Binary Access Read As #FileNo
    Clear 'clear old CRC value
    lngReadPos = Seek(FileNo)
    Do While lngReadPos < lngSize
        If (lngSize - lngReadPos) > lngChunkSize Then
            Do While lngReadPos < (lngSize - lngChunkSize)
                ReDim bArrayFile(lngChunkSize - 1)
                Get #FileNo, , bArrayFile()
                'Calculate the current chunk
                lngCRC32 = AddBytes(bArrayFile)
                Select Case m_Algorithm
                    Case CRC16
                      Value = m_CRC16 Or lngCRC32 'Setback last actual checksum
                    Case CRC32
                      Value = Not lngCRC32        'Setback last actual checksum
                End Select
                DoEvents
                If StopSearch = True Then 'flag to triminate the process
                    CalculateFileCRC = "00000000"
                    Close #FileNo
                    Exit Function
                End If
                lngReadPos = Seek(FileNo)
            Loop
        Else
            ReDim bArrayFile(lngSize - lngReadPos)
            Get #FileNo, , bArrayFile()
            lngCRC32 = AddBytes(bArrayFile)
            Select Case m_Algorithm
                    Case CRC16
                      Value = m_CRC16 Or lngCRC32 'Setback last actual checksum
                    Case CRC32
                      Value = Not lngCRC32        'Setback last actual checksum
            End Select
            'lngCRC32 = AddBytes(bArrayFile)
        End If
        DoEvents
    lngReadPos = Seek(FileNo)
    Loop

    Close #FileNo
    
    lngCRC32 = Value
NoData:
    CalculateFileCRC = Right$("00000000" & Hex$(lngCRC32), 8)
Else
    CalculateFileCRC = "00000000"
End If

End Function

Public Sub ClearPatData()

  'reset all variables
  m_PatData(0) = 0
  m_PatData(1) = 0
  m_PatData(2) = 0
End Sub


Private Sub InitializeBMHA32()

  Dim i As Long
  Dim sASM As String

  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E55756505351528B45088B750C8B7D108B500481FA010000007E3A31DB31C93B55147D738B45088B58048B0089D18A04188A240E38E0751181FB0000000074644B498B45088B00EBE531C08A04168B048701C2EBCAFC8B008B7D0C89FE8B4D1481FA01000000751E81F9020000007C27668B00F2AE67E31F3A2775F74F89F929F1E91E0000008A00F2AE67E30A4F89F929F1E90D0000008B450831DB895808E907000000418B45088948085A595B585E5F89EC5DC21000"

  ReDim m_BMHA32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_BMHA32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the BMH32 algorithm as initialized
  m_BMHA32Init = True

End Sub

Private Sub InitializeBYTE32()

  Dim i As Long
  Dim sASM As String

  'Create a bytearray to hold the
  'precompiled assembler code
   sASM = "5589E55756505351528B45088B750C8B4D148B388B500442FC89FB89F039CA7F1989D1F3A674204089C689F12B4D0C01D13B4D147F0489DFEBE78B450831DB895808E90C00000089C12B4D0C418B45088948085A595B585E5F89EC5DC21000"
  ReDim m_BYTE32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_BYTE32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the BYTE32 algorithm as initialized
  m_BYTE32Init = True

End Sub

Private Sub InitializeSTRC32()

  Dim i As Long
  Dim sASM As String

  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E55756505351528B45088B7D0C8B088B58044381FB010000007F2D31C089FE8B5D148A21FCAC38E0740E4B81FB0000000075F2E99500000089F12B4D0C8B4508894808E9A10000008A118A710189FE8B4D088B5904438B098A074638D0742589F02B450C01D83B45147F6231C08A064638D0741089F02B450C01D83B45147F4D31C0EBE98A064638F075E58D7EFF81FB0200000074448A61028A0681C60200000038E075A881FB03000000742D8A41038A66FF38E0759681C10200000081EB0200000081FB02000000740FEBC98B450831DB895808E90F0000008D47FF89C12B4D0C418B45088948085A595B585E5F89EC5DC21000"

  ReDim m_STRC32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_STRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the StrStr C algorithm as initialized
  m_STRC32Init = True

End Sub

Private Sub Class_Initialize()
    
  hFile = INVALIDHANDLE
  hFileMap = CLOSEDHANDLE
  hMapView = CLOSEDHANDLE
  mBaseAddr = CLOSEDHANDLE
  mFileSize = CLOSEDHANDLE
  Algorithm = CRC32 'The default algorithm is CRC32
  If (Not m_BYTE32Init) Then Call InitializeBYTE32
  If (Not m_STRC32Init) Then Call InitializeSTRC32
  If (Not m_BMHA32Init) Then Call InitializeBMHA32
  'InitializeHeaders
  SearchAlgorithm = Asm_BMHA 'The default algorithm is BMHA
End Sub

Private Sub Class_Terminate()
TerminateHeaders
End Sub

Public Sub GetMemStatus()
Dim lpBuffer As MEMORYSTATUS
Dim txtDisplay As String

GlobalMemoryStatus lpBuffer
With lpBuffer
txtDisplay = txtDisplay & "AvailPageFile=" & .dwAvailPageFile & vbCrLf
txtDisplay = txtDisplay & "AvailPhyscal =" & .dwAvailPhys & vbCrLf
txtDisplay = txtDisplay & "AvailVirtual =" & .dwAvailVirtual & vbCrLf
txtDisplay = txtDisplay & "Length =" & .dwLength & vbCrLf
txtDisplay = txtDisplay & "MemoryLoad =" & .dwMemoryLoad & vbCrLf
txtDisplay = txtDisplay & "TotalPageFile =" & .dwTotalPageFile & vbCrLf
txtDisplay = txtDisplay & "TotalPhysical Memory =" & .dwTotalPhys & vbCrLf
txtDisplay = txtDisplay & "TotalVirtual Memory =" & .dwTotalVirtual
MsgBox txtDisplay, vbOKOnly, "Memory Status"
End With
End Sub


'CRC Checksum Class
'------------------------------------
''(c) 2000, Fredrik Qvarfort
Public Function AddBytes(ByteArray() As Byte) As Variant

  Dim byteSize As Long
  
  'We need to add a simple error trapping
  'here because if the bytearray is not
  'dimensioned we want it to just skip
  'the assembler code call below
  On Local Error GoTo NoData
  
  'Precalculate the size of the byte array
  byteSize = UBound(ByteArray) - LBound(ByteArray) + 1
  
  'No error trapping needed, if something
  'goes bad below something is definitely
  'fishy with your computer
  On Local Error GoTo 0
  
  'Run the pre-compiled assembler code
  'for the current selected algorithm
  Select Case m_Algorithm
  Case CRC16
    Call CallWindowProc(VarPtr(m_CRC16Asm(0)), VarPtr(m_CRC16), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC16Table(0)), byteSize)
  Case CRC32
    Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC32Table(0)), byteSize)
  End Select
  
NoData:
  'Return the current CRC value
  AddBytes = Value
  
End Function
Public Function AddString(Text As String) As Variant

  'Convert the string into a byte array
  'and send it to the function that can
  'handle bytearrays
  AddString = AddBytes(StrConv(Text, vbFromUnicode))
  
End Function
Public Property Let Algorithm(New_Value As CRCAlgorithms)

  'Set the new algorithm
  m_Algorithm = New_Value

  'Make sure we have initialized the
  'current selected algorithm
  Select Case m_Algorithm
  Case CRC16
    If (Not m_CRC16Init) Then Call InitializeCRC16
  Case CRC32
    If (Not m_CRC32Init) Then Call InitializeCRC32
  End Select

  'Make sure we reset the data of the
  'current selected algorithm
  Call Clear
  
End Property
Public Property Get Algorithm() As CRCAlgorithms

  Algorithm = m_Algorithm
  
End Property

Public Function CalculateBytes(ByteArray() As Byte) As Variant

  'Reset the current CRC calculation
  Call Clear
  
  'Calculate the CRC from the bytearray
  'and return the current CRC value
  CalculateBytes = AddBytes(ByteArray)
  
End Function

Public Function CalculateFile(FileName As String) As Variant

  Dim Filenr As Integer
  Dim ByteArray() As Byte
  
  'Make sure the file contains data
  'to avoid errors later below
  If (FileLen(FileName) = 0) Then Exit Function
  
  'Open the file in binary mode, read
  'the data into a bytearray and then
  'close the file
  On Error GoTo CalcErrHandler
  Filenr = FreeFile
  Open FileName For Binary As #Filenr
  ReDim ByteArray(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  'MsgBox InBArrBM(ByteArray, "test", 0, 0, False)
  'Now send the bytearray to the function
  'that can calculate a CRC from it
  CalculateFile = CalculateBytes(ByteArray)
  Exit Function
  
CalcErrHandler:
  CalculateFile = "00000000"
End Function

Public Function CalculateString(Text As String)

  'Convert the string into a bytearray
  'and send it to the function that
  'calculates the CRC from a bytearray
  CalculateString = CalculateBytes(StrConv(Text, vbFromUnicode))
  
End Function
Public Property Get Value() As Variant

  Select Case m_Algorithm
  Case CRC16
    Value = (m_CRC16 And 65535)
  Case CRC32
    Value = (Not m_CRC32)
  End Select
  
End Property

Public Property Let Value(New_Value As Variant)

  Select Case m_Algorithm
  Case CRC16
    m_CRC16 = New_Value
  Case CRC32
    m_CRC32 = New_Value
  End Select
  
End Property

Private Sub InitializeCRC16()

  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim CRC As Long
  Dim sASM As String
  
  'Create the fixed lookup-table, this
  'is calculated because it won't take
  'long and is only done once
  For i = 0 To 255
    k = i * 256
    CRC = 0
    For j = 0 To 7
      If (((CRC Xor k) And 32768) = 32768) Then
        CRC = (CRC * 2) Xor &H1021
      Else
        CRC = (CRC * 2)
      End If
      k = k * 2
    Next
    m_CRC16Table(i) = CRC '(CRC And 65535)
  Next
  
  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E55756505351528B45088B008B750C8B7D108B4D1431DB8A1E30E3668B149F30C66689D0464975EF25FFFF00008B4D0889015A595B585E5F89EC5DC21000"
  ReDim m_CRC16Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_CRC16Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the CRC16 algorithm as initialized
  m_CRC16Init = True
  
End Sub
Public Sub Clear()

  'Here can be sloppy and reset both
  'crc variables (this procedure will
  'be more advanced when adding more
  'checksums algorithms..)
  m_CRC16 = 0
  m_CRC32 = &HFFFFFFFF
End Sub

Private Sub InitializeCRC32()

  Dim i As Long
  Dim sASM As String
  
  m_CRC32Table(0) = &H0
  m_CRC32Table(1) = &H77073096
  m_CRC32Table(2) = &HEE0E612C
  m_CRC32Table(3) = &H990951BA
  m_CRC32Table(4) = &H76DC419
  m_CRC32Table(5) = &H706AF48F
  m_CRC32Table(6) = &HE963A535
  m_CRC32Table(7) = &H9E6495A3
  m_CRC32Table(8) = &HEDB8832
  m_CRC32Table(9) = &H79DCB8A4
  m_CRC32Table(10) = &HE0D5E91E
  m_CRC32Table(11) = &H97D2D988
  m_CRC32Table(12) = &H9B64C2B
  m_CRC32Table(13) = &H7EB17CBD
  m_CRC32Table(14) = &HE7B82D07
  m_CRC32Table(15) = &H90BF1D91
  m_CRC32Table(16) = &H1DB71064
  m_CRC32Table(17) = &H6AB020F2
  m_CRC32Table(18) = &HF3B97148
  m_CRC32Table(19) = &H84BE41DE
  m_CRC32Table(20) = &H1ADAD47D
  m_CRC32Table(21) = &H6DDDE4EB
  m_CRC32Table(22) = &HF4D4B551
  m_CRC32Table(23) = &H83D385C7
  m_CRC32Table(24) = &H136C9856
  m_CRC32Table(25) = &H646BA8C0
  m_CRC32Table(26) = &HFD62F97A
  m_CRC32Table(27) = &H8A65C9EC
  m_CRC32Table(28) = &H14015C4F
  m_CRC32Table(29) = &H63066CD9
  m_CRC32Table(30) = &HFA0F3D63
  m_CRC32Table(31) = &H8D080DF5
  m_CRC32Table(32) = &H3B6E20C8
  m_CRC32Table(33) = &H4C69105E
  m_CRC32Table(34) = &HD56041E4
  m_CRC32Table(35) = &HA2677172
  m_CRC32Table(36) = &H3C03E4D1
  m_CRC32Table(37) = &H4B04D447
  m_CRC32Table(38) = &HD20D85FD
  m_CRC32Table(39) = &HA50AB56B
  m_CRC32Table(40) = &H35B5A8FA
  m_CRC32Table(41) = &H42B2986C
  m_CRC32Table(42) = &HDBBBC9D6
  m_CRC32Table(43) = &HACBCF940
  m_CRC32Table(44) = &H32D86CE3
  m_CRC32Table(45) = &H45DF5C75
  m_CRC32Table(46) = &HDCD60DCF
  m_CRC32Table(47) = &HABD13D59
  m_CRC32Table(48) = &H26D930AC
  m_CRC32Table(49) = &H51DE003A
  m_CRC32Table(50) = &HC8D75180
  m_CRC32Table(51) = &HBFD06116
  m_CRC32Table(52) = &H21B4F4B5
  m_CRC32Table(53) = &H56B3C423
  m_CRC32Table(54) = &HCFBA9599
  m_CRC32Table(55) = &HB8BDA50F
  m_CRC32Table(56) = &H2802B89E
  m_CRC32Table(57) = &H5F058808
  m_CRC32Table(58) = &HC60CD9B2
  m_CRC32Table(59) = &HB10BE924
  m_CRC32Table(60) = &H2F6F7C87
  m_CRC32Table(61) = &H58684C11
  m_CRC32Table(62) = &HC1611DAB
  m_CRC32Table(63) = &HB6662D3D
  m_CRC32Table(64) = &H76DC4190
  m_CRC32Table(65) = &H1DB7106
  m_CRC32Table(66) = &H98D220BC
  m_CRC32Table(67) = &HEFD5102A
  m_CRC32Table(68) = &H71B18589
  m_CRC32Table(69) = &H6B6B51F
  m_CRC32Table(70) = &H9FBFE4A5
  m_CRC32Table(71) = &HE8B8D433
  m_CRC32Table(72) = &H7807C9A2
  m_CRC32Table(73) = &HF00F934
  m_CRC32Table(74) = &H9609A88E
  m_CRC32Table(75) = &HE10E9818
  m_CRC32Table(76) = &H7F6A0DBB
  m_CRC32Table(77) = &H86D3D2D
  m_CRC32Table(78) = &H91646C97
  m_CRC32Table(79) = &HE6635C01
  m_CRC32Table(80) = &H6B6B51F4
  m_CRC32Table(81) = &H1C6C6162
  m_CRC32Table(82) = &H856530D8
  m_CRC32Table(83) = &HF262004E
  m_CRC32Table(84) = &H6C0695ED
  m_CRC32Table(85) = &H1B01A57B
  m_CRC32Table(86) = &H8208F4C1
  m_CRC32Table(87) = &HF50FC457
  m_CRC32Table(88) = &H65B0D9C6
  m_CRC32Table(89) = &H12B7E950
  m_CRC32Table(90) = &H8BBEB8EA
  m_CRC32Table(91) = &HFCB9887C
  m_CRC32Table(92) = &H62DD1DDF
  m_CRC32Table(93) = &H15DA2D49
  m_CRC32Table(94) = &H8CD37CF3
  m_CRC32Table(95) = &HFBD44C65
  m_CRC32Table(96) = &H4DB26158
  m_CRC32Table(97) = &H3AB551CE
  m_CRC32Table(98) = &HA3BC0074
  m_CRC32Table(99) = &HD4BB30E2
  m_CRC32Table(100) = &H4ADFA541
  m_CRC32Table(101) = &H3DD895D7
  m_CRC32Table(102) = &HA4D1C46D
  m_CRC32Table(103) = &HD3D6F4FB
  m_CRC32Table(104) = &H4369E96A
  m_CRC32Table(105) = &H346ED9FC
  m_CRC32Table(106) = &HAD678846
  m_CRC32Table(107) = &HDA60B8D0
  m_CRC32Table(108) = &H44042D73
  m_CRC32Table(109) = &H33031DE5
  m_CRC32Table(110) = &HAA0A4C5F
  m_CRC32Table(111) = &HDD0D7CC9
  m_CRC32Table(112) = &H5005713C
  m_CRC32Table(113) = &H270241AA
  m_CRC32Table(114) = &HBE0B1010
  m_CRC32Table(115) = &HC90C2086
  m_CRC32Table(116) = &H5768B525
  m_CRC32Table(117) = &H206F85B3
  m_CRC32Table(118) = &HB966D409
  m_CRC32Table(119) = &HCE61E49F
  m_CRC32Table(120) = &H5EDEF90E
  m_CRC32Table(121) = &H29D9C998
  m_CRC32Table(122) = &HB0D09822
  m_CRC32Table(123) = &HC7D7A8B4
  m_CRC32Table(124) = &H59B33D17
  m_CRC32Table(125) = &H2EB40D81
  m_CRC32Table(126) = &HB7BD5C3B
  m_CRC32Table(127) = &HC0BA6CAD
  m_CRC32Table(128) = &HEDB88320
  m_CRC32Table(129) = &H9ABFB3B6
  m_CRC32Table(130) = &H3B6E20C
  m_CRC32Table(131) = &H74B1D29A
  m_CRC32Table(132) = &HEAD54739
  m_CRC32Table(133) = &H9DD277AF
  m_CRC32Table(134) = &H4DB2615
  m_CRC32Table(135) = &H73DC1683
  m_CRC32Table(136) = &HE3630B12
  m_CRC32Table(137) = &H94643B84
  m_CRC32Table(138) = &HD6D6A3E
  m_CRC32Table(139) = &H7A6A5AA8
  m_CRC32Table(140) = &HE40ECF0B
  m_CRC32Table(141) = &H9309FF9D
  m_CRC32Table(142) = &HA00AE27
  m_CRC32Table(143) = &H7D079EB1
  m_CRC32Table(144) = &HF00F9344
  m_CRC32Table(145) = &H8708A3D2
  m_CRC32Table(146) = &H1E01F268
  m_CRC32Table(147) = &H6906C2FE
  m_CRC32Table(148) = &HF762575D
  m_CRC32Table(149) = &H806567CB
  m_CRC32Table(150) = &H196C3671
  m_CRC32Table(151) = &H6E6B06E7
  m_CRC32Table(152) = &HFED41B76
  m_CRC32Table(153) = &H89D32BE0
  m_CRC32Table(154) = &H10DA7A5A
  m_CRC32Table(155) = &H67DD4ACC
  m_CRC32Table(156) = &HF9B9DF6F
  m_CRC32Table(157) = &H8EBEEFF9
  m_CRC32Table(158) = &H17B7BE43
  m_CRC32Table(159) = &H60B08ED5
  m_CRC32Table(160) = &HD6D6A3E8
  m_CRC32Table(161) = &HA1D1937E
  m_CRC32Table(162) = &H38D8C2C4
  m_CRC32Table(163) = &H4FDFF252
  m_CRC32Table(164) = &HD1BB67F1
  m_CRC32Table(165) = &HA6BC5767
  m_CRC32Table(166) = &H3FB506DD
  m_CRC32Table(167) = &H48B2364B
  m_CRC32Table(168) = &HD80D2BDA
  m_CRC32Table(169) = &HAF0A1B4C
  m_CRC32Table(170) = &H36034AF6
  m_CRC32Table(171) = &H41047A60
  m_CRC32Table(172) = &HDF60EFC3
  m_CRC32Table(173) = &HA867DF55
  m_CRC32Table(174) = &H316E8EEF
  m_CRC32Table(175) = &H4669BE79
  m_CRC32Table(176) = &HCB61B38C
  m_CRC32Table(177) = &HBC66831A
  m_CRC32Table(178) = &H256FD2A0
  m_CRC32Table(179) = &H5268E236
  m_CRC32Table(180) = &HCC0C7795
  m_CRC32Table(181) = &HBB0B4703
  m_CRC32Table(182) = &H220216B9
  m_CRC32Table(183) = &H5505262F
  m_CRC32Table(184) = &HC5BA3BBE
  m_CRC32Table(185) = &HB2BD0B28
  m_CRC32Table(186) = &H2BB45A92
  m_CRC32Table(187) = &H5CB36A04
  m_CRC32Table(188) = &HC2D7FFA7
  m_CRC32Table(189) = &HB5D0CF31
  m_CRC32Table(190) = &H2CD99E8B
  m_CRC32Table(191) = &H5BDEAE1D
  m_CRC32Table(192) = &H9B64C2B0
  m_CRC32Table(193) = &HEC63F226
  m_CRC32Table(194) = &H756AA39C
  m_CRC32Table(195) = &H26D930A
  m_CRC32Table(196) = &H9C0906A9
  m_CRC32Table(197) = &HEB0E363F
  m_CRC32Table(198) = &H72076785
  m_CRC32Table(199) = &H5005713
  m_CRC32Table(200) = &H95BF4A82
  m_CRC32Table(201) = &HE2B87A14
  m_CRC32Table(202) = &H7BB12BAE
  m_CRC32Table(203) = &HCB61B38
  m_CRC32Table(204) = &H92D28E9B
  m_CRC32Table(205) = &HE5D5BE0D
  m_CRC32Table(206) = &H7CDCEFB7
  m_CRC32Table(207) = &HBDBDF21
  m_CRC32Table(208) = &H86D3D2D4
  m_CRC32Table(209) = &HF1D4E242
  m_CRC32Table(210) = &H68DDB3F8
  m_CRC32Table(211) = &H1FDA836E
  m_CRC32Table(212) = &H81BE16CD
  m_CRC32Table(213) = &HF6B9265B
  m_CRC32Table(214) = &H6FB077E1
  m_CRC32Table(215) = &H18B74777
  m_CRC32Table(216) = &H88085AE6
  m_CRC32Table(217) = &HFF0F6A70
  m_CRC32Table(218) = &H66063BCA
  m_CRC32Table(219) = &H11010B5C
  m_CRC32Table(220) = &H8F659EFF
  m_CRC32Table(221) = &HF862AE69
  m_CRC32Table(222) = &H616BFFD3
  m_CRC32Table(223) = &H166CCF45
  m_CRC32Table(224) = &HA00AE278
  m_CRC32Table(225) = &HD70DD2EE
  m_CRC32Table(226) = &H4E048354
  m_CRC32Table(227) = &H3903B3C2
  m_CRC32Table(228) = &HA7672661
  m_CRC32Table(229) = &HD06016F7
  m_CRC32Table(230) = &H4969474D
  m_CRC32Table(231) = &H3E6E77DB
  m_CRC32Table(232) = &HAED16A4A
  m_CRC32Table(233) = &HD9D65ADC
  m_CRC32Table(234) = &H40DF0B66
  m_CRC32Table(235) = &H37D83BF0
  m_CRC32Table(236) = &HA9BCAE53
  m_CRC32Table(237) = &HDEBB9EC5
  m_CRC32Table(238) = &H47B2CF7F
  m_CRC32Table(239) = &H30B5FFE9
  m_CRC32Table(240) = &HBDBDF21C
  m_CRC32Table(241) = &HCABAC28A
  m_CRC32Table(242) = &H53B39330
  m_CRC32Table(243) = &H24B4A3A6
  m_CRC32Table(244) = &HBAD03605
  m_CRC32Table(245) = &HCDD70693
  m_CRC32Table(246) = &H54DE5729
  m_CRC32Table(247) = &H23D967BF
  m_CRC32Table(248) = &HB3667A2E
  m_CRC32Table(249) = &HC4614AB8
  m_CRC32Table(250) = &H5D681B02
  m_CRC32Table(251) = &H2A6F2B94
  m_CRC32Table(252) = &HB40BBE37
  m_CRC32Table(253) = &HC30C8EA1
  m_CRC32Table(254) = &H5A05DF1B
  m_CRC32Table(255) = &H2D02EF8D

  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
  ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_CRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the CRC32 algorithm as initialized
  m_CRC32Init = True

End Sub



