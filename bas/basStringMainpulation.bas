Attribute VB_Name = "basStringMainpulation"
Option Explicit
'For best results, compile to Native code w/ the following optimizations:
'Remove Array Bounds Checks
'Remove Integer Overflow checks
'Remove Safe Pentium FDIV

'Obviously, don't use thes optimizations if you have additional code in the
'same component that you don't know would be safe.


'Do NOT use assume no aliasing, as there are many ByRef parameters

'SAFEARRAY Header, used in place of the real one to trick VB
'into letting us access string data in-place
Public Type tSafeArray1D
    Dimensions As Integer
    Attributes As Integer
    BytesPerElement As Long
    Locks As Long
    DataPointer As Long
    Elements As Long
    LBound As Long
End Type

'Safearray attributes to disallow redim I don't have redim in here anyway,
'but if you are copying the string map code, you will probably want to use these.

Public Const SAFEARRAY_AUTO = &H1
Public Const SAFEARRAY_FIXEDSIZE = &H10

'Used for unsigned addition
Private Const DWORDMostSignificantBit = &H80000000

'This is the header that will be used in place of the real header for myMap
Private mtHeader            As tSafeArray1D
Private miArrayPointer      As Long
Private miOldDescriptor     As Long

'Array of delimiters for the replace function, i.e. ". ;,?!"
Private myDelimiters()      As Byte
'Array of exclusions for before a match found in the replace function
Private msPreExclusions()   As String
'Array of exclusions for after a match found in the replace function
Private msPostExclusions()  As String

'Used to access string data in-place
Private myMap()             As Byte

'Used to avoid having to trap errors w/ UBound()
Private miDelimCount        As Long
Private miPreExclusions     As Long
Private miPostExclusions    As Long


'Quickly Allocate a string.  Thanks to Rde for this declare
Private Declare Function SysAllocStringByteLen Lib "oleaut32" ( _
                            ByVal pszStr As Long, _
                            ByVal lLenB As Long _
                         ) As Long

'Workaround b/c VB prefers to do a type check on AS ANY params!
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
                            ByRef ptr() As Any _
                         ) As Long

'The ever-useful but just plain simple copymemory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                            ByRef Dst As Any, _
                            ByRef Src As Any, _
                            ByVal ByteLen As Long _
                         )
                         
Public Function ReplaceString(ByRef psOriginal As String, _
                              ByRef psFind As String, _
                              ByRef psReplace As String, _
                     Optional ByVal piStart As Long, _
                     Optional ByRef piReplacementCount As Long, _
                     Optional ByVal piCompare As VbCompareMethod = vbBinaryCompare _
                ) As String
                'The three string arguments are not modified at all.
                'piReplacementCount is in/Out.  In: Maximum number of replacements
                '                               Out: Number of replacements made
                
    Dim liOriginalBytes As Long 'Number of bytes in the psOriginal argument
    Dim liFindBytes     As Long 'Number of bytes in the psFind argument
    Dim liReplaceBytes  As Long 'Number of bytes in the psReplace argument
    Dim liMaxReplace    As Long 'Maximum number of replacements to make
    Dim liTempLen       As Long 'Temporary length used to calculate offsets
    
    Dim liReplacedPtr   As Long 'Pointer to the current position in the replaced string
    Dim liOriginalPtr   As Long 'Pointer the the current position in the original string
    Dim liReplacePtr    As Long 'Pointer the the beginning of psReplace
    Dim liLastFind      As Long 'Byte position of the last match for psFind
    Dim liFound()       As Long 'Array to hold the positions that will be replaced
    Dim i               As Long 'Main counter, current place in myMap
    Dim j               As Long 'Secondary counter, loops through lyFind
    Dim k               As Long 'Triple duty:  counter through myDelimiters
                                '              Secondary counter for lyFind
                                '              byte diff between psFind & psReplace
    
    Dim liCopiedLen     As Long 'Total number of bytes copied from the original string
    
    Dim lyFind()        As Byte 'Byte array that holds a copy of psFind
    Dim lyFirstFind     As Byte 'First byte of psFind
    Dim lySecondFind    As Byte 'Second byte
    Dim lyByte          As Byte 'Temporary variable for byte comparison
    Dim lyByte1         As Byte 'Temporary variable for byte comparison
    Dim lyByte2         As Byte 'Temporary variable for byte comparison
    
    
    Dim lbIgnoreCase    As Boolean 'Whether text or binary compare is being done
    Dim lbValidate      As Boolean 'Indicates if the delimiter or exclusion was verified.
    
'####################################
'###  STEP 1: Initialization      ###
'####################################
    
    'Ignore case only if vbTextCompare was passed, otherwise do a binary compare
    lbIgnoreCase = piCompare = vbTextCompare
    
    'Set up the values for original lengths
    'These values are not modified again
    liOriginalBytes = LenB(psOriginal)
    liFindBytes = LenB(psFind)
    liReplaceBytes = LenB(psReplace)
    
    'Store the string pointers
    liOriginalPtr = StrPtr(psOriginal)
    liReplacePtr = StrPtr(psReplace)
    
    'Make sure that we aren't being given an incorrect starting point.
    If piStart <= 0& Then piStart = 0& Else piStart = (piStart - 1&) * 2&
    
    'Make sure that there's something to do.
    If liFindBytes < 2& Or liOriginalBytes < 2& Or piStart > liOriginalBytes Then
        'If not, this will be the easiest replace function ever!
        If liOriginalBytes > 0& Then
            liReplacedPtr = SysAllocStringByteLen(0&, liOriginalBytes)
            CopyMemory ByVal VarPtr(ReplaceString), liReplacedPtr, 4&
            CopyMemory ByVal liReplacedPtr, ByVal liOriginalPtr, liOriginalBytes
        End If
        Exit Function
    End If
    
    lyFind = psFind 'Initialize the bytes being looked for
    
    'Allocate the array to hold the positions to be replaced
    'with the maximum possible locations.
    ReDim liFound(0& To liOriginalBytes \ liFindBytes + 1&)
    
    If lbIgnoreCase Then
        'If we're ignoring case then we need to lcase$() all of the bytes that we're looking
        'for. Easier to do it once at the beginning then every time we compare against them
        For i = 0& To liFindBytes - 1& Step 2&
            CharLower lyFind(i), lyFind(i + 1)
        Next
    End If
    
    lyFirstFind = lyFind(0&) 'Initialize the first byte to look for
    lySecondFind = lyFind(1&)
    
    'Store the maximum replacements to be made
    If piReplacementCount > 0& Then liMaxReplace = piReplacementCount
    piReplacementCount = 0& 'Start counting replacements at 0&
    
    GetStringMap myMap, psOriginal, mtHeader, miOldDescriptor
    
'####################################
'###  STEP 2: Find Matches        ###
'####################################
    
    'Stepping by two b/c of unicode.
    'Did I mention that this function only works w/ unicode psOriginal and psFind strings?
    'Although it could be modified for ANSI relatively easily
    For i = piStart To liOriginalBytes - 2& Step 2&
        'Store the current byte
        lyByte = myMap(i)
        lyByte2 = myMap(i + 1)
        'If we're ignoring case then we need to lCase$() it
        If lbIgnoreCase Then CharLower lyByte, lyByte2
        
        'Could this be the beginning of what we're looking for?
        If lyByte = lyFirstFind And lyByte2 = lySecondFind Then
            'It Could!
            If miDelimCount > 0& Then
                'if there are delimiters, then we need to see if the current
                'byte is preceded by a valid delimiter
                If i >= 2& Then
                    lyByte1 = myMap(i - 2&)
                    lyByte2 = myMap(i - 1&)
                    For k = 0& To miDelimCount - 1&
                        lbValidate = lyByte1 = myDelimiters(k, 0&) _
                                                    And _
                                       lyByte2 = myDelimiters(k, 1&)
                        If lbValidate Then Exit For
                    Next
                Else
                    'If we're at the beginning of the string
                    'then no delimiter check is necessary.
                    lbValidate = True
                End If
            Else
                'No delimiter check necessary
                lbValidate = True
            End If
            
            If miPreExclusions > 0& And lbValidate Then lbValidate = ValidateExclusions(True, i, liOriginalBytes - 1, lbIgnoreCase)
            
            If lbValidate Then
                'We've matched the first byte we're looking for, and we have a valid
                'delimiter, so now we can see if we have the rest of psFind
                
                j = i + liFindBytes
                'if there aren't enough bytes left in the string, then it's no use
                If j <= liOriginalBytes Then
                    If lbIgnoreCase Then
                        'if we're ignoring case then we need to lcase$() all the bytes
                        'from the myMap to compare them with the bytes we're looking for.
                        k = 2&
                        For j = i + 2& To j - 2& Step 2&
                            lyByte = myMap(j)
                            lyByte2 = myMap(j + 1&)
                            CharLower lyByte, lyByte2
                            'If the bytes don't match, stop looking
                            If Not lyByte = lyFind(k) Or Not lyByte2 = lyFind(k + 1&) Then Exit For
                            k = k + 2&
                        Next
                    Else
                        'If we're doing a binary compare, there's no need to check
                        'the lbIgnoreCase through every iteration.
                        k = 2&
                        For j = i + 2& To j - 2& Step 2&
                            If Not myMap(j) = lyFind(k) Or Not myMap(j + 1) = lyFind(k + 1) Then Exit For
                            k = k + 2&
                        Next
                    End If
                    
                    'did we find a match?
                    If j >= i + liFindBytes Then
                        'Yes we did!
                        
                        If miDelimCount > 0& Then
                            If j + 3& <= liOriginalBytes Then
                                'if delimiters are set up, make sure that a valid delimiter
                                'appears after the current byte
                                lyByte1 = myMap(j)
                                lyByte2 = myMap(j + 1&)
                                For k = 0& To miDelimCount - 1&
                                    lbValidate = lyByte1 = myDelimiters(k, 0&) _
                                                                And _
                                                   lyByte2 = myDelimiters(k, 1&)
                                    If lbValidate Then Exit For
                                Next
                            Else
                                'If we're at the end of the string, there's no need to
                                'check for a valid delimiter
                                lbValidate = True
                            End If
                        Else
                            'No delimiter check necessary
                            lbValidate = True
                        End If
                        
                        If miPostExclusions > 0& And lbValidate Then lbValidate = ValidateExclusions(False, i + liFindBytes, liOriginalBytes - 1&, lbIgnoreCase)
                        
                        If lbValidate Then
                            'Now we've found a complete match that is enclosed in valid delimiters.
                            
                            'Store the current relative position
                            If piReplacementCount <> 0& Then _
                                liFound(piReplacementCount) = i - liLastFind - liFindBytes _
                            Else _
                                liFound(piReplacementCount) = i
                            'Inc the number of replacements
                            piReplacementCount = piReplacementCount + 1&
                            'Remember the last position
                            liLastFind = i
                            'Fool w/ the counter variable to skip the match that we just found
                            i = j - 2&
                            'Make sure we don't go over the limit that was provided.
                            If piReplacementCount = liMaxReplace Then Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    'Restore the original descriptor for the modular array.
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    
'####################################
'###  STEP 3: Build Return Value  ###
'####################################
    
    'Additional bytes added or removed for each replacement
    k = liReplaceBytes - liFindBytes
    liTempLen = k * piReplacementCount + liOriginalBytes
    If liTempLen = 0& Then Exit Function 'Exit if nothing to do!
    'Quickly allocate the string  (Thanks to Rde for the declare!)
    liReplacedPtr = SysAllocStringByteLen(0&, liTempLen)
    'Tell ReplaceString to point to the newly allocated string
    CopyMemory ByVal VarPtr(ReplaceString), liReplacedPtr, 4&
    
    For i = 0& To piReplacementCount - 1&
        liTempLen = liFound(i)
        
        If liTempLen > 0& Then
            CopyMemory ByVal liReplacedPtr, ByVal liOriginalPtr, liTempLen
            liCopiedLen = liCopiedLen + liTempLen
            'Workaround for VB's lack of an unsigned integer.
            'Check for the most significant bit, and adjust accordingly.
            'Add the bytes that we just copied to the return string pointer
            If liReplacedPtr And DWORDMostSignificantBit Then
               liReplacedPtr = liReplacePtr + liTempLen
            ElseIf (liReplacedPtr Or DWORDMostSignificantBit) < -liTempLen Then
               liReplacedPtr = liReplacedPtr + liTempLen
            Else
               liReplacedPtr = (liReplacedPtr + DWORDMostSignificantBit) + _
                               (liTempLen + DWORDMostSignificantBit)
            End If
        End If
        
        'We want to skip the bytes that are being replaced
        liTempLen = liTempLen + liFindBytes
        liCopiedLen = liCopiedLen + liFindBytes
        
        If liTempLen > 0& Then
            'Add the bytes we just copied plus the length of the matched bytes
            'to the psOriginal Pointer
            If liOriginalPtr And DWORDMostSignificantBit Then
               liOriginalPtr = liOriginalPtr + liTempLen
            ElseIf (liOriginalPtr Or DWORDMostSignificantBit) < -liTempLen Then
               liOriginalPtr = liOriginalPtr + liTempLen
            Else
               liOriginalPtr = (liOriginalPtr + DWORDMostSignificantBit) + _
                               (liTempLen + DWORDMostSignificantBit)
            End If
        End If
        
        If liReplaceBytes > 0& Then
            'Copy psReplaced to the next position in the string
            CopyMemory ByVal liReplacedPtr, ByVal liReplacePtr, liReplaceBytes
            
            'Add the bytes we just copied to the return string pointer
            If liReplacedPtr And DWORDMostSignificantBit Then
               liReplacedPtr = liReplacedPtr + liReplaceBytes
            ElseIf (liReplacedPtr Or DWORDMostSignificantBit) < -liReplaceBytes Then
               liReplacedPtr = liReplacedPtr + liReplaceBytes
            Else
               liReplacedPtr = (liReplacedPtr + DWORDMostSignificantBit) + _
                               (liReplaceBytes + DWORDMostSignificantBit)
            End If
        End If
    Next
    'Unless we replaced the very last bytes of the original string, we will
    'need to copy over the remainder of the original string
    liTempLen = liOriginalBytes - liCopiedLen
    If liTempLen > 0& Then CopyMemory ByVal liReplacedPtr, ByVal liOriginalPtr, liTempLen
    'Whew!  Does it really have to be so complicated?  (yes)
End Function

Private Function ValidateExclusions(ByVal pbBefore As Boolean, _
                                    ByVal piPlace As Long, _
                                    ByVal piMax As Long, _
                                    ByVal pbIgnoreCase As Boolean _
                ) As Boolean
'Helper function for ReplaceString
'The string being searched does not have to be passed b/c is it is a modular array
    
    Dim i As Long 'Counter through the modular array
    Dim j As Long 'Counter through myMap
    Dim k As Long 'Counter through lyFind
    Dim liStart As Long 'Start counting through myMap
    Dim liFinish As Long 'Finish counting through mymap
    Dim liLen As Long 'Array ubound then length of each string
    Dim lbVal As Boolean 'Whether there is sufficient room to bother checking
    
    Dim lyByte1 As Byte 'Bytes from myMap to compare
    Dim lyByte2 As Byte
    
    Dim lyExclude1 As Byte 'Bytes from lyFind to compare
    Dim lyExclude2 As Byte
    
    Dim ltHeader        As tSafeArray1D 'Custom header for lyFind
    Dim liOldDescriptor As Long 'Original Descriptor for lyFind
    Dim lyFind()        As Byte 'Bytes to point to each exclusion string
    
    'Get the correct array ubound
    If pbBefore Then liLen = miPreExclusions - 1& Else liLen = miPostExclusions - 1&
    
    For i = 0 To liLen
        'Get the correct string, and validate if there is enough room
        If pbBefore Then
            liLen = LenB(msPreExclusions(i))
            lbVal = liLen <= piPlace
        Else
            liLen = LenB(msPostExclusions(i))
            lbVal = liLen + piPlace <= piMax
        End If
        
        If lbVal Then
            'If there's enough room
            
            'Only store the original header from the first iteration
            If i > 0& Then
                'Get the string map to the correct string
                If pbBefore Then
                    GetStringMap lyFind, msPreExclusions(i), ltHeader, 0&
                Else
                    GetStringMap lyFind, msPostExclusions(i), ltHeader, 0&
                End If
            Else
                'Get the string map to the correct string
                If pbBefore Then
                    GetStringMap lyFind, msPreExclusions(i), ltHeader, liOldDescriptor
                Else
                    GetStringMap lyFind, msPostExclusions(i), ltHeader, liOldDescriptor
                End If
            End If

            'get the correct entry and exit points
            If pbBefore Then
                liStart = piPlace - liLen
                liFinish = piPlace - 2&
            Else
                liStart = piPlace
                liFinish = piPlace + liLen - 2&
            End If

            k = 0&
            'Loop through lyFind and myMap to see if they match
            For j = liStart To liFinish Step 2&
                lyByte1 = myMap(j)
                lyByte2 = myMap(j + 1&)
                lyExclude1 = lyFind(k)
                lyExclude2 = lyFind(k + 1&)
                If pbIgnoreCase Then
                    'Ignore case if we're supposed to
                    CharLower lyByte1, lyByte2
                    CharLower lyExclude1, lyExclude2
                End If
                'If the bytes don't match, then don't continue checking
                If lyByte1 <> lyExclude1 Or lyByte2 <> lyExclude2 Then Exit For
                'Inc the secondary counter
                k = k + 2&
            Next
            'If we found a match then get outta here
            If j > liFinish Then Exit Function
        End If
    Next
    'Return the original array descriptor
    CopyMemory ByVal ArrPtr(lyFind), liOldDescriptor, 4&
    'no matches found!
    ValidateExclusions = True

End Function

Public Sub SetReplaceDelimiters(psDelims As String)
    'Must call this w/ a unicode string.  Don't use StrConv(psDelims, vbFromUnicode)!
    'Call with a string defining the delimiters you would like to use, i.e. ". ;,?!"
    miDelimCount = LenB(psDelims) \ 2
    If miDelimCount = 0 Then
        'If no delimiters, then match all occurences of what we're looking for.
        Erase myDelimiters
    Else
        Dim i As Long 'Counter
        
        'Fool VB into letting us access the string in-place with a
        'byte array
        GetStringMap myMap, psDelims, mtHeader, miOldDescriptor
        
        'Loop through our delimiter array, assigning the bytes that were given in psDelims
        ReDim myDelimiters(0 To miDelimCount - 1, 0 To 1)
        For i = 0 To miDelimCount - 1
            myDelimiters(i, 0) = myMap(i * 2)
            myDelimiters(i, 1) = myMap(i * 2 + 1)
        Next
        'Restore the original descriptor
        CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    End If
End Sub

Public Sub SetReplaceExclusions(ByRef psStrings() As String, _
                                ByVal pbBefore As Boolean _
           )
                                'Strings are not modified
                                'Call with an unbound array to make no exclusions
    
    Dim liLBound As Long 'Ubound of psStrings
    Dim liUBound As Long 'Lbound of psStrings
    Dim liPtr    As Long 'Pointer to newly allocated string
    Dim liLen    As Long 'Length of the current string
    Dim i        As Long 'Counter
    
    On Error Resume Next
    liUBound = UBound(psStrings) 'Get the bounds
    liLBound = LBound(psStrings)
    
    If pbBefore Then
        If Err.Number = 0& Then 'If there are bounds
            miPreExclusions = liUBound - liLBound + 1& 'Get the number of strings
            ReDim msPreExclusions(0 To miPreExclusions - 1&) 'Dim the array
            For i = liLBound To liUBound 'Loop through the arrays
                liLen = LenB(psStrings(i)) 'Get the length
                liPtr = SysAllocStringByteLen(0&, liLen) 'Allocate our string
                'Point our string to the newly allocated string
                CopyMemory ByVal VarPtr(msPreExclusions(i - liLBound)), liPtr, 4&
                'Copy the string that was passed to our new string
                CopyMemory ByVal liPtr, ByVal StrPtr(psStrings(i)), liLen
            Next
        Else
            Erase msPreExclusions
            miPreExclusions = 0
        End If
    Else
        'same as above, but using miPostExclusions and msPostExclusions
        If Err.Number = 0& Then
            miPostExclusions = liUBound - liLBound + 1&
            ReDim msPostExclusions(0 To miPostExclusions - 1&)
            For i = liLBound To liUBound
                liLen = LenB(psStrings(i))
                liPtr = SysAllocStringByteLen(0&, liLen)
                CopyMemory ByVal VarPtr(msPostExclusions(i - liLBound)), liPtr, 4&
                CopyMemory ByVal liPtr, ByVal StrPtr(psStrings(i)), liLen
            Next
        Else
            Erase msPostExclusions
            miPostExclusions = 0
        End If
    End If

End Sub

Public Function InString(ByRef psStringSearch As String, _
                         ByRef psStringFind As String, _
                Optional ByVal piStart As Long = 1&, _
                Optional ByVal piCompare As VbCompareMethod = vbBinaryCompare _
                ) As Long
                        'String arguments are not modified
    
    Dim liLen        As Long 'Length of the search string
    Dim liFindLen    As Long 'Length of the find string
    Dim i            As Long 'counter
    
    Dim lbIgnoreCase As Boolean 'whether we are case-insensitive
    
    Dim lyFind()     As Byte 'Byte array for the find string
    Dim lyFindByte   As Byte 'first byte of the find string
    Dim lyFindByte2  As Byte 'second byte of the find string
    Dim lyByte       As Byte 'temp byte for comparison
    Dim lyByte2      As Byte 'temp byte for comparison
    
    'Initialization
    liLen = LenB(psStringSearch) 'Get then lengths of the strings
    liFindLen = LenB(psStringFind)
    
    If piStart <= 0& Then Err.Raise 5 'Same behavior as intrinsic function
    piStart = (piStart - 1&) * 2&     ' adjust for 0-based unicode byte array
    If piStart > liLen Or liFindLen < 2& Then Exit Function 'Make sure that there's something to do
    
    lbIgnoreCase = piCompare = vbTextCompare
    
    GetStringMap myMap, psStringSearch, mtHeader, miOldDescriptor
    
    lyFind = psStringFind 'we may modify this one, so we need a copy
    
    lyFindByte = lyFind(0&) 'initialize the first bytes to look for
    lyFindByte2 = lyFind(1&)
    
    If lbIgnoreCase Then
        'If we're case insensitive, then it's easier to lcase the find
        'string once instead of every time we compare to it.
        For i = 0& To liFindLen - 2& Step 2&
            CharLower lyFind(i), lyFind(i + 1&)
        Next
        'same for the first two bytes
        CharLower lyFindByte, lyFindByte2
    End If
    
    'Search the string for a match
    'step by two b/c of unicode
    For InString = piStart To liLen - 2& Step 2&
        lyByte = myMap(InString)
        lyByte2 = myMap(InString + 1&)
        
        'If case insensitive then lCase$() the bytes
        If lbIgnoreCase Then CharLower lyByte, lyByte2
        
        'Could this be the start of what we're looking for?
        If lyByte = lyFindByte And lyByte2 = lyFindByte2 Then
            'It could!
            If InString + liFindLen <= liLen Then
                'Step through psStringFind
                For i = 2& To liFindLen - 2& Step 2&
                    lyByte = myMap(InString + i)
                    lyByte2 = myMap(InString + i + 1&)
                    If lbIgnoreCase Then CharLower lyByte, lyByte2
                    If Not lyFind(i) = lyByte Then Exit For
                Next
                'If we found a match then stop looking
                If i >= liFindLen Then Exit For
            End If
        End If
    Next
                
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    If InString >= liLen - 1& Then InString = 0& Else InString = InString \ 2& + 1&
End Function

Public Function InStringRev(ByRef psStringSearch As String, _
                            ByRef psStringFind As String, _
                   Optional ByVal piStart As Long, _
                   Optional ByVal piCompare As VbCompareMethod = vbBinaryCompare _
                ) As Long
                        'String arguments are not modified
                        'If piStart is omitted, starts at the end of the string
    
    Dim liLen        As Long 'Length of the search string
    Dim liFindLen    As Long 'Length of the find string
    Dim i            As Long 'counter
    
    Dim lbIgnoreCase As Boolean 'whether we are case-insensitive
    
    Dim lyFind()     As Byte 'Byte array for the find string
    Dim lyFindByte   As Byte 'first byte of the find string
    Dim lyFindByte2  As Byte 'second byte of the find string
    Dim lyByte       As Byte 'temp byte for comparison
    Dim lyByte2      As Byte 'temp byte for comparison
    
    'Initialization
    liLen = LenB(psStringSearch) 'Get then lengths of the strings
    liFindLen = LenB(psStringFind)
    
    If liFindLen < 2& Then Exit Function 'Make sure that there's something to do
    
    piStart = (piStart - 1&) * 2&     ' adjust for 0-based unicode byte array
    
    If piStart <= 0& Or piStart > liLen - liFindLen Then _
        piStart = liLen - liFindLen

    lbIgnoreCase = piCompare = vbTextCompare
    
    GetStringMap myMap, psStringSearch, mtHeader, miOldDescriptor
    
    lyFind = psStringFind 'we may modify this one, so we need a copy
    
    lyFindByte = lyFind(0&) 'initialize the first bytes to look for
    lyFindByte2 = lyFind(1&)
    
    If lbIgnoreCase Then
        'If we're case insensitive, then it's easier to lcase the find
        'string once instead of every time we compare to it.
        For i = 0& To liFindLen - 2& Step 2&
            CharLower lyFind(i), lyFind(i + 1&)
        Next
        'same for the first two bytes
        CharLower lyFindByte, lyFindByte2
    End If
    
    'Search the string for a match
    'step by two b/c of unicode
    
    For InStringRev = piStart To 0& Step -2&
        lyByte = myMap(InStringRev)
        lyByte2 = myMap(InStringRev + 1&)
        
        'If case insensitive then lCase$() the bytes
        If lbIgnoreCase Then CharLower lyByte, lyByte2
        
        'Could this be the start of what we're looking for?
        If lyByte = lyFindByte And lyByte2 = lyFindByte2 Then
            'It could!
            If InStringRev + liFindLen <= liLen Then
                'Step through psStringFind
                For i = 2& To liFindLen - 2& Step 2&
                    lyByte = myMap(InStringRev + i)
                    lyByte2 = myMap(InStringRev + i + 1&)
                    If lbIgnoreCase Then CharLower lyByte, lyByte2
                    If Not lyFind(i) = lyByte Then Exit For
                Next
                'If we found a match then stop looking
                If i >= liFindLen Then Exit For
            End If
        End If
    Next
                
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    
    If InStringRev <= 0& Then InStringRev = 0& Else InStringRev = InStringRev \ 2& + 1&
End Function

Public Function StringReverse(ByRef psString As String) As String
                        'String argument is of course not modified
    
    Dim liLen              As Long 'Double duty: Length of the string
                                   '             Backwards counter
    Dim i                  As Long 'regular counter
    
    Dim ltHeader           As tSafeArray1D 'Custom header for lyReturn
    Dim lyReturn()         As Byte 'Byte array for the return string
    Dim liReturnDescriptor As Long 'Descriptor for to put back in the array
    Dim liReturnPtr        As Long 'Pointer to the return string
        
    'Initialization
    liLen = LenB(psString) 'Get then lengths of the strings
    
    If liLen = 0& Then Exit Function 'If no string, then that was easy!
    
    liReturnPtr = SysAllocStringByteLen(0&, liLen)
    CopyMemory ByVal VarPtr(StringReverse), liReturnPtr, 4&

    GetStringMap myMap, psString, mtHeader, miOldDescriptor
    GetStringMap lyReturn, StringReverse, ltHeader, liReturnDescriptor
    
    liLen = liLen - 2& 'adjust one byte for zero-base array, and one more for unicode
    For i = 0& To liLen Step 2&
        lyReturn(liLen) = myMap(i)
        lyReturn(liLen + 1&) = myMap(i + 1&)
        liLen = liLen - 2&
    Next
               
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    CopyMemory ByVal ArrPtr(lyReturn), liReturnDescriptor, 4&
End Function

Public Function TrimStr(ByRef psString As String) As String
    'psString is not modified
    
    Dim liLen As Long 'Length of psString, then length of return string
    Dim liPtr As Long 'Pointer to the return string
    Dim liPt2 As Long 'Pointer to first non=space char in psString
    Dim liR   As Long 'counter and rightmost non-space character
    Dim liL   As Long 'counter and leftmost non-space character
    
    Const lySpace As Byte = vbKeySpace 'Byte constants to save on implicit type conversions
    Const lyZero As Byte = 0
    
    liLen = LenB(psString) 'Get the length and exit if there's nothing to do
    If liLen = 0& Then Exit Function
    
    GetStringMap myMap, psString, mtHeader, miOldDescriptor
    
    'Step through each character RTL to see if it's a space
    For liR = liLen - 2& To 0& Step -2&
        If myMap(liR) <> lySpace Or myMap(liR + 1&) <> lyZero Then Exit For
    Next
    
    If liR > 0& Then
        'We will get here unless the string was filled with spaces
        
        'Step through each character LTR to see if it's a space
        For liL = 0& To liR Step 2&
            If myMap(liL) <> lySpace Or myMap(liL + 1&) <> lyZero Then Exit For
        Next
        
        
        liPt2 = StrPtr(psString)
        'Unsigned addition to get a pointer to the left-most non-space char
        If liPt2 And DWORDMostSignificantBit Then
           liPt2 = liPt2 + liL
        ElseIf (liPt2 Or DWORDMostSignificantBit) < -liL Then
           liPt2 = liPt2 + liL
        Else
           liPt2 = (liPt2 + DWORDMostSignificantBit) + _
                           (liL + DWORDMostSignificantBit)
        End If
        
        liR = liR + 2& ' make sure we count the very last non-space char
        liLen = liR - liL 'get then length in-between the two
        
        
        If liLen > 0& Then
            'quickly allocate the string
            liPtr = SysAllocStringByteLen(0&, liLen)
            'Point the return value to the newly allocated string
            CopyMemory ByVal VarPtr(TrimStr), liPtr, 4&
            'copy the characters to the return value
            CopyMemory ByVal liPtr, ByVal liPt2, liLen
        End If
    End If
    'destroy our string map
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
End Function

Public Function RTrimStr(ByRef psString As String) As String
    'psString is not modified
    
    Dim liLen As Long 'Length of the string
    Dim liPtr As Long 'Pointer to the return string
    Dim i     As Long 'counter
    
    Const lySpace As Byte = vbKeySpace 'Byte constants to save on implicit type conversions
    Const lyZero As Byte = 0
    
    liLen = LenB(psString) 'Get the length and exit if there's nothing to do
    If liLen = 0& Then Exit Function
    
    GetStringMap myMap, psString, mtHeader, miOldDescriptor
    
    'Step through chars RTL to see if it's a space
    For i = liLen - 2& To 0& Step -2&
        If myMap(i) <> lySpace Or myMap(i + 1&) <> lyZero Then Exit For
    Next
    
    If i > 0& Then
        'we'll get here unless the string was filled with spaces
        i = i + 2&
        'allocate the string
        liPtr = SysAllocStringByteLen(0&, i)
        'Point our return value to the allocated string
        CopyMemory ByVal VarPtr(RTrimStr), liPtr, 4&
        'Copy the characters to the return value
        CopyMemory ByVal liPtr, ByVal StrPtr(psString), i
    End If
    
    'destory our string map
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    
End Function

Public Function LTrimStr(ByRef psString As String) As String
    'psString is not modified
    
    Dim liLen As Long 'Length of psString
    Dim liPtr As Long 'Pointer to the return string
    Dim liPt2 As Long 'Pointer to first non=space char in psString
    Dim liL   As Long 'counter and leftmost non-space character
    
    Const lySpace As Byte = vbKeySpace 'Byte constants to save on implicit type conversions
    Const lyZero As Byte = 0
    
    liLen = LenB(psString) 'Get the length and exit if there's nothing to do
    If liLen = 0& Then Exit Function
    
    GetStringMap myMap, psString, mtHeader, miOldDescriptor
    
    'Step through chars LTR to see if it's a space
    For liL = 0& To liLen - 3& Step 2&
        If myMap(liL) <> lySpace Or myMap(liL + 1&) <> lyZero Then Exit For
    Next
    
    'Get a pointer to the string
    liPt2 = StrPtr(psString)
    
    'Unsigned addition to get a pointer to the left-most non-space char
    If liPt2 And DWORDMostSignificantBit Then
       liPt2 = liPt2 + liL
    ElseIf (liPt2 Or DWORDMostSignificantBit) < -liL Then
       liPt2 = liPt2 + liL
    Else
       liPt2 = (liPt2 + DWORDMostSignificantBit) + _
               (liL + DWORDMostSignificantBit)
    End If
    
    'get the length of the return string
    liLen = liLen - liL
    
    If liLen > 0& Then
        'allocate the return string
        liPtr = SysAllocStringByteLen(0&, liLen)
        'point our return value to the allocated string
        CopyMemory ByVal VarPtr(LTrimStr), liPtr, 4&
        'copy the chars to the return value
        CopyMemory ByVal liPtr, ByVal liPt2, liLen
    End If
    
    'destroy our string map
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
End Function

Public Function ReplicateString(ByRef psString As String, _
                                ByVal piTimes As Long _
                ) As String
                                'psString is not modified

    Dim liLen       As Long 'Length of psstring then remaining length to be copied last
    Dim liLenReturn As Long 'length of return string
    Dim liPtr       As Long 'pointer to position in return value
    Dim liPt2       As Long 'pointer to psString
    Dim liPtrStart  As Long 'pointer to beginning of psString
    Dim liCopied    As Long 'number of bytes that have been copied to the return value
    
    
    liLen = LenB(psString) 'Get the length
    
    If piTimes <= 0& Or liLen = 0& Then Exit Function 'exit if there's nothing to do
    liLenReturn = liLen * piTimes
    'allocate the string
    liPtrStart = SysAllocStringByteLen(0&, liLenReturn)
    liPtr = liPtrStart
    'point our return value to the allocated string
    CopyMemory ByVal VarPtr(ReplicateString), liPtrStart, 4&

    liPt2 = StrPtr(psString)
    
    'copy the string once to the return value
    CopyMemory ByVal liPtr, ByVal liPt2, liLen
    liCopied = liLen
    
    Do Until liLenReturn - liCopied < liCopied
        'Unsigned addition to get the next pointer position
        If liPtr And DWORDMostSignificantBit Then
           liPtr = liPtr + liLen
        ElseIf (liPtr Or DWORDMostSignificantBit) < -liLen Then
           liPtr = liPtr + liLen
        Else
           liPtr = (liPtr + DWORDMostSignificantBit) + _
                   (liLen + DWORDMostSignificantBit)
        End If
        'copy everything that has been copied to the return value again to the
        'return value, but at the next pointer position
        CopyMemory ByVal liPtr, ByVal liPtrStart, liCopied
        liLen = liCopied
        liCopied = liCopied + liCopied
    Loop
    
    liCopied = liLenReturn - liCopied
    If liCopied > 0& Then
        'Unless the piTimes argument is an even power of 2, then we need to copy the remaining
        'length of the string
        
        If liPtr And DWORDMostSignificantBit Then
           liPtr = liPtr + liLen
        ElseIf (liPtr Or DWORDMostSignificantBit) < -liLen Then
           liPtr = liPtr + liLen
        Else
           liPtr = (liPtr + DWORDMostSignificantBit) + _
                   (liLen + DWORDMostSignificantBit)
        End If
        CopyMemory ByVal liPtr, ByVal liPtrStart, liCopied
    End If
    
End Function

Public Function SplitString(ByRef psString As String, _
                            ByRef psDelim As String, _
                            ByRef psResult() As String, _
                   Optional ByVal piLimit As Long, _
                   Optional ByVal pbAllowZeroLength As Boolean = True, _
                   Optional ByVal piCompare As VbCompareMethod = vbBinaryCompare _
                ) As Long
                            'only psResult() is modified
                            'return value is the ubound of psResult
                            'raises an error of psResult is fixed or temporarily locked
                            'raises an error of lenb(psDelim) < 2

    Dim liDelimLen      As Long 'Length of psDelim
    Dim liStringLen     As Long 'Length of psString
    Dim liLen           As Long 'Length of data copied to the return string
    Dim liPlaces()      As Long 'Array to hold the psDelim matches found
    Dim liLastPlace     As Long 'Place of the last match
    Dim i               As Long 'counter through myMap
    Dim j               As Long 'counter through lyFind
    Dim liPtr           As Long 'Pointer to a position in psString
    Dim liTempPtr       As Long 'Pointer to a newly allocated string
    
    Dim lyFind()        As Byte 'byte characters in psDelim
    Dim lyFirstFind     As Byte 'First byte in psDelim
    Dim lySecondFind    As Byte 'Second Byte in psDelim
    Dim lyByte1         As Byte 'Temp var for byte comparison
    Dim lyByte2         As Byte 'Temp var for byte comparison
    
    Dim lbIgnoreCase    As Boolean 'whether we are case insensitive
    
    
    liDelimLen = LenB(psDelim)
    liStringLen = LenB(psString)
    
    'We're using unicode, so one character is two bytes
    If liStringLen < 2& Then
        Erase psResult
        Exit Function
    ElseIf liDelimLen < 2& Then
        Erase psResult
        Err.Raise 5
    End If
    
    'Store whether we are case insensitive
    lbIgnoreCase = piCompare = vbTextCompare

    GetStringMap myMap, psString, mtHeader, miOldDescriptor
    
    lyFind = psDelim 'We might modify this array, so don't use a string map
    
    If lbIgnoreCase Then
        'If we are case insensitive, then lcase() all the bytes in lyFind
        'instead of every time we compare against them.
        For i = 0 To liDelimLen - 2& Step 2&
            CharLower lyFind(i), lyFind(i + 1)
        Next
    End If
    
    'Store the first bytes to look for
    lyFirstFind = lyFind(0&)
    lySecondFind = lyFind(1&)
    
    ReDim liPlaces(0 To liStringLen \ liDelimLen + 1) 'Maximum number of matches that we may find
    
    For i = 0& To liStringLen - 2& Step 2&
        'Get the next two bytes in line
        lyByte1 = myMap(i)
        lyByte2 = myMap(i + 1&)

        'If case insensitive then lcase() them
        If lbIgnoreCase Then CharLower lyByte1, lyByte2
        
        
        If lyByte1 = lyFirstFind And lyByte2 = lySecondFind Then
            'If we got here, then this could be the start of a match
            
            'Step through the rest of the find bytes to see if they match
            For j = 2& To liDelimLen - 2& Step 2&
                lyByte1 = myMap(i + j)
                lyByte2 = myMap(i + j + 1&)
                
                'Lcase() if case insensitive
                If lbIgnoreCase Then CharLower lyByte1, lyByte2
                
                'If doesn't match, then no reason to keep looking
                If Not (lyByte1 = lyFind(j) And lyByte2 = lyFind(j + 1&)) Then Exit For
            Next
        
            If j > liDelimLen - 2& Then 'If we found a match
                
                If SplitString > 0& Then
                    'If it's not the first match, store the relative position
                    liPlaces(SplitString) = i - liLastPlace - liDelimLen
                Else
                    'If it is the first match, store the absolute position
                    liPlaces(SplitString) = i
                End If
                
                'remember the last position
                liLastPlace = i
                'Increment the count
                SplitString = SplitString + 1&
                'Skip the bytes that we just matched
                i = i + j - 2&
                'Make sure we don't exceed out limit
                If SplitString = piLimit Then Exit For
            
            End If
        End If
    Next
    
    'Destory our string map
    CopyMemory ByVal ArrPtr(myMap), miOldDescriptor, 4&
    
    liPtr = StrPtr(psString)
    
    If SplitString > 0& Then
        'If we found some matches
        ReDim psResult(0 To SplitString)
        i = SplitString
        SplitString = 0&
        
        For i = 0 To i - 1&
            
            j = liPlaces(i)
            
            If j > 0& Then
                liTempPtr = SysAllocStringByteLen(0&, j)
                CopyMemory ByVal VarPtr(psResult(SplitString)), liTempPtr, 4&
                CopyMemory ByVal liTempPtr, ByVal liPtr, j
                SplitString = SplitString + 1&
            Else
                If pbAllowZeroLength Then SplitString = SplitString + 1&
            End If
            
            liLastPlace = j + liDelimLen
            liLen = liLen + liLastPlace
            
            If liPtr And DWORDMostSignificantBit Then
               liPtr = liPtr + liLastPlace
            ElseIf (liPtr Or DWORDMostSignificantBit) < -liLastPlace Then
               liPtr = liPtr + liLastPlace
            Else
               liPtr = (liPtr + DWORDMostSignificantBit) + _
                       (liLastPlace + DWORDMostSignificantBit)
            End If
        Next
        j = liStringLen - liLen
        If j > 0& Then
            liTempPtr = SysAllocStringByteLen(0&, j)
            CopyMemory ByVal VarPtr(psResult(SplitString)), liTempPtr, 4&
            CopyMemory ByVal liTempPtr, ByVal liPtr, j
        End If
    Else
        ReDim psResult(0& To 0&)
        liTempPtr = SysAllocStringByteLen(0&, liStringLen)
        CopyMemory ByVal VarPtr(psResult(0&)), liTempPtr, 4&
        CopyMemory ByVal liTempPtr, ByVal liPtr, liStringLen
    End If
End Function
                
Public Function JoinString(ByRef psStrings() As String, _
                  Optional ByRef psDelim As String _
                ) As String

    Dim liUBound    As Long 'Ubound of psStrings
    Dim liLBound    As Long 'Lbound of psStrings
    
    Dim liTotalLen  As Long 'counter for total len of psStrings
    Dim liDelimLen  As Long 'Length of psDelim
    Dim liLen       As Long 'Length of current string
    Dim i           As Long 'counter
    
    Dim liPtr       As Long 'pointer to next position in the return string
    Dim liDelimPtr  As Long 'pointer to psDelim
    
    On Error Resume Next
    liUBound = UBound(psStrings) 'Get the bounds and exit if it is undefined
    liLBound = LBound(psStrings)
    If Err.Number <> 0& Then Exit Function
    On Error GoTo 0
    
    liDelimLen = LenB(psDelim) 'Get the length and ptr to psDelim
    If liDelimLen > 0& Then liDelimPtr = StrPtr(psDelim)
    
    For i = liLBound To liUBound 'Total the length of the strings
        liTotalLen = liTotalLen + LenB(psStrings(i)) + liDelimLen
    Next
    
    If liTotalLen > 0& Then 'If there were any non-zero length strings or the
        
        'Allocate the return value
        liPtr = SysAllocStringByteLen(0&, liTotalLen)
        'Point the return value to the newly allocated string
        CopyMemory ByVal VarPtr(JoinString), liPtr, 4&
        
        For i = liLBound To liUBound
            liLen = LenB(psStrings(i))
            'Copy the next string to the return value
            If liLen > 0& Then
                'if necessary, copy the string
                CopyMemory ByVal liPtr, ByVal StrPtr(psStrings(i)), liLen
                'Unsigned addition to inc the pointer
                If liPtr And DWORDMostSignificantBit Then
                   liPtr = liPtr + liLen
                ElseIf (liPtr Or DWORDMostSignificantBit) < -liLen Then
                   liPtr = liPtr + liLen
                Else
                   liPtr = (liPtr + DWORDMostSignificantBit) + _
                           (liLen + DWORDMostSignificantBit)
                End If
            End If
            
            If liDelimLen > 0& Then
                'If necessary, copy the delimiter
                CopyMemory ByVal liPtr, ByVal liDelimPtr, liDelimLen
                'Inc the pointer
                If liPtr And DWORDMostSignificantBit Then
                   liPtr = liPtr + liDelimLen
                ElseIf (liPtr Or DWORDMostSignificantBit) < -liDelimLen Then
                   liPtr = liPtr + liDelimLen
                Else
                   liPtr = (liPtr + DWORDMostSignificantBit) + _
                           (liDelimLen + DWORDMostSignificantBit)
                End If
            End If
        Next
    End If
End Function
                

Public Sub GetStringMap(ByRef pyMap() As Byte, _
                         ByRef psString As String, _
                         ByRef ptSafeArray As tSafeArray1D, _
                         ByRef piOldDescriptor As Long)
    'This is one of the few helper procedures in this module because it is not called
    'from inside loops, so this will only have a negligable effect on performance.
    
    
    Dim liArrPtr As Long
    
    With ptSafeArray
        .BytesPerElement = 1& 'This is a byte array
        .Dimensions = 1  '1 Dimensional
        .Attributes = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE 'Cannot REDIM the array
        .DataPointer = StrPtr(psString) 'Point to the string data as the first element
        .Elements = LenB(psString) 'As many elements as the string has bytes
    End With
    
    liArrPtr = ArrPtr(pyMap)
    CopyMemory piOldDescriptor, ByVal liArrPtr, 4& 'Store the original descriptor
    CopyMemory ByVal liArrPtr, VarPtr(ptSafeArray), 4& 'Replace original descriptor
    
End Sub

Private Sub CharLower(ByRef pyChar1 As Byte, ByRef pyChar2 As Byte)
    'This is the only procedure called from in the loops, b/c I don't want to put it inline!
    'If this this all seems completely random, well that's unicode for you!
    
    'The alternative to this sub is to convert the bytes to a string with ChrW$(),
    'Then call LCase$(), then call AscW().  To make these three calls is a bit faster
    'with P-Code and in the IDE, but is slower with native code.
    
    Const MakeLCase As Byte = 32
    
    Const UpperA As Byte = vbKeyA
    Const UpperZ As Byte = vbKeyZ
    
    Const Zero As Byte = 0
    Const One As Byte = 1
    Const Two As Byte = 2
    Const Three As Byte = 3
    Const Four As Byte = 4
    Const Five As Byte = 5
    Const Eight As Byte = 8
    Const Thirteen As Byte = 13
    Const Fifteen As Byte = 15
    Const Sixteen As Byte = 16
    Const TwentyTwo As Byte = 22
    Const TwentyFour As Byte = 24
    Const TwentySix As Byte = 26
    Const TwentyNine As Byte = 29
    Const Thirty  As Byte = 30
    Const ThirtyOne As Byte = 31
    Const ThirtyThree As Byte = 33
    Const ThirtySix As Byte = 36
    Const ThirtySeven As Byte = 37
    Const ThirtyEight As Byte = 38
    Const ThirtyNine As Byte = 39
    Const Forty As Byte = 40
    Const FortyTwo As Byte = 42
    Const FortyThree As Byte = 43
    Const FortyFive As Byte = 45
    Const FortySix As Byte = 46
    Const FortySeven As Byte = 47
    Const FortyEight As Byte = 48
    Const FortyNine As Byte = 49
    Const Fifty As Byte = 50
    Const FiftyOne As Byte = 51
    Const FiftyThree As Byte = 53
    Const FiftyFour As Byte = 54
    Const FiftySix As Byte = 56
    Const FiftySeven As Byte = 57
    Const FiftyEight As Byte = 58
    Const SixtyThree As Byte = 63
    Const SeventyOne As Byte = 71
    Const SeventyTwo As Byte = 72
    Const SeventyFour As Byte = 74
    Const SeventySeven As Byte = 77
    Const SeventyNine As Byte = 79
    Const Eighty As Byte = 80
    Const EightyTwo As Byte = 82
    Const EightyThree As Byte = 83
    Const EightyFour As Byte = 84
    Const EightySix As Byte = 86
    Const EightySeven As Byte = 87
    Const EightyNine As Byte = 89
    Const NinetyOne As Byte = 91
    Const NinetyThree As Byte = 93
    Const NinetyFive As Byte = 95
    Const NinetySix As Byte = 96
    Const NinetyNine As Byte = 99
    Const OneHundred As Byte = 100
    Const OneOhFour As Byte = 104
    Const OneOhFive As Byte = 105
    Const OneEleven As Byte = 111
    Const OneTwelve As Byte = 112
    Const OneFourteen As Byte = 114
    Const OneSeventeen As Byte = 117
    Const OneEightteen As Byte = 118
    Const OneNineteen As Byte = 119
    Const OneTwenty As Byte = 120
    Const OneTwentyOne As Byte = 121
    Const OneTwentyThree As Byte = 123
    Const OneTwentyFive As Byte = 125
    Const OneTwentySix As Byte = 126
    Const OneTwentyEight As Byte = 128
    Const OneTwentyNine As Byte = 129
    Const OneThirty As Byte = 130
    Const OneThirtyOne As Byte = 131
    Const OneThirtyTwo As Byte = 132
    Const OneThirtyFour As Byte = 134
    Const OneThirtyFive As Byte = 135
    Const OneThirtySix As Byte = 136
    Const OneThirtySeven As Byte = 137
    Const OneThirtyEight As Byte = 138
    Const OneThirtyNine As Byte = 139
    Const OneForty As Byte = 140
    Const OneFortyTwo As Byte = 142
    Const OneFortyThree As Byte = 143
    Const OneFortyFour As Byte = 144
    Const OneFortyFive As Byte = 145
    Const OneFortySix As Byte = 146
    Const OneFortySeven As Byte = 147
    Const OneFortyEight As Byte = 148
    Const OneFifty As Byte = 150
    Const OneFiftyOne As Byte = 151
    Const OneFiftyTwo As Byte = 152
    Const OneFiftySix As Byte = 156
    Const OneFiftySeven As Byte = 157
    Const OneFiftyNine As Byte = 159
    Const OneSixty As Byte = 160
    Const OneSixtyTwo As Byte = 162
    Const OneSixtyFour As Byte = 164
    Const OneSixtySeven As Byte = 167
    Const OneSixtyNine As Byte = 169
    Const OneSeventyOne As Byte = 171
    Const OneSeventyTwo As Byte = 172
    Const OneSeventyFour As Byte = 174
    Const OneSeventyFive As Byte = 175
    Const OneSeventySeven As Byte = 177
    Const OneSeventyEight As Byte = 178
    Const OneSeventyNine As Byte = 179
    Const OneEightyOne As Byte = 181
    Const OneEightyTwo As Byte = 182
    Const OneEightyThree As Byte = 183
    Const OneEightyFour As Byte = 184
    Const OneEightyFive As Byte = 185
    Const OneEightySix As Byte = 186
    Const OneEightySeven As Byte = 187
    Const OneEightyEight As Byte = 188
    Const OneNinety As Byte = 190
    Const OneNinetyTwo As Byte = 192
    Const OneNinetyThree As Byte = 193
    Const OneNinetyFive As Byte = 195
    Const OneNinetySix As Byte = 196
    Const OneNinetySeven As Byte = 197
    Const OneNinetyNine As Byte = 199
    Const TwoHundred As Byte = 200
    Const TwoOhTwo As Byte = 202
    Const TwoOhThree As Byte = 203
    Const TwoOhFour As Byte = 204
    Const TwoOhFive As Byte = 205
    Const TwoOhSix As Byte = 206
    Const TwoOhSeven As Byte = 207
    Const TwoOhEight As Byte = 208
    Const TwoFifteen As Byte = 215
    Const TwoSixteen As Byte = 216
    Const TwoSeventeen As Byte = 217
    Const TwoEightteen As Byte = 218
    Const TwoNineteen As Byte = 219
    Const TwoTwentyOne As Byte = 221
    Const TwoTwentyTwo As Byte = 222
    Const TwoTwentySix As Byte = 226
    Const TwoTwentyNine As Byte = 229
    Const TwoThirtyTwo As Byte = 232
    Const TwoThirtyThree As Byte = 233
    Const TwoThirtyFour As Byte = 234
    Const TwoThirtyFive As Byte = 235
    Const TwoThirtySix As Byte = 236
    Const TwoThirtyEight As Byte = 238
    Const TwoForty As Byte = 240
    Const TwoFortyOne As Byte = 241
    Const TwoFortyTwo As Byte = 242
    Const TwoFortyThree As Byte = 243
    Const TwoFortySix As Byte = 246
    Const TwoFortyEight As Byte = 248
    Const TwoFortyNine As Byte = 249
    Const TwoFifty As Byte = 250
    Const TwoFiftyOne As Byte = 251
    Const TwoFiftyFive As Byte = 255
    
    
    If pyChar2 = Zero Then
        If UpperA <= pyChar1 And pyChar1 <= UpperZ Then
            pyChar1 = pyChar1 + MakeLCase
        ElseIf OneNinetyTwo <= pyChar1 And pyChar1 <= TwoTwentyTwo Then
            If pyChar1 <> TwoFifteen Then pyChar1 = pyChar1 + MakeLCase
        End If
    ElseIf pyChar2 = One Then
        If pyChar1 = NinetySix Or pyChar1 = EightyTwo Or pyChar1 = OneTwentyFive Then
            pyChar1 = pyChar1 + 1
        ElseIf pyChar1 = OneTwenty Then
            pyChar1 = TwoFiftyFive
            pyChar2 = Zero
        ElseIf (Zero <= pyChar1 And pyChar1 <= FiftyFour) Or (SeventyFour <= pyChar1 And pyChar1 <= OneEightteen) Then
            If pyChar1 Mod Two = Zero And pyChar1 <> FortyEight Then pyChar1 = pyChar1 + 1
        ElseIf FiftySeven <= pyChar1 And pyChar1 <= SeventyOne Then
            If pyChar1 Mod Two = One Then pyChar1 = pyChar1 + 1
        ElseIf TwoOhFive <= pyChar1 And pyChar1 <= TwoNineteen Then
            If pyChar1 Mod Two = One Then pyChar1 = pyChar1 + 1
        ElseIf TwoTwentyTwo <= pyChar1 And pyChar1 < TwoFiftyFive Then
            If Not (pyChar1 = TwoForty Or pyChar1 = TwoFortyTwo Or pyChar1 = TwoFortySix Or pyChar1 = TwoFortyEight) Then
                If pyChar1 <> TwoFortyOne Then
                    If pyChar1 Mod Two = Zero Then pyChar1 = pyChar1 + 1
                Else
                    pyChar1 = TwoFortyThree
                End If
            End If
        Else
            If pyChar1 = OneTwenty Then
                pyChar1 = TwoFiftyFive
                pyChar2 = Zero
            ElseIf pyChar1 = OneTwentyOne Or pyChar1 = OneTwentyThree Or pyChar1 = OneTwentyFive Or pyChar1 = OneThirty Or pyChar1 = OneThirtyTwo Or pyChar1 = OneThirtyFive Or pyChar1 = OneThirtyNine Or pyChar1 = OneFortyFive Or pyChar1 = OneFiftyTwo Or pyChar1 = OneSixty Or pyChar1 = OneSixtyTwo Or pyChar1 = OneSixtyFour Or pyChar1 = OneSixtySeven Or pyChar1 = OneSeventyTwo Or pyChar1 = OneSeventyFive Or pyChar1 = OneSeventyNine Or pyChar1 = OneEightyOne Or pyChar1 = OneEightyFour Or pyChar1 = OneEightyEight Then
                pyChar1 = pyChar1 + One
            ElseIf pyChar1 = OneTwentyNine Then
                pyChar1 = EightyThree
                pyChar2 = Two
            ElseIf pyChar1 = OneNinetySix Or pyChar1 = OneNinetyNine Or pyChar1 = TwoOhTwo Or pyChar1 = TwoFortyOne Then
                pyChar1 = pyChar1 + Two
            ElseIf pyChar1 = OneFortyTwo Then
                pyChar1 = TwoTwentyOne
            Else
                If pyChar1 = OneThirtyFour Then
                    pyChar1 = EightyFour
                    pyChar2 = Two
                ElseIf pyChar1 = OneThirtySeven Then
                    pyChar1 = EightySix
                    pyChar2 = Two
                ElseIf pyChar1 = OneThirtyEight Then
                    pyChar1 = EightySeven
                    pyChar2 = Two
                ElseIf pyChar1 = OneFortyThree Then
                    pyChar1 = EightyNine
                    pyChar2 = Two
                ElseIf pyChar1 = OneFortyFour Then
                    pyChar1 = NinetyOne
                    pyChar2 = Two
                ElseIf pyChar1 = OneFortySeven Then
                    pyChar1 = NinetySix
                    pyChar2 = Two
                ElseIf pyChar1 = OneFortyEight Then
                    pyChar1 = NinetyNine
                    pyChar2 = Two
                ElseIf pyChar1 = OneFifty Then
                    pyChar1 = OneOhFive
                    pyChar2 = Two
                ElseIf pyChar1 = OneFiftyOne Then
                    pyChar1 = OneOhFour
                    pyChar2 = Two
                ElseIf pyChar1 = OneFiftySix Then
                    pyChar1 = OneEleven
                    pyChar2 = Two
                ElseIf pyChar1 = OneFiftySeven Then
                    pyChar1 = OneFourteen
                    pyChar2 = Two
                ElseIf pyChar1 = OneFiftyNine Then
                    pyChar1 = OneSeventeen
                    pyChar2 = Two
                ElseIf pyChar1 = OneSixtyNine Then
                    pyChar1 = OneThirtyOne
                    pyChar2 = Two
                ElseIf pyChar1 = OneSeventyFour Then
                    pyChar1 = OneThirtySix
                    pyChar2 = Two
                ElseIf pyChar1 = OneSeventySeven Then
                    pyChar1 = OneThirtyEight
                    pyChar2 = Two
                ElseIf pyChar1 = OneSeventyEight Then
                    pyChar1 = OneThirtyNine
                    pyChar2 = Two
                ElseIf pyChar1 = OneEightyThree Then
                    pyChar1 = OneFortySix
                    pyChar2 = Two
                End If
            End If
        End If
    ElseIf pyChar2 = Two Then
        If Zero <= pyChar1 And pyChar1 <= TwentyTwo Then
            If pyChar1 Mod Two = Zero Then pyChar1 = pyChar1 + 1
        End If
    ElseIf pyChar2 = Three Then
        If OneFortyFive <= pyChar1 And pyChar1 <= OneSeventyOne Then
            If pyChar1 <> OneSixtyTwo Then pyChar1 = pyChar1 + MakeLCase
        ElseIf TwoTwentySix <= pyChar1 And pyChar1 <= TwoThirtyEight Then
            If pyChar1 Mod Two = Zero Then pyChar1 = pyChar1 + One
        ElseIf pyChar1 = OneThirtyFour Then
            pyChar1 = OneSeventyTwo
        ElseIf pyChar1 = OneThirtySix Or pyChar1 = OneThirtySeven Or pyChar1 = OneThirtyEight Then
            pyChar1 = pyChar1 + ThirtySeven
        ElseIf pyChar1 = OneForty Then
            pyChar1 = TwoOhFour
        ElseIf pyChar1 = OneFortyTwo Then
            pyChar1 = TwoOhFive
        ElseIf pyChar1 = OneFortyThree Then
            pyChar1 = TwoOhSix
        End If
    ElseIf pyChar2 = Four Then
        If One <= pyChar1 And pyChar1 <= Fifteen Then
            If pyChar1 <> Thirteen Then pyChar1 = pyChar1 + Eighty
        ElseIf Fifteen < pyChar1 And pyChar1 <= FortySeven Then
            pyChar1 = pyChar1 + MakeLCase
        ElseIf (NinetySix <= pyChar1 And pyChar1 <= OneTwentyEight) Or (OneFortyFour <= pyChar1 And pyChar1 <= OneNinety) Or (TwoOhEight <= pyChar1 And pyChar1 <= TwoFortyEight) Then
            If pyChar1 <> TwoThirtySix And pyChar1 <> TwoFortySix Then
                If pyChar1 Mod Two = Zero Then pyChar1 = pyChar1 + 1
            End If
        Else
            If pyChar1 = OneNinetyThree Or pyChar1 = OneNinetyFive Or pyChar1 = OneNinetyNine Or pyChar1 = TwoOhThree Then pyChar1 = pyChar1 + One
        End If
    ElseIf pyChar2 = Five Then
        If FortyNine <= pyChar1 And pyChar1 <= EightySix Then pyChar1 = pyChar1 + FortyEight
    ElseIf pyChar2 = Sixteen Then
        If OneSixty <= pyChar1 And pyChar1 <= OneNinetySeven Then pyChar1 = pyChar1 + FortyEight
    ElseIf pyChar2 = Thirty Then
        If pyChar1 <= TwoFortyEight Then
            If pyChar1 Mod 2 = Zero Then
                If Not (OneFifty <= pyChar1 And pyChar1 <= OneFiftyNine) Then pyChar1 = pyChar1 + 1
            End If
        End If
    ElseIf pyChar2 = ThirtyOne Then
        If (Eight <= pyChar1 And pyChar1 <= Fifteen) Or (TwentyFour <= pyChar1 And pyChar1 <= TwentyNine) Or (Forty <= pyChar1 And pyChar1 <= FortySeven) Or (FiftySix <= pyChar1 And pyChar1 <= SixtyThree) Or (SeventyTwo <= pyChar1 And pyChar1 <= SeventySeven) Or (OneOhFour <= pyChar1 And pyChar1 <= OneEleven) Then
            pyChar1 = pyChar1 - Eight
        ElseIf pyChar1 = EightyNine Or pyChar1 = NinetyOne Or pyChar1 = NinetyThree Or pyChar1 = NinetyFive Or pyChar1 = OneEightyFour Or pyChar1 = OneEightyFive Or pyChar1 = TwoSixteen Or pyChar1 = TwoSeventeen Or pyChar1 = TwoThirtyTwo Or pyChar1 = TwoThirtyThree Then
            pyChar1 = pyChar1 - Eight
        ElseIf pyChar1 = OneEightySix Or pyChar1 = OneEightySeven Then
            pyChar1 = pyChar1 - SeventyFour
        ElseIf TwoHundred <= pyChar1 And pyChar1 <= TwoOhThree Then
            pyChar1 = pyChar1 - EightySix
        ElseIf pyChar1 = TwoEightteen Or pyChar1 = TwoNineteen Then
            pyChar1 = pyChar1 - OneHundred
        ElseIf pyChar1 = TwoThirtyFour Or pyChar1 = TwoThirtyFive Then
            pyChar1 = pyChar1 - OneTwelve
        ElseIf pyChar1 = TwoThirtySix Then
            pyChar1 = TwoTwentyNine
        ElseIf pyChar1 = TwoFortyEight Or pyChar1 = TwoFortyNine Then
            pyChar1 = pyChar1 - OneTwentyEight
        ElseIf pyChar1 = TwoFifty Or pyChar1 = TwoFiftyOne Then
            pyChar1 = pyChar1 - OneTwentySix
        End If
    ElseIf pyChar2 = ThirtyThree Then
        If NinetySix <= pyChar1 And pyChar1 <= OneEleven Then pyChar1 = pyChar1 + Sixteen
    ElseIf pyChar2 = ThirtySix Then
        If OneEightyTwo <= pyChar1 And pyChar1 <= TwoOhSeven Then pyChar1 = pyChar1 + TwentySix
    ElseIf pyChar2 = TwoFiftyFive Then
        If ThirtyThree <= pyChar1 And pyChar1 <= FiftyEight Then pyChar1 = pyChar1 + MakeLCase
    End If
End Sub

'These test procedures are the only places that cPerformanceTimer is used.
'You do not need to include it if you use the rest of this module in another project.
'Sub Main()
'    'Sorry, I'm not going to comment this one!
'    SetReplaceDelimiters " "
'
'    Dim lsTest As String
'    Dim lsReplaced As String
'    Dim lsMsg As String
'    Dim liPtr As Long
'
'    lsTest = "This is a stringThis THIS is a string. This"
'    lsMsg = "Compile w/ the optimizations described at the top of this module for best results." & vbNewLine & "Remember that the delimiters are set to a space, so whole words only are replaced." & vbNewLine & "Original: " & lsTest & vbNewLine & "VB's Replace: " & Replace(lsTest, "This", "Replaced", , , vbTextCompare) & vbNewLine & "This Replace: " & ReplaceString(lsTest, "This", "Replaced", , , vbTextCompare) & vbNewLine
    'Debug.Print lsMsg
    'Exit Sub
    
'    CopyMemory ByVal VarPtr(lsReplaced), SysAllocStringByteLen(0&, 10000& * LenB(lsTest)), 4&
'    liPtr = StrPtr(lsReplaced)
    
'    Dim i As Long
'    For i = 1& To 10000&
'        CopyMemory ByVal liPtr, ByVal StrPtr(lsTest), LenB(lsTest)
'
'        If liPtr And DWORDMostSignificantBit Then
'           liPtr = liPtr + LenB(lsTest)
'        ElseIf (liPtr Or DWORDMostSignificantBit) < -LenB(lsTest) Then
'           liPtr = liPtr + LenB(lsTest)
'        Else
'           liPtr = (liPtr + DWORDMostSignificantBit) + (LenB(lsTest) + DWORDMostSignificantBit)
'        End If
'
'    Next
'
'    'Swap the two strings
'    liPtr = StrPtr(lsReplaced)
'    CopyMemory ByVal VarPtr(lsReplaced), StrPtr(lsTest), 4&
'    CopyMemory ByVal VarPtr(lsTest), liPtr, 4&
'
'    'lsTest = loAppend.ToString

'    Dim loTimer As cPerformanceTimer
'    Set loTimer = New cPerformanceTimer
'
'    loTimer.TimerStart
'    For i = 1 To 10
'        lsReplaced = Replace$(lsTest, "This", "Replaced")
'    Next
'    loTimer.TimerStop
'    lsMsg = lsMsg & "VB's Replace for original appended 10000x replaced 10x in a row: " & loTimer.TimerElapsed & vbNewLine
'
'    loTimer.TimerStart
'    For i = 1 To 10
'        lsReplaced = ReplaceString(lsTest, "This", "Replaced")
'    Next
'    loTimer.TimerStop
'    lsMsg = lsMsg & "This Replace for original appended 10000x replaced 10x in a row: " & loTimer.TimerElapsed
'    Set loTimer = Nothing
'    lsTest = vbNullString
'    lsReplaced = vbNullString
'    MsgBox lsMsg
'
'End Sub

'Public Sub TestInString()
'    'my results:  in IDE or compiled to P-Code: 4 times slower
'    '             Native Code:  3 times faster
'
'    Dim ls As String
'    Dim i As Long
'    Dim l As Long
'    Dim loTimer As cPerformanceTimer
'    Set loTimer = New cPerformanceTimer
'    Dim lsMsg As String
'
'    ls = String$(50000, "x") & "FIND" & String$(100, "l")
'
'    loTimer.TimerStart
'    For i = 1 To 100
'        l = InStr(1, ls, "find", vbTextCompare)
'    Next
'    loTimer.TimerStop
'    lsMsg = loTimer.TimerElapsed
'
'
'    loTimer.TimerStart
'    For i = 1 To 100
'        l = InString(1, ls, "find", vbTextCompare)
'    Next
'    loTimer.TimerStop
'    lsMsg = lsMsg & vbNewLine & loTimer.TimerElapsed
'    MsgBox lsMsg
'
'End Sub
'
'Public Function IsGoodReplace(Optional fLigaturesToo As Boolean) As Boolean
''test function from http://www.xbeat.net/vbspeed/
'' verify correct Replace returns, 20020929
'' returns True if all tests are passed
'  Dim fFailed As Boolean
'
'  ' replace "replacestring" with the name of your function to test
'  '
'  ' ! note the differences to VB6's native Replace function!
'  'With New CStrProcessorV2
'  'With New CStrProcessor
'
'  If ReplaceString("aaa", "a", "baa") <> "baabaabaa" Then Stop: fFailed = True
'  If ReplaceString("", "b", "XX") <> "" Then Stop: fFailed = True
'  If ReplaceString("abc", "", "XX") <> "abc" Then Stop: fFailed = True
'  If ReplaceString("abc", "b", "") <> "ac" Then Stop: fFailed = True
'  If ReplaceString("abc", "c", "") <> "ab" Then Stop: fFailed = True
'  If ReplaceString("abc", "b", "X") <> "aXc" Then Stop: fFailed = True
'  If ReplaceString("abc", "b", "XX") <> "aXXc" Then Stop: fFailed = True
'  If ReplaceString("blah", "blah", "ha") <> "ha" Then Stop: fFailed = True
'
'  ' text compare
'  If ReplaceString("abc", "B", "X", , , vbTextCompare) <> "aXc" Then Stop: fFailed = True
'  If ReplaceString("aBc", "b", "XX", , , vbTextCompare) <> "aXXc" Then Stop: fFailed = True
'  If ReplaceString("abc", "B", "XX", , , vbTextCompare) <> "aXXc" Then Stop: fFailed = True
'  If ReplaceString("aüc", "Ü", "XX", , , vbTextCompare) <> "aXXc" Then Stop: fFailed = True
'  If ReplaceString("aþÞc", "Þþ", "XX", , , vbTextCompare) <> "aXXc" Then Stop: fFailed = True
'
'  ' the 4 stooges: š/Š, œ/Œ, ž/Ž, ÿ/Ÿ (154/138, 156/140, 158/142, 255/159)
'  If ReplaceString("Hašiš", "Š", "sch", , , vbTextCompare) <> "Haschisch" Then Stop: fFailed = True
'  ' ligatures  textcompare (VBspeed entries do NOT have to pass this test)
'  If fLigaturesToo Then
'    ' ligatures, a digraphemic fun house: ss/ß, ae/æ, oe/œ, th/þ
'    If ReplaceString("Straße", "ss", "f", , , vbTextCompare) <> "Strafe" Then Stop: fFailed = True
'  End If
'
'  ' non-textual chars in vbTextCompare mode:
'  ' using &HDF (223) to convert to uppercase is problematic:
'  ' it works with textual chars like "a"
'    ' chr$(97) -> "a"           97=01100001
'    '                      AND 223=11011111
'    ' chr$(97 and 223) -> "A"   65=01000001
'  ' but it fails with these for example:
'    ' "[" = 91,   91 And 223 = 91
'    ' "{" = 123, 123 And 223 = 91
'  If ReplaceString("[[{{", "{", "x", , , vbTextCompare) <> "[[xx" Then Stop: fFailed = True
'  If ReplaceString("[[{{", "[", "x", , , vbTextCompare) <> "xx{{" Then Stop: fFailed = True
'
'
'  ' unicode
'  If ReplaceString("€", "€", "x") <> "x" Then Stop: fFailed = True
'  If ReplaceString("x", "x", "€") <> "€" Then Stop: fFailed = True
'
'  ' hard core unicode + text compare
'  ' high unicode textwise, test too hard for now => TestReplace
'  If ReplaceString(ChrW$(400) & " Tag!", ChrW$(603), "Guten", , , vbTextCompare) <> "Guten Tag!" Then Stop: fFailed = True
'
'
'  ' Start param
'  If ReplaceString("abc", "a", "XX", 1) <> "XXbc" Then Stop: fFailed = True
'  ' ! VB6 Replace returns "bc":
'  If ReplaceString("abc", "a", "XX", 2) <> "abc" Then Stop: fFailed = True
'  ' ! VB6 Replace returns "c":
'  If ReplaceString("abc", "a", "XX", 3) <> "abc" Then Stop: fFailed = True
'  ' ! VB6 Replace returns "":
'  If ReplaceString("abc", "a", "XX", 4) <> "abc" Then Stop: fFailed = True
'  ' ! VB6 Replace returns "bcXabcXabcabc":
'  If ReplaceString("abcabcabcabc", "a", "Xa", 2, 2) <> "abcXabcXabcabc" Then Stop: fFailed = True
'
'  ' very large inputs
'  If ReplaceString("x" & Space$(10000) & "x", "x", Space$(10000)) <> Space$(30000) Then Stop: fFailed = True
'
'  'End With
'
'  ' well done
'  IsGoodReplace = Not fFailed
'
'End Function
'
'Public Function IsGoodLCase() As Boolean
'    Dim i As Long
'    Dim lyBytes() As Byte
'    Dim lyBytes2() As Byte
'    Dim lyByte(0 To 3) As Byte
'
'    For i = -32768 To 65535
'        lyBytes = ChrW$(i)
'        lyBytes2 = LCase$(ChrW$(i))
'
'        lyByte(0) = lyBytes(0) 'First regular
'        lyByte(1) = lyBytes(1) 'Second regular
'        lyByte(2) = lyBytes2(0) 'First Lcase
'        lyByte(3) = lyBytes2(1) 'Second Lcase
'
'        CharLower lyByte(0), lyByte(1)
'
'        IsGoodLCase = lyByte(0) = lyByte(2) And lyByte(1) = lyByte(3)
'        Debug.Assert IsGoodLCase
'        If Not IsGoodLCase Then Exit Function
'    Next
'
'End Function


