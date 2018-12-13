Attribute VB_Name = "basSafeBound"
Public Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

    Dim lAddress&, cElements&, lLbound&, cDims%

    If Dimension < 1 Then
        SafeUBound = -1
        Exit Function
    End If

    CopyMemory lAddress, ByVal lpArray, 4

    If lAddress = 0 Then
        ' The array isn't initilized
        SafeUBound = -1
        Exit Function
    End If

    ' Calculate the dimensions
    CopyMemory cDims, ByVal lAddress, 2
    Dimension = cDims - Dimension + 1
    ' Obtain the needed data
    CopyMemory cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4
    CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
    SafeUBound = cElements + lLbound - 1
End Function


Public Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

    Dim lAddress&, cElements&, lLbound&, cDims%

    If Dimension < 1 Then
        SafeLBound = -1
        Exit Function
    End If

    CopyMemory lAddress, ByVal lpArray, 4

    If lAddress = 0 Then
        ' The array isn't initilized
        SafeLBound = -1
        Exit Function
    End If

    ' Calculate the dimensions
    CopyMemory cDims, ByVal lAddress, 2
    Dimension = cDims - Dimension + 1
    ' Obtain the needed data
    CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
    SafeLBound = lLbound
End Function


Public Function ArrayDims(ByVal lpArray As Long) As Integer

    Dim lAddress As Long
    CopyMemory lAddress, ByVal lpArray, 4

    If lAddress = 0 Then
        ' The array isn't initilized
        ArrayDims = -1
        Exit Function
    End If

    CopyMemory ArrayDims, ByVal lAddress, 2
End Function
