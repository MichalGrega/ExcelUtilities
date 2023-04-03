Function CtoLng(ByVal Value As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function accepts a Value and tries to convert it to Long type
' regardless of the decimal separator used.
' Solves the problem of two possible decimal separators "." and ",".
' First "." is used. If error is raised, "." is replaced with "," and
' the conversion is attempted once more. If the conversion fails, an error is raised.

    On Error Resume Next
    CtoLng = CLng(Value)
    If Err.number <> 0 Then
        Err.Clear
        CtoLng = CLng(Replace(Value, ".", ","))
        If Err.number <> 0 Then Err.Raise vbObjectError + 100, , "fCtoLng: wrong value with error " & Err.number & ": " & Err.Description
    End If
    
End Function

Function convert(ByVal InputArray As Variant, ByVal number As Long, Optional ByVal reversed As Boolean = False) As Variant
    Dim bounds() As Variant
    numDims = NumberOfArrayDimensions(InputArray)
    
    If numDims = 0 Or numDims > 2 Then GoTo WrongDimensions
    
    If numDims = 2 Then
        On Error Resume Next
            bounds = InputArray
            If Err.number = 13 Then
                ReDim bounds(LBound(InputArray, 1) To UBound(InputArray), 0 To 1) As Variant
                bottom = LBound(InputArray, 2)
                For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                    bounds(i, 0) = CtoLng(InputArray(i, bottom))
                    bounds(i, 1) = CtoLng(InputArray(i, bottom + 1))
                Next i
            End If
        On Error GoTo 0
    Else
        ReDim bounds(LBound(InputArray) To UBound(InputArray), 0 To 1) As Variant
        
        For i = LBound(InputArray) To UBound(InputArray)
            If Not IsArray(InputArray(i)) Then
                bounds(i, 0) = 0
                bounds(i, 1) = CtoLng(InputArray(i))
            Else
                Ndx = NumberOfArrayDimensions(InputArray(i))
                If Ndx <> 1 Then GoTo WrongDimensions
                bottom = LBound(InputArray(i))
                bounds(i, 0) = CtoLng(InputArray(i)(bottom))
                bounds(i, 1) = CtoLng(InputArray(i)(bottom + 1))
            End If
        Next i
    End If
       
    ReDim result(LBound(InputArray, 1) To UBound(InputArray, 1)) As Variant
    
    solved = False
    If reversed Then
        stt = Array(LBound(result), UBound(result), 1)
    Else
        stt = Array(UBound(result), LBound(result), -1)
    End If
    
    For i = stt(0) To stt(1) Step stt(2)
        If Not solved Then
            bottom = LBound(bounds, 2)
            result(i) = bounds(i, bottom) + (number Mod (bounds(i, bottom + 1) - bounds(i, bottom) + 1))
            number = WorksheetFunction.Quotient(number, bounds(i, bottom + 1) - bounds(i, bottom) + 1)
        Else
            result(i) = 0
        End If
        If number = 0 Then
            solved = True
        End If
    Next i
    If Not solved Then Err.Raise vbObjectError + 101, , "fConvert2: unsolved. Number is greater than the settings allow."
    
    convert = result
    Exit Function
WrongDimensions:
    Err.Raise vbObjectError - 1, , "fConvert2: wrong dimensions of the input array."
    Exit Function
End Function
