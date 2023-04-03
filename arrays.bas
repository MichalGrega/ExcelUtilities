Function fCLng(ByVal Value As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' fCLng
' Function accepts a Value and tries to convert it to the Long type
' regardless of the decimal separator used.
' Solves the problem of two possible decimal separators "." and ",".
' First "." is used. If error is raised, "." is replaced with "," and
' the conversion is attempted once more. If the conversion fails, an error is raised.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error Resume Next
    fCLng = CLng(Value)
    If Err.number = 13 Then
        Err.Clear
        fCLng = CLng(Replace(Value, ".", ","))
        If Err.number <> 0 Then Err.Raise vbObjectError + 100, , "fCLng: wrong value with error " & Err.number & ": " & Err.Description
    End If
    
End Function
'_________________________________________________________________________________________________________________________________________________________________

Function fConv(ByVal InputArray As Variant, ByVal number As Long, _
                Optional ByVal reversed As Boolean = False)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function accepts an InputArray that represents settings for the conversion and converts
' the number input accordingly. Basicaly it is the same type of conversion like DEC2BIN or
' DEC2HEX but with a limited number of digits and variable values. It returns a one dimensional
' array with the individual digits for the converted number. The resulting array has same bounds
' as the first dimension of the input array.
'
' This function can be used as a special counter. Let say we have a counter diveded into three
' sections that can have possible values: 2 to 5, 3 to 6, 15 to 17 and we want to count points
' or laps or something similar. Every roundthe rightmost value is increased and after it reaches
' its maximu value, next section (i.e. the middle one) is increased and so on like:
' round 0: 2,3,15
' round 1: 2,3,16
' round 2: 2,3,17
' round 3: 2,4,15
' ...
' round 10: 2,6,16
' ...
' round 48: 5,6,17
'
' * InputArray - 1. can be one dimensional array declared as Dim arr(n), Dim arr(m,n) or Array(x, y, ...).
'                in this case, the lower bound for the section will be set to 0 and the upper
'                bound to the value of the InputArray.
'                2. Or it can be one dimensional array with another one dimensional array as its items
'                like Array(Array(2,5), Array(3,6), Array(15,17)) or even a combination of 1Dim arrays
'                and primitive values.
'                3. A two dimensional array defining lower and upper bounds of the counter on each
'                position can be also used.
' * Number - any number of Long type. If the number is greater than the largest possible number allowed
'            by the input array, an error is raised.
' * Optional reversed - if true, the counting goes from left to right.
'                       reversed = false
'                           0: 0,0,0
'                           1: 0,0,1
'                           2: 0,0,2
'                       reversed = True
'                           0: 0,0,0
'                           1: 1,0,0
'                           2: 2,0,0
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim bounds() As Variant
    
    '''''''''''''''''''''''''''''''''''''''''''
    ' Check the dimensions of the Input array.
    ' If the number of dimensions is not 1 or 2
    ' raise an error
    '''''''''''''''''''''''''''''''''''''''''''
    numDims = NumberOfArrayDimensions(InputArray)
    If numDims = 0 Or numDims > 2 Then GoTo WrongDimensions
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Prepare an array for the calculation.
    ' Process different types of input arrays (1, 2 or 3 in the description)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If numDims = 2 Then
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' 2 dimensional array. If types differ,
        ' new array is created and values copied to it
        ''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
            bounds = InputArray
            If Err.number = 13 Then
                ReDim bounds(LBound(InputArray, 1) To UBound(InputArray), 0 To 1) As Variant
                bottom = LBound(InputArray, 2)
                For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                    bounds(i, 0) = fCLng(InputArray(i, bottom))
                    bounds(i, 1) = fCLng(InputArray(i, bottom + 1))
                Next i
            End If
        On Error GoTo 0
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' 1 dimensional array with either primitives or other
        ' one dimensional arrays on individual positions
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ReDim bounds(LBound(InputArray) To UBound(InputArray), 0 To 1) As Variant
        
        For i = LBound(InputArray) To UBound(InputArray)
            If Not IsArray(InputArray(i)) Then
                bounds(i, 0) = 0
                bounds(i, 1) = fCLng(InputArray(i))
            Else
                Ndx = NumberOfArrayDimensions(InputArray(i))
                If Ndx <> 1 Then GoTo WrongDimensions
                bottom = LBound(InputArray(i))
                bounds(i, 0) = fCLng(InputArray(i)(bottom))
                bounds(i, 1) = fCLng(InputArray(i)(bottom + 1))
            End If
        Next i
    End If

    ReDim result(LBound(bounds, 1) To UBound(bounds, 1)) As Variant
    
    ''''''''''''''''''''''''''''''''''''''
    ' Setting the mode of the calculation:
    ' normal or reversed
    ''''''''''''''''''''''''''''''''''''''
    If reversed Then
        stt = Array(LBound(result), UBound(result), 1)
    Else
        stt = Array(UBound(result), LBound(result), -1)
    End If
    
    ''''''''''''''''''''''''''''''''''''
    ' Calculation of individual sections
    ''''''''''''''''''''''''''''''''''''
    solved = False
    For i = stt(0) To stt(1) Step stt(2)
        bottom = LBound(bounds, 2)
        If Not solved Then
            result(i) = bounds(i, bottom) + (number Mod (bounds(i, bottom + 1) - bounds(i, bottom) + 1))
            number = WorksheetFunction.Quotient(number, bounds(i, bottom + 1) - bounds(i, bottom) + 1)
        Else
            result(i) = bounds(i, bottom) + 0
        End If
        If number = 0 Then
            solved = True
        End If
    Next i
    If Not solved Then Err.Raise vbObjectError + 101, , "fConvert2: unsolved. Number is greater than the settings allow."
    
    fConv = result
    Exit Function
WrongDimensions:
    Err.Raise vbObjectError - 1, , "fConvert2: wrong dimensions of the input array."
    Exit Function
End Function
'_________________________________________________________________________________________________________________________________________________________________