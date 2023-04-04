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

Sub fConv_Examples()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Example usage of fConv function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim arr(1 To 3, 1) As String
    arr(1, 0) = 0
    arr(1, 1) = 2
    arr(2, 0) = 1
    arr(2, 1) = 3
    arr(3, 0) = 2
    arr(3, 1) = 4
    
    Debug.Print Join(fConv(arr, 7), ", ")
    '>>> 0, 3, 3
    
    Debug.Print Join(fConv(Array(15, 15, 15, 15, 15, 15), 1553460), ", ")
    '>>> 1, 7, 11, 4, 3, 4
    
    
    Dim arr2(8 To 19) As Variant
    arr2(8) = 9
    arr2(9) = 16
    arr2(10) = "a"
    Debug.Print Join(fConv(Array(8, Array("4", "15.89"), arr2), 60, True), ", ")
    '>>> 6, 10, 9
    
End Sub
'_________________________________________________________________________________________________________________________________________________________________


Function DUMP(ByVal Variable As Variant, Optional ByVal Deepness As Integer = 0, _
                Optional ByVal LineBreaks = False, Optional ByVal ShowArrayIndexes As Boolean = False)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DUMP
' Function serializes (dumps) a variable into a string output. It started with a need to see
' items of a Scripting.Dictionary since in the Locals window only keys are shown, but because
' an item in a dictionary can be anything, also this function needs to dump anything.
'
' Oh, did I mention it is recursive? :)
'
' * Variable as Variant - Variable to be dumped. It can be almost anything. It is then dumped
'                         if it is a string, it is enclosed in quotes ("like this"), if array
'                         it is in brackets (), if a collection in a square brackets [],
'                         if dictionary in curly brackets {}, other primitives are converted
'                         into string with CStr function, if it is something other, TypeName is
'                         shown.
' * Deepness as integer - a level of recursion. It will indent the items correctly. This doesn't
'                         need to be used. It is used automatically in case of recursion.
' * LineBreaks - if true, line breaks are added after individual items.
' * ShowArrayIndexes - if true, array item indexes will be shown to better orientate easier.
'
' Requirements:
'       - NumberOfArrayDimensions function
'       - fConv function
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set default values of output and constants
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Const primitives As String = "String, Long, Integer, Boolean, Single, Double, Byte, Currency, Decimal, Date, Error"
    Const itemSeparator As String = ", "
    Const indentation As Variant = "  "
    output = ""
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 1. Primitive value
    ' If variable holds a primitive value from const - primitives
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If primitives Like "*" & TypeName(Variable) & "*" Then
        If TypeName(Variable) = "String" Then
            '''''''''''''''''''''''''''''''
            ' Wrap string with quotes
            '''''''''''''''''''''''''''''''
            output = """" & Variable & """"
        Else
            '''''''''''''''''''''''''''''''
            ' If not string then convert
            ' to string
            '''''''''''''''''''''''''''''''
            output = CStr(Variable)
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Decimal separator is a pain in the ass so replace comma
        ' with point
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If IsNumeric(output) Then output = Replace(output, ",", ".")
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 2. Arrays
    ' Check if Variable is an array
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf IsArray(Variable) Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Count array dimensions
        ''''''''''''''''''''''''''''''''''''''''''''''
        dimensions = NumberOfArrayDimensions(Variable)
        
        
        If dimensions = 0 Then
            ''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions means an empty array
            ''''''''''''''''''''''''''''''''''''''''''
            output = output & "Array(Empty)"
        Else
            ''''''''''''''''''''''''''''''''''''''''''
            ' One- or Multidimensional array
            '
            ' Values in an array with unknown number
            ' of dimesions will be copied to a helper
            ' collection with an index as a key. Index
            ' is calculated with an fConv function to
            ' which bounds of dimensions is passed.
            ''''''''''''''''''''''''''''''''''''''''''
            
            ''''''''''''''''''''''''''''''''''''''''''
            ' Declaration of helper collection to
            ' correctly sort values in an array with
            ' unknown dimensions.
            ''''''''''''''''''''''''''''''''''''''''''
            Dim keyVals As New Collection
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Create an array with number of dimensions of
            ' the input array as its first dimension and lower
            ' and upper bound of the corresponding dimension
            ' in its second dimension.
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            ReDim dimsDef(1 To dimensions, 1 To 2) As Integer
            For i = 1 To dimensions
                dimsDef(i, 1) = LBound(Variable, i)
                dimsDef(i, 2) = UBound(Variable, i)
            Next i
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Transfer each value in the array to a helper
            ' collection with item's index as its key.
            ' The index is calculated with fConv function in a form
            ' of dimension indexes joined with a dash like e.g.:
            ' "2-4-6-1"
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            itemIndex = 0
            For Each v In Variable
                keyVals.Add v, Join(fConv(InputArray:=dimsDef, number:=itemIndex, reversed:=True), "-")
                itemIndex = itemIndex + 1
            Next v
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Traverse through the helper collection by calling the
            ' keys in a correct order and build a string output.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For i = 0 To itemIndex - 1
                coordinates = fConv(dimsDef, i) ' Calculate coordinates for the current item.
                
                ''''''''''''''''''''''''''''''''''''''''''
                ' Determine if a bracket has to be opened.
                ' Check each dimension.
                ''''''''''''''''''''''''''''''''''''''''''
                For dimensionIndex = LBound(coordinates) To UBound(coordinates)
                    
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' Check if all following dimensions  are at lower bounds.
                    ' A bracket for the current dimension has to be opened if
                    ' all consecutive coordinates are at their lower bounds.
                    ' If at least one of them is not, then bracket for that
                    ' dimension is not opened.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    openBracket = True
                    For consecutiveDimensionsIndex = dimensionIndex To UBound(coordinates)
                        If coordinates(consecutiveDimensionsIndex) <> dimsDef(consecutiveDimensionsIndex, 1) Then
                            openBracket = False
                            Exit For
                        End If
                    Next consecutiveDimensionsIndex
                    
                    '''''''''''''''''''''''''''''''''''''''
                    ' Open a bracket
                    '''''''''''''''''''''''''''''''''''''''
                    If openBracket Then
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''
                        ' Create string for indexes for the bracket in
                        ' a form n.m.d. ... => for previous dimensions.
                        ' If not enabled, set empty string.
                        '''''''''''''''''''''''''''''''''''''''''''''''
                        indexes = ""
                        If ShowArrayIndexes And dimensionIndex > LBound(coordinates) Then
                            For previousDimensions = LBound(coordinates) To dimensionIndex - 1
                                indexes = indexes & coordinates(previousDimensions) & "."
                            Next previousDimensions
                            If Right(indexes, 1) = "." Then indexes = Left(indexes, Len(indexes) - 1) & "=>"
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' Insert line break after first bracket if enabled
                        ' and increase indentation level i.e. Deepness.
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If LineBreaks Then
                            If output <> "" Then output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
                            If dimensionIndex < UBound(coordinates) Then
                                Deepness = Deepness + 1
                            End If
                        End If
                        
                        '''''''''''''''''''''''''''''''''
                        ' Now add indexes and the bracket
                        '''''''''''''''''''''''''''''''''
                        output = output & indexes & "("
                    End If
                Next dimensionIndex
                
                '''''''''''''''''''''''''''''''''''''''''''''''
                ' Create string for indexes for the item in
                ' a form n.m.d. ... => for previous dimensions.
                ' If not enabled, set empty string.
                '''''''''''''''''''''''''''''''''''''''''''''''
                indexes = ""
                If ShowArrayIndexes Then
                    For allDimensionsIndex = LBound(coordinates) To UBound(coordinates)
                        indexes = indexes & coordinates(allDimensionsIndex) & "."
                    Next allDimensionsIndex
                    If Right(indexes, 1) = "." Then indexes = Left(indexes, Len(indexes) - 1) & "=>"
                End If
                
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Recursively call the DUMP function to convert
                ' the value called by the key to a string and add
                ' the returned value to the output string.
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                output = output & _
                         indexes & _
                         DUMP(keyVals(Join(coordinates, "-")), Deepness, LineBreaks, ShowArrayIndexes) & _
                         itemSeparator
                
                
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Determine if a bracket has to be closed.
                ' Check each dimension from the last to the first one.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                For leadingDimensionIndex = UBound(coordinates) To LBound(coordinates) Step -1
                    
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' Check if all following dimensions  are at upper bounds.
                    ' A bracket for the current dimension has to be closed if
                    ' all consecutive coordinates are at their upper bounds.
                    ' If at least one of them is not, then bracket for that
                    ' dimension is not closed.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    cls = True
                    For consecutiveDimensionIndex = leadingDimensionIndex To UBound(coordinates)
                        If coordinates(consecutiveDimensionIndex) <> dimsDef(consecutiveDimensionIndex, 2) Then
                            cls = False
                            Exit For
                        End If
                    Next consecutiveDimensionIndex
                    
                    ''''''''''''''''''''''''''''''''''''
                    ' Cloase a bracket if necessary
                    ''''''''''''''''''''''''''''''''''''
                    If cls Then
                    
                        ''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' Remove item separator before closing the bracket
                        ''''''''''''''''''''''''''''''''''''''''''''''''''
                        If Right(output, 2) = itemSeparator Then output = Left(output, Len(output) - 2)
                        
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' Reduce indentation before closing a bracket exept
                        ' the last dimension and insert line break if enabled.
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If leadingDimensionIndex < UBound(coordinates) Then
                            Deepness = Deepness - 1
                            If LineBreaks Then output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
                        End If
                        
                        '''''''''''''''''''
                        ' Close the bracket
                        '''''''''''''''''''
                        output = output & ")" & Application.Trim(itemSeparator)
                    End If
                Next leadingDimensionIndex
            Next i
            
            ''''''''''''''''''''''''''''''''''''''''''
            ' Remove item separator after last bracket
            ''''''''''''''''''''''''''''''''''''''''''
            If Right(output, Len(Application.Trim(itemSeparator))) = Application.Trim(itemSeparator) Then output = Left(output, Len(output) - Len(Application.Trim(itemSeparator)))
            
            '''''''''''''''''''''''''''''''''''''''''''
            ' If there are no dimensions, return info
            ' about the array being empty.
            '''''''''''''''''''''''''''''''''''''''''''
            If output = "" Then output = "Array(Empty)"
            
        End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 3. Dictionaries
    ' Check if Variable is a dictionary instance
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf "Dictionary" Like "*" & TypeName(Variable) & "*" Then
        '''''''''''''''''''''''''''
        ' Open a dictionary bracket
        '''''''''''''''''''''''''''
        output = "{"
        
        ''''''''''''''''''''''''''''''''''
        ' If line breaks are enabled
        ' add a line break and indentation
        ''''''''''''''''''''''''''''''''''
        If LineBreaks Then
            Deepness = Deepness + 1
            output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
        End If
        
        ''''''''''''''''''''''''''
        ' Traverse dictionary keys
        ''''''''''''''''''''''''''
        For Each currentKey In Variable.Keys
            
            ''''''''''''''''''''''''''''''
            ' Retrieve value corresponding
            ' to the current key.
            ''''''''''''''''''''''''''''''
            If IsObject(Variable(currentKey)) Then
                Set currentValue = Variable(currentKey)
            Else
                currentValue = Variable(currentKey)
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''
            ' Format key
            ' if the key is a string, wrap it in quotes.
            ''''''''''''''''''''''''''''''''''''''''''''
            showKy = currentKey
            If TypeName(currentKey) = "String" Then showKy = """" & currentKey & """"
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Recursively call the DUMP function for the value with
            ' an increased indentation level
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            currentValue = DUMP(currentValue, Deepness + 1, LineBreaks, ShowArrayIndexes)
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Add key: value pair string to the output string and add an
            ' item separator.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            output = output & showKy & ": " & currentValue & itemSeparator
            
            '''''''''''''''''''''''''''''
            ' Add a line break if enabled
            '''''''''''''''''''''''''''''
            If LineBreaks Then output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
        Next currentKey
        
        ''''''''''''''''''''''''''''''''''
        ' Remove a line break from the end
        ''''''''''''''''''''''''''''''''''
        If LineBreaks Then output = Left(output, Len(output) - 1)
        
        ''''''''''''''''''''''''''''''''''''''''
        ' Remove the item separator from the end
        ''''''''''''''''''''''''''''''''''''''''
        If Right(output, 2) = itemSeparator Then output = Left(output, Len(output) - 2)
        
        ''''''''''''''''''''''''''''''''
        ' Close the dictionary's bracket
        ''''''''''''''''''''''''''''''''
        output = output & "}"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 4. Collections
    ' Check if the Variable is a collection
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf "Collection" Like "*" & TypeName(Variable) & "*" Then
        '''''''''''''''''''''''''''
        ' Open a collection bracket
        '''''''''''''''''''''''''''
        output = "["
        
        ''''''''''''''''''''''''''''''''''
        ' If line breaks are enabled
        ' add a line break and indentation
        ''''''''''''''''''''''''''''''''''
        If LineBreaks Then
            Deepness = Deepness + 1
            output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Recursively call DUMP on each item in the
        ' collection and add a line break if enabled.
        '''''''''''''''''''''''''''''''''''''''''''''
        For Each itm In Variable
            output = output & DUMP(itm, Deepness, LineBreaks, ShowArrayIndexes) & itemSeparator
            If LineBreaks Then output = output & vbNewLine & WorksheetFunction.Rept(indentation, Deepness)
        Next itm
        
        '''''''''''''''''''''''''''''''
        ' Close the collection bracket.
        '''''''''''''''''''''''''''''''
        output = Left(output, Len(output) - 2) & "]"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 5. Other values
    ' If not any of the previous types, return type name.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        output = TypeName(Variable)
    End If
    DUMP = output
End Function
'_________________________________________________________________________________________________________________________________________________________________