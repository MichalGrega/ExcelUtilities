Attribute VB_Name = "DUMP"

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
    Const MaxItems As Long = 200 'if number items in collection or dictionary exceeds this limit, show only number
    Dim objects As New Collection
    output = ""
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Collection of properties of objects. Keys are the TypeName
    ' of those objects.
    ' Can be expanded as needed.
    ' If TypeName of an object is not included here, object TypeName is
    ' dumped.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With objects
        .Add Array("FirstIndex", "Length", "Value"), "IMAtch2"
    End With
    
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
                Set CurrentValue = Variable(currentKey)
            Else
                CurrentValue = Variable(currentKey)
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
            CurrentValue = DUMP(CurrentValue, Deepness + 1, LineBreaks, ShowArrayIndexes)
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Add key: value pair string to the output string and add an
            ' item separator.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            output = output & showKy & ": " & CurrentValue & itemSeparator
            
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
    ElseIf "Collection, IMatchCollection2" Like "*" & TypeName(Variable) & "*" Then
        
        If Variable.Count <= MaxItems Then
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
        Else
            If TypeName(Variable) <> "Collection" Then output = TypeName(Variable)
            output = output & "[" & Variable.Count & " items]"
        End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 5. Other values
    ' If not any of the previous types, return type name.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Any other variable
        ' If a TypeName is included in the objects collection keys,
        ' its properties which names are stored in the collection
        ' are dumped.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set objectMemberNames = Nothing
        On Error Resume Next
        objectMemberNames = objects(TypeName(Variable))
        On Error GoTo 0
        If IsArray(objectMemberNames) Then
            output = TypeName(Variable) & "("
            For Each memberName In objectMemberNames
                CurrentValue = CallByName(Variable, memberName, VbGet)
                If TypeName(CurrentValue) = "String" Then CurrentValue = """" & CurrentValue & """"
                output = output & memberName & ": " & CurrentValue & ", "
            Next memberName
            output = Left(output, Len(output) - 2) & ")"
        Else
            '''''''''''''''''''''''''''
            ' Everything else
            ' Just a TypeName is dumped
            '''''''''''''''''''''''''''
            output = TypeName(Variable)
        End If
    End If
    DUMP = output
End Function