Attribute VB_Name = "modStringFunctions"

Option Explicit
Option Compare Text


'_________________________________________________________________________________________________________________________________________________________________

Function SEPAR(ByVal CommaSeparatedText As String, _
               Optional ByVal ColumnDelimiter As String = ";", _
               Optional ByVal RowDelimiter As String = vbCrLf, _
               Optional ByVal TextSpecifier As String = """", _
               Optional ByVal FunctionEvaluateStatement As String = "", _
               Optional ByVal ValuePlaceholderForFunction As String = "{value}", _
               Optional ByVal OutputFormat As Integer = 1001, _
               Optional ByVal HeaderRowNumber As Long = 1, _
               Optional ByVal ColumnNumberForRowKey As Long = 0, _
               Optional ByVal ForceEventsCompletion As Boolean = False) As Variant
               
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function: SEPAR
' Function separates text separated by specified delimiter to the specified output format.
' Delimiters inside text specifier characters are ignored.
'
' Arguments:
'   - CommaSeparatedText as String - the source text to be separated
'   - ColumnDelimiter as String - a string that delimits individual columns in the
'                                 source text. Defaults to ";".
'   - RowDelimiter as String - a string that means a new line. Defaults to te new line
'                              character vbCrLf. Needs to be changed to vbLf if unix new
'                              line character is used.
'   - TextSpecifier as String - a string that enclose a textual part of the text i.e.
'                               delimiters inside of two text specifiers will be ignored.
'   - FunctionEvaluateStatement as String - a string representation of a function that
'                                           will be used for every value. It can be
'                                           a function in a module. In this case it needs
'                                           to be called e.g. "Module1.myFunction". The
'                                           return value will be used in the output.
'                                           It can also be a text representation of a
'                                           worksheet function like "=TRIM({value})". The
'                                           ValuePlaceholderForFunction (in this case
'                                           "{value}" will be replaced by the values from
'                                           the text.
'   - ValuePlaceholderForFunction As String - a string used as a placeholder for a value
'                                             in FunctionEvaluateStatement. Defaults to
'                                             "{value}". It has to be present in the
'                                             FunctionEvaluateStatement or the function
'                                             won't be run.
'   - OutputFormat as Integer - specifies a returned output format of the function.
'                               Function can return a collection, dictionary or array.
'                               Value can be any integer. Defaults to 1001.
'                               Function returns for
'                               -----------------------------
'                               |   Value   | Returned type |
'                               -----------------------------
'                               |   1001    |  Collection   |
'                               |   1002    |  Dictionary   |
'                               | Any other |  Array        |
'                               -----------------------------
'                               Custom constants from modConstants can be used i.e.:
'                               cARRAY, cCOLLECTION, cDICTIONARY.
'                               Since for dictionary all keys have to be created, it takes
'                               some time. If you value performance, choose either array
'                               or collection.
'   - HeaderRowNumber as Long - row index that is to be used as keys for items in a row.
'                               Only relevant when dictionary output is chosen. Defaults
'                               to 1. If set to less than 1, indexes of items are used.
'   - ColumnNumberForKey as Long - column index that holds keys for individual rows.
'                                  Defaults to 0. If set to less than 1, indexes of rows
'                                  are used.
'   - ForceEventsCompletion as Boolean - if set to true, DoEvents command is executed after
'                                        every item. It is helpful when you are trying to
'                                        process a very long text and you want an option to
'                                        stop/pause the script and prevent an "not
'                                        responding" problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               
    Dim _
        rowCll As New Collection, _
        outputCll As New Collection, _
        RowCollection As Collection, _
        headerRow As Collection, _
        colItem As Variant

    Dim _
        textPartsRgx, _
        delimitersRgx, _
        delimitersMatches, _
        textPartsMatches, _
        delimiterMatch
    Dim _
        cursor As Long, _
        maxLength As Long, _
        textPartMatchInd As Long, _
        firstIndex As Long, _
        lastIndex As Long, _
        textFirstIndex As Long, _
        textLastIndex As Long, _
        iRow As Long, _
        iCol As Long, _
        columnDeduplicator As Long, _
        rowDeduplicator As Long
    Dim _
        rowKey As String, _
        parsedText As Variant, _
        processedParsedText As Variant, _
        columnkey As String
    Dim _
        outputArray() As Variant, _
        caller As Variant
    Dim _
        separate As Boolean
    Dim _
        outputDic, rowDic
        
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'CALLER Check - If called from Range, output format is set to vbArray
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
        Set caller = Application.caller
        If Not caller Is Nothing Then
            If TypeName(caller) = "Range" Then
                OutputFormat = 0
            End If
        End If
    On Error GoTo 0
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Prepare regex for text parts and delimiters and execute search
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set textPartsRgx = CreateObject("VBScript.RegExp")
    With textPartsRgx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = TextSpecifier & ".*?" & TextSpecifier
    End With
    Set delimitersRgx = CreateObject("VBScript.RegExp")
    With delimitersRgx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = Replace("(" & ColumnDelimiter & ")|(" & RowDelimiter & ")|($)", "\", "\\")
    End With
    Set textPartsMatches = textPartsRgx.Execute(CommaSeparatedText)
    Set delimitersMatches = delimitersRgx.Execute(CommaSeparatedText)
    
    
    '''''''''''''''''''''''''''''''''
    ' Set default values for counters
    '''''''''''''''''''''''''''''''''
    cursor = 1
    textPartMatchInd = 0
    maxLength = 0
    
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through found delimiters
    ''''''''''''''''''''''''''''''''''''''''''''
    For Each delimiterMatch In delimitersMatches
        
        separate = True 'predespose positive separation
        
        ''''''''''''''''''''''''''''''''
        ' Set delimiter properties:
        ' first and last index
        ''''''''''''''''''''''''''''''''
        With delimiterMatch
            firstIndex = .firstIndex
            lastIndex = .firstIndex + .length
        End With
        
        '''''''''''''''''''''''''''''''''''''''''
        ' Check if the delimiter is not inside of
        ' some textual part of the string - loop
        ' through textual parts until the textual
        ' part starts after the delimiter. Remember
        ' last checked textual part.
        '''''''''''''''''''''''''''''''''''''''''
        For textPartMatchInd = textPartMatchInd To textPartsMatches.Count - 1
            ''''''''''''''''''''''''''
            ' Get first and last index
            ' of a textual part.
            ''''''''''''''''''''''''''
            With textPartsMatches(textPartMatchInd)
                textFirstIndex = .firstIndex
                textLastIndex = .firstIndex + .length
            End With
            
            ''''''''''''''''''''''''''''''''''
            ' Check if the delimiter is in the
            ' current textual part - if it is,
            ' set separation to false and exit
            ' loop.
            ''''''''''''''''''''''''''''''''''
            If textFirstIndex < lastIndex And _
                textLastIndex > firstIndex Then
                    separate = False
                    Exit For
            ''''''''''''''''''''''''''''''''''
            ' If current textual part starts
            ' after current delimeter, stop
            ' the loop and remember textual
            ' part index.
            ''''''''''''''''''''''''''''''''''
            ElseIf textFirstIndex >= lastIndex Then
                Exit For
            End If
        Next textPartMatchInd
    
        ''''''''''''''''''''''''''''''''''''
        ' Separate if the delimiter is valid
        '
        ''''''''''''''''''''''''''''''''''''
        If separate Then
            
            '''''''''''''''''''''''''''''''''''''''''
            ' Get the text between the cursor and the
            ' delimeter and remove textual specifier
            ' character.
            '
            '''''''''''''''''''''''''''''''''''''''''
            parsedText = Replace(Mid(CommaSeparatedText, cursor, firstIndex - cursor + 1), TextSpecifier, "")
            
            
            ''''''''''''''''''''''''''''
            ' Execute Function Statement
            ''''''''''''''''''''''''''''
            If FunctionEvaluateStatement <> "" Then
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Check if the statement references module function
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                On Error Resume Next
                processedParsedText = Application.Run(FunctionEvaluateStatement, parsedText)
                If Err.Number = 0 Then
                    On Error GoTo 0
                    parsedText = processedParsedText
                Else
                    
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' If it is not a module function,execute worksheet function
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Err.Clear
                    processedParsedText = Application.Evaluate(Replace(FunctionEvaluateStatement, ValuePlaceholderForFunction, """" & parsedText & """"))
                    If Not IsError(processedParsedText) Then
                        parsedText = processedParsedText
                    Else
                        Debug.Print "Error: PARSE.FunctionEvaluate returned " & CStr(processedParsedText) & " for '" & Replace(FunctionEvaluateStatement, ValuePlaceholderForFunction, """" & parsedText & """")
                    End If
                End If
            End If
            
            '''''''''''''''''''''
            ' Add processed text
            ' to row collection
            '''''''''''''''''''''
            rowCll.Add parsedText
            
            ''''''''''''''''''''''
            ' Move cursor
            ''''''''''''''''''''''
            cursor = lastIndex + 1
            
            '''''''''''''''''''''''''''''''''''''''''''''
            ' End row and add it to the output collection
            ' if row delimiter encountered or at the end
            ' of the input text.
            '''''''''''''''''''''''''''''''''''''''''''''
            If delimiterMatch.Value = RowDelimiter Or _
                firstIndex = Len(CommaSeparatedText) Then
                outputCll.Add rowCll
                
                ''''''''''''''''''''''''''
                ' Store maximal row length
                ' for output array
                ''''''''''''''''''''''''''
                If rowCll.Count > maxLength Then
                    maxLength = rowCll.Count
                End If
                Set rowCll = New Collection
            End If
        End If
        If ForceEventsCompletion Then DoEvents
    Next delimiterMatch
    
    
    
    If OutputFormat = 1001 Then
    ''''''''''''''''''''
    'COLLECTION output
    ''''''''''''''''''''
        Set SEPAR = outputCll
    
    ElseIf OutputFormat = 1002 Then
    ''''''''''''''''''''
    'DICTIONARY output
    ''''''''''''''''''''
        
        Set outputDic = CreateObject("Scripting.Dictionary")
        
        '''''''''''''''''''''''
        ' Set default values of
        ' helping vars
        '''''''''''''''''''''''
        iRow = 1
        columnDeduplicator = 1
        rowDeduplicator = 1
        
        ''''''''''''''''''''''''''''''''
        ' If HeaderRowNumber argument is
        ' set, get the header row
        ''''''''''''''''''''''''''''''''
        If HeaderRowNumber > 0 Then Set headerRow = outputCll(HeaderRowNumber)
        
        
        For Each RowCollection In outputCll
            
            '''''''''''''''''''''''''''''''''
            ' Loop through rows that are not
            ' header row
            '''''''''''''''''''''''''''''''''
'            If HeaderRowNumber > 0 And iRow <> HeaderRowNumber Then
            If iRow <> HeaderRowNumber Then
                
                ''''''''''''''''''''''''''
                ' Set default column index
                ' and create dic object
                ''''''''''''''''''''''''''
                iCol = 1
                Set rowDic = CreateObject("Scripting.Dictionary")
                
                ''''''''''''''''''''''''''''''''''
                ' Loop through row colletion items
                ''''''''''''''''''''''''''''''''''
                For Each colItem In RowCollection
                    
                    
                    ''''''''''''''''''''''''''''''
                    ' Pick dictionary key for the
                    ' current item
                    ''''''''''''''''''''''''''''''
                    If HeaderRowNumber > 0 Then
                        columnkey = headerRow(iCol)
                    Else
                        If iCol > ColumnNumberForRowKey Then
                            columnkey = iCol - 1
                        Else
                            columnkey = iCol
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''
                    ' Check for key duplicates and
                    ' ensure uniqueness
                    ''''''''''''''''''''''''''''''
                    If rowDic.Exists(columnkey) Then
                        columnkey = columnkey & columnDeduplicator
                        columnDeduplicator = columnDeduplicator + 1
                    End If
                    
                    '''''''''''''''''''''''''''''''
                    ' Add key: item pair to the row
                    ' dictionary
                    '''''''''''''''''''''''''''''''
                    If iCol <> ColumnNumberForRowKey Then rowDic.Add columnkey, colItem
                    
                    '''''''''''''''''''''''
                    ' Increase column index
                    '''''''''''''''''''''''
                    iCol = iCol + 1
                Next colItem
                
                ''''''''''''''''''''''
                ' Pick key for the row
                ''''''''''''''''''''''
                If ColumnNumberForRowKey > 0 And ColumnNumberForRowKey < RowCollection.Count Then
                    rowKey = RowCollection(ColumnNumberForRowKey)
                Else
                    If iRow > HeaderRowNumber Then
                        rowKey = iRow - 1
                    Else
                        rowKey = iRow
                    End If
                End If
                
                ''''''''''''''''''''''''''''''
                ' Check for row key duplicates
                ' and ensure uniqueness
                ''''''''''''''''''''''''''''''
                If outputDic.Exists(rowKey) Then
                    rowKey = rowKey & rowDeduplicator
                    rowDeduplicator = rowDeduplicator + 1
                End If
                
                '''''''''''''''''''''''''''''''''
                ' Add row dictionary with its key
                ' to the output dictionary
                '''''''''''''''''''''''''''''''''
                outputDic.Add rowKey, rowDic
                
            End If
            
            ''''''''''''''''''''
            ' Increase row index
            ''''''''''''''''''''
            iRow = iRow + 1
        Next RowCollection
        
        Set SEPAR = outputDic
    Else
    '''''''''''''''''''
    ' Dictionary output
    '''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''
        ' Define output array dimensions
        ' and default row index
        ''''''''''''''''''''''''''''''''
        If maxLength = 0 Then maxLength = outputCll(1).Count
        ReDim outputArray(outputCll.Count - 1, maxLength - 1)
        iRow = 0
        
        '''''''''''''''''''''''''''''''''''''
        ' Loop through row collections in the
        ' output collection
        '''''''''''''''''''''''''''''''''''''
        For Each RowCollection In outputCll
            
            ''''''''''''''''''''''''''''''''''''''
            ' Loop through items in row collection
            ' and add them to the output array
            ''''''''''''''''''''''''''''''''''''''
            iCol = 0
            For Each colItem In RowCollection
                If iCol <= UBound(outputArray, 2) Then
                    outputArray(iRow, iCol) = colItem
                    iCol = iCol + 1
                End If
            Next colItem
            iRow = iRow + 1
        Next RowCollection
        
        SEPAR = outputArray
    End If
End Function

'_________________________________________________________________________________________________________________________________________________________________

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
    If Err.Number = 13 Then
        Err.Clear
        fCLng = CLng(Replace(Value, ".", ","))
        If Err.Number <> 0 Then Err.Raise vbObjectError + 100, , "fCLng: wrong value with error " & Err.Number & ": " & Err.Description
    End If
    
End Function
'_________________________________________________________________________________________________________________________________________________________________

Function Conv(ByVal Number As Long, ParamArray DimensionSettings() As Variant) As Variant

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Wrapper function for Convert function. Makes use of parameter array so the
' input of settings is easier.
'
' Parameters:
'     * Number as Long - number to be converted
'     * DimensionSettings - gattering array for dimension settings. Every dimension
'                           is a separate argument or you can use two dimensional
'                           array.
'
' Dependancies:
'     * Convert
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Conv = Convert(DimensionSettings, Number, False)
End Function

Function ConvRev(ByVal Number As Long, ParamArray DimensionSettings() As Variant) As Variant
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Wrapper function for Convert function but in a reversed order. Makes use of
' parameter array so the input of settings is easier.
'
' Parameters:
'     * Number as Long - number to be converted
'     * DimensionSettings - gattering array for dimension settings. Every dimension
'                           is a separate argument or you can use two dimensional
'                           array.
'
' Dependancies:
'     * Convert
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ConvRev = Convert(DimensionSettings, Number, True)
End Function


Function Convert(ByVal InputArray As Variant, ByVal Number As Long, _
                Optional ByVal Reversed As Boolean = False) As Variant

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
' Dependancies:
'   - fCLng
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim bounds() As Variant, _
        stt() As Variant
    Dim NumDims As Integer, _
        NumDims0 As Integer, _
        Ndx As Integer, _
        bottom As Long, _
        i As Long
    Dim solved As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''
    ' Check the dimensions of the Input array.
    ' If the number of dimensions is not 1 or 2
    ' raise an error
    '''''''''''''''''''''''''''''''''''''''''''
    
    NumDims = NumberOfArrayDimensions(InputArray)
    If NumDims = 0 Or NumDims > 2 Then GoTo WrongDimensions
    If NumDims = 1 Then
        NumDims0 = NumberOfArrayDimensions(InputArray(LBound(InputArray)))
        If NumDims0 = 2 Then
            InputArray = InputArray(LBound(InputArray))
            NumDims = NumDims0
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Prepare an array for the calculation.
    ' Process different types of input arrays (1, 2 or 3 in the description)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If NumDims = 2 Then
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' 2 dimensional array. If types differ,
        ' new array is created and values copied to it
        ''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
            bounds = InputArray
            If Err.Number = 13 Then
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

    ReDim Result(LBound(bounds, 1) To UBound(bounds, 1)) As Variant
    
    ''''''''''''''''''''''''''''''''''''''
    ' Setting the mode of the calculation:
    ' normal or reversed
    ''''''''''''''''''''''''''''''''''''''
    If Reversed Then
        stt = Array(LBound(Result), UBound(Result), 1)
    Else
        stt = Array(UBound(Result), LBound(Result), -1)
    End If
    
    ''''''''''''''''''''''''''''''''''''
    ' Calculation of individual sections
    ''''''''''''''''''''''''''''''''''''
    solved = False
    For i = stt(0) To stt(1) Step stt(2)
        bottom = LBound(bounds, 2)
        If Not solved Then
            Result(i) = bounds(i, bottom) + (Number Mod (bounds(i, bottom + 1) - bounds(i, bottom) + 1))
            Number = WorksheetFunction.Quotient(Number, bounds(i, bottom + 1) - bounds(i, bottom) + 1)
        Else
            Result(i) = bounds(i, bottom) + 0
        End If
        If Number = 0 Then
            solved = True
        End If
    Next i
    If Not solved Then Err.Raise vbObjectError + 101, , "fConvert2: unsolved. Number is greater than the settings allow."
    
    Convert = Result
    Exit Function
WrongDimensions:
    Err.Raise vbObjectError - 1, , "fConvert2: wrong dimensions of the input array."
    Exit Function
End Function

'_________________________________________________________________________________________________________________________________________________________________

Function DUMP(ByVal Variable As Variant, _
                Optional ByVal LineBreaks As Boolean = False, _
                Optional ByVal ShowArrayIndexes As Boolean = False, _
                Optional ByVal Deepness As Integer = 0)
    
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
' * LineBreaks - if true, line breaks are added after individual items.
' * ShowArrayIndexes - if true, array item indexes will be shown to better orientate easier.
' * Deepness as Integer - a level of recursion. It will indent the items correctly. This doesn't
'                         need to be used. It is used automatically in case of recursion.
'
' Requirements:
'       - NumberOfArrayDimensions function
'       - Convert function
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim objects As New Collection
    Dim coordinates() As Variant, _
        objectMemberNames As Variant
    Dim output As String, _
        v As Variant, _
        Dimensions As Integer, _
        i As Long, _
        ItemIndex As Long, _
        dimensionIndex As Long, _
        consecutiveDimensionsIndex As Long, _
        previousDimensions As Long, _
        allDimensionsIndex As Long, _
        leadingDimensionIndex As Long, _
        currentKey As Variant, _
        CurrentValue As Variant, _
        showKy As Variant, _
        itm As Variant, _
        indexes As String, _
        memberName As Variant
    Dim openBracket As Boolean, _
        cls As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set default values of output and constants
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Const primitives As String = "String, Long, Integer, Boolean, Single, Double, Byte, Currency, Decimal, Date, Error"
    Const itemSeparator As String = ", "
    Const indentation As Variant = "  "
    Const MaxItems As Long = 200 'if number items in collection or dictionary exceeds this limit, show only number
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
        Dimensions = NumberOfArrayDimensions(Variable)
        
        
        If Dimensions = 0 Then
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
            ' is calculated with an Convert function to
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
            ReDim dimsDef(1 To Dimensions, 1 To 2) As Integer
            For i = 1 To Dimensions
                dimsDef(i, 1) = LBound(Variable, i)
                dimsDef(i, 2) = UBound(Variable, i)
            Next i
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Transfer each value in the array to a helper
            ' collection with item's index as its key.
            ' The index is calculated with Convert function in a form
            ' of dimension indexes joined with a dash like e.g.:
            ' "2-4-6-1"
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ItemIndex = 0
            For Each v In Variable
                keyVals.Add v, Join(Convert(InputArray:=dimsDef, Number:=ItemIndex, Reversed:=True), "-")
                ItemIndex = ItemIndex + 1
            Next v
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Traverse through the helper collection by calling the
            ' keys in a correct order and build a string output.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For i = 0 To ItemIndex - 1
                coordinates = Convert(dimsDef, i) ' Calculate coordinates for the current item.
                
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
                         DUMP(keyVals(Join(coordinates, "-")), LineBreaks, ShowArrayIndexes, Deepness) & _
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
                    For consecutiveDimensionsIndex = leadingDimensionIndex To UBound(coordinates)
                        If coordinates(consecutiveDimensionsIndex) <> dimsDef(consecutiveDimensionsIndex, 2) Then
                            cls = False
                            Exit For
                        End If
                    Next consecutiveDimensionsIndex
                    
                    ''''''''''''''''''''''''''''''''''''
                    ' Cloase a bracket if necessary
                    ''''''''''''''''''''''''''''''''''''
                    If cls Then
                    
                        ''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' Remove item separator before closing the bracket
                        ''''''''''''''''''''''''''''''''''''''''''''''''''
                        If Right(output, Len(itemSeparator)) = itemSeparator Then output = Left(output, Len(output) - Len(itemSeparator))
                        
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
                        output = output & ")" & itemSeparator
                    End If
                Next leadingDimensionIndex
            Next i
            
            ''''''''''''''''''''''''''''''''''''''''''
            ' Remove item separator after last bracket
            ''''''''''''''''''''''''''''''''''''''''''
            If Right(output, Len(itemSeparator)) = itemSeparator Then output = Left(output, Len(output) - Len(itemSeparator))
            
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
            CurrentValue = DUMP(CurrentValue, LineBreaks, ShowArrayIndexes, Deepness + 1)
            
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
                output = output & DUMP(itm, LineBreaks, ShowArrayIndexes, Deepness) & itemSeparator
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
'_________________________________________________________________________________________________________________________________________________________________



