Attribute VB_Name = "fCLng"

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