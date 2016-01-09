Attribute VB_Name = "RangeToArray"
Const ARRAYNAME = "MyArray"
Const ARRAYTYPE = "String"

Sub RangeToArray()
    Dim i As Integer
    Dim j As Integer
    Dim R As Variant

    If TypeName(Selection) <> "Range" Then
        Exit Sub
    End If

    Set R = ActiveWindow.Selection

    ' 1dim
    If R.Rows.Count = 1 Or R.Columns.Count = 1 Then
        Dim max
        If R.Rows.Count >= R.Columns.Count Then
            max = R.Rows.Count
        Else
            max = R.Columns.Count
        End If
        
        Debug.Print "Dim " & ARRAYNAME & "(" & max & ") As " & ARRAYTYPE
        For i = 0 To max - 1
            Debug.Print ARRAYNAME & "(" & i & ") = " & """" & R.Cells(i + 1) & """"
        Next
        Exit Sub
    End If
    
    ' 2dim
    Debug.Print "Dim " & ARRAYNAME & "(" & R.Rows.Count & "," & R.Columns.Count & ") As " & ARRAYTYPE

    For i = 0 To R.Rows.Count - 1
        For j = 0 To R.Columns.Count - 1

            Debug.Print ARRAYNAME & "(" & i & "," & j & ") = " & """" & R.Cells(i + 1, j + 1) & """"

        Next
    Next

End Sub
