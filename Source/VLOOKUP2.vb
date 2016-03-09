'This function will look up a search key in an array.
'Then if the key is found in the array it will return
'a value in the same row as the found key.
'
'As the name implies its inspired by the VLOOKUP function.
'Unlike that function it will caste every type into a string
'and it will compare the strings for similarity.  


Public Function VLOOKUP2(look_up As String, sh As String, _
    r1 As String, c As Integer) As String

    Dim r As Range
    Set r = ActiveWorkbook.Sheets(sh).Range(r1)
    found = False
    VLOOKUP2 = "#NA"
    
    'perform linear search through array
    For Each i In r
        If found = False Then
            If i.Value = look_up Then
		'c is relative to the array colum
		'if c is positive it will return a column to the right
		'if c is negative it will return a column to the left
		'zero just returns the key if found in the array
                VLOOKUP2 = i.Offset(0, c).Value
                found = True
            End If
        End If
    Next

End Function
