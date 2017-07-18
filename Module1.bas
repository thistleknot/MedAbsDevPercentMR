Attribute VB_Name = "Module1"
Function DeriveMAD(myArray As Variant)
    
    Dim myArray2() As Variant
    
    Dim MAD As Double
    
    'Get Column size
    Dim Rows As Integer
    Rows = Application.WorksheetFunction.Count(myArray)
    
    'set myArray2 to 1 to # of Rows
    ReDim myArray2(1 To Rows)
    
    'Derive distance from Median
    Dim i As Integer
    For i = 1 To Rows
        myArray2(i) = Abs(myArray(i) - Application.WorksheetFunction.Median(myArray))
    Next i
    
    MAD = Application.WorksheetFunction.Median(myArray2)
    
    DeriveMAD = MAD
    
End Function

Function DeriveMADZPct(value As Double, myArray As Variant)

    Dim myArray3() As Variant
    
    Dim MAD As Double
    
    MAD = DeriveMAD(myArray)

    Dim Rows As Integer
    
    Rows = Application.WorksheetFunction.Count(myArray)
    
    'set myArray2 to 1 to # of Rows
    ReDim myArray3(1 To Rows)
    
    Dim i As Integer
    For i = 1 To Rows
        myArray3(i) = (myArray(i) - Application.WorksheetFunction.Median(myArray)) / MAD
    Next i
    
    value = (value - Application.WorksheetFunction.Median(myArray)) / MAD
    
    If (Abs(value) <= 1) Then
        DeriveMADZPct = ((value + 1) / 4) + 0.25
    ElseIf (value > 1) Then
        DeriveMADZPct = ((value - 1) / (Application.WorksheetFunction.Max(myArray3) - 1)) * 0.25 + 0.75
    ElseIf (value < -1) Then
        DeriveMADZPct = 0.25 - ((value + 1) / (Application.WorksheetFunction.Min(myArray3) + 1) * 0.25)
    ElseIf value = 0 Then
        DeriveMADZPct = 0.5
    End If
    
    
    'DeriveMADZ = myArray2
    
End Function

'Returns array as Percents based on MAD Z Score normalization around 1 MAD.
Function DeriveMADZPercents(myArray As Variant)

    Dim myArray4() As Variant
    
    Dim MAD As Double
    
    Dim MaxZ As Double
    
    Dim MinZ As Double
    
    MAD = DeriveMAD(myArray)

    Dim Rows As Integer
    
    Rows = Application.WorksheetFunction.Count(myArray)
    
    'set myArray2 to 1 to # of Rows
    ReDim myArray4(1 To Rows)
        
    Dim i As Integer
    
    'Assign Z's first
    For i = 1 To Rows
        
        value = myArray(i)
        value = (value - Application.WorksheetFunction.Median(myArray)) / MAD
        myArray4(i) = value
    
    Next i
    
    'get Min
    'get Max
    
    MaxZ = Application.WorksheetFunction.Max(myArray4)
    MinZ = Application.WorksheetFunction.Min(myArray4)
    
    For i = 1 To Rows
    
        value = myArray4(i)
        
        If (Abs(value) <= 1) Then
            value = ((value + 1) / 4) + 0.25
        ElseIf (value > 1) Then
            
            value = (value - 1) / (MaxZ - 1) * 0.25 + 0.75
            '(value - 1) '/ (Application.WorksheetFunction.Max(myArray3) - 1) '* 0.25 + 0.75
        ElseIf (value < -1) Then
            value = 0.25 - ((value + 1) / (MinZ + 1) * 0.25)
        ElseIf value = 0 Then
            value = 0.5
        End If
        myArray4(i) = value
    
    Next i
    
    'Assign %'s second (need Z's to derive Max's/Min's)
    
    DeriveMADZPercents = Application.WorksheetFunction.Transpose(myArray4())
    
End Function
