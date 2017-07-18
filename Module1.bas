Attribute VB_Name = "Module1"
Function DeriveMAD(myArray As Variant)
    
    Dim myArray2() As Variant
    
    Dim MAD As Double
    
    'Get Column size
    Dim Columns As Integer
    Columns = Application.WorksheetFunction.Count(myArray)
    
    'set myArray2 to 1 to # of Columns
    ReDim myArray2(1 To Columns)
    
    'Derive distance from Median
    Dim i As Integer
    For i = 1 To Columns
        myArray2(i) = Abs(myArray(i) - Application.WorksheetFunction.Median(myArray))
    Next i
    
    MAD = Application.WorksheetFunction.Median(myArray2)
    
    DeriveMAD = MAD
    
End Function

Function DeriveMADZPct(value As Double, myArray As Variant)

    Dim myArray3() As Variant
    
    Dim MAD As Double
    
    MAD = DeriveMAD(myArray)

    Dim Columns As Integer
    
    Columns = Application.WorksheetFunction.Count(myArray)
    
    'set myArray2 to 1 to # of Columns
    ReDim myArray3(1 To Columns)
    
    Dim i As Integer
    For i = 1 To Columns
        myArray3(i) = (myArray(i) - Application.WorksheetFunction.Median(myArray)) / MAD
    Next i
    
    value = (value - Application.WorksheetFunction.Median(myArray)) / MAD
    
    If (Abs(value) <= 1) Then
        DeriveMADZPct = ((value + 1) / 4) + 0.25
    ElseIf (value > 1) Then
        DeriveMADZPct = (value - 1) / (Application.WorksheetFunction.Max(myArray3) - 1) * 0.25 + 0.75
    ElseIf (value < -1) Then
        DeriveMADZPct = 0.25 - ((value + 1) / (Application.WorksheetFunction.Min(myArray3) + 1) * 0.25)
    ElseIf value = 0 Then
        DeriveMADZPct = 0.5
    End If
    
    
    'DeriveMADZ = myArray2
    
End Function
