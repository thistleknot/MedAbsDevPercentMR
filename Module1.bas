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
    
    Dim theMedian As Double
    
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
    
End Function

Function ReturnColumns(myArray As Range)
    ReturnColumns = myArray.Columns.Count
End Function


Function ReturnArray(myArray As Range)

    Dim myArray4() As Variant

    Dim Rows As Integer
    
    Dim Columns As Integer
    
    Rows = Application.WorksheetFunction.Count(myArray)
    
    Columns = ReturnColumns(myArray)
    
    'set myArray4 to 1 to # of Rows
    ReDim myArray4(1 To Columns, 1 To Rows)
    
    myArray4 = myArray

    ReturnArray = myArray4()
    
End Function

'Returns array as Percents based on MAD Z Score normalization around 1 MAD.
Function DeriveMADZPercents(myArray As Range)

    Dim myArray4() As Variant
    
    Dim MAD As Double
    
    Dim MaxZ As Double
    
    Dim MinZ As Double
    
    MAD = DeriveMAD(myArray)

    Dim Rows As Integer
    
    Dim Columns As Integer
    
    Dim theMedian As Double
    
    theMedian = Application.WorksheetFunction.Median(myArray)
    
    Rows = Application.WorksheetFunction.Count(myArray)
    
    Columns = ReturnColumns(myArray)
    
    'Destination array. Set myArray4 to 1 to # of Rows
    ReDim myArray4(1 To Columns, 1 To Rows)
        
    Dim h As Integer
    
    For h = 1 To Columns
        
        Dim i As Integer
        
        MinZ = myArray(1, 1)
        MaxZ = myArray(1, 1)
        
        'Assign Z's first, derive max/min z's
        For i = 1 To Rows
            
            value = myArray(h, i)
            'value = (value - Application.WorksheetFunction.Median(myArray)) / MAD
            value = (value - theMedian) / MAD
            myArray4(h, i) = value
            
            'max/min z's
            If (value > MaxZ) Then
                MaxZ = value
            End If
            
            If (value < MinZ) Then
                MinZ = value
            End If
        
        Next i
        
        'necessary to have max and minz derived per column ahead of time.
        
        For i = 1 To Rows
        
            value = myArray4(h, i)
            
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
            myArray4(h, i) = value
        
        Next i
        
    Next h
        
    'Assign %'s second (need Z's to derive Max's/Min's)
    
    DeriveMADZPercents = Application.WorksheetFunction.Transpose(myArray4())
    
End Function
