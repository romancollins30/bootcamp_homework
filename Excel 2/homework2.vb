Sub DisplayAll()
    Dim i As Double
    Dim a As Double
    Dim vol As Double
    Dim lastrow As Double
    Dim change As Double
    Dim percent As Double
    Dim op As Double
    Dim cl As Double
    Dim maxinc As Double
    Dim mininc As Double
    
    i = 0
    a = 1
    vol = 0
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    change = 0
    percent = 0
    op = 0
    cl = 0
    maxinc = 0
    mininc = 0
    
'Cell labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greates % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    For i = 2 To lastrow
 
 'if the ticker cell at row i is different from the one above, the program moves down a row in my new table,
 'displays the new name, and sets volume to zero before adding the initial volume
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            a = a + 1
            Cells(a, 9).Value = Cells(i, 1).Value
            vol = 0
            Cells(a, 10).Value = Cells(i, 7).Value
        End If
 
 'if the ticker cell at row i is the same as the above row, it adds the volume to my running count and sets my
 'cell value for total volume to the new number
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            vol = vol + Cells(i, 7).Value
            Cells(a, 12) = vol
        End If
            
    Next i
    
    i = 0
    a = 1
    
'finds the yearly change and percent change
    For i = 2 To lastrow

'only takes the open from the first row
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i - 1, 2).Value = "<date>" Then
            op = Cells(i, 3).Value
'skips rows with opens of 0
        ElseIf Cells(i, 3).Value = 0 Then
            a = a + 1
'does math for the previous stock when it hits a new ticker value
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i - 1, 2).Value <> "<date>" Then
            a = a + 1
            cl = Cells(i - 1, 6).Value
            change = cl - op
            percent = change / op
            Cells(a, 10).Value = change
            Cells(a, 11).Value = percent
            op = Cells(i, 3).Value
        End If
            
    Next i
   
'sets the lastrow value to the last row in column 10
    lastrow = Cells(Rows.Count, 10).End(xlUp).Row
    i = 0
        
'Color Codes the yearly change column
    For i = 2 To lastrow
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
        
        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

    vol = 0

'finds the max increase, max decrease, and max total volume
    For i = 2 To lastrow

'finds if the current row is greater than the previous one
        If Cells(i, 12).Value > vol Then
            vol = Cells(i, 12).Value
            Cells(4, 16).Value = vol
            Cells(4, 15).Value = Cells(i, 9).Value
        End If
        
'finds if the curent row has a greater increase than the previous one and saves the value
        If Cells(i, 11).Value > maxinc Then
            maxinc = Cells(i, 11).Value
            Cells(2, 16).Value = maxinc
            Cells(2, 15).Value = Cells(i, 9).Value
        End If
        
'finds if the curent row has a greater increase than the previous one and saves the value
        If Cells(i, 11).Value < mininc Then
            mininc = Cells(i, 11).Value
            Cells(3, 16).Value = mininc
            Cells(3, 15).Value = Cells(i, 9).Value
        End If
    Next i
    
    
End Sub
