
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call tickerTime
    Next
    Application.ScreenUpdating = True
End Sub



Sub tickerTime()


Dim i As Long
Dim totalVolume As Double
Dim rowCounter As Long
Dim Lastrow As Long
Dim ticker As String
Dim openValue As Double
Dim closeValue As Double
Dim firstrow As Long
Dim percentageChange As Double
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVol As Double


rowCounter = 1
ticker = Cells(2, 1).Value
totalVolume = 0
firstrow = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Ticker Volume"
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Looping through all the rows
For i = 2 To Lastrow

    'Checking to see if the ticker has changed from the previous value
    If Cells(i, 1).Value = ticker Then
    
    totalVolume = totalVolume + Cells(i, 7).Value

    'This counts how many rows there are with the same ticker value 
    firstrow = firstrow + 1
    
    Else
        'This variable is used to count how many times a new ticker has found so that we can input the values into another cell
        rowCounter = rowCounter + 1
        
        Cells(rowCounter, 12).Value = totalVolume
        Cells(rowCounter, 9).Value = ticker

        'Once we see that the ticker has changed, we re-assign the ticker variable to the new one it has just found 
        ticker = Cells(i, 1).Value

        'Once the total volume of the previous ticker has been calculated we reset the value for totalVolume to be 0
        totalVolume = 0

        'This if statement is to calculate the yearly and percent changes and to make sure that we are not dividing any 0 values 
        If Cells(i - firstrow, 3).Value <> 0 And Cells(i - 1, 6).Value <> 0 Then
        openValue = Cells(i - firstrow, 3).Value
        closeValue = Cells(i - 1, 6).Value
        
        
        yearlyChange = closeValue - openValue
        percentageChange = (closeValue / openValue) - 1
        Cells(rowCounter, 10) = yearlyChange
        Cells(rowCounter, 10).NumberFormat = "0.00000000"
        Cells(rowCounter, 11) = Format(percentageChange, "Percent")
        
        'If the value of a cell in column 11 is = 0, we will clear the cell under percent change so that we avoid having an error value
        Else
            
            Cells(rowCounter, 11).ClearContents
        
        End If
        
        'This resets the counter variable to 0 so it can keep counting how many rows are until the next new ticker
        firstrow = 0
        
 
        If yearlyChange > 0 Then
        
        Cells(rowCounter, 10).Interior.ColorIndex = 4
        
        Else
        
        Cells(rowCounter, 10).Interior.ColorIndex = 3
        
           End If
                
                
    End If
    
    Next i
    
    maxDecrease = 0
    maxIncrease = 0
    maxVol = 0
    
    Cells(2, 14).Value = "Maximum % Increase"
    Cells(3, 14).Value = " Maximum % Decrease"
    Cells(4, 14).Value = "Maximum Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    'This is to loop through all the results that we have calculated
    For j = 2 To Lastrow

        'starting condition is to check to see that the value in 'Percentage Changed' column is positive
        If Cells(j, 11).Value > maxIncrease Then

            'If the cell is more than 0 we will record it as the new maxIncrease value and this will then be the new value for the condition above
            maxIncrease = Cells(j, 11).Value
            Cells(2, 16) = Format(maxIncrease, "Percent")
            Cells(2, 15).Value = Cells(j, 1).Value
            
            'starting condition is to check to see that the value in 'Percentage Changed' column is negative
            ElseIf Cells(j, 11).Value < maxDecrease Then
            
            'If the cell is less than 0 we will record it as the new maxDecrease value and this will then be the new value for the condition above
            maxDecrease = Cells(j, 11).Value
            
            Cells(3, 16) = Format(maxDecrease, "Percent")
            Cells(3, 15).Value = Cells(j, 1).Value
            
            End If
            'starting condition is to check to see that the value in 'TotalVolume' column is positive
            If Cells(j, 12).Value > maxVol Then
                'If the cell is more than 0 we will record it as the new maxVol value and this will then be the new value for the condition above
                maxVol = Cells(j, 12).Value
                Cells(4, 16).Value = maxVol
                Cells(4, 15).Value = Cells(j, 1).Value
                
            Else
            
            End If
            
            Next j
        
    
   
End Sub


