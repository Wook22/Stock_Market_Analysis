Sub ConditionalLoop()
    
    ' Loop through all worksheets in the workbook
    Dim alphabetic  As Worksheet
    
    ' Turn off screen updating to speed up the macro
    Application.ScreenUpdating = FALSE
    
    ' Declare variables
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim RowCount1   As Integer
    Dim RowCount2   As Integer
    Dim Count1      As Integer
    Dim Count2      As Integer
    Dim Percentage  As Double
    Dim Total       As Double
    Dim GreatInc    As Double
    Dim GreatDec    As Double
    Dim GreatVol    As Double
    
    ' Loop through each worksheet
    For Each alphabetic In Worksheets
        
        ' Add headers to the worksheet
        With alphabetic
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(1, 10).Value = "Yearly Change"
        End With
        
        ' Find the last row in the worksheet
        RowCount1 = alphabetic.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize counters
        Count1 = 2
        Count2 = 2
        
        ' Loop through each row in the worksheet
        For i = 2 To RowCount1
            
            ' Check if the ticker symbol has changed
            If alphabetic.Cells(i + 1, 1).Value <> alphabetic.Cells(i, 1).Value Then
                
                ' Add the ticker symbol to the summary table
                alphabetic.Cells(Count2, 9).Value = alphabetic.Cells(i, 1).Value
                
                ' Calculate the yearly change for the ticker symbol
                alphabetic.Cells(Count2, 10).Value = alphabetic.Cells(i, 6).Value - alphabetic.Cells(Count1, 3).Value
                
                ' Calculate the total stock volume for the ticker symbol
                Total = Total + alphabetic.Cells(i, 7).Value
                alphabetic.Cells(Count2, 12).Value = Total
                
                ' Color the yearly change cell based on its value
                If alphabetic.Cells(Count2, 10).Value < 0 Then
                    alphabetic.Cells(Count2, 10).Interior.ColorIndex = 3
                Else
                    alphabetic.Cells(Count2, 10).Interior.ColorIndex = 4
                End If
                
                ' Calculate the percentage change for the ticker symbol
                If alphabetic.Cells(Count1, 3).Value <> 0 Then
                    Percentage = ((alphabetic.Cells(i, 6).Value - alphabetic.Cells(Count1, 3).Value) / alphabetic.Cells(Count1, 3).Value)
                    alphabetic.Cells(Count2, 11).Value = Format(Percentage, "Percent")
                Else
                    alphabetic.Cells(Count2, 11).Value = Format(0, "Percent")
                End If
                
                ' Reset the total stock volume and counter variables
                Total = 0
                Count1 = i + j
                Count2 = Count2 + 1
                
            Else
                
                Total = Total + alphabetic.Cells(i, 7).Value
                
            End If
            
        Next i
        'Find the last row of data in column I
        RowCount2 = alphabetic.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Initialize variables to track greatest increase, greatest decrease, and greatest volume
        GreatInc = alphabetic.Cells(2, 11).Value
        GreatDec = alphabetic.Cells(2, 11).Value
        GreatVol = alphabetic.Cells(2, 12).Value
        
        'Loop through each row of data in the worksheet
        For k = 2 To RowCount2
            
            'Check if the value in column K (change in price) is greater than the current greatest increase
            If alphabetic.Cells(k, 11).Value > GreatInc Then
                'If so, update the value of the greatest increase
                GreatInc = alphabetic.Cells(k, 11).Value
                
                'Update the values in columns P and Q to show the stock with the greatest increase and the percentage of the increase
                alphabetic.Cells(2, 16).Value = alphabetic.Cells(k, 9).Value
                alphabetic.Cells(2, 17).Value = GreatInc
                alphabetic.Cells(2, 17).Value = Format(GreatInc, "Percent")
                
            End If
            
            'Check if the value in column K (change in price) is less than the current greatest decrease
            If alphabetic.Cells(k, 11).Value < GreatDec Then
                'If so, update the value of the greatest decrease
                GreatDec = alphabetic.Cells(k, 11).Value
                
                'Update the values in columns P and Q to show the stock with the greatest decrease and the percentage of the decrease
                alphabetic.Cells(3, 16).Value = alphabetic.Cells(k, 9).Value
                alphabetic.Cells(3, 17).Value = GreatDec
                alphabetic.Cells(3, 17).Value = Format(GreatDec, "Percent")
                
            End If
            
            'Check if the value in column L (total stock volume) is greater than the current greatest volume
            If alphabetic.Cells(k, 12).Value > GreatVol Then
                'If so, update the value of the greatest volume
                GreatVol = alphabetic.Cells(k, 12).Value
                
                'Update the values in columns P and Q to show the stock with the greatest volume and the volume in scientific notation
                alphabetic.Cells(4, 16).Value = alphabetic.Cells(k, 9).Value
                alphabetic.Cells(4, 17).Value = GreatVol
                alphabetic.Cells(4, 17).Value = Format(GreatVol, "Scientific")
                
            End If
            
        Next k
        
        'Adjust the width of columns J-O to make sure all data is visible
        alphabetic.Columns("J:O").ColumnWidth = 20
        
    Next alphabetic
    
    'Ensure screen updating is turned back on
    Application.ScreenUpdating = TRUE
    
End Sub
