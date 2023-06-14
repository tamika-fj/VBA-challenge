Dim LastRow As Long
Dim Opening_Price As Double
Dim Closing_Price As Double
            
Sub GetWS()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS For ALL SUBROUTINES
    'ALL SUBROUTINES CAN BE RUN FROM THIS SUBROUTINE
    ' --------------------------------------------
    
    'Set Variable type for worksheet
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        Call DisplayHeadings(ws)
        Call SumTickerColumn(ws)
        Call Formatting(ws)
        Call Bonus(ws)
        
        
        
    Next ws
    
End Sub

Sub DisplayHeadings(ws As Worksheet)

'for each worksheet
    With ws
       
     'set new coloumn headings
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Yearly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 15
        
    End With
    
End Sub

Sub SumTickerColumn(ws As Worksheet)
   
    'set up variable for ticker
    Dim ticker As String

    'set varible for yearly change,total open,close and difference
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    Dim Total_Open As Double
    Total_Open = 0
    
    Dim Total_Close As Double
    Total_Close = 0
    
    Dim Total_Diff As Double
    Total_Diff = 0
    
    'set varible for percentage change and stock volume
    Dim PCent_Change As Double
    PCent_Change = 0
    
    Dim Total_Stock_Vol As LongLong
    Total_Stock_Vol = 0

    'set location for calculated data
    Dim New_column As Integer
    New_column = 2
 
    
    With ws
    
    'get last row
        LastRow = 0
        LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all values
        For i = 2 To LastRow

            'get values for all unique ticker value
            If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then

                'set ticker name
                ticker = .Cells(i, 1).Value

                'Set opening and closing price variable
                Dim Opening_Price As Double
                Dim Closing_Price As Double
                Dim Stock_Volume As LongLong
    
                'set value for opening and closing price
                Opening_Price = .Cells(i, 3).Value
                Closing_Price = .Cells(i, 6).Value
                Stock_Volume = .Cells(i, 7).Value

                'add to yearly change
                Yearly_Change = Yearly_Change + (Closing_Price - Opening_Price)
                
                'formula for Percentage change
                Total_Open = Total_Open + Opening_Price
                Total_Close = Total_Close + Closing_Price
                Total_Diff = Total_Close - Total_Open
                
                'add to Percentage Change
                PCent_Change = Total_Diff / Total_Open
                
                'add to total stock volume
                Total_Stock_Vol = Stock_Volume + Total_Stock_Vol
                
                'put yearly change in column J
                .Range("I" & New_column).Value = ticker
                .Range("J" & New_column).Value = Yearly_Change
                .Range("K" & New_column).Value = PCent_Change
                .Range("L" & New_column).Value = Total_Stock_Vol
                
                'Format Percentage Change Column
                .Range("K" & New_column).NumberFormat = "0.00%"
                
                'Set column width for Total Stock Value to display full number
                .Range("L" & New_column).ColumnWidth = 18

                'Add 1 to the New Column
                New_column = New_column + 1
                
                'Reset Yearly change
                Yearly_Change = 0
                Total_Open = 0
                Total_Close = 0
                
                'Reset Total Stock Volume
                Total_Stock_Vol = 0

            Else

                'set values for opening and closing price and stock volume
                Opening_Price = .Cells(i, 3).Value
                Closing_Price = .Cells(i, 6).Value
                
                'set values for stock volume
                Stock_Volume = .Cells(i, 7).Value

                'Add to yearly change
                Yearly_Change = Yearly_Change + (.Cells(i, 6).Value - .Cells(i, 3).Value)
                
                'formula for Percentage change
                Total_Open = Total_Open + .Cells(i, 3).Value
                Total_Close = Total_Close + .Cells(i, 6).Value
                Total_Diff = .Cells(i, 6).Value - .Cells(i, 3).Value
                
                'add to Percentage Change
                PCent_Change = .Cells(i, 6).Value - .Cells(i, 3).Value / Total_Open
                
                'Add to stock volume
                Total_Stock_Vol = .Cells(i, 7).Value + Total_Stock_Vol
                
            End If
        
        Next i

    End With

End Sub

Sub Formatting(ws As Worksheet)

'for all worksheets
With ws
    
    'get last row
        LastRow = 0
        LastRow = .Cells(Rows.Count, 10).End(xlUp).Row
        
'Loop through all values
For i = 2 To LastRow

'Set cell colour for Yearly Change Values, red for negaitve and green for positive values
    If .Cells(i, 10).Value >= 0 Then
    .Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf .Cells(i, 10).Value < 0 Then
    .Cells(i, 10).Interior.ColorIndex = 3
    

    
End If

Next i

End With

End Sub



Sub Bonus(ws As Worksheet)

'run code for all worksheets
With ws

'set summary table headings
.Cells(1, 16) = "Ticker"
.Cells(1, 17) = "Value"
.Cells(2, 15) = "Greatest % Increase"
.Cells(3, 15) = "Greatest % Decrease"
.Cells(4, 15) = "Greatest Total Volume"
.Columns("O").ColumnWidth = 20

'set variables for min and max values
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As LongLong

'set values for greatest total volume
GreatestTotalVolume = Application.WorksheetFunction.Max(.Columns("L"))

'get last row
        LastRow = 0
        LastRow = .Cells(Rows.Count, 10).End(xlUp).Row
        
'Loop through all values
For i = 2 To LastRow

'Set column to search for greatest total volume Values
If .Cells(i, 12).Value = GreatestTotalVolume Then

'input values in summary table
.Cells(4, 16).Value = .Cells(i, 9).Value
.Cells(4, 17).Value = .Cells(i, 12).Value
.Cells(4, 17).ColumnWidth = 18



End If
Next i

For i = 2 To LastRow

'set column to search for greatest percentage change increase and decrease and input values in summary table
If .Cells(i, 11).Value = Application.WorksheetFunction.Max(.Columns("K")) Then
.Cells(2, 16).Value = .Cells(i, 9).Value
.Cells(2, 17).Value = .Cells(i, 11).Value
.Cells(2, 17).NumberFormat = "0.00%"

ElseIf .Cells(i, 11).Value = Application.WorksheetFunction.Min(.Columns("K")) Then
.Cells(3, 16).Value = .Cells(i, 9).Value
.Cells(3, 17).Value = .Cells(i, 11).Value
.Cells(3, 17).NumberFormat = "0.00%"


End If
Next i
End With
End Sub

