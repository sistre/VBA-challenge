Sub Stock_Analysis()

' Loop for each worksheet
For Each ws In Worksheets

' Add the column labels to the necessary columns

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume (K)"

' Add the column/row labels to the necessary ws.cells for the challenges

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume (K)"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Set the new colums to autofit
ws.Columns("I:Q").AutoFit

' Create a variable to hold the total volume
Dim TotalVol As Double

' Create a variable to hold the opening value and sets the opening value
Dim OpenVal As Double
OpenVal = ws.Cells(2, 3).Value

' Create a variable to hold the output row
Dim OutputRow As Long
OutputRow = 2

' For loop to cycle through the ticker column
For Row = 2 To ws.Range("A1").End(xlDown).Row

' Sets the initial value of the total volume to the volume of the first entry
    TotalVol = TotalVol + (ws.Cells(Row, 7).Value / 1000)
    
' Conditional to check if the ticker name has changed, writes the ticker name and total stock volume, and runs the yearly change and % change calculations
    If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
 
 ' Writes the ticker name to the appropriate cell
        ws.Cells(OutputRow, 9).Value = ws.Cells(Row, 1).Value
 
 ' Writes the total stock volume to the appropriate cell
        ws.Cells(OutputRow, 12).Value = TotalVol
        
' Calulates the yearly change and writes it to the appropriate cell and sets the background color based on if it is positive or negative
        ws.Cells(OutputRow, 10).Value = (ws.Cells(Row, 6).Value - OpenVal)
        If ws.Cells(OutputRow, 10).Value < 0 Then
            ws.Cells(OutputRow, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(OutputRow, 10).Value > 0 Then
                ws.Cells(OutputRow, 10).Interior.ColorIndex = 4
            Else: ws.Cells(OutputRow, 10).Interior.ColorIndex = 6
        End If

' Calulates the percent change and writes it to the appropriate cell and sets the result to two decimal places and percentage formatting
        If OpenVal = 0 Then
            ws.Cells(OutputRow, 11).Value = "NaN"
        Else
        ws.Cells(OutputRow, 11).Value = ((ws.Cells(Row, 6).Value - OpenVal) / OpenVal)
        ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
        End If
        
 ' Sets the opening value to the new value
        OpenVal = ws.Cells(Row + 1, 3).Value
        
' Cycles to the next output row
        OutputRow = OutputRow + 1
        
 ' Resets the total volume
        TotalVol = 0
        
    End If

Next Row

' Creates variables for row placeholders for the greatest % increase, greatest % decrease, greatest total volume
Dim MaxPercentRow As Long
MaxPercentRow = 2

Dim MinPercentRow As Long
MinPercentRow = 2

Dim MaxVolRow As Long
MaxVolRow = 2

' For loop to determine the greatest % increase, greatest % decrease, greatest total volume
For MinMax = 2 To ws.Range("K1").End(xlDown).Row

' Conditional to find the greatest % increase and capture what row it is on and filter out "NaN" values
    
    If ws.Cells(MinMax, 11).Value <> "NaN" And ws.Cells(MinMax, 11).Value > ws.Cells(MaxPercentRow, 11).Value Then
        MaxPercentRow = MinMax
    End If

' Conditional to find the greatest % decrease and capture what row it is on
    If ws.Cells(MinMax, 11).Value < ws.Cells(MinPercentRow, 11).Value Then
        MinPercentRow = MinMax
    End If

' Conditional to find the greatest total volume and capture what row it is on
    If ws.Cells(MinMax, 12).Value > ws.Cells(MaxVolRow, 12).Value Then
        MaxVolRow = MinMax
    End If

Next MinMax

' Sets the values of the greatest % increase, greatest % decrease, greatest total volume to the appropriate ws.cells
ws.Cells(2, 16).Value = ws.Cells(MaxPercentRow, 9).Value
ws.Cells(2, 17).Value = ws.Cells(MaxPercentRow, 11).Value

ws.Cells(3, 16).Value = ws.Cells(MinPercentRow, 9).Value
ws.Cells(3, 17).Value = ws.Cells(MinPercentRow, 11).Value

' Sets the greatest % increase and greatest % decrease results to two decimal places and percentage formatting
ws.Range("Q2:Q3").NumberFormat = "0.00%"

ws.Cells(4, 16).Value = ws.Cells(MaxVolRow, 9).Value
ws.Cells(4, 17).Value = ws.Cells(MaxVolRow, 12).Value

Next ws

End Sub