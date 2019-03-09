Attribute VB_Name = "Module1"
Sub StockMarket()

'Starter variable to make this work through a workbook
Dim ws As Worksheet

'Variable to make the index long enough to include all data
Dim LastRow As Long

'Set variable for holding the ticker
Dim ticker As String

'Set variable to hold value of total volume
Dim total_volume As Double

'Set volume total at start point
total_volume = 0

'Keep track of the location for each ticker in a Totals Table
Dim Total_Table_Row As Long
Total_Table_Row = 2

'loop thru each worksheet
For Each ws In Worksheets

    'Calculate the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through the years to total the volume
    For i = 2 To LastRow
    
        'check to see if we are at the same ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Set ticker symbol
            ticker = ws.Cells(i, 1).Value
    
            'add to volume total
            total_volume = total_volume + ws.Cells(i, 7).Value
    
            'print the ticker in column 10
            ws.Range("J" & Total_Table_Row).Value = ticker
        
            'Print the volume total to column 11
            ws.Range("K" & Total_Table_Row).Value = total_volume
        
            'Add one row to the Total Table
            Total_Table_Row = Total_Table_Row + 1
            
            'Reset total volume to zero
            total_volume = 0
        
        'If the cell immediately following the row is the same ticker.
        Else
            
            'Add the volume to the total
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
        
        Next i
        
        'reset total volume to zero
        total_volume = 0
        
        'reset table first row
        Total_Table_Row = 2
        
    Next ws
    
End Sub
