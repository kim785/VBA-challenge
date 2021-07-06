Attribute VB_Name = "Module1"
Sub Stock_data()

Dim ws As Worksheet
Dim worksheet_count As Integer
Dim last_row As Long
Dim ticker_type As Integer
Dim open_value As Double
Dim total_volume As Double
Dim last_row_2 As Long
Dim G_increase As Double
Dim G_decrease As Double
Dim G_volume As String

worksheet_count = ThisWorkbook.Worksheets.Count
' Counts the number of worksheets that are present in the workbook. MsgBox (worksheet_count) returns 3

For Each ws In Worksheets

    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' This function counts the number of rows in each worksheet
    
    ticker_type = 1
    ' stores the row number for the summary table
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ' The above 4 functions display the title/header for each summary column
    
    open_value = ws.Cells(2, 3).Value
    ' This function sets the initial open value for calculation of the "Yearly Change"
    
    total_volume = 0
    ' This function sets the initial total volume for calculation of the "Total Stock Volume"
    
    For i = 2 To last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker_type = ticker_type + 1
        ws.Cells(ticker_type, 9).Value = ws.Cells(i, 1).Value
        ' This fucntion will enter the ticker symbol on column I
        
        ws.Cells(ticker_type, 10).Value = ws.Cells(i, 6).Value - open_value
        ' The above function will calculate the change from the opening price at the beginning of the given year to the closing price at the end of that year
        
            If open_value = 0 Then
            ws.Cells(ticker_type, 11).Value = 0
            ' Some open values are 0. Computing percent change with open_value 0 gives an error
        
            Else
            ws.Cells(ticker_type, 11).Value = ws.Cells(ticker_type, 10).Value / open_value
            ' Calculates the percent change
        
            End If
        
        open_value = ws.Cells(i + 1, 3).Value
        ' Resets the open_value to the next ticker symbol's opening value of the year
        
            If ws.Cells(ticker_type, 10).Value < 0 Then
            ' a loop to color-code the change column
            
            ws.Cells(ticker_type, 10).Interior.ColorIndex = 3
            'color-code red for negative yearly change
            
            Else
            ws.Cells(ticker_type, 10).Interior.ColorIndex = 4
            'color-code green for positive yearly change
            
            End If
        
        total_volume = total_volume + ws.Cells(i, 7).Value
        ' adds tolume for each entries
        
        ws.Cells(ticker_type, 12).Value = total_volume
        ' finally prints the final total volume to the column L
        
        total_volume = 0
        ' resets total volume for the next ticker symbol
        
        Else
        total_volume = total_volume + ws.Cells(i, 7).Value
        ' If the ticker symbol is the same, this function keeps adding the volume of the same ticker symbol
        
        End If
        
    Next i

    ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
    ' Formats the "Percent Change" column to display percent values
    
    last_row_2 = ws.Range("I1").End(xlDown).Row
    ' This function counts the number of rows of ther summary table. This will be used for the second forloop for calculating Greatest % increase, decrease and total volume
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ' These functions display titles/headers for the bonus table
    
    G_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_2))
    G_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row_2))
    G_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_2))
    ' these 3 functions will find the greatest percent increase, decrease, and total volume from the first summary table I created
    
    ws.Cells(2, 17).Value = G_increase
    ws.Cells(3, 17).Value = G_decrease
    ws.Cells(4, 17).Value = G_volume
    ' prints the G values on the following cells
    
    For j = 2 To last_row_2
    
        If G_increase = ws.Cells(j, 11).Value Then
        ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
        ' prints the ticker symbol of the G increase
    
        ElseIf G_decrease = ws.Cells(j, 11).Value Then
        ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
        ' prints the ticker symbol of the G decrease
        
        ElseIf G_volume = ws.Cells(j, 12).Value Then
        ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
        ' prints the ticker symbol of the G volume
        
        End If
        
    Next j
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
Next ws

End Sub

