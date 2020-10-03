Sub Summarize_Ticker_Data()

    ' declare variables
    Dim ws As Worksheet
    Dim Summary_Table_Row As Long
    Dim Last_Row As Long
    Dim ticker As String
    Dim Year_Open As Double
    Dim Stock_Volume As Double ' this is a big integer
    
            
    'Loop through every worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' For every worksheet, do this:
    
        ' sort the data by ticker and period -
        ' (if ever the data weren't sorted by ticker and date, the code wouldn't work)
        ws.Range("A1", ws.Range("G1").End(xlDown)).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Key2:=ws.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
        ' track the row of the summary table

        Summary_Table_Row = 1
            
        ' get the last row of the spreadsheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' loop through rows in the spreadsheet, first to last
        ' except we need to go to one row past the last row to get the last row to write
        ' the summary table write and the start of the new ticker are both triggered by the change in ticker symbol
        For i = 2 To Last_Row + 1
            
            ' set the variable ticker to the value of ticker for the current row
            ticker = ws.Cells(i, 1).Value
                    
            ' if the ticker symbol at the current row doesn't match the summary row, it's a new row, so closing data needs to be written
            ' else if the ticker symbol equals the current row, continue adding the stock volume data
            If ws.Cells(Summary_Table_Row, 9).Value <> ticker Then
                
                ' before we move to the next row, we need to write the total stock volume
                ' get the close amount unless it's the first row!
                If Summary_Table_Row <> 1 Then
                    ' write the stock volume to the summary table
                    ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume
                    
                    ' write the yearly change
                    ws.Cells(Summary_Table_Row, 10).Value = ws.Cells(i - 1, 6).Value - Year_Open
                    
                    ' write the percentage change
                    If Year_Open <> 0 Then
                        ws.Cells(Summary_Table_Row, 11).Value = 100 * (ws.Cells(i - 1, 6).Value - Year_Open) / Year_Open
                                        Else
                        'if the open value is 0, set the growth to 0
                        ws.Cells(Summary_Table_Row, 11).Value = 0
                    End If

                End If
                
                ' move the summary table row to the next row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' add the ticker symbol to the summary table
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                            
                ' since we've sorted by ticker and date, and we only have one
                ' year of ticker data in each sheet, the first row is the
                ' open for the year
                Year_Open = ws.Cells(i, 3).Value
                                        
                ' start tracking the stock volume from 0,
                ' this is the first row, so the total will equal the volume
                Stock_Volume = ws.Cells(i, 7).Value
    
            Else
                
                ' still summarizing the same stock, so keep adding the volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        ' write column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Value"
        ws.Cells(1, 16).Value = "Ticker"
                
        ' find the greatest % increase/% decrease/volume table
        ' using max/min on the columns is faster than using a tracker variable in the row loop
        ws.Cells(2, 15).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row - 1))
        ws.Cells(3, 15).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row - 1))
        ws.Cells(4, 15).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row - 1))
        
        ' use match to find the first ticker that matches the greatest % increase/decrease and greatest volume
        ' this is much faster than using 3 tracker variables in the row loop
        
        ' greatest % increase
        ws.Cells(2, 16).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(2, 15).Value, ws.Range("K1:K" & Summary_Table_Row), 0), 9).Value
        
        ' greatest % decrease
        ws.Cells(3, 16).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(3, 15).Value, ws.Range("K1:K" & Summary_Table_Row), 0), 9).Value
        
        ' greatest volume
        ws.Cells(4, 16).Value = ws.Cells(Application.WorksheetFunction.Match(ws.Cells(4, 15).Value, ws.Range("L1:L" & Summary_Table_Row), 0), 9).Value
        
        ' delete any previous formatting from the cells - just in case
        ws.Range("K2:K" & Summary_Table_Row).FormatConditions.Delete
        
        ' apply formatting to summary table: green when % >= 0%, red < %0
        ' first rule: red < 0%
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions(1).Interior.Color = vbRed
        
        ' second rule: green >= 0%
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions(2).Interior.Color = vbGreen

        ' adjust column widths
        ws.Columns("I:P").EntireColumn.AutoFit
        
    Next ws
        
    MsgBox "Done"

End Sub


