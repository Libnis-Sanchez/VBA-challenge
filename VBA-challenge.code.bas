Attribute VB_Name = "Module1"
Sub multiple_year_stock()
    'run through each worksheet
    For Each ws In Worksheets
        'set headers
        ws.Cells(1, 9).value = "Ticker"
        ws.Cells(1, 10).value = "Yearly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Cells(1, 15).value = "Ticker"
        ws.Cells(1, 16).value = "Value"
        ws.Cells(2, 14).value = "Greatest % Increase"
        ws.Cells(3, 14).value = "Greatest % Decrease"
        ws.Cells(4, 14).value = "Greatest Total Volume"
        
        'variables
        Dim WorksheetName As String
        Dim ticker As String
        ticker = " "
        Dim ticker_number As Integer
        ticker_number = 0
        Dim vol As Integer
        vol = 0
        Dim year_open As Double
        year_open = 0
        Dim year_close As Double
        year_close = 0
        Dim yearly_change As Double
        yearly_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim max_increase As Double
        max_increase = 0
        Dim max_decrease As Double
        max_decrease = 0
        Dim ticker_max_increase As Double
        ticker_max_increase = 0
        Dim ticker_max_decrease As Double
        ticker_max_decrease = 0
        Dim max_volume As Double
        max_volume = 0
        Dim ticker_max_volume As String
        Dim value As Double
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim date_ As Date

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'this prevents my overflow error
On Error Resume Next

    'loop
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then

            'find all the values
            ticker = ws.Cells(i, 1).value
            vol = vol + ws.Cells(i, 7).value
            
            ticker_number = ticker_number + 1
            Cells(ticker_number + 1, 9) = ticker
            
            year_open = ws.Cells(i, 3).value
            year_close = ws.Cells(i, 6).value
        
            yearly_change = (year_close - year_open)
            
            Cells(ticker_number + 1, 10).value = yearly_change
            
            If year_close <> 0 Then
                percent_change = (year_close - year_open) / year_close * 100
            End If

            'insert values into summary
            ws.Range("I" & Summary_Table_Row).value = ticker
            ws.Range("J" & Summary_Table_Row).value = yearly_change
            ws.Range("K" & Summary_Table_Row).value = percent_change
            ws.Range("L" & Summary_Table_Row).value = vol
            
            'add 1 to summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
    
            End If
        
            max_increase = Cells(i, 11).value
            ticker_max_increase = Cells(i, 9).value
            max_decrease = Cells(i, 11).value
            ticker_max_decrease = Cells(i, 9).value
            max_volume = Cells(i, 15).value
            ticker_max_volume = Cells(i, 9).value
            
    
            'return highest % increase / decrease
            If Cells(i, 11).value > max_increase Then
                max_increase = Cells(i, 11).value
                ticker_max_increase = ticker
            ElseIf Cells(i, 9).value < max_decrease Then
                max_decrease = percent_change
                ticker_max_decrease = Cells(i, 9).value
            End If
            
            If Cells(i, 12).value > max_volume Then
                max_volume = Cells(i, 12).value
                ticker_max_volume = Cells(i, 9).value
                
                'insert values into summary
            ws.Range("O2").value = ticker_max_increase
            ws.Range("O3").value = ticker_max_decrease
            ws.Range("O4").value = max_volume
            
            End If
            
            
'finish loop
    Next i
    
    
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Columns("Q").NumberFormat = "0.00%"

    'format columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range

    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count

    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g


'move to next worksheet
Next ws
End Sub
