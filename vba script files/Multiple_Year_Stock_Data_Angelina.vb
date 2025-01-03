'Create a script that loops through all the stocks for each quarter and outputs the following information:

'1-The ticker symbol

'2-Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

'3-The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

'4-The total stock volume of the stock. The result should match the following image:

'5-Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.

'note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'-------------------------------------

Sub MultipleYearStock()

    'create variables to store data
    'double = decimal/numerical value , integer=whole number , string = text , long = interger w/ long range
    Dim ws As Worksheet
    Dim open_price, close_price, qtly_change, percent_change, volume As Double
    Dim row As Long
    Dim ticker_column As Integer ' column A
    Dim LastRow As Long
    Dim change_lastrow As Long
    Dim ticker As String
    Dim i, j, k As Integer
'    Dim greatest_volume as double
'    Dim found_cell As Range
'    Dim rng As Range
'    Dim found_row As Integer

    ' --------------------------------------------
    ' LOOP THROUGH ALL WORKSHEETS
    ' --------------------------------------------
        'set ticker column to 1, column A
        ticker_column = 1
            
       For Each ws In ActiveWorkbook.Worksheets
        
        'find the last row of data in column A
         LastRow = ws.Cells(Rows.Count, "A").End(xlUp).row
    
        'create headers for the output in columns I to L
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'sets location of starting row for output to row 2 for the output
        row = 2
          
        'set location of inital open_price (row 2, column c)
        open_price = ws.Cells(2, ticker_column + 2).Value
        

        ' --------------------------------------------
        ' LOOP THROUGH EACH ROW TO PROCESS DATA
        ' --------------------------------------------

        'begins loop at row 2 and runs until the last row
        For i = 2 To LastRow
        
         'checks if ticker changes in the next row
            If ws.Cells(i + 1, ticker_column).Value <> ws.Cells(i, ticker_column).Value Then
            
                'if value changes, then
                ticker = ws.Cells(i, ticker_column).Value 'store ticker value
                ws.Cells(row, ticker_column + 8).Value = ticker '& print the ticker symbol to the output in column I
                                                       
               'if value changes, then
                volume = volume + ws.Cells(i, ticker_column + 6).Value 'calculate total volume for ticker, column G
                ws.Cells(row, ticker_column + 11).Value = volume '& print total volumne to column L after ticker changes
                
                'resets the volume to 0 for next ticker, set before row processing loops starts to reset for each new ticker
                volume = 0
             
                'if value changes, then
                close_price = ws.Cells(i, ticker_column + 5).Value 'closing price at the point of the loop
                qtly_change = close_price - open_price 'calculate qtly change
                percent_change = (qtly_change / open_price) ' calculate percent change
                
                'output results
                ws.Cells(row, ticker_column + 9).Value = qtly_change ' output qtly change in J
                ws.Cells(row, ticker_column + 10).Value = percent_change ' output percent change in K
                ws.Cells(row, ticker_column + 10).NumberFormat = "0.00%" ' change format in K to %
          
                'increments row variable by 1, moving to the next row
                row = row + 1
                
                'reset open_price for next ticker
                open_price = ws.Cells(i + 1, ticker_column + 2)
                
            Else
                'if value hasn't changed, then continue adding volume of current row to running total until reaches a different ticker
                volume = volume + ws.Cells(i, ticker_column + 6).Value
                
            'close the If/Else statement
            End If
     'call the next iteration
       Next i
        
        ' --------------------------------------------
        ' HIGHLIGHT POSITIVE AND NEGATIVE CHANGES W/ CONDITIONAL FORMATTING
        ' --------------------------------------------
        
        'finds the last row of column J, quarterly change
        change_lastrow = ws.Cells(Rows.Count, "J").End(xlUp).row
        
        'loop through rows to apply conditional formatting based on percent change.
        For j = 2 To change_lastrow
            If (ws.Cells(j, 10).Value >= 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 10 'green for positive=10
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3 'red for negative =3
            End If
        Next j
        
        
        ' --------------------------------------------
        ' IDENTIFY GREATEST INCREASE, DECREASE
        ' --------------------------------------------

            
'       greatest_volume = 0
'       greatest_volume_ticker = ""
'       ticker = ""
            
        'add headers for greatest increase/decrease
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'begins loops starting at row 2
        For k = 2 To change_lastrow
        'find the greatest percent increase by max
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & change_lastrow)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value 'update output. the code structure -- left side = cell you are updating, right side = cell you are pulling data from
               ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
               ws.Cells(2, 17).NumberFormat = "0.00%"
                
            'find the greatest percent decrease by min
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & change_lastrow)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value 'ticker with greatest % increase
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"

            'find the greatest total volume
            ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & change_lastrow)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value 'ticker w/ greatest total volume
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value ' return value to output

            End If
            Next k
            

                
        'find the max value in column L (Total Stock Volume)
'         greatest_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & change_lastrow))


        'set the range to search for the greatest total volume
'        Set rng = ws.Columns("L")

        'set search for the cell that contains the greatest total volume
'        Set found_cell = rng.Find(What:=greatest_volume, LookIn:=xlValues, LookAt:=xlWhole)

        'check if a matching cell was found
'        If Not found_cell Is Nothing Then
'            found_row = found_cell.row 'get the row of the found cell
            
            'output the ticker with the greatest total volume and the volume itself
'            ws.Cells(4, 16).Value = ws.Cells(found_row, 9).Value 'ticker in column 9
'           ws.Cells(4, 17).Value = greatest_volume 'greatest total volume in column L
'           ws.Cells(4, 17).NumberFormat = "General"
        
    
'        End If
        
        'auto fit columns I to Q for the current worksheet
        ws.Range("I:Q").EntireColumn.AutoFit
        
    'move to next worksheet
    Next ws

'message box notification when script is complete
MsgBox ("Multiple Year Stock Data Calculation Complete!")


End Sub


'resets all sheets to pre-analysis state
Sub RestButton()
    Dim i As Integer
    
    'loop to cycle through all workbook sheets and delete columns I through Q - This also resets formating
    For i = 1 To Sheets.Count
        With Sheets(i)
            .Columns("I:Q").Delete
        End With
    Next i
End Sub



