# VBA-scrpting-

Sub total()

'Set Variables to keep track on worksheet

    Dim ticker As String
    Dim percentChange As Double
    Dim yearlyChange As Double
    Dim Totalstock As Double
    Dim stock_open As Double
    Dim stock_close As Double
    Dim Lastrow As Double
    Dim num_ticker As String

'Start Loop
    For Each ws In Worksheets
    

'Finds the last row of each worksheet

    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Assign title to columns

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"

    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"

    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    
    Totalstock = 0
    numb_ticker = 0
    year_open = 0
    

    'Loops through worksheet excepy ticker column
        For i = 2 To Lastrow
    
        stock_open = ws.Cells(i, 3).Value
        stock_close = ws.Cells(i, 6).Value
        Totalstock = ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
    
    'Grabs opening price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                year_open = ws.Cells(i, 3).Value
                
              End If
    
    'Conditional for ticker
            If ws.Cells(i + 1, 1) <> ticker Then
                numb_ticker = numb_ticker + 1
                
        
            End If
        
    'Takes the total volume for each row
        Totalstock = Totalstock + ws.Cells(i, 7).Value
    
    
    
    'Calculates price change for the year
        yearlyChange = stock_close - stock_open
     
    
    'Calculates percentage change
            If stock_open = 0 Then
                percentChange = 0
                
                Else
                    percentChange = (yearlyChange) / stock_open
     
            End If
     
    'Assigning values to appropriate column
    
    ws.Cells(numb_ticker + 1, 9).Value = ticker
    ws.Cells(numb_ticker + 1, 10).Value = yearlyChange
    ws.Cells(numb_ticker + 1, 11).Value = Format(percentChange, "Percent")
    ws.Cells(numb_ticker + 1, 12).Value = Totalstock
    
    


        
    'For any value greater than zero, the color green is assigned to that cell
    
    If yearlyChange > 0 Then
        ws.Cells(numb_ticker + 1, 10).Interior.ColorIndex = 4

    'For any value less than zero, the color red is assigned to that cell
    
    ElseIf yearlyChange < 0 Then
            ws.Cells(numb_ticker + 1, 10).Interior.ColorIndex = 3


    End If

    
    
        
    

    
Next i

Next ws




End Sub
