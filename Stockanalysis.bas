Sub stockanalysis()

'Set an initial variable for holding the brand name
Dim Ticker_Name As String

'Set an initial variable for holding total stock volume by ticket
Dim Total_stock_Volume As Double

'Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
  
'Declare Hold variables
Dim Percentage_Change as Double
Dim Stock_Open As Double
Dim Stock_Close As Double

'loop thru each worksheet
For Each ws In Worksheets
    'initialize variables  
    Total_stock_Volume = 0
    Summary_Table_Row = 2
    Greatest_Percent_Decrease = 0
    Greatest_Percent_Increast = 0
    Greatest_total_Volume = 0

    Stock_Open = ws.Range("C2").Value
    'get the number of rows in the spreadsheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Add headerings for column I and J
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest%Increase"
    ws.Range("O3").Value = "Greatest%Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q2").Value = 0
    ws.Range("Q3").Value = 0
    ws.Range("Q4").Value = 0

    'Loop thru all tickers
    For i = 2 To LastRow
        'Check and accumulate totals for same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
            'save ticker
            Ticker_Name = ws.Cells(i, 1).Value
       
            'add to total volume
            Total_stock_Volume = Total_stock_Volume + ws.Cells(i, 7).Value
       
            'store close price
            Stock_Close = ws.Cells(i, 6).Value
       
            ' Print Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
            ' Print Total stock volume to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Stock_Close - Stock_Open
      
            'color code cell based on value
            If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
      
            ' Print Percentage change
            ' handle error condition for divide by zero situation, default percentage change to 0
            On Error Resume Next
            Percentage_Change = (Stock_Close - Stock_Open) / Stock_Open
            If Err.Number <> 0 Then
                ws.Range("K" & Summary_Table_Row).Value = 0
                Percentage_Change = 0
            Else
                ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
            End If
      
            'Format to percentage    
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            ' Print Total stock volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_stock_Volume

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Check % increase
            If Percentage_Change > ws.Range("Q2").Value  Then
                ws.Range("Q2").Value = Percentage_Change
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = Ticker_Name
            End If

            ' Check % Decrease
            If Percentage_Change < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = Percentage_Change
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = Ticker_Name
            End If

            ' check greatest total volume
            If Total_stock_Volume > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = Total_stock_Volume
                ws.Range("P4").Value = Ticker_Name
            End If

            ' Reset work variables
            Total_stock_Volume = 0
            Stock_Close = 0
            Stock_Open = ws.Cells(i + 1, 3).Value
            Percentage_Change = 0
        Else
            'add to total volume
            Total_stock_Volume = Total_stock_Volume + ws.Cells(i, 7).Value
        End If
    
    Next i
Next ws

End Sub

