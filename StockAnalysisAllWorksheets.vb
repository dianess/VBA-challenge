Attribute VB_Name = "StockAnalysisAllWorksheets"
Private Sub StockAnalysis()

'Describe variable types
Dim addStockVolume As Double
Dim closingStockPrice As Double
Dim column As Integer
Dim counterForResults As Integer
Dim greatestStockTotal As Double
Dim greatestStockTotalTicker As String
Dim previousStockVolume As Double
Dim openingStockPrice As Double
Dim percentChangeStockPrice As Double
Dim startingStockVolume As Double
Dim symbolNotChanged As String
Dim tickerSymbol As String
Dim yearlyChangeStockPrice As Double

'print top row of summary table information
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'state beginning values
column = 1
counterForResults = 1
previousStockVolume = 0
greatestStockTotal = 0

'set the starting ticker symbol and opening stock price
symbolNotChanged = Range("A2").Value
openingStockPrice = Range("C2").Value

'allows for finding the last row
lastrow = Cells(Rows.Count, column).End(xlUp).Row

'scanning through each row to collect the needed data
For Row = 2 To lastrow + 1
    
    startingStockVolume = Cells(Row, 7).Value
    tickerSymbol = Cells(Row, 1).Value
    
    If tickerSymbol = symbolNotChanged Then
        addStockVolume = startingStockVolume + previousStockVolume
        previousStockVolume = addStockVolume
        symbolNotChanged = Cells(Row, 1).Value
    Else
        'print ticker symbol and total stock volume
        Cells(counterForResults + 1, 12).Value = addStockVolume
        Cells(counterForResults + 1, 9).Value = symbolNotChanged
        
        'while printing these in column L, also keep track to find the greatest total volume and associated ticker
            If greatestStockTotal < addStockVolume Then
            'set new value for greatestStockTotal and it's associated ticker
            greatestStockTotal = addStockVolume
            greatestStockTotalTicker = symbolNotChanged
            End If
            
        'calculate and print yearly change in stock price
        closingStockPrice = Cells(Row - 1, 6).Value
        yearlyChangeStockPrice = closingStockPrice - openingStockPrice 'calculates yearly change
        Cells(counterForResults + 1, 10).Value = yearlyChangeStockPrice 'prints yearly change in stock price
        
        'formats yearly change column cells with green if positive and red with white font if negative
        If yearlyChangeStockPrice >= 0 Then
            Range(Cells(counterForResults + 1, 10), Cells(counterForResults + 1, 10)).Interior.ColorIndex = 4

        Else
            Range(Cells(counterForResults + 1, 10), Cells(counterForResults + 1, 10)).Interior.ColorIndex = 3
            Range(Cells(counterForResults + 1, 10), Cells(counterForResults + 1, 10)).Font.ColorIndex = 2
             
        End If
        
        'calculate and print percent change in stock price
        If openingStockPrice = 0 Then
            Cells(counterForResults + 1, 11).Value = "N/A" 'to avoid meaningless calculation (division by zero)
            
        Else
            percentChangeStockPrice = yearlyChangeStockPrice / openingStockPrice 'calculates percent change
            Cells(counterForResults + 1, 11).Value = percentChangeStockPrice 'prints percent change in stock price
            Cells(counterForResults + 1, 11).NumberFormat = "0.00%" 'creates % number format

        End If
        
        'reset values for next round
        counterForResults = counterForResults + 1
        symbolNotChanged = Cells(Row, 1).Value
        openingStockPrice = Cells(Row, 3).Value
        previousStockVolume = 0
        addStockVolume = 0
        Row = Row - 1
        
    End If
        
Next Row

Cells(4, 17).Value = greatestStockTotal
Cells(4, 17).NumberFormat = "#"
Cells(4, 16).Value = greatestStockTotalTicker

GreatestOnStockAnalysis

'autofit column width to data contained within it
Cells.Columns.AutoFit

End Sub

Private Sub GreatestOnStockAnalysis()
'This sub finds and prints the Greatest % Increase and Greatest % Decrease, both Ticker and Value

Dim maxValue As Double
Dim maxValueTicker As String
Dim minValue As Double
Dim minValueTicker As String

'Set range from which to determine the largest percent value change
Set colK = ActiveSheet.Range("K:K")
Set tickerRange = ActiveSheet.Range("I:I")

'Worksheet function MAX returns the largest value in a range
maxValue = Application.WorksheetFunction.Max(colK)
'Worksheet function Index and Match are used to find the associated ticker symbol for the maximum value
maxValueTicker = Application.WorksheetFunction.Index(tickerRange, Application.WorksheetFunction.Match(maxValue, colK, 0))

'Worksheet function MIN returns the smallest value in a range
minValue = Application.WorksheetFunction.Min(colK)
'Worksheet function Index and Match are used to find the associated ticker symbol for the minimum value
minValueTicker = Application.WorksheetFunction.Index(tickerRange, Application.WorksheetFunction.Match(minValue, colK, 0))

'print the maximum and minimum values with % format and their associated tickers in the cells specified
Cells(2, 17).Value = maxValue
Cells(2, 17).NumberFormat = "0.00%" 'creates % number format
Cells(2, 16).Value = maxValueTicker

Cells(3, 17).Value = minValue
Cells(3, 17).NumberFormat = "0.00%" 'creates % number format
Cells(3, 16).Value = minValueTicker

End Sub

Sub loop_through_all_worksheets()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    StockAnalysis 'run sub StockAnalysis which runs GreatestOnStockAnalysis
Next

starting_ws.Activate 'activate the worksheet that was originally active

End Sub

