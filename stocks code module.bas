Attribute VB_Name = "Module1"
Sub stocks()
  Dim WorksheetName As String
  Dim ticker As String
  Dim nextTicker As String
  ticker = ""
  nextTicker = ""
  Dim openPrice As Double
  Dim closePrice As Double
  Dim dailyVol As Double
  Dim yearlyVol As Double
  Dim yearlyChange As Double
  Dim percentChange As Double
  Dim tickerRow As Integer
  Dim firstLine As Boolean
  Dim greatestPercentIncrease As Double
  Dim greatestPercentDecrease As Double
  Dim greatestVolume As Double
  Dim greatestPercentIncreaseTicker As String
  Dim greatestPercentDecreaseTicker As String
  Dim greatestVolumeTicker As String
  
  
    
    For Each ws In Worksheets
        ws.Activate
        tickerRow = 1
        yearlyVol = 0
        dailyVol = 0
        firstLine = True
        percentChange = 0
        yearlyChange = 0
        openPrice = 0
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        greatestVolume = 0
                     

        ' Determine the Last Row for stocks
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '' Determine the Last Column Number
        'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        For i = 2 To LastRow
          
          If firstLine = True Then
            openPrice = Cells(i, 3).Value
          End If
          
          closePrice = Cells(i, 6).Value
          dailyVol = Cells(i, 7).Value
          yearlyVol = dailyVol + yearlyVol
          ticker = Cells(i, 1).Value
          nextTicker = Cells(i + 1, 1).Value
          firstLine = False

          If ticker <> nextTicker Then
            'print ticker, yearly change, percent change, yearly volume
            Cells(1 + tickerRow, 9).Value = ticker
            yearlyChange = closePrice - openPrice
            If openPrice <> 0 Then
              percentChange = yearlyChange / openPrice
            Else
              percentChange = 0
            End If
            Cells(1 + tickerRow, 10).Value = yearlyChange
            Cells(1 + tickerRow, 11).Value = FormatPercent(percentChange, 2)
            Cells(1 + tickerRow, 12).Value = yearlyVol
            
            'color yearlyChange cell
            If yearlyChange > 0 Then
              Cells(1 + tickerRow, 10).Interior.ColorIndex = 4 'green if positive
            ElseIf yearlyChange < 0 Then
              Cells(1 + tickerRow, 10).Interior.ColorIndex = 3 'red if negative
            End If
            
                       
            'increase tickerRow Counter
            tickerRow = tickerRow + 1
            firstLine = True
            
            'reset volume
            dailyVol = 0
            yearlyVol = 0
          End If
          
        Next i
        
        'determine last row for output data in Column I
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'challenge calculations
        For i = 2 To LastRow
          ticker = Cells(i, 9).Value
          percentChange = Cells(i, 11).Value
          yearlyVolume = Cells(i, 12).Value
          If yearlyVolume > greatestVolume Then
            greatestVolume = yearlyVolume
            greatestVolumeTicker = ticker
          End If
          If percentChange > 0 And percentChange > greatestPercentIncrease Then
            greatestPercentIncrease = percentChange
            greatestPercentIncreaseTicker = ticker
          End If
          If percentChange < 0 And percentChange < greatestPercentDecrease Then
            greatestPercentDecrease = percentChange
            greatestPercentDecreaseTicker = ticker
          End If
        Next i
        
        'print results of second calculations
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P2").Value = greatestPercentIncreaseTicker
        Range("Q2").Value = FormatPercent(greatestPercentIncrease, 2)
        
        Range("P3").Value = greatestPercentDecreaseTicker
        Range("Q3").Value = FormatPercent(greatestPercentDecrease, 2)
        
        Range("P4").Value = greatestVolumeTicker
        Range("Q4").Value = greatestVolume
        
        
        
    Next ws

    MsgBox ("Complete")


End Sub


