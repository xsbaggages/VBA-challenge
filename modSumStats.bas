Attribute VB_Name = "Module1"
Sub CallSummaryStats()

    ' To find out the number of sheets in a workbook
    For i = 1 To Sheets.Count
    
        ' To call the method that perform the summary stats
        SummaryStats Sheets(i).Name
    
    Next i
    
    MsgBox "Done"
    
End Sub
' This is doing all the hardwork. It takes in 1 parameter which is the sheet name.
Sub SummaryStats(sheetname As String)

  ' Specify column position
  Dim columnTicker As Integer
  Dim columnDate As Integer
  Dim columnOpen As Integer
  Dim columnHigh As Integer
  Dim columnLow As Integer
  Dim columnClose As Integer
  Dim columnVol As Integer
  
  ' variables for calculations
  Dim Ticker As String
  Dim firstOpen As Double
  Dim lastClose As Double
  Dim YearlyChange As Double
  Dim PercentChange As Double
  Dim TotalStock As Double
  Dim GreatInc As Double
  Dim GreatDec As Double
  Dim GreatTotVol As Double
  Dim GreatIncTic As String
  Dim GreatDecTic As String
  Dim GreatTotVolTic As String
   
  ' Defining data columns
  columnTicker = 1
  columnDate = 2
  columnOpen = 3
  columnHigh = 4
  columnLow = 5
  columnClose = 6
  columnVol = 7

  
  'Defining worksheet and setting currentSheet
  Dim currentSheet As Worksheet
  Set currentSheet = Worksheets(sheetname)
   
   'Define last row in the dataset. So For loop does not have to re calculate last row every loop
   Dim lastrow As Long
    lastrow = currentSheet.Cells(1, "A").End(xlDown).Row
    
    ' Declare the coordinates of the summary stats
    Dim sumTickerCol As Integer
    Dim sumTickerRow As Integer
    Dim sumPriceDifCol As Integer
    Dim sumPerChangeCol As Integer
    Dim sumTotStockVolCol As Integer
    Dim BonusHeaderCol As Integer
    Dim BonusTickerCol As Integer
    Dim BonusValueCol As Integer

    ' Defining summary columns
    sumTickerCol = 9
    sumPriceDifCol = 10 'yearly change
    sumPerChangeCol = 11
    sumTotStockVolCol = 12
    sumTickerRow = 2
    BonusHeaderCol = 14
    BonusTickerCol = 15
    BonusValueCol = 16
    
    ' Adding headers
    currentSheet.Cells(1, sumTickerCol).Value = "Ticker"
    currentSheet.Cells(1, sumPriceDifCol).Value = "Yearly Change"
    currentSheet.Cells(1, sumPerChangeCol).Value = "Percent Change"
    currentSheet.Cells(1, sumTotStockVolCol).Value = "Total Stock Volume"
    currentSheet.Cells(1, BonusTickerCol).Value = "Ticker"
    currentSheet.Cells(1, BonusValueCol).Value = "Value"
    currentSheet.Cells(2, BonusHeaderCol).Value = "Greatest % Increase"
    currentSheet.Cells(3, BonusHeaderCol).Value = "Greatest % Decrease"
    currentSheet.Cells(4, BonusHeaderCol).Value = "Greatest Total Volume"
    
    ' Percentage formatting
    currentSheet.Columns(sumPerChangeCol).NumberFormat = "0.00%" ' for entire column
    currentSheet.Cells(2, BonusValueCol).NumberFormat = "0.00%"  ' for cell only
    currentSheet.Cells(3, BonusValueCol).NumberFormat = "0.00%"  ' for cell only
    
    
   
    ' This is to assign opening price of first ticker
    firstOpen = currentSheet.Cells(2, columnOpen).Value
    
    
  ' Loop through rows in the column
  Dim i As Long
  For i = 2 To lastrow

        ' Assigning a variable for incrementing TotalStock value.
        TotalStock = TotalStock + (currentSheet.Cells(i, columnVol).Value)

        ' Searches for when the value of the next cell is different than that of the current cell
        If currentSheet.Cells(i + 1, columnTicker).Value <> currentSheet.Cells(i, columnTicker).Value Then
            
                    
            ' Assigning the value of lastClose of current ticker
            lastClose = currentSheet.Cells(i, columnClose).Value
            
            ' Assigning ticker symbol
            Ticker = currentSheet.Cells(i, columnTicker).Value
            
            ' To calculate yearly price diff (lastClose minus previously assigned firstOpen) before the previous
            ' firstOpen is overwritten
            YearlyChange = lastClose - firstOpen
            
            ' To calculate percent change in price
            If firstOpen = 0 Then ' To handle 0 values in firstopen
                PercentChange = 0
            Else
                PercentChange = YearlyChange / firstOpen
            End If
            
            ' Assigns value of opening price of next ticker
            firstOpen = currentSheet.Cells(i + 1, columnOpen).Value
            
            ' Assigning the cell value to input ticker of current row
            currentSheet.Cells(sumTickerRow, sumTickerCol).Value = Ticker
            
            ' Assigning the cell value for YearlyChange
            currentSheet.Cells(sumTickerRow, sumPriceDifCol).Value = YearlyChange
            
            ' Conditional formatting of yearly change
            If YearlyChange > 0 Then
                currentSheet.Cells(sumTickerRow, sumPriceDifCol).Interior.Color = vbGreen
            ElseIf YearlyChange < 0 Then
                currentSheet.Cells(sumTickerRow, sumPriceDifCol).Interior.Color = vbRed
            End If
            
            ' Assigning the cell value for PercentChange
            currentSheet.Cells(sumTickerRow, sumPerChangeCol).Value = PercentChange
            
            ' Assigning value for total stock volume
            currentSheet.Cells(sumTickerRow, sumTotStockVolCol).Value = TotalStock
            
            ' To find out greatest % increase
            If PercentChange > GreatInc Then
                GreatInc = PercentChange
                GreatIncTic = Ticker
            End If
            
            ' To find out greatest % decrease
             If PercentChange < GreatDec Then
                GreatDec = PercentChange
                GreatDecTic = Ticker
            End If
            
            ' To find out greatest stock volume
            If TotalStock > GreatTotVol Then
                GreatTotVol = TotalStock
                GreatTotVolTic = Ticker
            End If
            
            ' To reset TotalStock for next ticker
            TotalStock = 0
            
            ' To select the next empty row
            sumTickerRow = sumTickerRow + 1
            
        End If

    Next i
    
    ' Assigning the greatest variable values to cells
    currentSheet.Cells(2, BonusTickerCol).Value = GreatIncTic
    currentSheet.Cells(2, BonusValueCol).Value = GreatInc
    
    currentSheet.Cells(3, BonusTickerCol).Value = GreatDecTic
    currentSheet.Cells(3, BonusValueCol).Value = GreatDec
    
    currentSheet.Cells(4, BonusTickerCol).Value = GreatTotVolTic
    currentSheet.Cells(4, BonusValueCol).Value = GreatTotVol

    ' To auto fit all columns
    currentSheet.Cells.EntireColumn.AutoFit
    
    'Debug.Print sheetname, "done", Now

End Sub

