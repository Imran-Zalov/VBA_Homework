Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

'Headers based on the given data

Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly change"
Range("L1").Value = "Percent change"
Range("M1").Value = "Total Stock Volume"


'Organazing columns sizes based on the capacity

Columns("J:K").ColumnWidth = 15
Columns("L:M").ColumnWidth = 20
Columns("O").ColumnWidth = 20
Columns("Q").ColumnWidth = 15

Range("J1:M1").Font.Bold = True
Range("O2:O4").Font.Bold = True
Range("P1:Q1").Font.Bold = True

'Placing our data in center

Range("P1:Q1").HorizontalAlignment = xlCenter


End Sub

Sub ForWSh()

'Dim i As Integer


'Looping data in order to apply all changes to each worksheet

'For i = 1 To Worksheets.Count

 '       Worksheets(i).Select
        
  '      Multiple_year_stock_data
   '     yearlyChange
    '    totalStockVolume
     '   other_info
'Next i

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
        ws.Select
    
    Multiple_year_stock_data
    yearlyChange
    totalStockVolume
    other_info
    
    End If
    
Next ws

        Worksheets("Home").Select
        MsgBox "All required data has been calculated and ready to show."
        
End Sub
Sub yearlyChange()

'Calculation yearly change
'clPrice = close price
'opPrice = open price

Dim opPrice, clPrice, yearlyChange, percentChange, lastrow As Double
Dim ticker As String
Dim numberTickers As Integer

'accepting zero as default

opPrice = 0
yearlyChange = 0
percentChange = 0
ticker = ""
numberTickers = 0

'Below function will give us last row number of the column A

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through opening price data

For i = 2 To lastrow
    
    ticker = Cells(i, 1).Value
    
    If opPrice = 0 Then
    opPrice = Cells(i, 3).Value
    End If

'Loop through closing price data

If Cells(i + 1, 1).Value <> ticker Then

'numberTicker changes each time when ticker is different from -- ticker = Cells(i, 1).Value

    numberTickers = numberTickers + 1
    Cells(numberTickers + 1, 10) = ticker
    
    clPrice = Cells(i, 6)

'Finding the differences between  opening / closing price of each ticker seperately

    yearlyChange = clPrice - opPrice
    
    Cells(numberTickers + 1, 11).Value = yearlyChange
    
    If opPrice = 0 Then
        percentChange = Format(0, "Percent")
    Else
        percentChange = Round(yearlyChange / opPrice, 4)
    End If
    
    Cells(numberTickers + 1, 12).Value = Format(percentChange, "Percent")
    
    'If condition matches requirement, change the color based on result
    
    If yearlyChange > 0 Then
    Cells(numberTickers + 1, 11).Interior.ColorIndex = 4
    
    ElseIf yearlyChange < 0 Then
        Cells(numberTickers + 1, 11).Interior.ColorIndex = 3
        
  
          
        End If
        
        opPrice = Cells(i + 1, 3).Value
        End If


    Next i
    
    
End Sub


Sub totalStockVolume()

'Assign variables

Dim ticker As String
Dim numberTickers As Integer
Dim lastrow, totalStockVolume As Double

'accepting zero and"" as default

ticker = ""
numberTickers = 0
totalStockVolume = 0

'Below function will give us last row number of the column A

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through the total stock volume

For i = 2 To lastrow

ticker = Cells(i, 1).Value

totalStockVolume = totalStockVolume + Cells(i, 7).Value

If Cells(i + 1, 1).Value <> ticker Then

'numberTicker changes each time when ticker is different from -- ticker = Cells(i, 1).Value

    numberTickers = numberTickers + 1
    Cells(numberTickers + 1, 10) = ticker
    
    Cells(numberTickers + 1, 13).Value = totalStockVolume
    
    totalStockVolume = 0

 End If
 

 Next i

End Sub

Sub other_info()

'Assigning variables to required fields

Dim i As Integer
Dim great As Double
Dim ticker1 As String
Dim ticker2 As String
Dim ticker3 As String
Dim smallest As Double
Dim total_volume As Double
Dim lastCell As Integer

'Assigning to ranges

Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

'Below function will give us last row number of the column J

lastCell = Cells(Rows.Count, "J").End(xlUp).Row
great = Cells(2, 12).Value
smallest = Cells(2, 12).Value
great_volume = Cells(2, 13).Value

'Looping through the data based on percentage change max/min and total stock volume

For i = 2 To lastCell

If Cells(i, 12).Value > great Then
 
great = Cells(i, 12).Value
ticker1 = Cells(i, 10).Value
 
  
End If

If Cells(i, 12).Value < smallest Then

smallest = Cells(i, 12).Value
ticker2 = Cells(i, 10).Value

End If

If Cells(i, 13).Value > great_volume Then

great_volume = Cells(i, 13).Value
ticker3 = Cells(i, 10).Value

End If


Next i

'Passing the data to appropriate places after finding reqired value

Cells(2, 16).Value = ticker1

'Changing the format to the percent

Cells(2, 17).Value = Format(great, "Percent")
 
Cells(3, 16).Value = ticker2
Cells(3, 17).Value = Format(smallest, "Percent")

Cells(4, 16).Value = ticker3

Cells(4, 17).Value = great_volume

'Changing format of the value to scientific value for four decimal

Cells(4, 17).NumberFormat = "0.0000E+00"


 End Sub
 
Sub Raw_data()

Dim ws As Worksheet

'Below loop will clean all calculated data and will keep the only raw data

For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
        ws.Select
        
   'in this case data was gathered in J-Q column, and we will clean only indicated columns
   
    Range("J:Q").Clear
    
    End If
    
Next ws

'Whenever file is opened it will first go "Home" sheet

Worksheets("Home").Select

'Below message will pop up as soon as we click on Raw_data button

MsgBox "All calculated data was erased and only raw data remained."
End Sub

