VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockData()

' create a script that loops through all the stocks for each quarter and outputs the following information


' vairables

    'define active worksheet
    Dim ws As Worksheet

    ' variable for ticker symbol
    Dim Ticker_sym As String

    ' variable for holding Qtr opening stock price & closing stock price
    Dim Open_Price As Double
    Dim Close_Price As Double

    ' vairable for Stock Quarterly Change
    Dim Quarter_Change As Double

    ' variable for percent change
    Dim Percent_Change As Double

    ' variable for Total Stock Volume
    Dim Stock_Vol As Double

    ' variable for counter
    Dim i As Long

    ' Variable for tracking ticker
    Dim Ticker_Summary_Row As Long
    Dim Ticker_First_Row As Long


    'variable for Greatest % Increase, Decrease, ticker symbols and volume
    Dim Max_Percent As Double
    Dim Max_Ticker As String
    Dim Min_Percent As Double
    Dim Min_Ticker As String
    Dim Max_Volume As Double
    Dim Max_Vol_Ticker As String


' Loop through all worksheets
For Each ws In ThisWorkbook.Worksheets


'Resets
    ' Track each ticker
    Ticker_Summary_Row = 2
    Ticker_First_Row = 2

    ' Set Stock Volume to 0
    Stock_Vol = 0
    
    ' Set Max and min percent values to lowest opposing values
    Max_Percent = -999999
    Min_Percent = 999999

    ' Set Greatest Stock Volume to 0
    Max_Volume = 0


' Headers and Labels

    ' name Cell I1 "Ticker"
    ws.Cells(1, 9).Value = "Ticker"

    ' name Cell J1 "Quarterly Change"
    ws.Cells(1, 10).Value = "Quarterly Change"

    ' name Cell I1 "Percent Change"
    ws.Cells(1, 11).Value = "Percent Change"

    ' name Cell L1 "Total Stock Volume"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' name Cell P1 "Ticker"
    ws.Cells(1, 16).Value = "Ticker"

    ' name Cell Q1 "Value"
    ws.Cells(1, 17).Value = "Value"

    ' name Cell 02 "Greatest % Increase"
    ws.Cells(2, 15).Value = "Greatest % Increase"

    ' name Cell 03 "Greatest % Decrease"
    ws.Cells(3, 15).Value = "Greatest % Decrease"

    ' name Cell 05 "Greatest Total Volume"
    ws.Cells(4, 15).Value = "Greatest Total Volume"




' loop through all tickers
For i = 2 To 96001


' Calculate Total stock volume for each Ticker - part of number 4
Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value


' 1) the ticker symbol
    
    ' check to see if we are still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ' set Ticker symbol name
    Ticker_sym = ws.Cells(i, 1).Value

    ' Print Ticker symbol name in column I
    ws.Cells(Ticker_Summary_Row, 9).Value = Ticker_sym



' 2) Quarterly Change - opening to closing price at end of quarter

    ' Set Open Price for Ticker Symbol
    Open_Price = ws.Cells(Ticker_First_Row, 3).Value

    ' Set Close Price for Ticker Symbol
    Close_Price = ws.Cells(i, 6).Value

    ' Calculate Quarterly change
    Quarter_Change = Close_Price - Open_Price

    ' Print Ticker Quarterly Change in column J
    ws.Cells(Ticker_Summary_Row, 10).Value = Quarter_Change



' 3) Percent Change - opening to closing price at end of quarter

    ' Calculate Percent change
    Percent_Change = (Quarter_Change / Open_Price)

    ' Print Ticker Quarterly Change in column K
    ws.Cells(Ticker_Summary_Row, 11).Value = Percent_Change

    ' Format values as %
    ws.Cells(Ticker_Summary_Row, 11).NumberFormat = "0.00%"



' 4) Total stock volume of each stock
    ' Print total stock volume in column L
     ws.Cells(Ticker_Summary_Row, 12).Value = Stock_Vol



' 5) Check for Greatest % Increase, Decrease and volume
    ' Check % Increase
    If Percent_Change > Max_Percent Then
    Max_Percent = Percent_Change
    Max_Ticker = Ticker_sym
    End If

    ' Print greatest % increase
    ws.Cells(2, 16).Value = Max_Ticker
    ws.Cells(2, 17).Value = Max_Percent
    ws.Cells(2, 17).NumberFormat = "0.00%"

    ' Check % Decrease
    If Percent_Change < Min_Percent Then
    Min_Percent = Percent_Change
    Min_Ticker = Ticker_sym
    End If

    ' Print greatest % decrease
    ws.Cells(3, 16).Value = Min_Ticker
    ws.Cells(3, 17).Value = Min_Percent
    ws.Cells(3, 17).NumberFormat = "0.00%"

    ' Check greatest volume
    If Stock_Vol > Max_Volume Then
    Max_Volume = Stock_Vol
    Max_Vol_Ticker = Ticker_sym
    End If
    
    'Print greatest volume
    ws.Cells(4, 16).Value = Max_Vol_Ticker
    ws.Cells(4, 17).Value = Max_Volume



' Reset Total Stock Volume for next ticker
Stock_Vol = 0
    
' Add one to Ticker_summary_row
Ticker_Summary_Row = Ticker_Summary_Row + 1

' Set first row for next ticker
Ticker_First_Row = i + 1




End If

Next i


' conditional formating for quaterly change
Call ConFormat(ws)

Next ws

End Sub


Sub ConFormat(ws As Worksheet)

' Script for conditional formatting - Quarterly Change and Percent Change per instructions


'variable for Quarterly change column and percent change column
Dim Qtr_col As Range
Dim Per_col As Range

'variable for last row
Dim lastrow As Long

'variable for holding conditional formatting
Dim Cond1 As FormatCondition
Dim Cond2 As FormatCondition

' Find last row
lastrow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row


' Column for formatting - J2 to last J row & K2 to last K row
Set Qtr_col = ws.Range("J2:J" & lastrow)
Set Per_col = ws.Range("K2:K" & lastrow)

' clear existing conditions
Qtr_col.FormatConditions.Delete
Per_col.FormatConditions.Delete


' Apply green for values >0 to columns J and K
    ' J
    Set Cond1 = Qtr_col.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
    Cond1.Interior.Color = RGB(144, 238, 144)

    ' K
    Set Cond1 = Per_col.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
    Cond1.Interior.Color = RGB(144, 238, 144)


' Apply red values <0 to columns J and K
    ' J
    Set Cond2 = Qtr_col.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
    Cond2.Interior.Color = RGB(255, 99, 71)
    
    ' K
    Set Cond2 = Per_col.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
    Cond2.Interior.Color = RGB(255, 99, 71)

End Sub

