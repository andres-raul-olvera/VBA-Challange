VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker()

Dim ws As Worksheet

'Min and Max table
Dim volumerange As Range
Dim percentrange As Range

For Each ws In Worksheets

'ticker symbol
Dim ticker_symbol As String

Dim StartRow As Double
StartRow = 2

'stock volume
Dim stock_volume As Double
stock_volume = 0


'Table Row Tracker
Dim results_table_row As Integer
results_table_row = 2

'Open and Close Data Values
Dim opendatevalue As Double
Dim closedatevalue As Double
Dim changevalue As Double
Dim percentchange As Double

'open date value
opendatevalue = ws.Cells(2, 3)

'new columns
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Cgange"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"

'set range
Set volumerange = ws.Range("L:L")
Set percentrange = ws.Range("K:K")

Dim LR As String
LR = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LR

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker_symbol = ws.Cells(i, 1).Value
    
    closedatevalue = ws.Cells(i, 6).Value
    changevalue = closedatevalue - opendatevalue
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    
    ws.Range("I" & results_table_row).Value = ticker_symbol
    ws.Range("J" & results_table_row).Value = changevalue
     If ws.Range("J" & results_table_row).Value < FormatPercent(0) Then
     
     ws.Range("J" & results_table_row).Interior.ColorIndex = 3
    Else: ws.Range("J" & results_table_row).Interior.ColorIndex = 4
    End If
     
     
    ws.Range("L" & results_table_row).Value = stock_volume
    ws.Range("K" & results_table_row).Value = FormatPercent((changevalue / opendatevalue))
    
    results_table_row = results_table_row + 1
    
    stock_volume = 0
    opendatevalue = ws.Cells(i + 1, 3)
      
      Else
        stock_volume = stock_volume + ws.Cells(i, 7).Value
    
End If


Next i
'Find Max Increase and Ticker
For i = 2 To LR
    If ws.Cells(i, 11) = Application.WorksheetFunction.Max(percentrange) Then
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = ws.Cells(i, 9)
    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(percentrange)
        Exit For
        
        End If
        
        Next i
        
  'Find Max Decrease and Ticker
For i = 2 To LR
    If ws.Cells(i, 11) = Application.WorksheetFunction.Min(percentrange) Then
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = ws.Cells(i, 9)
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(percentrange)
        Exit For
        
        End If
        
        Next i
'Find Max Vol and Ticker
For i = 2 To LR
    If ws.Cells(i, 12) = Application.WorksheetFunction.Max(volumerange) Then
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = ws.Cells(i, 9)
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(volumerange)
        Exit For
        
        End If
        
        Next i
        
  Next ws
  
End Sub

