Sub AlphaTest()

'declare variables

Dim tickerid As String
Dim RowCount As Long
Dim stockvolume As Double
Dim yearopen As Double
Dim yearclose As Double
Dim yearlychange As Double
Dim percentchange As Long
Dim summary_table_row As Integer
Dim tickertotal As Long
Dim i As Long
Dim j As Integer
Dim ws As Worksheet

'run thru all worksheets
For Each ws In Worksheets

' define summary_table and determine last row
summary_table_row = 2

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

' set column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' loop 1 (ticker symbols)

Start = 2
tickertotal = 0
j = 1
    For i = 2 To RowCount
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerid = ws.Cells(i, 1).Value
        
    'Print Ticker Total to Summary Table
        ws.Range("I" & summary_table_row).Value = tickerid
    ' add 1 to summary table
        summary_table_row = summary_table_row + 1
    ' if cell immediatey following is the same ticker value...
    Else
    End If
  ' yearly change + conditional formatting
    yearclose = Cells(i, 6).Value
    yearopen = Cells(i, 3).Value
    yearlychange = yearopen - yearclose
    
' print year change to summary table

    ws.Range("J" & summary_table_row).Value = yearlychange
    summary_table_row = summary_table_row + 1
   
    If yearlychange >= 0 Then
        ws.Range("J" & summary_table_row + 1).Interior.ColorIndex = 4
    ElseIf yearlychange < 0 Then
       ws.Range("J" & summary_table_row + 1).Interior.ColorIndex = 3
        End If
        
' percent change
    yearclose = Cells(i, 6).Value
    yearopen = Cells(i, 3).Value

    If yearclose <> 0 Or yearlychange <> 0 Then
        percentchage = (yearlychange / yearclose) * 100
    ElseIf yearclose = 0 Or yearlychange = 0 Then
        percentchange = 0
        
    ' print percent change to summary table
    ws.Range("K" & summary_table_row).Value = percentchange
    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
    End If
    
' total stock volume
stockvolume = 0
ws.Range("G" & summary_table_row).Value = totalstockvalue
    
    Next i

Next ws

End Sub