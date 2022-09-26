Attribute VB_Name = "Module1"
Sub stock_analysis():

Dim ws As Worksheet

'Walk through each worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate


'declare variables
Dim j As Long
Dim rcount As Long
Dim lrow As Long
Dim sopen As Double
Dim sclose As Double
Dim soutput As Double
Dim volume As LongLong

lrow = Cells(Rows.Count, 1).End(xlUp).Row
rcount = 2
r_count = 2
RowCount = 2
volume = 0
'create column labels
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Volume"

'for loop that goes through stocks
For j = 2 To lrow
   'list the ticker symbols
If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
    Cells(rcount, 10).Value = Cells(j, 1).Value
    rcount = rcount + 1
End If

 'yearly change from opening price at beginning to closing price
If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
sopen = Cells(j, 3).Value
ElseIf Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
sclose = Cells(j, 6).Value
Cells(r_count, 11) = sclose - sopen

'percent change from opening price to closing price
Cells(r_count, 12) = (Format(((sclose - sopen) / sopen) * 100, "#.00") & " %")
r_count = r_count + 1
End If

'total stock volume
If Cells(j, 1).Value = Cells(j + 1, 1).Value Then
    volume = Cells(j, 7).Value + volume
ElseIf Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
Cells(RowCount, 13).Value = volume
RowCount = RowCount + 1
End If



Next j
For j = 2 To lrow

'conditional formatting for yearly change
If Cells(j, 11).Value > 0 Then
Cells(j, 11).Interior.ColorIndex = 4
ElseIf Cells(j, 11).Value < 0 Then
Cells(j, 11).Interior.ColorIndex = 3
End If
Next j

Next ws

End Sub
