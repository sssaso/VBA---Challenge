Attribute VB_Name = "VBA_HW"
Sub Run_All()

Call Uniq_Ticker
Call Yearly_Change
Call Total_Stock_volum
Call Greatest_performer

End Sub

'Part 1: list up All unique tickers from all sheets in sheet(1) column I

Sub Uniq_Ticker()

r = Range("a1").End(xlDown).Row

Range("a2:a" & r).Copy Range("i2")
Range("i:i").RemoveDuplicates 1

Range("i1") = "Tickers"
Range("j1") = "Yearly_Change"
Range("k1") = "Yearly_Change %"
Range("l1") = "Total_Volume"
Range("N1") = ("Greatest Perfomer from " + ActiveSheet.Name)
Range("O1") = "Ticker"
Range("P1") = "Value"

Range("i1:P1").Font.Bold = True
Range("i1:P1").Interior.ColorIndex = 5
Range("i1:P1").Font.ColorIndex = 2
Range("i1:P1").Columns.AutoFit

End Sub

'Part 3 Yearly change $ and %

Sub Yearly_Change()

Dim OpenP, CloseP, yChange As Double
Dim rOpenD, rCloseD As Long

'total count of all tickers +1  or  last row number of column I
k = Range("I1").End(xlDown).Row

For a = 2 To k
    
    rOpenD = Application.WorksheetFunction.Match(Cells(a, 9), Range("A:A"), 0)
    OpenP = Cells(rOpenD, 3).Value
    
    rCloseD = rOpenD + Application.WorksheetFunction.CountIf(Range("A:A"), Cells(a, 9)) - 1
    CloseP = Cells(rCloseD, 6).Value
    
    yChange = CloseP - OpenP
    Cells(Rows.Count, 10).End(xlUp).Offset(1, 0).Value = yChange
    
    yChange_Percent = Format((CloseP / OpenP - 1), "percent")
    Cells(Rows.Count, 11).End(xlUp).Offset(1, 0).Value = yChange_Percent
    
Next a

Range("J2:J" & k).FormatConditions.Add(xlCellValue, xlGreater, "0").Interior.Color = vbGreen
Range("J2:J" & k).FormatConditions.Add(xlCellValue, xlLess, "0").Interior.Color = vbRed


End Sub

Sub Total_Stock_volum()

'Part4: Total trade volum of the stock
Dim k, r As Long

r = Range("a1").End(xlDown).Row
k = Range("I1").End(xlDown).Row

For i = 2 To k
    
    Cells(i, 9).Offset(0, 3) = Application.WorksheetFunction.SumIf _
(Range("A2:A" & r), Cells(i, 9), Range("G2:G" & r))

Next i

Range("L2:L" & k).NumberFormat = "0,0"


End Sub

Sub Greatest_performer()

'Challenge :"Greatest % increase", "Greatest % decrease" and "Greatest total volume"

Dim k As Long

k = Range("I1").End(xlDown).Row

ChgRng = Range("K:K")
TotVolRng = Range("L:L")


Range("N2") = "Greatest % Increase"
'get the max value then find the corresponding rows and ticker
Range("P2") = Application.WorksheetFunction.Max(ChgRng)
Range("P2").NumberFormat = "0.0%"
Range("O2") = Cells(Application.WorksheetFunction.Match(Range("P2"), ChgRng, 0), 9)


Range("N3") = "Greatest % Decrease"
'get the min value then find the corresponding rows and ticker
Range("P3") = Application.WorksheetFunction.Min(ChgRng)
Range("P3").NumberFormat = "0.0%"
Range("O3") = Cells(Application.WorksheetFunction.Match(Range("P3"), ChgRng, 0), 9)

Range("N4") = "Greatest Total Volume"
'get the max Vol value then find the corresponding rows and ticker
Range("P4") = Application.WorksheetFunction.Max(TotVolRng)
Range("P4").NumberFormat = "0,0"
Range("O4") = Cells(Application.WorksheetFunction.Match(Range("P4"), TotVolRng, 0), 9)

End Sub

