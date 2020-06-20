Sub stockanalysis():
Dim total As Double
Range("I1").Value = "Ticker"
Range("J1").Value = "yearlychange"
Range("K1").Value = "yearlypercentchange"
Range("L1").Value = "Total Stock Volume"
j = 0
Start = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        total = total + Cells(i, 7).Value
        If Cells(Start, 3) = 0 Then
                    For find_value = Start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                            Exit For
                        End If
                     Next find_value
                     End If
                     
       yearlychange = Cells(i, 6) - Cells(Start, 3)
       yearlypercentchange = Round((yearlychange / Cells(Start, 3) * 100), 2)
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = total
        Range("J" & 2 + j).Value = Round(yearlychange, 2)
        Range("K" & 2 + j).Value = "%" & yearlypercentchange
        Select Case yearlychange
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
       total = 0
       j = j + 1
       yearlychange = 0
    Else
       total = total + Cells(i, 7).Value
    End If
Next i

