Sub Hard()

Dim ticker As String, volume As Double
Dim i, j, a, w, ws_count As Integer
Dim Rows_count  As Long

j = 2
a = 0
volume = 0
ws_count = ActiveWorkbook.Worksheets.Count
opendate = 30000000
closedate = 0
largest = 0
smallest = 5000
max_volume = 0

For w = 1 To ws_count
    ActiveWorkbook.Worksheets(w).Range("I1:M1,Q1:R1") = [{"Ticker","Total Stock Volume","Count","Yearly Change","Percent Change"}]
    ActiveWorkbook.Worksheets(w).Range("Q1:R1") = [{"Ticker","Value"}]
    ActiveWorkbook.Worksheets(w).Range("P2") = "Greatest % Increase"
    ActiveWorkbook.Worksheets(w).Range("P3") = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(w).Range("P4") = "Greatest Total Volume"
    Rows_count = ActiveWorkbook.Worksheets(w).Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To Rows_count
        If ActiveWorkbook.Worksheets(w).Cells(i, 1) = ActiveWorkbook.Worksheets(w).Cells(i + 1, 1) Then
            a = a + 1
            ticker = ActiveWorkbook.Worksheets(w).Cells(i, 1)
            volume = volume + ActiveWorkbook.Worksheets(w).Cells(i, 7)
            If ActiveWorkbook.Worksheets(w).Cells(i, 2) < opendate Then
                opendate = ActiveWorkbook.Worksheets(w).Cells(i, 2)
                openprice = ActiveWorkbook.Worksheets(w).Cells(i, 3)
            End If
            If ActiveWorkbook.Worksheets(w).Cells(i + 1, 2) > closedate Then
                closedate = ActiveWorkbook.Worksheets(w).Cells(i + 1, 2)
                closeprice = ActiveWorkbook.Worksheets(w).Cells(i + 1, 6)
            End If

        Else
            ActiveWorkbook.Worksheets(w).Cells(j, 9) = ticker
            volume = volume + ActiveWorkbook.Worksheets(w).Cells(i, 7)
            ActiveWorkbook.Worksheets(w).Cells(j, 10) = volume
            ActiveWorkbook.Worksheets(w).Cells(j, 11) = a + 1
            Change = closeprice - openprice
            If openprice <> 0 Then
                Percent = Change / openprice
            Else
                Percent = Null
            End If
            ActiveWorkbook.Worksheets(w).Cells(j, 12) = Change
            If ActiveWorkbook.Worksheets(w).Cells(j, 12) > 0 Then
                ActiveWorkbook.Worksheets(w).Cells(j, 12).Interior.ColorIndex = 4
            ElseIf ActiveWorkbook.Worksheets(w).Cells(j, 12) < 0 Then
                ActiveWorkbook.Worksheets(w).Cells(j, 12).Interior.ColorIndex = 3
            End If
            ActiveWorkbook.Worksheets(w).Cells(j, 13) = Percent
            If Percent > largest Then
                largest = Percent
                largest_ticker = ticker
            End If
            If Percent < smallest Then
                smallest = Percent
                smallest_ticker = ticker
            End If
            If volume > max_volume Then
                max_volume = volume
                max_vol_ticker = ticker
            End If
            a = 0
            volume = 0
            j = j + 1
            opendate = 30000000
            closedate = 0
        End If
    Next i
    
    ActiveWorkbook.Worksheets(w).Range("Q2") = largest_ticker
    ActiveWorkbook.Worksheets(w).Range("R2") = largest
    ActiveWorkbook.Worksheets(w).Range("Q3") = smallest_ticker
    ActiveWorkbook.Worksheets(w).Range("R3") = smallest
    ActiveWorkbook.Worksheets(w).Range("Q4") = max_vol_ticker
    ActiveWorkbook.Worksheets(w).Range("R4") = max_volume
    
    j = 2
    largest = 0
    smallest = 5000
    max_volume = 0
    ActiveWorkbook.Worksheets(w).Range("M:M").NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(w).Range("R2:R3").NumberFormat = "0.00%"

    ActiveWorkbook.Worksheets(w).Columns("A:R").AutoFit
Next w

End Sub

