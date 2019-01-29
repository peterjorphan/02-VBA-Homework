Sub Moderate()

Dim ticker As String, volume As Double
Dim i, j, a, w, ws_count As Integer
Dim Rows_count  As Long

j = 2
a = 0
volume = 0
ws_count = ActiveWorkbook.Worksheets.Count
opendate = 30000000
closedate = 0

For w = 1 To ws_count
    ActiveWorkbook.Worksheets(w).Range("I1:M1") = [{"Ticker","Total Stock Volume","Count","Yearly Change","Percent Change"}]
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
            a = 0
            volume = 0
            j = j + 1
            opendate = 30000000
            closedate = 0
        End If
    Next i
    j = 2

    ActiveWorkbook.Worksheets(w).Range("M:M").NumberFormat = "0.00%"

Next w

End Sub
