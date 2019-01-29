Sub easy()

Dim ticker As String, volume As Double
Dim i, j, a, w, ws_count As Integer
Dim Rows_count As Long

j = 2
a = 0
volume = 0
ws_count = ActiveWorkbook.Worksheets.Count

For w = 1 To ws_count

    ActiveWorkbook.Worksheets(w).Range("I1:K1") = [{"Ticker","Total Stock Volume","Count"}]
    Rows_count = ActiveWorkbook.Worksheets(w).Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Rows_count
        If ActiveWorkbook.Worksheets(w).Cells(i, 1) = ActiveWorkbook.Worksheets(w).Cells(i + 1, 1) Then
            a = a + 1
            ticker = ActiveWorkbook.Worksheets(w).Cells(i, 1)
            volume = volume + ActiveWorkbook.Worksheets(w).Cells(i, 7)
        Else
            ActiveWorkbook.Worksheets(w).Cells(j, 9) = ticker
            volume = volume + ActiveWorkbook.Worksheets(w).Cells(i, 7)
            ActiveWorkbook.Worksheets(w).Cells(j, 10) = volume
            ActiveWorkbook.Worksheets(w).Cells(j, 11) = a + 1
            a = 0
            volume = 0
            j = j + 1
        End If
    Next i
    
    j = 2
Next w

End Sub
