Attribute VB_Name = "Module3"
Sub ZamienNaDateTypuData()
    Dim rng As Range
    Dim komorka As Range
    Dim tekstData As String
    Dim poprawionaData As Date

    ' Ustawienie zakresu danych w kolumnie H (gdzie mamy tekstowe daty)
    Set rng = Range("H2:H" & Cells(Rows.Count, "H").End(xlUp).Row)

    ' Iteracja przez ka¿d¹ komórkê w kolumnie H
    For Each komorka In rng
        tekstData = komorka.Value

        ' Sprawdzamy, czy komórka zawiera tekst w formie daty
        If IsDate(tekstData) Then
            ' Przekszta³camy tekstow¹ datê na typ daty
            poprawionaData = CDate(tekstData)

            ' Zapisujemy datê w nowej kolumnie (I) jako typ daty
            komorka.Offset(0, 1).Value = poprawionaData

            ' Formatujemy komórkê w kolumnie I jako datê w formacie "d-m-yyyy"
            komorka.Offset(0, 1).NumberFormat = "d-m-yyyy"
        Else
            ' Jeœli nie jest to poprawna data, wpisujemy "B³¹d"
            komorka.Offset(0, 1).Value = "B³¹d"
        End If
    Next komorka
End Sub

