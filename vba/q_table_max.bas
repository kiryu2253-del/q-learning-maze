Sub QTable()

    Dim i As Integer
    Dim j As Integer
    Dim max As Integer
    Dim j_max As Integer

    Randomize

    Range("A1:M11").Clear

    Cells(1, 1).Value = "i\j"

    For j = 1 To 10
        Cells(1, j + 1).Value = j
    Next j

    For i = 1 To 10
        Cells(i + 1, 1).Value = i
    Next i

    Cells(1, 12).Value = "max"
    Cells(1, 13).Value = "j_max"

    For i = 1 To 10

        For j = 1 To 10
            Cells(i + 1, j + 1).Value = Int(Rnd * 10) + 1
        Next j

    Next i

    For i = 1 To 10

        max = -1
        j_max = -1

        For j = 1 To 10

            If Cells(i + 1, j + 1).Value > max Then
                max = Cells(i + 1, j + 1).Value
                j_max = j
            End If

        Next j

        Cells(i + 1, 12).Value = max
        Cells(i + 1, 13).Value = j_max

        Cells(i + 1, j_max + 1).Interior.Color = vbYellow

    Next i

    Range("A1:M11").Borders.LineStyle = xlContinuous

End Sub
