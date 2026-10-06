Sub QTable()

    Dim i As Integer
    Dim j As Integer
    Dim max As Integer
    Dim j_max As Integer

    Randomize

    Range("A1:J10").Interior.ColorIndex = xlNone

    For i = 1 To 10

        For j = 1 To 10
            Cells(i, j).Value = Int(Rnd * 10) + 1
        Next j

    Next i

    For i = 1 To 10

        max = -1
        j_max = -1

        For j = 1 To 10

            If Cells(i, j).Value > max Then
                max = Cells(i, j).Value
                j_max = j
            End If

        Next j

        Cells(i, 11).Value = max
        Cells(i, 12).Value = j_max

        Cells(i, j_max).Interior.Color = vbYellow

    Next i

End Sub
