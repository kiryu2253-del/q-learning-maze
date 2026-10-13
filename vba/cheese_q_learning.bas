Function max_Qval(S As Integer, Q() As Double) As Double

    Dim j As Integer
    Dim max As Double

    max = Q(S, 1)

    For j = 1 To 2
        If Q(S, j) > max Then
            max = Q(S, j)
        End If
    Next j

    max_Qval = max

End Function


Sub CheeseQLearning()

    Dim Q(1 To 2, 1 To 2) As Double
    Dim S As Integer, a As Integer, nextS As Integer
    Dim i As Integer, j As Integer
    Dim r As Double

    Range("A1:M110").Clear

    Randomize
    S = 2

    For i = 1 To 100

        If Rnd < 0.1 Then
            a = Int(Rnd * 2) + 1
        ElseIf Q(S, 1) > Q(S, 2) Then
            a = 1
        ElseIf Q(S, 2) > Q(S, 1) Then
            a = 2
        Else
            a = Int(Rnd * 2) + 1
        End If

        If a = 1 Then
            nextS = 3 - S
        Else
            nextS = 2
        End If

        r = 0
        If S = 1 And a = 2 Then r = 10

        Q(S, a) = 0.5 * Q(S, a) _
                + 0.5 * (r + 0.9 * max_Qval(nextS, Q))

        S = nextS

    Next i

    Cells(1, 1) = "状態 / 行動"
    Cells(1, 2) = "a1 電源"
    Cells(1, 3) = "a2 チーズ"

    Cells(2, 1) = "S1 ON"
    Cells(3, 1) = "S2 OFF"

    For i = 1 To 2
        For j = 1 To 2
            Cells(i + 1, j + 1) = Q(i, j)
        Next j

        For j = 1 To 2
            If Q(i, j) = max_Qval(i, Q) Then
                Cells(i + 1, j + 1).Interior.Color = vbYellow
            End If
        Next j
    Next i

    Range("A1:C3").Borders.LineStyle = xlContinuous
    Range("B2:C3").NumberFormat = "0.000"
    Columns("A:C").AutoFit

End Sub
