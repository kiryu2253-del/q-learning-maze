Function max_Qval(S As Integer, Q() As Double) As Double

    Dim j As Integer

    max_Qval = Q(S, 1)

    For j = 2 To 2
        If Q(S, j) > max_Qval Then
            max_Qval = Q(S, j)
        End If
    Next j

End Function


Sub CheeseQLearningVisual()

    Dim Q(1 To 2, 1 To 2) As Double
    Dim S As Integer, a As Integer, nextS As Integer
    Dim i As Long, steps As Long
    Dim r As Double, oldQ As Double, nextMax As Double
    Dim t As Double, elapsed As Double, waitSec As Double
    Dim choiceType As String
    Dim ws As Worksheet

    steps = 30
    waitSec = 0.5

    Set ws = ActiveSheet
    ws.Range("A1:I120").Clear
    Application.ScreenUpdating = True

    Randomize
    S = 2

    ws.Range("A1") = "状態 / 行動"
    ws.Range("B1") = "a1 電源"
    ws.Range("C1") = "a2 チーズ"
    ws.Range("A2") = "S1 ON"
    ws.Range("A3") = "S2 OFF"
    ws.Range("B2:C3").Value = 0

    ws.Range("A5") = "学習回数"
    ws.Range("A6") = "現在の状態"
    ws.Range("A7") = "行動"
    ws.Range("A8") = "選択方法"
    ws.Range("A9") = "報酬"
    ws.Range("A10") = "更新前Q"
    ws.Range("A11") = "次の状態"
    ws.Range("A12") = "次状態maxQ"
    ws.Range("A13") = "更新後Q"

    ws.Range("A16:I16").Value = Array("回数", "状態", "行動", _
        "選択方法", "報酬", "次状態", "更新前Q", "次状態maxQ", "更新後Q")

    ws.Range("A1:C3").Borders.LineStyle = xlContinuous
    ws.Range("A16:I" & (steps + 16)).Borders.LineStyle = xlContinuous
    ws.Range("B2:C3").NumberFormat = "0.000"
    ws.Range("G17:I" & (steps + 16)).NumberFormat = "0.000"
    ws.Columns("A:I").AutoFit

    For i = 1 To steps

        If Rnd < 0.1 Then
            a = Int(Rnd * 2) + 1
            choiceType = "探索"
        ElseIf Q(S, 1) > Q(S, 2) Then
            a = 1
            choiceType = "活用"
        ElseIf Q(S, 2) > Q(S, 1) Then
            a = 2
            choiceType = "活用"
        Else
            a = Int(Rnd * 2) + 1
            choiceType = "同点"
        End If

        If a = 1 Then
            nextS = 3 - S
        Else
            nextS = 2
        End If

        r = 0
        If S = 1 And a = 2 Then r = 10

        oldQ = Q(S, a)
        nextMax = max_Qval(nextS, Q)
        Q(S, a) = oldQ + 0.5 * (r + 0.9 * nextMax - oldQ)

        ws.Range("B2") = Q(1, 1)
        ws.Range("C2") = Q(1, 2)
        ws.Range("B3") = Q(2, 1)
        ws.Range("C3") = Q(2, 2)

        ws.Range("B2:C3").Interior.Pattern = xlNone
        ws.Cells(S + 1, a + 1).Interior.Color = vbYellow

        ws.Range("B5") = i
        ws.Range("B6") = S
        ws.Range("B7") = a
        ws.Range("B8") = choiceType
        ws.Range("B9") = r
        ws.Range("B10") = oldQ
        ws.Range("B11") = nextS
        ws.Range("B12") = nextMax
        ws.Range("B13") = Q(S, a)

        ws.Cells(i + 16, 1) = i
        ws.Cells(i + 16, 2) = S
        ws.Cells(i + 16, 3) = a
        ws.Cells(i + 16, 4) = choiceType
        ws.Cells(i + 16, 5) = r
        ws.Cells(i + 16, 6) = nextS
        ws.Cells(i + 16, 7) = oldQ
        ws.Cells(i + 16, 8) = nextMax
        ws.Cells(i + 16, 9) = Q(S, a)

        DoEvents

        t = Timer
        Do
            DoEvents
            elapsed = Timer - t
            If elapsed < 0 Then elapsed = elapsed + 86400
        Loop While elapsed < waitSec

        S = nextS

    Next i

    MsgBox "学習が終了しました"

End Sub
