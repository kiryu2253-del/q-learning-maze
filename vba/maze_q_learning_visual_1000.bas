Option Explicit

Function MazeMaxQ(ByVal s As Long, Q() As Double) As Double
    Dim a As Long
    MazeMaxQ = Q(s, 1)
    For a = 2 To 4
        If Q(s, a) > MazeMaxQ Then MazeMaxQ = Q(s, a)
    Next a
End Function

Function MazeBestAction(ByVal s As Long, Q() As Double) As Long
    Dim a As Long
    MazeBestAction = 1
    For a = 2 To 4
        If Q(s, a) > Q(s, MazeBestAction) Then MazeBestAction = a
    Next a
End Function

Sub MazeQLearningVisual1000()
    Dim ws As Worksheet
    Dim walls As Variant, dr As Variant, dc As Variant
    Dim Q(1 To 25, 1 To 4) As Double
    Dim history(1 To 1000, 1 To 4) As Variant
    Dim trail(1 To 101) As Long
    Dim ep As Long, k As Long, s As Long, ns As Long, a As Long
    Dim r As Long, c As Long, nr As Long, nc As Long
    Dim nTrail As Long, reward As Double, totalReward As Double
    Dim oldQ As Double, target As Double, epsilon As Double
    Dim successes As Long, reached As Boolean
    Dim viewEvery As Long, delaySec As Double
    Dim lastState As Long, lastAction As Long

    ' GitHub上の5×5迷路と同じ壁の配置
    walls = Array("00100", "10101", "00000", "01110", "00000")
    dr = Array(-1, 0, 1, 0)  ' 上、右、下、左
    dc = Array(0, 1, 0, -1)

    viewEvery = 10         ' 10エピソードごとに画面更新
    delaySec = 0.03        ' 画面更新後の待ち時間（秒）
    epsilon = 1           ' 探索率は最初1.0、徐々に0.05まで下げる

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Maze_Q_Visual")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Maze_Q_Visual"
    End If

    ws.Activate
    ws.Cells.Clear        ' 専用シートの前回結果だけ消す
    Application.ScreenUpdating = True
    Randomize

    ws.Range("A1") = "5×5迷路 Q学習 / 1000エピソード"
    ws.Range("A10") = "エピソード"
    ws.Range("A11") = "進み具合"
    ws.Range("A12") = "ε（探索率）"
    ws.Range("A13") = "直近の移動回数"
    ws.Range("A14") = "直近の合計報酬"
    ws.Range("A15") = "累計ゴール到達"
    ws.Range("A16") = "更新したQ値"
    ws.Range("A18") = "学習後の経路"
    ws.Range("H3:M3").Value = Array("状態", "↑上", "→右", "↓下", "←左", "最大の行動")
    ws.Range("P3:S3").Value = Array("エピソード", "移動回数", "合計報酬", "ゴール")

    For s = 1 To 25
        ws.Cells(s + 3, 8) = s - 1  ' Pythonと同じ0〜24の状態番号
    Next s
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:F").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 9
    ws.Columns("I:L").ColumnWidth = 11
    ws.Columns("M").ColumnWidth = 15
    ws.Columns("P:S").ColumnWidth = 13
    ws.Rows("3:7").RowHeight = 31
    ws.Range("B3:F7").HorizontalAlignment = xlCenter
    ws.Range("B3:F7").VerticalAlignment = xlCenter
    ws.Range("B3:F7").Font.Size = 15
    ws.Range("B3:F7").Borders.LineStyle = xlContinuous
    ws.Range("H3:M28").Borders.LineStyle = xlContinuous
    ws.Range("I4:L28").NumberFormat = "0.00"
    ws.Range("H3:M3").Interior.Color = RGB(212, 229, 248)
    ws.Range("P3:S3").Interior.Color = RGB(212, 229, 248)

    For ep = 1 To 1000
        r = 1: c = 1
        totalReward = 0
        reached = False
        nTrail = 1
        trail(1) = 1

        For k = 1 To 100
            s = (r - 1) * 5 + c

            ' ε-greedy: 探索、または最大Q値の行動
            If Rnd < epsilon Then
                a = Int(Rnd * 4) + 1
            Else
                a = MazeBestAction(s, Q)
            End If

            nr = r + dr(a - 1)
            nc = c + dc(a - 1)
            If nr < 1 Or nr > 5 Or nc < 1 Or nc > 5 Then
                nr = r: nc = c
                reward = -5
            ElseIf Mid$(walls(nr - 1), nc, 1) = "1" Then
                nr = r: nc = c
                reward = -5
            ElseIf nr = 5 And nc = 5 Then
                reward = 100
                reached = True
            Else
                reward = -1
            End If

            ns = (nr - 1) * 5 + nc
            oldQ = Q(s, a)
            If reached Then
                target = reward
            Else
                target = reward + 0.9 * MazeMaxQ(ns, Q)
            End If
            Q(s, a) = oldQ + 0.5 * (target - oldQ)
            lastState = s: lastAction = a

            r = nr: c = nc
            totalReward = totalReward + reward
            nTrail = nTrail + 1
            trail(nTrail) = ns
            If reached Then Exit For
        Next k
        If k > 100 Then k = 100

        If reached Then successes = successes + 1
        history(ep, 1) = ep
        history(ep, 2) = k
        history(ep, 3) = totalReward
        If reached Then history(ep, 4) = "到達" Else history(ep, 4) = "未到達"

        epsilon = epsilon * 0.995
        If epsilon < 0.05 Then epsilon = 0.05

        If ep = 1 Or ep Mod viewEvery = 0 Then
            Call MazeDraw(ws, walls, trail, nTrail, Q, lastState, lastAction)
            ws.Range("B10") = ep
            ws.Range("B11") = Format$(ep / 1000, "0.0%")
            ws.Range("B12") = Format$(epsilon, "0.000")
            ws.Range("B13") = k
            ws.Range("B14") = totalReward
            ws.Range("B15") = successes & " / " & ep
            ws.Range("B16") = (lastState - 1) & ", 行動" & lastAction
            DoEvents
            Call MazeShortWait(delaySec)
        End If
    Next ep

    ws.Range("P4:S1003").Value = history

    ' 学習後は最大Q値の行動だけでゴールに向かう様子を表示
    r = 1: c = 1
    nTrail = 1
    trail(1) = 1
    ws.Range("B18") = "再生中"
    Call MazeDraw(ws, walls, trail, nTrail, Q, 0, 0)
    For k = 1 To 30
        s = (r - 1) * 5 + c
        a = MazeBestAction(s, Q)
        nr = r + dr(a - 1)
        nc = c + dc(a - 1)
        If nr < 1 Or nr > 5 Or nc < 1 Or nc > 5 Then Exit For
        If Mid$(walls(nr - 1), nc, 1) = "1" Then Exit For
        r = nr: c = nc
        nTrail = nTrail + 1
        trail(nTrail) = (r - 1) * 5 + c
        Call MazeDraw(ws, walls, trail, nTrail, Q, s, a)
        DoEvents
        Call MazeShortWait(0.12)
        If r = 5 And c = 5 Then Exit For
    Next k

    If r = 5 And c = 5 Then
        ws.Range("B18") = "ゴール到達 / " & (nTrail - 1) & " 歩"
    Else
        ws.Range("B18") = "ゴール未到達（要再学習）"
    End If
    ws.Range("A20") = "青: 通った道 / 黄色: 更新したQ値 / 黒: 壁"
    MsgBox "1000エピソードの学習が終了しました。"
End Sub

Private Sub MazeDraw(ws As Worksheet, walls As Variant, _
                     trail() As Long, nTrail As Long, _
                     Q() As Double, lastState As Long, lastAction As Long)
    Dim r As Long, c As Long, i As Long, s As Long, a As Long
    Dim rr As Long, cc As Long, pos As Long
    Dim values(1 To 25, 1 To 4) As Variant
    Dim names As Variant
    Dim cell As Range

    names = Array("上", "右", "下", "左")

    For r = 1 To 5
        For c = 1 To 5
            Set cell = ws.Cells(r + 2, c + 1)
            If Mid$(walls(r - 1), c, 1) = "1" Then
                cell.Value = ""
                cell.Interior.Color = RGB(52, 64, 81)
            Else
                cell.Value = ""
                cell.Interior.Color = RGB(241, 245, 249)
            End If
        Next c
    Next r

    For i = 1 To nTrail
        pos = trail(i)
        rr = (pos - 1) \ 5 + 1
        cc = (pos - 1) Mod 5 + 1
        ws.Cells(rr + 2, cc + 1).Interior.Color = RGB(185, 220, 255)
    Next i

    ws.Cells(3, 2).Value = "S"
    ws.Cells(7, 6).Value = "G"
    pos = trail(nTrail)
    rr = (pos - 1) \ 5 + 1
    cc = (pos - 1) Mod 5 + 1
    ws.Cells(rr + 2, cc + 1).Value = "●"
    ws.Cells(rr + 2, cc + 1).Interior.Color = RGB(255, 217, 128)

    For s = 1 To 25
        For a = 1 To 4
            values(s, a) = Q(s, a)
        Next a
        ws.Cells(s + 3, 13).Value = names(MazeBestAction(s, Q) - 1)
    Next s
    ws.Range("I4:L28").Value = values
    ws.Range("I4:L28").Interior.ColorIndex = xlNone
    If lastState >= 1 And lastAction >= 1 Then
        ws.Cells(lastState + 3, lastAction + 8).Interior.Color = vbYellow
    End If
End Sub

Private Sub MazeShortWait(ByVal seconds As Double)
    Dim startTime As Double, passed As Double
    startTime = Timer
    Do
        DoEvents
        passed = Timer - startTime
        If passed < 0 Then passed = passed + 86400
    Loop While passed < seconds
End Sub
