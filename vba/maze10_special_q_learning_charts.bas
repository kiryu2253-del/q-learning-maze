Option Explicit

' 10x10迷路: 状態100個、行動4個
' 1=上、2=右、3=下、4=左
Private Function Maze10MaxQ(ByVal s As Long, Q() As Double) As Double
    Dim a As Long
    Maze10MaxQ = Q(s, 1)
    For a = 2 To 4
        If Q(s, a) > Maze10MaxQ Then Maze10MaxQ = Q(s, a)
    Next a
End Function

Private Function Maze10BestAction(ByVal s As Long, Q() As Double) As Long
    Dim a As Long
    Maze10BestAction = 1
    For a = 2 To 4
        If Q(s, a) > Q(s, Maze10BestAction) Then
            Maze10BestAction = a
        End If
    Next a
End Function

' 落とし穴・固定ワープ・ランダムワープを含む移動処理
' 戻り値 True = ゴール到達
Private Function Maze10Move(ByVal r As Long, ByVal c As Long, _
    ByVal a As Long, walls As Variant, ByRef nr As Long, _
    ByRef nc As Long, ByRef reward As Double, _
    ByRef eventName As String) As Boolean

    Dim dr As Variant, dc As Variant, dest As Long
    dr = Array(-1, 0, 1, 0)
    dc = Array(0, 1, 0, -1)
    nr = r + dr(a - 1)
    nc = c + dc(a - 1)
    reward = -1
    eventName = "通常移動"
    Maze10Move = False

    If nr < 1 Or nr > 10 Or nc < 1 Or nc > 10 Then
        nr = r: nc = c: reward = -5: eventName = "迷路の外"
    ElseIf Mid$(walls(nr - 1), nc, 1) = "1" Then
        nr = r: nc = c: reward = -5: eventName = "壁に衝突"
    ElseIf (nr = 5 And nc = 6) Or (nr = 8 And nc = 8) _
        Or (nr = 9 And nc = 3) Then
        nr = 1: nc = 1: reward = -30
        eventName = "落とし穴 → スタート"
    ElseIf nr = 5 And nc = 8 Then
        nr = 9: nc = 8
        eventName = "固定ワープ → (9,8)"
    ElseIf nr = 2 And nc = 6 Then
        dest = Int(Rnd * 3) + 1
        Select Case dest
            Case 1: nr = 9: nc = 10
            Case 2: nr = 8: nc = 2
            Case 3: nr = 1: nc = 2
        End Select
        reward = -3
        eventName = "ランダムワープ → (" & nr & "," & nc & ")"
    ElseIf nr = 10 And nc = 10 Then
        reward = 100: eventName = "ゴール！"
        Maze10Move = True
    End If
End Function

Sub Maze10SpecialQLearningCharts()
    Const episodes As Long = 1000
    Const maxSteps As Long = 160
    Const viewEvery As Long = 20
    Const delaySec As Double = 0.02

    Dim ws As Worksheet, wsData As Worksheet, walls As Variant
    Dim Q(1 To 100, 1 To 4) As Double
    Dim stepsHistory(1 To 1000) As Long
    Dim rewardHistory(1 To 1000) As Double
    Dim path(1 To 161) As Long
    Dim ep As Long, k As Long, nPath As Long
    Dim r As Long, c As Long, nr As Long, nc As Long
    Dim s As Long, ns As Long, a As Long
    Dim reward As Double, oldQ As Double, target As Double
    Dim totalReward As Double, epsilon As Double
    Dim eventName As String, lastEvent As String
    Dim reached As Boolean, reachedTotal As Long
    Dim fixedCount As Long, randomCount As Long, pitCount As Long
    Dim runningSteps As Double, runningReward As Double
    Dim avgSteps As Double, avgReward As Double
    Dim countAvg As Long, i As Long

    ' 0=通路、1=壁。スタート(1,1)、ゴール(10,10)
    walls = Array( _
        "0000001001", "1111000100", "1100011000", _
        "0101000001", "1110000000", "0000000000", _
        "0001111101", "0000000010", "0101101000", _
        "0010000010")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Maze10_Special")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Maze10_Special"
    End If
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("Maze10_Data")
    On Error GoTo 0
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add(After:=ws)
        wsData.Name = "Maze10_Data"
    End If
    ws.Activate
    ws.Cells.Clear
    wsData.Cells.Clear
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    Application.ScreenUpdating = True
    Randomize
    epsilon = 1

    ws.Range("A1") = "10×10 Q学習：落とし穴・固定ワープ・ランダムワープ"
    ws.Range("A15:A23").Value = Application.Transpose(Array( _
        "学習回数", "進捗", "探索率 ε", "直近の歩数", _
        "直近の報酬", "ゴール累計", "固定ワープ累計", _
        "ランダムワープ累計", "落とし穴累計"))
    ws.Range("A25") = "直近の出来事"
    ws.Range("A27") = "学習後の確認"
    ws.Range("M3:R3").Value = Array("状態", "上", "右", "下", "左", "最大Qの行動")
    wsData.Range("A3:G3").Value = Array("回数", "ステップ数", "獲得報酬", _
        "30回平均ステップ", "30回平均報酬", "ゴール", "最後の出来事")
    wsData.Range("A3:G3").Interior.Color = RGB(210, 231, 248)
    wsData.Columns("A:G").ColumnWidth = 20

    ws.Columns("A").ColumnWidth = 19
    ws.Columns("B:K").ColumnWidth = 6
    ws.Columns("M").ColumnWidth = 9
    ws.Columns("N:Q").ColumnWidth = 11
    ws.Columns("R").ColumnWidth = 14
    ws.Columns("T:W").ColumnWidth = 12
    ws.Columns("X").ColumnWidth = 29
    ws.Rows("3:12").RowHeight = 30
    With ws.Range("B3:K12")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 13
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("M3:R103").Borders.LineStyle = xlContinuous
    ws.Range("N4:Q103").NumberFormat = "0.00"
    ws.Range("M3:R3").Interior.Color = RGB(210, 231, 248)
    ws.Range("A32") = "上: 獲得報酬 / 下: ステップ数。線は学習に応じて増える。"
    ws.Range("A33") = "生データと直近30回平均を重ねて表示"
    Maze10AddChart ws, wsData, "T3", "獲得報酬の推移", "C", "E"
    Maze10AddChart ws, wsData, "T22", "ステップ数の推移", "B", "D"
    ws.Range("A29") = "S=開始、G=ゴール、X=穴、F=固定ワープ、W=到着先、R=ランダムワープ"
    ws.Range("A30") = "固定 F(5,8)→(9,8) / ランダム R(2,6)→(9,10),(8,2),(1,2)"

    For ep = 1 To episodes
        r = 1: c = 1: nPath = 1: path(1) = 1
        totalReward = 0: reached = False: lastEvent = ""

        For k = 1 To maxSteps
            s = (r - 1) * 10 + c

            If Rnd < epsilon Then
                a = Int(Rnd * 4) + 1
            Else
                a = Maze10BestAction(s, Q)
            End If

            reached = Maze10Move(r, c, a, walls, nr, nc, reward, eventName)
            ns = (nr - 1) * 10 + nc

            oldQ = Q(s, a)
            If reached Then
                target = reward          ' ゴールは終端状態
            Else
                target = reward + 0.9 * Maze10MaxQ(ns, Q)
            End If
            Q(s, a) = oldQ + 0.5 * (target - oldQ)

            If Left$(eventName, 5) = "固定ワープ" Then fixedCount = fixedCount + 1
            If Left$(eventName, 7) = "ランダムワープ" Then randomCount = randomCount + 1
            If Left$(eventName, 4) = "落とし穴" Then pitCount = pitCount + 1

            r = nr: c = nc: totalReward = totalReward + reward
            nPath = nPath + 1: path(nPath) = ns
            lastEvent = eventName
            If reached Then Exit For
        Next k
        If k > maxSteps Then k = maxSteps
        If reached Then reachedTotal = reachedTotal + 1

        stepsHistory(ep) = k
        rewardHistory(ep) = totalReward
        runningSteps = runningSteps + k
        runningReward = runningReward + totalReward
        If ep > 30 Then
            runningSteps = runningSteps - stepsHistory(ep - 30)
            runningReward = runningReward - rewardHistory(ep - 30)
        End If
        countAvg = ep
        If countAvg > 30 Then countAvg = 30
        avgSteps = runningSteps / countAvg
        avgReward = runningReward / countAvg

        wsData.Cells(ep + 3, 1) = ep
        wsData.Cells(ep + 3, 2) = k
        wsData.Cells(ep + 3, 3) = totalReward
        wsData.Cells(ep + 3, 4) = avgSteps
        wsData.Cells(ep + 3, 5) = avgReward
        If reached Then
            wsData.Cells(ep + 3, 6) = 1
        Else
            wsData.Cells(ep + 3, 6) = 0
        End If
        wsData.Cells(ep + 3, 7) = lastEvent

        epsilon = epsilon * 0.995
        If epsilon < 0.05 Then epsilon = 0.05

        If ep = 1 Or ep Mod viewEvery = 0 Then
            Maze10Draw ws, walls, path, nPath, Q, s, a
            Maze10RefreshCharts ws, wsData, ep
            ws.Range("B15") = ep
            ws.Range("B16") = Format$(ep / episodes, "0.0%")
            ws.Range("B17") = Format$(epsilon, "0.000")
            ws.Range("B18") = k
            ws.Range("B19") = totalReward
            ws.Range("B20") = reachedTotal
            ws.Range("D15") = "直近30回平均"
            ws.Range("D16") = "報酬"
            ws.Range("E16") = Round(avgReward, 1)
            ws.Range("D17") = "ステップ"
            ws.Range("E17") = Round(avgSteps, 1)
            ws.Range("B21") = fixedCount
            ws.Range("B22") = randomCount
            ws.Range("B23") = pitCount
            ws.Range("B25") = lastEvent
            DoEvents
            Maze10Wait delaySec
        End If
    Next ep

    wsData.Range("A3:G1003").Borders.LineStyle = xlContinuous
    wsData.Range("C4:E1003").NumberFormat = "0.0"

    ' 学習後は ε=0 として最もQ値が高い行動だけで移動する
    r = 1: c = 1: nPath = 1: path(1) = 1
    ws.Range("B27") = "最終経路を再生中"
    For k = 1 To 60
        s = (r - 1) * 10 + c
        a = Maze10BestAction(s, Q)
        reached = Maze10Move(r, c, a, walls, nr, nc, reward, eventName)
        r = nr: c = nc
        nPath = nPath + 1: path(nPath) = (r - 1) * 10 + c
        Maze10Draw ws, walls, path, nPath, Q, s, a
        ws.Range("B25") = eventName
        DoEvents
        Maze10Wait 0.08
        If reached Then Exit For
    Next k

    If reached Then
        ws.Range("B27") = "ゴール到達 / " & (nPath - 1) & "歩"
    Else
        ws.Range("B27") = "60歩以内にゴールできず"
    End If
    MsgBox "1000エピソードの学習が終了しました。"
End Sub

Private Sub Maze10Draw(ByVal ws As Worksheet, walls As Variant, _
    path() As Long, ByVal nPath As Long, Q() As Double, _
    ByVal lastState As Long, ByVal lastAction As Long)

    Dim row As Long, col As Long, i As Long, pos As Long
    Dim rr As Long, cc As Long, s As Long, a As Long
    Dim qValues(1 To 100, 1 To 4) As Variant
    Dim best(1 To 100, 1 To 1) As Variant
    Dim actionNames As Variant
    Dim targetCell As Range

    actionNames = Array("上", "右", "下", "左")

    For row = 1 To 10
        For col = 1 To 10
            Set targetCell = ws.Cells(row + 2, col + 1)
            targetCell.Value = ""
            targetCell.Interior.Color = RGB(244, 247, 250)
            If Mid$(walls(row - 1), col, 1) = "1" Then
                targetCell.Interior.Color = RGB(49, 61, 76)
            ElseIf (row = 5 And col = 6) Or (row = 8 And col = 8) _
                Or (row = 9 And col = 3) Then
                targetCell.Value = "X"
                targetCell.Interior.Color = RGB(255, 172, 172)
            ElseIf row = 5 And col = 8 Then
                targetCell.Value = "F"
                targetCell.Interior.Color = RGB(194, 168, 238)
            ElseIf row = 9 And col = 8 Then
                targetCell.Value = "W"
                targetCell.Interior.Color = RGB(224, 205, 250)
            ElseIf row = 2 And col = 6 Then
                targetCell.Value = "R"
                targetCell.Interior.Color = RGB(143, 225, 219)
            ElseIf (row = 9 And col = 10) Or (row = 8 And col = 2) _
                Or (row = 1 And col = 2) Then
                targetCell.Value = "?"
                targetCell.Interior.Color = RGB(198, 239, 236)
            End If
        Next col
    Next row

    For i = 1 To nPath
        pos = path(i)
        rr = (pos - 1) \ 10 + 1
        cc = (pos - 1) Mod 10 + 1
        If ws.Cells(rr + 2, cc + 1).Value = "" Then
            ws.Cells(rr + 2, cc + 1).Interior.Color = RGB(177, 216, 248)
        End If
    Next i

    ws.Cells(3, 2).Value = "S"
    ws.Cells(3, 2).Interior.Color = RGB(175, 227, 186)
    ws.Cells(12, 11).Value = "G"
    ws.Cells(12, 11).Interior.Color = RGB(168, 230, 184)

    pos = path(nPath)
    rr = (pos - 1) \ 10 + 1
    cc = (pos - 1) Mod 10 + 1
    ws.Cells(rr + 2, cc + 1).Value = "●"
    ws.Cells(rr + 2, cc + 1).Interior.Color = RGB(255, 212, 109)

    For s = 1 To 100
        For a = 1 To 4
            qValues(s, a) = Q(s, a)
        Next a
        best(s, 1) = actionNames(Maze10BestAction(s, Q) - 1)
    Next s
    ws.Range("M4:M103").Formula = "=ROW()-4"
    ws.Range("N4:Q103").Value = qValues
    ws.Range("R4:R103").Value = best
    ws.Range("N4:Q103").Interior.Pattern = xlNone
    If lastState >= 1 And lastAction >= 1 Then
        ws.Cells(lastState + 3, lastAction + 13).Interior.Color = vbYellow
    End If
End Sub

Private Sub Maze10Wait(ByVal seconds As Double)
    Dim beginTime As Double, elapsed As Double
    beginTime = Timer
    Do
        DoEvents
        elapsed = Timer - beginTime
        If elapsed < 0 Then elapsed = elapsed + 86400
    Loop While elapsed < seconds
End Sub

Private Sub Maze10AddChart(ByVal ws As Worksheet, ByVal dataWs As Worksheet, _
    ByVal anchor As String, ByVal heading As String, _
    ByVal rawCol As String, ByVal avgCol As String)

    Dim co As ChartObject
    Dim ser As Series

    Set co = ws.ChartObjects.Add(Left:=ws.Range(anchor).Left, _
        Top:=ws.Range(anchor).Top, Width:=610, Height:=260)
    With co.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = heading
        .HasLegend = True
        Set ser = .SeriesCollection.NewSeries
        ser.Name = "毎回の値"
        ser.XValues = dataWs.Range("A4:A4")
        ser.Values = dataWs.Range(rawCol & "4:" & rawCol & "4")
        Set ser = .SeriesCollection.NewSeries
        ser.Name = "直近30回平均"
        ser.XValues = dataWs.Range("A4:A4")
        ser.Values = dataWs.Range(avgCol & "4:" & avgCol & "4")
        .DisplayBlanksAs = xlNotPlotted
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "エピソード"
    End With
End Sub

Private Sub Maze10RefreshCharts(ByVal ws As Worksheet, _
    ByVal dataWs As Worksheet, ByVal ep As Long)

    Dim rowEnd As Long
    rowEnd = ep + 3
    With ws.ChartObjects(1).Chart
        .SeriesCollection(1).XValues = dataWs.Range("A4:A" & rowEnd)
        .SeriesCollection(1).Values = dataWs.Range("C4:C" & rowEnd)
        .SeriesCollection(2).XValues = dataWs.Range("A4:A" & rowEnd)
        .SeriesCollection(2).Values = dataWs.Range("E4:E" & rowEnd)
    End With
    With ws.ChartObjects(2).Chart
        .SeriesCollection(1).XValues = dataWs.Range("A4:A" & rowEnd)
        .SeriesCollection(1).Values = dataWs.Range("B4:B" & rowEnd)
        .SeriesCollection(2).XValues = dataWs.Range("A4:A" & rowEnd)
        .SeriesCollection(2).Values = dataWs.Range("D4:D" & rowEnd)
    End With
End Sub
