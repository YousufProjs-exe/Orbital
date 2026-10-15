
Option Explicit

Private Const FIRST_LANE As Long = 1
Private Const LAST_LANE As Long = 3
Private Const PLAYER_ROW As Long = 16
Private Const FIRST_DEBRIS_ROW As Long = 6
Private Const TICK_SECONDS As Long = 1

Private debrisLane As Long
Private debrisRow As Long
Private gameOver As Boolean
Private gameRunning As Boolean
Private nextTick As Date

Private Const PLAYER_SYMBOL As String = "A"
Private Const DEBRIS_SYMBOL As String = "X"

Public Sub SetupBoard()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    StopAutoGame

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 3
    ws.Rows.RowHeight = 20

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B2:J26").Interior.Color = RGB(10, 15, 28)

    With ws.Range("B2:J2")
        .Merge
        .Value = "ORBITAL"
        .Font.Name = "Consolas"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(100, 200, 255)
        .HorizontalAlignment = xlCenter
    End With

    DrawLanes
    CreateGameButtons
End Sub

Public Sub DrawLanes()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("D6:D16,F6:F16,H6:H16").Interior.Color = RGB(20, 35, 55)

    With ws.Range("E5")
        .Value = "ORBITAL LANES"
        .Font.Name = "Consolas"
        .Font.Color = RGB(100, 200, 255)
    End With
End Sub

Public Sub PlaceSatellite()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("B20").Value = 2
    ws.Range("E18").Value = "SATELLITE"
    ws.Range("E18").Font.Name = "Consolas"
    ws.Range("E18").Font.Color = RGB(80, 255, 190)

    RenderSatellite ws, 2
End Sub

Public Sub MoveLeft()
    If gameOver Or Not gameRunning Then Exit Sub
    ChangeLane -1
End Sub

Public Sub MoveRight()
    If gameOver Or Not gameRunning Then Exit Sub
    ChangeLane 1
End Sub

Private Sub ChangeLane(ByVal direction As Long)
    Dim ws As Worksheet
    Dim lane As Long

    Set ws = ThisWorkbook.Worksheets("Orbital")

    lane = Val(ws.Range("B20").Value)
    If lane = 0 Then lane = 2

    lane = lane + direction

    If lane < FIRST_LANE Then lane = FIRST_LANE
    If lane > LAST_LANE Then lane = LAST_LANE

    ws.Range("B20").Value = lane
    RenderSatellite ws, lane
End Sub

Private Sub RenderSatellite(ByVal ws As Worksheet, ByVal lane As Long)
    Dim positions As Variant
    Dim i As Long

    positions = Array("D16", "F16", "H16")

    For i = LBound(positions) To UBound(positions)
        With ws.Range(positions(i))
            .ClearContents
            .Font.Name = "Consolas"
            .Font.Color = RGB(80, 255, 190)
            .Font.Size = 16
            .HorizontalAlignment = xlCenter
        End With
    Next i

    With ws.Range(positions(lane - 1))
        .Value = PLAYER_SYMBOL
        .Font.Bold = True
    End With
End Sub

Public Sub StartGame()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    StopAutoGame

    gameOver = False
    debrisLane = 0
    debrisRow = 0

    ws.Range("D6:D16,F6:F16,H6:H16").ClearContents
    ws.Range("B21").Value = "MISSION ACTIVE"
    ws.Range("B21").Font.Name = "Consolas"
    ws.Range("B21").Font.Color = RGB(80, 255, 190)

    DrawLanes
    PlaceSatellite

    Randomize
    SpawnDebris
End Sub

Private Sub SpawnDebris()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    debrisLane = Int(Rnd * 3) + 1
    debrisRow = FIRST_DEBRIS_ROW

    DrawDebris
End Sub

Private Sub DrawDebris()
    Dim ws As Worksheet
    Dim positions As Variant

    Set ws = ThisWorkbook.Worksheets("Orbital")
    positions = Array("D", "F", "H")

    With ws.Range(positions(debrisLane - 1) & debrisRow)
        .Value = DEBRIS_SYMBOL
        .Font.Name = "Consolas"
        .Font.Color = RGB(255, 100, 100)
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub GameTick()
    Dim ws As Worksheet
    Dim positions As Variant

    If gameOver Or debrisLane = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("Orbital")
    positions = Array("D", "F", "H")

    ws.Range(positions(debrisLane - 1) & debrisRow).ClearContents
    debrisRow = debrisRow + 1

    If debrisRow >= PLAYER_ROW Then
        If debrisLane = Val(ws.Range("B20").Value) Then
            gameOver = True
            ws.Range("B21").Value = "MISSION FAILED"
            ws.Range("B21").Font.Color = RGB(255, 100, 100)

            MsgBox "Satellite hit by debris!", vbExclamation, "ORBITAL"
            Exit Sub
        End If

        SpawnDebris
    Else
        DrawDebris
    End If
End Sub

Public Sub StartAutoGame()
    Dim ws As Worksheet

    StopAutoGame
    StartGame

    Set ws = ThisWorkbook.Worksheets("Orbital")

    gameRunning = True
    ScheduleNextTick
End Sub

Private Sub ScheduleNextTick()
    If Not gameRunning Then Exit Sub

    nextTick = Now + TimeSerial(0, 0, TICK_SECONDS)

    Application.OnTime _
        EarliestTime:=nextTick, _
        Procedure:="AutoGameTick", _
        Schedule:=True
End Sub

Public Sub AutoGameTick()
    If Not gameRunning Then Exit Sub

    GameTick

    If gameOver Then
        gameRunning = False
    Else
        ScheduleNextTick
    End If
End Sub

Public Function IsGameOver() As Boolean
    IsGameOver = gameOver
End Function

Public Sub StopAutoGame()
    If gameRunning Then
        gameRunning = False

        On Error Resume Next
        Application.OnTime _
            EarliestTime:=nextTick, _
            Procedure:="AutoGameTick", _
            Schedule:=False
        On Error GoTo 0
    End If
End Sub

Public Sub CreateGameButtons()
    Dim ws As Worksheet
    Dim btn As Shape

    Set ws = ThisWorkbook.Worksheets("Orbital")

    On Error Resume Next
    ws.Shapes("btnStart").Delete
    ws.Shapes("btnStop").Delete
    On Error GoTo 0

    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("B24").Left, ws.Range("B24").Top, 145, 35)

    With btn
        .Name = "btnStart"
        .TextFrame.Characters.Text = "START / RESTART"
        .OnAction = "'" & ThisWorkbook.Name & "'!StartAutoGame"
        .Fill.ForeColor.RGB = RGB(20, 35, 55)
        .Line.ForeColor.RGB = RGB(80, 255, 190)
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(80, 255, 190)
    End With

    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("F24").Left, ws.Range("F24").Top, 100, 35)

    With btn
        .Name = "btnStop"
        .TextFrame.Characters.Text = "STOP"
        .OnAction = "'" & ThisWorkbook.Name & "'!StopAutoGame"
        .Fill.ForeColor.RGB = RGB(20, 35, 55)
        .Line.ForeColor.RGB = RGB(255, 100, 100)
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 100, 100)
    End With
End Sub
