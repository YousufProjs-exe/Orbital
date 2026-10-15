
Option Explicit

Private Const FIRST_LANE As Long = 1
Private Const LAST_LANE As Long = 3
Private Const PLAYER_ROW As Long = 20
Private Const FIRST_DEBRIS_ROW As Long = 7
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
    ws.Cells.ColumnWidth = 10
    ws.Rows.RowHeight = 24

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85

    ws.Range("B2:N25").Interior.Color = RGB(8, 13, 27)

    With ws.Range("B2:N2")
        .Merge
        .Value = "O R B I T A L"
        .Font.Name = "Consolas"
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = RGB(100, 200, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 48
    End With

    With ws.Range("B3:N3")
        .Merge
        .Value = "SATELLITE SURVIVAL SYSTEM  //  MISSION CONTROL"
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(80, 255, 190)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With

    With ws.Range("B4:N4")
        .Merge
        .Value = "SYSTEM READY  |  AVOID INCOMING DEBRIS"
        .Font.Name = "Consolas"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(100, 200, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    With ws.Range("C6:F6")
        .Merge
        .Value = "LANE 01"
    End With

    With ws.Range("G6:J6")
        .Merge
        .Value = "LANE 02"
    End With

    With ws.Range("K6:N6")
        .Merge
        .Value = "LANE 03"
    End With

    With ws.Range("C6:N6")
        .Font.Name = "Consolas"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(100, 200, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    DrawLanes
    CreateGameButtons

    With ws.Range("B25:N25")
        .Merge
        .Value = "ORBITAL  //  SURVIVE THE ORBIT"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(100, 130, 160)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub DrawLanes()
    Dim ws As Worksheet
    Dim laneRange As Range

    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("C7:F20").Interior.Color = RGB(16, 29, 48)
    ws.Range("G7:J20").Interior.Color = RGB(16, 29, 48)
    ws.Range("K7:N20").Interior.Color = RGB(16, 29, 48)

    For Each laneRange In ws.Range("C7:F20,G7:J20,K7:N20").Areas
        With laneRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(45, 75, 105)
            .Weight = xlThin
        End With
    Next laneRange

    ws.Range("D7:D20,H7:H20,L7:L20").HorizontalAlignment = xlCenter
    ws.Range("D7:D20,H7:H20,L7:L20").VerticalAlignment = xlCenter
End Sub

Public Sub PlaceSatellite()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("B21").Value = 2
    RenderSatellite ws, 2
End Sub

Private Sub RenderSatellite(ByVal ws As Worksheet, ByVal lane As Long)
    Dim positions As Variant
    Dim i As Long

    positions = Array("D20", "H20", "L20")

    For i = LBound(positions) To UBound(positions)
        With ws.Range(positions(i))
            .ClearContents
            .Font.Name = "Consolas"
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = RGB(80, 255, 190)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    With ws.Range(positions(lane - 1))
        .Value = PLAYER_SYMBOL
        .Font.Size = 26
        .Font.Color = RGB(80, 255, 190)
    End With
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

    lane = Val(ws.Range("B21").Value)
    If lane = 0 Then lane = 2

    lane = lane + direction

    If lane < FIRST_LANE Then lane = FIRST_LANE
    If lane > LAST_LANE Then lane = LAST_LANE

    ws.Range("B21").Value = lane
    RenderSatellite ws, lane
End Sub

Public Sub StartGame()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")

    StopAutoGame

    gameOver = False
    debrisLane = 0
    debrisRow = 0

    ws.Range("D7:D20,H7:H20,L7:L20").ClearContents

    DrawLanes
    PlaceSatellite

    With ws.Range("B4")
        .Value = "MISSION ACTIVE  |  DODGE THE DEBRIS"
        .Font.Color = RGB(80, 255, 190)
    End With

    Randomize
    SpawnDebris
End Sub

Private Sub SpawnDebris()
    debrisLane = Int(Rnd * 3) + 1
    debrisRow = FIRST_DEBRIS_ROW

    DrawDebris
End Sub

Private Sub DrawDebris()
    Dim ws As Worksheet
    Dim positions As Variant

    Set ws = ThisWorkbook.Worksheets("Orbital")
    positions = Array("D", "H", "L")

    With ws.Range(positions(debrisLane - 1) & debrisRow)
        .Value = DEBRIS_SYMBOL
        .Font.Name = "Consolas"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 100, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub GameTick()
    Dim ws As Worksheet
    Dim positions As Variant

    If gameOver Or debrisLane = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("Orbital")
    positions = Array("D", "H", "L")

    ws.Range(positions(debrisLane - 1) & debrisRow).ClearContents
    debrisRow = debrisRow + 1

    If debrisRow >= PLAYER_ROW Then
        If debrisLane = Val(ws.Range("B21").Value) Then
            gameOver = True

            With ws.Range("B4")
                .Value = "MISSION FAILED  |  SATELLITE IMPACT"
                .Font.Color = RGB(255, 100, 100)
            End With

            MsgBox "Satellite hit by debris!", vbExclamation, "ORBITAL"
            Exit Sub
        End If

        SpawnDebris
    Else
        DrawDebris
    End If
End Sub

Public Sub StartAutoGame()
    StopAutoGame
    StartGame

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
        ws.Range("C22").Left, ws.Range("C22").Top, 220, 48)

    With btn
        .Name = "btnStart"
        .TextFrame.Characters.Text = "START / RESTART MISSION"
        .OnAction = "'" & ThisWorkbook.Name & "'!StartAutoGame"
        .Fill.ForeColor.RGB = RGB(16, 40, 55)
        .Line.ForeColor.RGB = RGB(80, 255, 190)
        .Line.Weight = 2
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(80, 255, 190)
        .Placement = xlMove
    End With

    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("I22").Left, ws.Range("I22").Top, 150, 48)

    With btn
        .Name = "btnStop"
        .TextFrame.Characters.Text = "STOP MISSION"
        .OnAction = "'" & ThisWorkbook.Name & "'!StopAutoGame"
        .Fill.ForeColor.RGB = RGB(40, 22, 30)
        .Line.ForeColor.RGB = RGB(255, 100, 100)
        .Line.Weight = 2
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 100, 100)
        .Placement = xlMove
    End With
End Sub
