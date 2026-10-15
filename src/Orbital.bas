
Option Explicit

Private Const FIRST_LANE As Long = 1
Private Const LAST_LANE As Long = 3
Private Const PLAYER_ROW As Long = 16
Private Const FIRST_DEBRIS_ROW As Long = 6

Private debrisLane As Long
Private debrisRow As Long
Private gameOver As Boolean

Public Sub SetupBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 3
    ws.Rows.RowHeight = 20

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B2:J22").Interior.Color = RGB(10, 15, 28)

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
    ws.Range("E18").Font.Color = RGB(80, 255, 190)

    RenderSatellite ws, 2
End Sub

Public Sub MoveLeft()
    If gameOver Then Exit Sub
    ChangeLane -1
End Sub

Public Sub MoveRight()
    If gameOver Then Exit Sub
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

    ws.Range(positions(lane - 1)).Value = "◆"
End Sub

Public Sub StartGame()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    gameOver = False
    debrisLane = 0
    debrisRow = 0

    ws.Range("D6:D16,F6:F16,H6:H16").ClearContents
    ws.Range("B21").Value = "MISSION ACTIVE"
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
        .Value = "●"
        .Font.Name = "Consolas"
        .Font.Color = RGB(255, 100, 100)
        .Font.Size = 14
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
