Attribute VB_Name = "Module1"

Option Explicit

Private Const FIRST_LANE As Long = 1
Private Const LAST_LANE As Long = 3
Private Const PLAYER_ROW As Long = 16

Sub SetupBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Cells.ColumnWidth = 3
    ws.Rows.RowHeight = 20
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B2:J18").Interior.Color = RGB(10, 15, 28)
    ws.Range("B2:J2").Merge

    With ws.Range("B2")
        .Value = "ORBITAL"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(100, 200, 255)
    End With
End Sub

Sub DrawLanes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("D6:D16,F6:F16,H6:H16").Interior.Color = RGB(20, 35, 55)

    With ws.Range("E5")
        .Value = "ORBITAL LANES"
        .Font.Color = RGB(100, 200, 255)
    End With
End Sub

Sub PlaceSatellite()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Range("B20").Value = 2
    ws.Range("E18").Value = "SATELLITE"
    ws.Range("E18").Font.Color = RGB(80, 255, 190)

    RenderSatellite ws, 2
End Sub

Public Sub MoveLeft()
    ChangeLane -1
End Sub

Public Sub MoveRight()
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
            .Font.Color = RGB(80, 255, 190)
            .Font.Size = 16
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ws.Range(positions(lane - 1)).Value = "?"
End Sub

