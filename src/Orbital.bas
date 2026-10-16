Option Explicit

Private Const SHEET_NAME As String = "Orbital"
Private Const FIRST_LANE As Long = 1
Private Const PLAYER_ROW As Long = 20
Private Const SAFE_ROW As Long = 19
Private Const FIRST_DEBRIS_ROW As Long = 7
Private Const LEVEL_SECONDS As Long = 15

Private Const PLAYER_SYMBOL As String = "A"
Private Const DEBRIS_SYMBOL As String = "X"

Private debrisLanes(1 To 4) As Long
Private debrisRows(1 To 4) As Long
Private debrisCount As Long
Private spawnCounter As Long

Private playerLane As Long
Private currentLevel As Long
Private survivalSeconds As Long
Private currentStreak As Long
Private bestStreak As Long

Private gameOver As Boolean
Private gameRunning As Boolean
Private nextTick As Date
Private tickScheduled As Boolean
Private randomInitialized As Boolean

Private Function LaneColumn(ByVal lane As Long) As Long
    LaneColumn = 4 + (lane - 1) * 2
End Function

Private Function LaneCount() As Long
    If currentLevel >= 6 Then
        LaneCount = 4
    Else
        LaneCount = 3
    End If
End Function

Public Sub SetupBoard()
    Dim ws As Worksheet
    Dim cell As Range
    Dim i As Long

    On Error GoTo SetupError

    StopAutoGame
    DisableShipControls

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    For i = ws.Shapes.Count To 1 Step -1
        Select Case ws.Shapes(i).Name
            Case "btnStartMission", "btnStart", "btnStop"
                ws.Shapes(i).Delete
        End Select
    Next i

    For Each cell In ws.Range("B2:N25").Cells
        If cell.MergeCells Then
            cell.MergeArea.UnMerge
        End If
    Next cell

    With ws.Range("B2:N25")
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    ws.Columns("B:N").ColumnWidth = 8
    ws.Rows("2:25").RowHeight = 25

    ws.Columns("D").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 10

    ws.Rows(2).RowHeight = 40
    ws.Rows(3).RowHeight = 26
    ws.Rows(4).RowHeight = 30
    ws.Rows("7:20").RowHeight = 30

    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 90

    With ws.Range("B2:J2")
        .Merge
        .Value = "O R B I T A L"
        .Interior.Color = RGB(13, 29, 47)
        .Font.Name = "Consolas"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(73, 217, 154)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:J3")
        .Merge
        .Value = "SATELLITE SURVIVAL"
        .Interior.Color = RGB(23, 40, 59)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(145, 169, 198)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B4:J4")
        .Merge
        .Value = "MISSION READY"
        .Interior.Color = RGB(23, 40, 59)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(73, 217, 154)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    DrawLanes
    CreateGameButtons

    ws.Range("B22:J22").Merge
    With ws.Range("B22")
        .Value = "LEFT / A       MOVE       RIGHT / D"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(145, 169, 198)
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("B24:J24").Merge
    With ws.Range("B24")
        .Value = "ORBITAL // SURVIVE THE ORBIT"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(100, 130, 160)
        .HorizontalAlignment = xlCenter
    End With

    Exit Sub

SetupError:
    Application.ScreenUpdating = True
    MsgBox "Board setup failed: " & Err.Description, _
           vbExclamation, "ORBITAL"
End Sub

Public Sub DrawLanes()
    Dim ws As Worksheet
    Dim lane As Long
    Dim r As Long
    Dim c As Long
    Dim laneCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    For lane = 1 To 4
        c = LaneColumn(lane)

        For r = FIRST_DEBRIS_ROW To PLAYER_ROW
            Set laneCell = ws.Cells(r, c)
            laneCell.Interior.Color = RGB(16, 29, 48)
            laneCell.HorizontalAlignment = xlCenter
            laneCell.VerticalAlignment = xlCenter
            laneCell.Font.Name = "Consolas"
            laneCell.Font.Size = 18

            With laneCell.Borders
                .LineStyle = xlContinuous
                .Color = RGB(45, 75, 105)
                .Weight = xlThin
            End With

            If r = SAFE_ROW Then
                laneCell.Interior.Color = RGB(22, 48, 65)

                With laneCell.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = RGB(65, 190, 220)
                    .Weight = xlMedium
                End With
            End If
        Next r

        Set laneCell = ws.Cells(6, c)

        If lane <= LaneCount() Then
            laneCell.Value = "L" & Format$(lane, "00")
            laneCell.Interior.Color = RGB(23, 40, 59)
            laneCell.Font.Name = "Consolas"
            laneCell.Font.Size = 10
            laneCell.Font.Bold = True
            laneCell.Font.Color = RGB(73, 217, 154)
            laneCell.HorizontalAlignment = xlCenter
            laneCell.VerticalAlignment = xlCenter
        Else
            laneCell.ClearContents
            laneCell.Interior.Pattern = xlNone
            laneCell.Borders.LineStyle = xlNone
        End If
    Next lane
End Sub

Public Sub PlaceSatellite()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    playerLane = 2
    RenderSatellite ws
End Sub

Private Sub RenderSatellite(ByVal ws As Worksheet)
    Dim lane As Long
    Dim spriteCell As Range

    For lane = 1 To 4
        Set spriteCell = ws.Cells(PLAYER_ROW, LaneColumn(lane))
        spriteCell.ClearContents

        If lane <= LaneCount() Then
            spriteCell.Interior.Color = RGB(16, 29, 48)
        End If
    Next lane

    Set spriteCell = ws.Cells(PLAYER_ROW, LaneColumn(playerLane))

    spriteCell.Value = PLAYER_SYMBOL
    spriteCell.Font.Name = "Consolas"
    spriteCell.Font.Size = 22
    spriteCell.Font.Bold = True
    spriteCell.Font.Color = RGB(73, 217, 154)
    spriteCell.HorizontalAlignment = xlCenter
    spriteCell.VerticalAlignment = xlCenter
End Sub

Public Sub MoveLeft()
    If Not gameRunning Or gameOver Then Exit Sub

    If playerLane > FIRST_LANE Then
        playerLane = playerLane - 1
        RenderSatellite ThisWorkbook.Worksheets(SHEET_NAME)
    End If
End Sub

Public Sub MoveRight()
    If Not gameRunning Or gameOver Then Exit Sub

    If playerLane < LaneCount() Then
        playerLane = playerLane + 1
        RenderSatellite ThisWorkbook.Worksheets(SHEET_NAME)
    End If
End Sub

Public Sub EnableShipControls()
    Application.OnKey "{LEFT}", "'" & ThisWorkbook.Name & "'!MoveLeft"
    Application.OnKey "{RIGHT}", "'" & ThisWorkbook.Name & "'!MoveRight"
    Application.OnKey "a", "'" & ThisWorkbook.Name & "'!MoveLeft"
    Application.OnKey "d", "'" & ThisWorkbook.Name & "'!MoveRight"
End Sub

Public Sub DisableShipControls()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "a"
    Application.OnKey "d"
End Sub

Public Sub StartGame()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    gameOver = False
    debrisCount = 0
    currentLevel = 1
    survivalSeconds = 0
    currentStreak = 0
    spawnCounter = 0
    playerLane = 2

    Erase debrisLanes
    Erase debrisRows

    DrawLanes
    PlaceSatellite
    UpdateStatus

    If Not randomInitialized Then
        Randomize
        randomInitialized = True
    End If

    SpawnDebris
End Sub

Public Sub StartAutoGame()
    On Error GoTo StartError

    StopAutoGame
    DisableShipControls

    gameRunning = False
    SetupBoard
    StartGame

    gameRunning = True
    EnableShipControls
    ScheduleNextTick

    Exit Sub

StartError:
    gameRunning = False
    DisableShipControls

    MsgBox "Could not start ORBITAL: " & Err.Description, _
           vbExclamation, "ORBITAL"
End Sub

Public Sub StopAutoGame()
    gameRunning = False
    DisableShipControls

    If tickScheduled Then
        On Error Resume Next
        Application.OnTime _
            EarliestTime:=nextTick, _
            Procedure:="'" & ThisWorkbook.Name & "'!AutoGameTick", _
            Schedule:=False
        On Error GoTo 0
    End If

    tickScheduled = False
End Sub

Private Sub ScheduleNextTick()
    If Not gameRunning Or gameOver Then Exit Sub

    nextTick = Now + TimeSerial(0, 0, 1)

    Application.OnTime _
        EarliestTime:=nextTick, _
        Procedure:="'" & ThisWorkbook.Name & "'!AutoGameTick", _
        Schedule:=True

    tickScheduled = True
End Sub

Public Sub AutoGameTick()
    tickScheduled = False

    If Not gameRunning Or gameOver Then Exit Sub

    GameTick

    If gameRunning And Not gameOver Then
        ScheduleNextTick
    End If
End Sub

Private Sub SpawnDebris()
    Dim lane As Long
    Dim i As Long
    Dim attempts As Long
    Dim laneUsed(1 To 4) As Boolean
    Dim wanted As Long
    Dim spawnEvery As Long

    spawnEvery = 3
    If currentLevel >= 3 Then spawnEvery = 2
    If currentLevel >= 10 Then spawnEvery = 1

    spawnCounter = spawnCounter + 1
    If spawnCounter < spawnEvery Then Exit Sub
    spawnCounter = 0

    wanted = 1
    If currentLevel >= 3 Then wanted = 2
    If currentLevel >= 8 Then wanted = 3
    If currentLevel >= 15 Then wanted = 4
    If wanted > LaneCount() Then wanted = LaneCount()

    For i = 1 To debrisCount
        If debrisRows(i) = FIRST_DEBRIS_ROW Then
            laneUsed(debrisLanes(i)) = True
        End If
    Next i

    For i = 1 To wanted
        If debrisCount >= 4 Then Exit For

        attempts = 0

        Do
            lane = Int(Rnd * LaneCount()) + 1
            attempts = attempts + 1
        Loop While laneUsed(lane) And attempts < 20

        If Not laneUsed(lane) And debrisCount < 4 Then
            debrisCount = debrisCount + 1
            debrisLanes(debrisCount) = lane
            debrisRows(debrisCount) = FIRST_DEBRIS_ROW
            laneUsed(lane) = True
            DrawDebris debrisCount
        End If
    Next i
End Sub

Private Sub DrawDebris(ByVal index As Long)
    Dim ws As Worksheet
    Dim debrisCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Set debrisCell = ws.Cells( _
        debrisRows(index), LaneColumn(debrisLanes(index)))

    debrisCell.Value = DEBRIS_SYMBOL
    debrisCell.Font.Name = "Consolas"
    debrisCell.Font.Size = 18
    debrisCell.Font.Bold = True
    debrisCell.Font.Color = RGB(255, 105, 105)
    debrisCell.HorizontalAlignment = xlCenter
    debrisCell.VerticalAlignment = xlCenter
End Sub

Private Sub ClearDebrisCell(ByVal lane As Long, ByVal rowNumber As Long)
    Dim ws As Worksheet
    Dim debrisCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set debrisCell = ws.Cells(rowNumber, LaneColumn(lane))
    debrisCell.ClearContents
End Sub

Public Sub GameTick()
    Dim ws As Worksheet
    Dim i As Long
    Dim lane As Long
    Dim rowNumber As Long
    Dim stepCount As Long
    Dim stepIndex As Long

    If gameOver Or Not gameRunning Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    survivalSeconds = survivalSeconds + 1
    currentStreak = currentStreak + 1

    If currentStreak > bestStreak Then
        bestStreak = currentStreak
    End If

    If survivalSeconds Mod LEVEL_SECONDS = 0 Then
        currentLevel = currentLevel + 1

        ' for init - lane 4 befor e le 6 
        If currentLevel = 6 Then
            MsgBox "ORBITAL UPDATE" & vbCrLf & vbCrLf & _
                   "NEW LANE UNLOCKED" & vbCrLf & _
                   "4-LANE MODE ACTIVATED", _
                   vbInformation, "ORBITAL"
        End If

        DrawLanes
        RenderSatellite ws
    End If

    stepCount = 1
    If currentLevel >= 9 Then stepCount = 2
    If currentLevel >= 17 Then stepCount = 3

    For i = debrisCount To 1 Step -1
        lane = debrisLanes(i)
        rowNumber = debrisRows(i)

        ClearDebrisCell lane, rowNumber

        For stepIndex = 1 To stepCount
            rowNumber = rowNumber + 1

            If rowNumber >= PLAYER_ROW Then
                If lane = playerLane Then
                    EndGame
                    Exit Sub
                End If

                Exit For
            End If
        Next stepIndex

        If rowNumber >= PLAYER_ROW Then
            RemoveDebris i
        Else
            debrisRows(i) = rowNumber
            DrawDebris i
        End If
    Next i

    SpawnDebris
    UpdateStatus
End Sub

Private Sub RemoveDebris(ByVal index As Long)
    Dim i As Long

    For i = index To debrisCount - 1
        debrisLanes(i) = debrisLanes(i + 1)
        debrisRows(i) = debrisRows(i + 1)
    Next i

    If debrisCount > 0 Then
        debrisLanes(debrisCount) = 0
        debrisRows(debrisCount) = 0
        debrisCount = debrisCount - 1
    End If
End Sub

Private Sub EndGame()
    Dim ws As Worksheet

    gameOver = True
    gameRunning = False
    currentStreak = 0

    DisableShipControls

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    With ws.Range("B4")
        .Value = "MISSION FAILED  |  SURVIVAL: " & FormatTime(survivalSeconds)
        .Font.Color = RGB(255, 105, 105)
    End With

    MsgBox "Satellite lost!" & vbCrLf & _
           "Level reached: " & currentLevel & vbCrLf & _
           "Survival time: " & FormatTime(survivalSeconds) & vbCrLf & _
           "Best streak: " & FormatTime(bestStreak), _
           vbExclamation, "ORBITAL"
End Sub

Private Sub UpdateStatus()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    With ws.Range("B4")
        .Value = "LEVEL " & currentLevel & _
                 "  |  SURVIVAL " & FormatTime(survivalSeconds) & _
                 "  |  STREAK " & FormatTime(currentStreak)
        .Font.Color = RGB(73, 217, 154)
    End With
End Sub

Private Function FormatTime(ByVal totalSeconds As Long) As String
    FormatTime = Format$(totalSeconds \ 60, "00") & ":" & _
                 Format$(totalSeconds Mod 60, "00")
End Function

Public Function IsGameOver() As Boolean
    IsGameOver = gameOver
End Function

Public Sub CreateGameButtons()
    Dim ws As Worksheet
    Dim btn As Shape

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    On Error Resume Next
    ws.Shapes("btnStart").Delete
    ws.Shapes("btnStop").Delete
    On Error GoTo 0

    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("C23").Left, ws.Range("C23").Top, 180, 34)

    With btn
        .Name = "btnStart"
        .TextFrame.Characters.Text = "RESTART MISSION"
        .OnAction = "'" & ThisWorkbook.Name & "'!StartAutoGame"
        .Fill.ForeColor.RGB = RGB(16, 40, 55)
        .Line.ForeColor.RGB = RGB(73, 217, 154)
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(73, 217, 154)
        .Placement = xlMove
    End With

    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("G23").Left, ws.Range("G23").Top, 140, 34)

    With btn
        .Name = "btnStop"
        .TextFrame.Characters.Text = "STOP"
        .OnAction = "'" & ThisWorkbook.Name & "'!StopAutoGame"
        .Fill.ForeColor.RGB = RGB(40, 22, 30)
        .Line.ForeColor.RGB = RGB(255, 105, 105)
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 105, 105)
        .Placement = xlMove
    End With
End Sub