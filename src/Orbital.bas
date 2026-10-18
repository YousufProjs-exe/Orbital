
Option Explicit

Private Const SHEET_NAME As String = "Orbital"
Private Const PLAYER_ROW As Long = 20
Private Const SAFE_ROW As Long = 19
Private Const FIRST_DEBRIS_ROW As Long = 7
Private Const LEVEL_SECONDS As Long = 15

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

Private Function AssetPath(ByVal fileName As String) As String
    Dim rootPath As String

    rootPath = ThisWorkbook.Path
    If Len(rootPath) = 0 Then rootPath = CurDir$

    AssetPath = rootPath & Application.PathSeparator & "assets" & _
                Application.PathSeparator & fileName
End Function

Public Function AssetsAvailable() As Boolean
    Dim requiredFiles As Variant
    Dim fileName As Variant
    Dim fullPath As String

    requiredFiles = Array("rocket1.png", "rocket2.png", _
                          "stone1.png", "stone2.png", "space_bg.png")

    For Each fileName In requiredFiles
        fullPath = AssetPath(CStr(fileName))

        If Len(Dir$(fullPath)) = 0 Then
            MsgBox "Missing ORBITAL asset:" & vbCrLf & fullPath & vbCrLf & _
                   "Extract the assets folder beside ORBITAL.xlsm.", _
                   vbExclamation, "ORBITAL Assets"
            AssetsAvailable = False
            Exit Function
        End If
    Next fileName

    AssetsAvailable = True
End Function

Public Sub ApplySpaceBackground()
    Dim ws As Worksheet
    Dim fullPath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    fullPath = AssetPath("space_bg.png")

    If Len(Dir$(fullPath)) > 0 Then
        On Error Resume Next
        ws.SetBackgroundPicture fullPath
        On Error GoTo 0
    End If
End Sub

Private Function LaneCount() As Long
    If currentLevel >= 6 Then
        LaneCount = 2
    Else
        LaneCount = 3
    End If
End Function

Private Function LaneColumn(ByVal lane As Long) As Long
    If currentLevel >= 6 Then
        LaneColumn = 5 + (lane - 1) * 2
    Else
        LaneColumn = 4 + (lane - 1) * 2
    End If
End Function

Private Sub DeleteShapeIfExists(ByVal shapeName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub ClearDebrisPictures()
    Dim i As Long

    For i = 1 To 4
        DeleteShapeIfExists "sprDebris" & CStr(i)
    Next i
End Sub

Private Sub AddPictureToCell(ByVal shapeName As String, _
                             ByVal fileName As String, _
                             ByVal targetCell As Range, _
                             ByVal maxWidth As Single, _
                             ByVal maxHeight As Single)
    Dim ws As Worksheet
    Dim pictureShape As Shape
    Dim fullPath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    fullPath = AssetPath(fileName)

    If Len(Dir$(fullPath)) = 0 Then Exit Sub

    DeleteShapeIfExists shapeName

    Set pictureShape = ws.Shapes.AddPicture( _
        Filename:=fullPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=-1, _
        Height:=-1)

    With pictureShape
        .Name = shapeName
        .LockAspectRatio = msoTrue
        .Height = maxHeight

        If .Width > maxWidth Then .Width = maxWidth

        .Left = targetCell.Left + (targetCell.Width - .Width) / 2
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
        .Placement = xlMove
        .ZOrder msoBringToFront
    End With
End Sub

Public Sub SetupBoard()
    Dim ws As Worksheet
    Dim i As Long

    On Error GoTo SetupError
    Application.ScreenUpdating = False

    StopAutoGame
    DisableShipControls

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Application.Run "'" & ThisWorkbook.Name & "'!RestoreAllGridlines"

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ws.Range("B2:N25").UnMerge
    ws.Range("B2:N25").Clear

    With ws.Range("B2:J24")
        .Interior.Pattern = xlNone
        .Font.Name = "Consolas"

        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(39, 53, 65)
            .Weight = xlThin
        End With
    End With

    ws.Range("K2:N25").Interior.Pattern = xlNone

    ws.Columns("B:N").ColumnWidth = 9
    ws.Rows("2:25").RowHeight = 25
    ws.Rows(2).RowHeight = 40
    ws.Rows(3).RowHeight = 26
    ws.Rows(4).RowHeight = 30
    ws.Rows(5).RowHeight = 12
    ws.Rows(6).RowHeight = 24
    ws.Rows("7:20").RowHeight = 25
    ws.Rows(21).RowHeight = 12
    ws.Rows(22).RowHeight = 22
    ws.Rows(23).RowHeight = 26
    ws.Rows(24).RowHeight = 20

    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 90

    With ws.Range("D2:J2")
        .Merge
        .Value = "ORBITAL"
        .Font.Name = "Consolas"
        .Font.Size = 23
        .Font.Bold = True
        .Font.Color = RGB(196, 207, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:J3")
        .Merge
        .Value = "SATELLITE SURVIVAL"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(151, 165, 173)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B4:J4")
        .Merge
        .Value = "MISSION READY"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(128, 164, 145)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B22:J22")
        .Merge
        .Value = "LEFT / A       MOVE       RIGHT / D"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(151, 165, 173)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B24:J24")
        .Merge
        .Value = "ORBITAL / SURVIVE THE ORBIT"
        .Font.Name = "Consolas"
        .Font.Size = 8
        .Font.Color = RGB(100, 119, 128)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    DrawLanes
    CreateGameButtons

    Application.ScreenUpdating = True
    Exit Sub

SetupError:
    Application.ScreenUpdating = True
    MsgBox "Board setup failed: " & Err.Description, vbExclamation, "ORBITAL"
End Sub

Public Sub DrawLanes()
    Dim ws As Worksheet
    Dim lane As Long
    Dim r As Long
    Dim c As Long
    Dim laneCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    ws.Range("D6:J20").UnMerge
    ws.Range("D6:J20").ClearContents

    With ws.Range("D6:J20")
        .Interior.Pattern = xlNone

        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(39, 53, 65)
            .Weight = xlThin
        End With
    End With

    For lane = 1 To LaneCount()
        c = LaneColumn(lane)
        Set laneCell = ws.Cells(6, c)

        With laneCell
            .Value = "L" & Format$(lane, "00")
            .Interior.Color = RGB(28, 42, 51)
            .Font.Name = "Consolas"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(128, 164, 145)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With laneCell.Borders
            .LineStyle = xlContinuous
            .Color = RGB(58, 77, 87)
            .Weight = xlThin
        End With

        For r = FIRST_DEBRIS_ROW To PLAYER_ROW
            Set laneCell = ws.Cells(r, c)

            With laneCell
                .Interior.Color = RGB(16, 29, 40)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Name = "Consolas"
                .Font.Size = 18
            End With

            With laneCell.Borders
                .LineStyle = xlContinuous
                .Color = RGB(45, 65, 78)
                .Weight = xlThin
            End With

            If r = SAFE_ROW Then
                laneCell.Interior.Color = RGB(25, 43, 52)

                With laneCell.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = RGB(91, 139, 143)
                    .Weight = xlMedium
                End With
            End If
        Next r
    Next lane

    If currentLevel >= 6 Then
        With ws.Range("H6:J6")
            .Merge
            .Value = "LANE 3 CLOSED"
            .Font.Name = "Consolas"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(225, 174, 131)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

Public Sub PlaceSatellite()
    playerLane = 2
    RenderSatellite ThisWorkbook.Worksheets(SHEET_NAME)
End Sub

Private Sub RenderSatellite(ByVal ws As Worksheet)
    Dim targetCell As Range
    Dim fileName As String

    If playerLane < 1 Then playerLane = 1
    If playerLane > LaneCount() Then playerLane = LaneCount()

    DeleteShapeIfExists "sprPlayer"
    Set targetCell = ws.Cells(PLAYER_ROW, LaneColumn(playerLane))

    If currentLevel >= 6 Then
        fileName = "rocket2.png"
    Else
        fileName = "rocket1.png"
    End If

    AddPictureToCell "sprPlayer", fileName, targetCell, _
                     targetCell.Width - 8, targetCell.Height - 2
End Sub

Public Sub MoveLeft()
    If Not gameRunning Or gameOver Then Exit Sub

    If playerLane > 1 Then
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
    gameOver = False
    debrisCount = 0
    currentLevel = 1
    survivalSeconds = 0
    currentStreak = 0
    spawnCounter = 0
    playerLane = 2

    Erase debrisLanes
    Erase debrisRows
    ClearDebrisPictures

    DrawLanes
    PlaceSatellite
    UpdateStatus

    If Not randomInitialized Then
        Randomize
        randomInitialized = True
    End If

    SpawnDebris
    DrawAllDebrisPictures
End Sub

Public Sub StartAutoGame()
    On Error GoTo StartError

    If Not AssetsAvailable() Then Exit Sub

    StopAutoGame
    DisableShipControls
    gameRunning = False
    gameOver = False
    currentLevel = 1

    SetupBoard
    StartGame

    gameRunning = True
    EnableShipControls
    ScheduleNextTick
    Exit Sub

StartError:
    gameRunning = False
    DisableShipControls
    MsgBox "Could not start ORBITAL: " & Err.Description, vbExclamation, "ORBITAL"
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

Public Sub TogglePauseResume()
    Dim ws As Worksheet
    Dim btn As Shape

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If gameOver Then Exit Sub

    If gameRunning Then
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
        Set btn = Nothing

        On Error Resume Next
        Set btn = ws.Shapes("btnStop")
        On Error GoTo 0

        If Not btn Is Nothing Then
            btn.TextFrame.Characters.Text = "RESUME"
            btn.Fill.ForeColor.RGB = RGB(63, 75, 55)
            btn.TextFrame.Characters.Font.Color = RGB(211, 211, 165)
        End If

        With ws.Range("B4")
            .Value = "MISSION PAUSED / SURVIVAL " & FormatTime(survivalSeconds)
            .Font.Color = RGB(225, 174, 131)
        End With
    Else
        gameRunning = True
        EnableShipControls

        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes("btnStop")
        On Error GoTo 0

        If Not btn Is Nothing Then
            btn.TextFrame.Characters.Text = "STOP"
            btn.Fill.ForeColor.RGB = RGB(61, 47, 41)
            btn.TextFrame.Characters.Font.Color = RGB(225, 174, 131)
        End If

        UpdateStatus
        ScheduleNextTick
    End If
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

    If gameRunning And Not gameOver Then ScheduleNextTick
End Sub

Private Sub SpawnDebris()
    Dim lane As Long
    Dim i As Long
    Dim attempts As Long
    Dim laneUsed(1 To 3) As Boolean
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
    If wanted > LaneCount() Then wanted = LaneCount()

    For i = 1 To debrisCount
        If debrisRows(i) = FIRST_DEBRIS_ROW Then
            If debrisLanes(i) >= 1 And debrisLanes(i) <= LaneCount() Then
                laneUsed(debrisLanes(i)) = True
            End If
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
        End If
    Next i
End Sub

Private Sub DrawDebris(ByVal index As Long)
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim fileName As String
    Dim lane As Long

    lane = debrisLanes(index)
    If lane < 1 Or lane > LaneCount() Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set targetCell = ws.Cells(debrisRows(index), LaneColumn(lane))

    If currentLevel >= 6 Then
        fileName = "stone2.png"
    Else
        fileName = "stone1.png"
    End If

    AddPictureToCell "sprDebris" & CStr(index), fileName, _
                     targetCell, targetCell.Width - 8, targetCell.Height - 3
End Sub

Private Sub DrawAllDebrisPictures()
    Dim i As Long

    ClearDebrisPictures

    For i = 1 To debrisCount
        DrawDebris i
    Next i
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

Private Sub RemoveInactiveLaneDebris()
    Dim i As Long

    For i = debrisCount To 1 Step -1
        If debrisLanes(i) > LaneCount() Then RemoveDebris i
    Next i
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

    If currentStreak > bestStreak Then bestStreak = currentStreak

    If survivalSeconds Mod LEVEL_SECONDS = 0 Then
        currentLevel = currentLevel + 1

        If currentLevel = 6 Then
            If playerLane > 2 Then playerLane = 2
            RemoveInactiveLaneDebris
            ClearDebrisPictures
            DrawLanes
            RenderSatellite ws
            DrawAllDebrisPictures
        End If
    End If

    stepCount = 1
    If currentLevel >= 9 Then stepCount = 2
    If currentLevel >= 17 Then stepCount = 3

    For i = debrisCount To 1 Step -1
        DeleteShapeIfExists "sprDebris" & CStr(i)
        lane = debrisLanes(i)
        rowNumber = debrisRows(i)

        For stepIndex = 1 To stepCount
            rowNumber = rowNumber + 1

            If rowNumber >= PLAYER_ROW Then
                If lane = playerLane Then
                    DrawAllDebrisPictures
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
        End If
    Next i

    SpawnDebris
    DrawAllDebrisPictures
    UpdateStatus
End Sub

Private Sub EndGame()
    Dim ws As Worksheet
    Dim btn As Shape

    gameOver = True
    gameRunning = False
    currentStreak = 0
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
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set btn = Nothing

    On Error Resume Next
    Set btn = ws.Shapes("btnStop")
    On Error GoTo 0

    If Not btn Is Nothing Then btn.TextFrame.Characters.Text = "MISSION ENDED"

    With ws.Range("B4")
        .Value = "MISSION FAILED / SURVIVAL " & FormatTime(survivalSeconds)
        .Font.Color = RGB(225, 125, 112)
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
                 " / LANES " & LaneCount() & _
                 " / SURVIVAL " & FormatTime(survivalSeconds) & _
                 " / STREAK " & FormatTime(currentStreak)
        .Font.Name = "Consolas"
        .Font.Size = 8
        .Font.Color = RGB(128, 164, 145)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
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

    DeleteShapeIfExists "btnMainMenu"
    DeleteShapeIfExists "btnStart"
    DeleteShapeIfExists "btnStop"

    Set btn = ws.Shapes.AddShape( _
        msoShapeRectangle, ws.Range("B2").Left + 3, _
        ws.Range("B2").Top + 7, ws.Range("B2:C2").Width - 6, 25)

    With btn
        .Name = "btnMainMenu"
        .TextFrame.Characters.Text = "< MENU"
        .OnAction = "'" & ThisWorkbook.Name & "'!Orbital_ReturnToMenu"
        .Fill.ForeColor.RGB = RGB(34, 44, 51)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Color = RGB(196, 207, 204)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Placement = xlMove
    End With

    Set btn = ws.Shapes.AddShape( _
        msoShapeRectangle, ws.Range("C23").Left, _
        ws.Range("C23").Top + 1, 155, 23)

    With btn
        .Name = "btnStart"
        .TextFrame.Characters.Text = "RESTART MISSION"
        .OnAction = "'" & ThisWorkbook.Name & "'!StartAutoGame"
        .Fill.ForeColor.RGB = RGB(128, 164, 145)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(19, 27, 35)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Placement = xlMove
    End With

    Set btn = ws.Shapes.AddShape( _
        msoShapeRectangle, ws.Range("G23").Left, _
        ws.Range("G23").Top + 1, 115, 23)

    With btn
        .Name = "btnStop"
        .TextFrame.Characters.Text = "STOP"
        .OnAction = "'" & ThisWorkbook.Name & "'!TogglePauseResume"
        .Fill.ForeColor.RGB = RGB(61, 47, 41)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(225, 174, 131)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Placement = xlMove
    End With
End Sub
