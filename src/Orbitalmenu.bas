
Option Explicit

Private Const MENU_SHEET As String = "Orbital"

Public Sub RestoreAllGridlines()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Orbital")
    ws.Activate

    With ws.Cells
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    ActiveWindow.DisplayGridlines = True
End Sub

Public Sub ShowMainMenu()
    Dim ws As Worksheet
    Dim cell As Range
    Dim shp As Shape
    Dim i As Long

    On Error GoTo MenuError
    Application.ScreenUpdating = False

    Application.Run "'" & ThisWorkbook.Name & "'!StopAutoGame"
    Application.Run "'" & ThisWorkbook.Name & "'!DisableShipControls"

    Set ws = ThisWorkbook.Worksheets(MENU_SHEET)
    RestoreAllGridlines

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    For Each cell In ws.Range("B2:N25").Cells
        If cell.MergeCells Then cell.MergeArea.UnMerge
    Next cell

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
    ws.Rows("2:24").RowHeight = 24
    ws.Rows(2).RowHeight = 12
    ws.Rows(3).RowHeight = 38
    ws.Rows(4).RowHeight = 10
    ws.Rows(5).RowHeight = 24
    ws.Rows(6).RowHeight = 12
    ws.Rows(7).RowHeight = 25
    ws.Rows("8:10").RowHeight = 23
    ws.Rows(11).RowHeight = 10
    ws.Rows("12:14").RowHeight = 23
    ws.Rows(15).RowHeight = 12
    ws.Rows(16).RowHeight = 24
    ws.Rows(17).RowHeight = 22
    ws.Rows("18:20").RowHeight = 16
    ws.Rows(21).RowHeight = 22
    ws.Rows(22).RowHeight = 10
    ws.Rows(23).RowHeight = 22
    ws.Rows(24).RowHeight = 16

    With ws.Range("B3:J3")
        .Merge
        .Value = "ORBITAL"
        .Interior.Pattern = xlNone
        .Font.Name = "Consolas"
        .Font.Size = 27
        .Font.Bold = True
        .Font.Color = RGB(196, 207, 204)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With ws.Range("B5:J5")
        .Merge
        .Value = "FLIGHT SYSTEM / MISSION 001"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(128, 164, 145)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With ws.Range("B7:J7")
        .Merge
        .Value = "MISSION BRIEFING"
        .Font.Name = "Consolas"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(225, 174, 131)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With ws.Range("B8:J10")
        .Merge
        .Value = "Debris field detected." & vbCrLf & _
                 "Navigate the satellite through the orbit." & vbCrLf & _
                 "Survive. Adapt. Keep moving."
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(196, 207, 204)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .WrapText = True
    End With

    With ws.Range("B12:J12")
        .Merge
        .Value = "CRAFT     /     A"
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(128, 164, 145)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With ws.Range("B13:J13")
        .Merge
        .Value = "CONTROL   /   ARROW KEYS OR A / D"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(151, 165, 173)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With ws.Range("B14:J14")
        .Merge
        .Value = "CONFIG    /   3 LANES, THEN 2 AT LEVEL 6"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(151, 165, 173)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    Set shp = ws.Shapes.AddShape( _
        msoShapeRectangle, _
        ws.Range("B16").Left + 5, _
        ws.Range("B16").Top + 1, _
        190, 36)

    With shp
        .Name = "btnStartMission"
        .TextFrame.Characters.Text = "START MISSION    >"
        .OnAction = "'" & ThisWorkbook.Name & "'!Orbital_StartMission"
        .Fill.ForeColor.RGB = RGB(128, 164, 145)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(19, 27, 35)
        .TextFrame.HorizontalAlignment = xlHAlignLeft
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Placement = xlMove
    End With

    With ws.Range("B21:J21")
        .Merge
        .Value = "ORBITAL / SATELLITE SURVIVAL"
        .Font.Name = "Consolas"
        .Font.Size = 8
        .Font.Color = RGB(100, 119, 128)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 90

    Application.Run "'" & ThisWorkbook.Name & "'!ApplySpaceBackground"
    Exit Sub

MenuError:
    Application.ScreenUpdating = True
    MsgBox "Menu error: " & Err.Description, vbExclamation, "ORBITAL"
End Sub

Public Sub Orbital_StartMission()
    On Error GoTo LaunchError

    Application.Run "'" & ThisWorkbook.Name & "'!StartAutoGame"
    Exit Sub

LaunchError:
    MsgBox "Could not launch ORBITAL: " & Err.Description, _
           vbExclamation, "ORBITAL"
End Sub

Public Sub Orbital_ReturnToMenu()
    On Error GoTo ReturnError

    Application.Run "'" & ThisWorkbook.Name & "'!StopAutoGame"
    Application.Run "'" & ThisWorkbook.Name & "'!DisableShipControls"
    ShowMainMenu
    Exit Sub

ReturnError:
    MsgBox "Could not return to the main menu: " & Err.Description, _
           vbExclamation, "ORBITAL"
End Sub


Public Sub LaunchOrbital()
    ThisWorkbook.Worksheets("Orbital").Activate
    ShowMainMenu
End Sub

Public Sub CreateAboutLaunchButton()
    Dim ws As Worksheet
    Dim btn As Shape

    Set ws = ThisWorkbook.Worksheets("About Orbital")

    On Error Resume Next
    ws.Shapes("btnLaunchOrbital").Delete
    On Error GoTo 0

    Set btn = ws.Shapes.AddShape( _
        msoShapeRectangle, _
        ws.Range("B20").Left, _
        ws.Range("B20").Top, _
        220, 48)

    With btn
        .Name = "btnLaunchOrbital"
        .Fill.ForeColor.RGB = RGB(25, 91, 115)
        .Line.Visible = msoFalse
        .OnAction = "'" & ThisWorkbook.Name & "'!LaunchOrbital"

        With .TextFrame2
            .TextRange.Text = "LAUNCH ORBITAL"
            .TextRange.Font.Name = "Aptos"
            .TextRange.Font.Size = 15
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
End Sub
