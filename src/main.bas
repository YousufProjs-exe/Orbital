
Option Explicit

Private Const MENU_SHEET As String = "Orbital"

Public Sub ShowMainMenu()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn As Shape
    Dim i As Long

    On Error GoTo MenuError

    Set ws = ThisWorkbook.Worksheets(MENU_SHEET)
    Application.ScreenUpdating = False

    'Remove only the previous menu button.
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = "btnStartMission" Then
            ws.Shapes(i).Delete
        End If
    Next i

    'Clear only the menu area.
    With ws.Range("B2:N23")
        .UnMerge
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.Name = "Consolas"
        .Font.Color = RGB(235, 244, 255)
    End With

    'Menu panel only.
    With ws.Range("B2:N23")
        .Interior.Color = RGB(13, 29, 47)
    End With

    ws.Columns("B:N").ColumnWidth = 10
    ws.Rows("2:23").RowHeight = 24
    ws.Rows(2).RowHeight = 12

    'Title.
    With ws.Range("C4:M5")
        .Merge
        .Value = "ORBITAL"
        .Font.Size = 28
        .Font.Bold = True
        .Font.Color = RGB(73, 217, 154)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Subtitle.
    With ws.Range("C6:M6")
        .Merge
        .Value = "SATELLITE SURVIVAL PROGRAM"
        .Font.Size = 10
        .Font.Color = RGB(145, 169, 198)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Mission introduction.
    With ws.Range("D9:L10")
        .Merge
        .Value = "DODGE THE DEBRIS." & vbLf & _
                 "SURVIVE ALL SIX LEVELS."
        .WrapText = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Sprite preview placeholder.
    With ws.Range("F12:J15")
        .Merge
        .Value = "[ SATELLITE PREVIEW ]"
        .Font.Size = 11
        .Font.Color = RGB(73, 217, 154)
        .Interior.Color = RGB(23, 40, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Start button.
    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("F17").Left, _
        ws.Range("F17").Top, _
        ws.Range("F17:J18").Width, _
        ws.Range("F17:J18").Height)

    With btn
        .Name = "btnStartMission"
        .TextFrame2.TextRange.Text = "START MISSION"
        .TextFrame2.TextRange.Font.Name = "Consolas"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(13, 29, 47)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = RGB(73, 217, 154)
        .Line.Visible = msoFalse
        .OnAction = "'" & ThisWorkbook.Name & "'!StartMission"
        .Placement = xlMove
    End With

    'Controls.
    With ws.Range("C21:M22")
        .Merge
        .Value = "CONTROLS     LEFT / A: MOVE LEFT     RIGHT / D: MOVE RIGHT"
        .Font.Size = 9
        .Font.Color = RGB(145, 169, 198)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 90
    ws.Range("B2").Select

    Application.ScreenUpdating = True
    Exit Sub

MenuError:
    Application.ScreenUpdating = True
    MsgBox "Could not display the ORBITAL menu:" & vbCrLf & _
           Err.Description, vbExclamation, "ORBITAL"

End Sub

Public Sub StartMission()

    On Error GoTo StartError

    Application.Run "'" & ThisWorkbook.Name & "'!StartAutoGame"
    Exit Sub

StartError:
    MsgBox "Mission launch failed." & vbCrLf & _
           "Check that StartAutoGame exists in a standard VBA module." & _
           vbCrLf & vbCrLf & "Details: " & Err.Description, _
           vbExclamation, "ORBITAL"

End Sub
