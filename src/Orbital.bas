
Attribute VB_Name = "Orbital"
Option Explicit

Sub SetupBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Cells.ColumnWidth = 3
    ws.Rows.RowHeight = 20
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B2:J18").Interior.Color = RGB(10, 15, 28)
    ws.Range("B2:J2").Merge
    ws.Range("B2").Value = "ORBITAL"
    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(100, 200, 255)
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

    ws.Range("D16,F16,H16").ClearContents

    With ws.Range("F16")
        .Value = "◆"
        .Font.Color = RGB(80, 255, 190)
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("E18")
        .Value = "SATELLITE"
        .Font.Color = RGB(80, 255, 190)
    End With
End Sub
