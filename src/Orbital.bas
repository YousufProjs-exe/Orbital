
Option Explicit

Sub SetupBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Orbital")

    ws.Cells.ColumnWidth = 3
    ws.Rows.RowHeight = 20
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Range("B2:J2").Merge
    ws.Range("B2").Value = "ORBITAL"

    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Bold = True

    ws.Range("B2:J18").Interior.Color = RGB(10, 15, 28)
    ws.Range("B2").Font.Color = RGB(100, 200, 255)

End Sub
