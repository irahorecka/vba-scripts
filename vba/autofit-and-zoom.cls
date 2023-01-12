' Autofit and set zoom for every sheet in the Excel workbook

Sub zoomLevel()
    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Select
        ActiveWindow.Zoom = 170 ' change as per your requirements
        ws.Columns.AutoFit
    Next ws
End Sub

