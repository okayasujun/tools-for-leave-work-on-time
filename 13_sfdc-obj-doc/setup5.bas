Attribute VB_Name = "setup5"
Function revival0()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    ActiveSheet.Name = "log"
    Cells(1, 1).Value = "No."
    Cells(1, 1).Font.Name = "���S�V�b�N"
    Cells(1, 1).Font.Bold = True
    Cells(1, 2).Value = "���e"
    Cells(1, 2).Font.Name = "���S�V�b�N"
    Cells(1, 2).Font.Bold = True
    Rows(1).RowHeight = 18.75
    Columns(1).ColumnWidth = 8.38
    Columns(2).ColumnWidth = 8.38
    Dim onShape As Object
End Function
Sub revival()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    CALL revival0()
    Worksheets(1).select
end sub