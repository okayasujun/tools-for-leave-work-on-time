Attribute VB_Name = "setup2"
Function revival0()
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 100
    ActiveSheet.Name = "log"
    Cells(1, 1).Value = ""
    Cells(1, 1).Font.Name = "游ゴシック"
    Cells(1, 2).Value = ""
    Cells(1, 2).Font.Name = "游ゴシック"
    Cells(1, 3).Value = ""
    Cells(1, 3).Font.Name = "游ゴシック"
    Cells(1, 4).Value = ""
    Cells(1, 4).Font.Name = "游ゴシック"
    Cells(1, 5).Value = ""
    Cells(1, 5).Font.Name = "游ゴシック"
    Cells(1, 6).Value = ""
    Cells(1, 6).Font.Name = "游ゴシック"
    Cells(1, 7).Value = ""
    Cells(1, 7).Font.Name = "游ゴシック"
    Cells(1, 8).Value = ""
    Cells(1, 8).Font.Name = "游ゴシック"
    Cells(1, 9).Value = ""
    Cells(1, 9).Font.Name = "游ゴシック"
    Cells(1, 10).Value = ""
    Cells(1, 10).Font.Name = "游ゴシック"
    Cells(1, 11).Value = ""
    Cells(1, 11).Font.Name = "游ゴシック"
    Cells(1, 12).Value = ""
    Cells(1, 12).Font.Name = "游ゴシック"
    Cells(1, 13).Value = ""
    Cells(1, 13).Font.Name = "游ゴシック"
    Cells(1, 14).Value = ""
    Cells(1, 14).Font.Name = "游ゴシック"
    Cells(1, 15).Value = ""
    Cells(1, 15).Font.Name = "游ゴシック"
    Cells(1, 16).Value = ""
    Cells(1, 16).Font.Name = "游ゴシック"
    Rows(1).RowHeight = 18.75
    Columns(1).ColumnWidth = 3.75
    Columns(2).ColumnWidth = 41.63
    Columns(3).ColumnWidth = 25.75
    Columns(4).ColumnWidth = 101.5
    Columns(5).ColumnWidth = 10.38
    Columns(6).ColumnWidth = 8.38
    Columns(7).ColumnWidth = 20
    Columns(8).ColumnWidth = 8.38
    Columns(9).ColumnWidth = 8.38
    Columns(10).ColumnWidth = 8.38
    Columns(11).ColumnWidth = 8.38
    Columns(12).ColumnWidth = 8.38
    Columns(13).ColumnWidth = 8.38
    Columns(14).ColumnWidth = 8.38
    Columns(15).ColumnWidth = 8.38
    Columns(16).ColumnWidth = 8.38
    Dim onShape As Object
End Function
Sub revival()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    CALL revival0()
    Worksheets(1).select
end sub