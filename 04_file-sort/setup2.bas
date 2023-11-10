Attribute VB_Name = "setup2"
Function revival0()
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 100
    ActiveSheet.Name = "log"
    Cells(2, 7).Value = ""
    Cells(2, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(2, 7).Font.Name = "游ゴシック"
    Cells(3, 7).Value = ""
    Cells(3, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(3, 7).Font.Name = "游ゴシック"
    Rows(2).RowHeight = 18.75
    Rows(3).RowHeight = 18.75
    Columns(7).ColumnWidth = 15.25
    Dim onShape As Object
End Function
Sub revival()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    CALL revival0()
    Worksheets(1).select
end sub