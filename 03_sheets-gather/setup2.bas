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
    Cells(4, 7).Value = ""
    Cells(4, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(4, 7).Font.Name = "游ゴシック"
    Cells(5, 7).Value = ""
    Cells(5, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(5, 7).Font.Name = "游ゴシック"
    Cells(6, 7).Value = ""
    Cells(6, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(6, 7).Font.Name = "游ゴシック"
    Cells(7, 7).Value = ""
    Cells(7, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(7, 7).Font.Name = "游ゴシック"
    Cells(8, 7).Value = ""
    Cells(8, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(8, 7).Font.Name = "游ゴシック"
    Cells(9, 7).Value = ""
    Cells(9, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(9, 7).Font.Name = "游ゴシック"
    Cells(10, 7).Value = ""
    Cells(10, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(10, 7).Font.Name = "游ゴシック"
    Cells(11, 7).Value = ""
    Cells(11, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(11, 7).Font.Name = "游ゴシック"
    Cells(12, 7).Value = ""
    Cells(12, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(12, 7).Font.Name = "游ゴシック"
    Cells(13, 7).Value = ""
    Cells(13, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(13, 7).Font.Name = "游ゴシック"
    Cells(14, 7).Value = ""
    Cells(14, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(14, 7).Font.Name = "游ゴシック"
    Cells(15, 7).Value = ""
    Cells(15, 7).NumberFormatLocal = "yyyy/m/d h:mm"
    Cells(15, 7).Font.Name = "游ゴシック"
    Rows(2).RowHeight = 18.75
    Rows(3).RowHeight = 18.75
    Rows(4).RowHeight = 18.75
    Rows(5).RowHeight = 18.75
    Rows(6).RowHeight = 18.75
    Rows(7).RowHeight = 18.75
    Rows(8).RowHeight = 18.75
    Rows(9).RowHeight = 18.75
    Rows(10).RowHeight = 18.75
    Rows(11).RowHeight = 18.75
    Rows(12).RowHeight = 18.75
    Rows(13).RowHeight = 18.75
    Rows(14).RowHeight = 18.75
    Rows(15).RowHeight = 18.75
    Columns(7).ColumnWidth = 14.13
    Dim onShape As Object
End Function
Sub revival()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    CALL revival0()
    Worksheets(1).select
end sub