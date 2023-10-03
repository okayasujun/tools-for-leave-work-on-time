Attribute VB_Name = "F_投入CSV様式作成"
Dim itemSheet As Worksheet
Dim csvSheet As Worksheet
Sub F_投入CSV様式作成()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'データローダ用のCSVファイルを吐き出すマクロ。また腰据えてやろう・・(∩´∀｀)∩
    Set itemSheet = Sheets(ITEM_SHEET)
    
    Worksheets.Add
    'シート名を変更（追加されたシートはアクティブとなる）
    ActiveSheet.Name = "dataloader_format"
    Set csvSheet = ActiveSheet
    
    Dim writeCol As Integer: writeCol = 1
    With itemSheet
        For i = 5 To .Cells(4, 1).End(xlDown).row
            If .Cells(i, 2) = "〇" And .Cells(i, 7) <> "自動採番" And .Cells(i, 8) = "" Then
                  
                'ラベル名
                csvSheet.Cells(1, writeCol) = .Cells(i, 3).Value
                'API名
                csvSheet.Cells(2, writeCol) = .Cells(i, 5).Value
                'データ型
                csvSheet.Cells(3, writeCol) = .Cells(i, 7).Value
                '選択リスト
                csvSheet.Cells(5, writeCol) = .Cells(i, 14).Value
                '列幅調整うまくいかない
                csvSheet.Cells.EntireColumn.AutoFit
                '必須マーク
                If .Cells(i, 17) = "〇" Then
                    csvSheet.Cells(4, writeCol) = "必須！"
                End If
                If .Cells(i, 18) = "〇" Then
                    csvSheet.Cells(4, writeCol) = csvSheet.Cells(4, writeCol) & "一意！"
                End If
                writeCol = writeCol + 1
            End If
        Next
    End With
    'うまくいかない列幅調整
    csvSheet.Cells.EntireColumn.AutoFit
    
    '保存
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim saveDir As String: saveDir = ThisWorkbook.Path & "\" & csvSheet.Name & Format(Now, "yyyyddmm-hhmmss") & ".csv"
    Sheets(csvSheet.Name).Copy
    ActiveWorkbook.SaveAs saveDir, FileFormat:=xlCSV, Local:=True
    ActiveWorkbook.Close
    csvSheet.Delete
    
    MsgBox "完了しました。"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
