Attribute VB_Name = "setup2"
Function revival0()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ActiveSheet.Name = "オブジェクト"
    Cells(1, 1) = "■オブジェクト情報"
    Cells(1, 1).Font.Name = "游ゴシック"
    Cells(1, 2).Font.Name = "游ゴシック"
    Cells(1, 3).Font.Name = "游ゴシック"
    Cells(1, 4).Font.Name = "游ゴシック"
    Cells(1, 5) = "下記以降は開発用カラム。最終的には不要だけど。それは完成してから。"
    Cells(1, 5).Font.Name = "游ゴシック"
    Cells(1, 6).Font.Name = "游ゴシック"
    Cells(1, 7).Font.Name = "游ゴシック"
    Cells(2, 1) = "カスタムオブジェクトの情報"
    Cells(2, 1).Font.Name = "游ゴシック"
    Cells(2, 1).Interior.Color = 13431551
    Cells(2, 1).Font.Bold = True
    Cells(2, 2).Font.Name = "游ゴシック"
    Cells(2, 2).Interior.Color = 13431551
    Cells(2, 2).Font.Bold = True
    Cells(2, 3).Font.Name = "游ゴシック"
    Cells(2, 3).Interior.Color = 13431551
    Cells(2, 4) = "システム用（編集不要）"
    Cells(2, 4).Font.Name = "游ゴシック"
    Cells(2, 4).Interior.Color = 13431551
    Cells(2, 5).Font.Name = "游ゴシック"
    Cells(2, 6).Font.Name = "游ゴシック"
    Cells(2, 7) = "設定タイプ"
    Cells(2, 7).Font.Name = "游ゴシック"
    Cells(3, 1) = "表示ラベル"
    Cells(3, 1).Font.Name = "游ゴシック"
    Cells(3, 2) = "テスト02"
    Cells(3, 2).Font.Name = "游ゴシック"
    Cells(3, 3).Font.Name = "游ゴシック"
    Cells(3, 4) = "テスト02"
    Cells(3, 4).Font.Name = "游ゴシック"
    Cells(3, 5) = "OK"
    Cells(3, 5).Font.Name = "游ゴシック"
    Cells(3, 6).Font.Name = "游ゴシック"
    Cells(3, 7) = "テキスト"
    Cells(3, 7).Font.Name = "游ゴシック"
    Cells(4, 1) = "オブジェクト名"
    Cells(4, 1).Font.Name = "游ゴシック"
    Cells(4, 2) = "TestTestTest02"
    Cells(4, 2).Font.Name = "游ゴシック"
    Cells(4, 3).Font.Name = "游ゴシック"
    Cells(4, 4) = "TestTestTest02__c"
    Cells(4, 4).Font.Name = "游ゴシック"
    Cells(4, 5) = "ファイル内には表れない"
    Cells(4, 5).Font.Name = "游ゴシック"
    Cells(4, 6).Font.Name = "游ゴシック"
    Cells(4, 7) = "テキスト"
    Cells(4, 7).Font.Name = "游ゴシック"
    Cells(5, 1) = "説明"
    Cells(5, 1).Font.Name = "游ゴシック"
    Cells(5, 2) = "おためし中・・・！(∩´∀｀)∩"
    Cells(5, 2).WrapText  = True
    Cells(5, 2).Font.Name = "游ゴシック"
    Range("B5:B6").Merge
    Cells(5, 3).Font.Name = "游ゴシック"
    Cells(5, 4) = "おためし中・・・！(∩´∀｀)∩"
    Cells(5, 4).WrapText  = True
    Cells(5, 4).Font.Name = "游ゴシック"
    Range("D5:D6").Merge
    Cells(5, 5) = "OK"
    Cells(5, 5).Font.Name = "游ゴシック"
    Cells(5, 6).Font.Name = "游ゴシック"
    Cells(5, 7) = "テキスト"
    Cells(5, 7).Font.Name = "游ゴシック"
    Cells(6, 1).Font.Name = "游ゴシック"
    Cells(6, 2).WrapText  = True
    Cells(6, 2).Font.Name = "游ゴシック"
    Range("B5:B6").Merge
    Cells(6, 3).Font.Name = "游ゴシック"
    Cells(6, 4).WrapText  = True
    Cells(6, 4).Font.Name = "游ゴシック"
    Range("D5:D6").Merge
    Cells(6, 5).Font.Name = "游ゴシック"
    Cells(6, 6).Font.Name = "游ゴシック"
    Cells(6, 7).Font.Name = "游ゴシック"
    Cells(7, 1) = "カスタムベルプの設定"
    Cells(7, 1).Font.Name = "游ゴシック"
    Cells(7, 2) = "Salesforce 標準の [ヘルプ & トレーニング] ウィンドウを開く"
    Cells(7, 2).Font.Name = "游ゴシック"
    Cells(7, 2).Validation.Delete
    Cells(7, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$I$2:$I$3"
    Cells(7, 3).Font.Name = "游ゴシック"
    Cells(7, 4).Font.Name = "游ゴシック"
    Cells(7, 5) = "保留"
    Cells(7, 5).Font.Name = "游ゴシック"
    Cells(7, 6) = "metaDataファイルに反映されない"
    Cells(7, 6).Font.Name = "游ゴシック"
    Cells(7, 7) = "リスト"
    Cells(7, 7).Font.Name = "游ゴシック"
    Cells(8, 1) = "コンテンツ名"
    Cells(8, 1).Font.Name = "游ゴシック"
    Cells(8, 2).Font.Name = "游ゴシック"
    Cells(8, 3).Font.Name = "游ゴシック"
    Cells(8, 4).Font.Name = "游ゴシック"
    Cells(8, 5) = "保留"
    Cells(8, 5).Font.Name = "游ゴシック"
    Cells(8, 6) = "metaDataファイルに反映されない"
    Cells(8, 6).Font.Name = "游ゴシック"
    Cells(8, 7) = "テキスト"
    Cells(8, 7).Font.Name = "游ゴシック"
    Cells(9, 1).Font.Name = "游ゴシック"
    Cells(9, 2).Font.Name = "游ゴシック"
    Cells(9, 3).Font.Name = "游ゴシック"
    Cells(9, 4).Font.Name = "游ゴシック"
    Cells(9, 5).Font.Name = "游ゴシック"
end Function
Function revival1()
    Cells(9, 6).Font.Name = "游ゴシック"
    Cells(9, 7).Font.Name = "游ゴシック"
    Cells(10, 1) = "レコード名の表示ラベルと型を入力"
    Cells(10, 1).Font.Name = "游ゴシック"
    Cells(10, 1).Interior.Color = 13431551
    Cells(10, 1).Font.Bold = True
    Cells(10, 2).Font.Name = "游ゴシック"
    Cells(10, 2).Interior.Color = 13431551
    Cells(10, 2).Font.Bold = True
    Cells(10, 3).Font.Name = "游ゴシック"
    Cells(10, 3).Interior.Color = 13431551
    Cells(10, 4).Font.Name = "游ゴシック"
    Cells(10, 4).Interior.Color = 13431551
    Cells(10, 5).Font.Name = "游ゴシック"
    Cells(10, 6).Font.Name = "游ゴシック"
    Cells(10, 7).Font.Name = "游ゴシック"
    Cells(11, 1) = "NAME項目の名称"
    Cells(11, 1).Font.Name = "游ゴシック"
    Cells(11, 2) = "管理番号"
    Cells(11, 2).Font.Name = "游ゴシック"
    Cells(11, 3).Font.Name = "游ゴシック"
    Cells(11, 4) = "管理番号"
    Cells(11, 4).Font.Name = "游ゴシック"
    Cells(11, 5) = "OK"
    Cells(11, 5).Font.Name = "游ゴシック"
    Cells(11, 6).Font.Name = "游ゴシック"
    Cells(11, 7) = "テキスト"
    Cells(11, 7).Font.Name = "游ゴシック"
    Cells(12, 1) = "履歴追跡"
    Cells(12, 1).Font.Name = "游ゴシック"
    Cells(12, 2) = "する"
    Cells(12, 2).Font.Name = "游ゴシック"
    Cells(12, 2).Validation.Delete
    Cells(12, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(12, 3).Font.Name = "游ゴシック"
    Cells(12, 4) = "True"
    Cells(12, 4).Font.Name = "游ゴシック"
    Cells(12, 5) = "OK"
    Cells(12, 5).Font.Name = "游ゴシック"
    Cells(12, 6).Font.Name = "游ゴシック"
    Cells(12, 7) = "boolean"
    Cells(12, 7).Font.Name = "游ゴシック"
    Cells(13, 1) = "データ型"
    Cells(13, 1).Font.Name = "游ゴシック"
    Cells(13, 2) = "自動採番"
    Cells(13, 2).Font.Name = "游ゴシック"
    Cells(13, 2).Validation.Delete
    Cells(13, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$C$2:$C$3"
    Cells(13, 3).Font.Name = "游ゴシック"
    Cells(13, 4) = "AutoNumber"
    Cells(13, 4).Font.Name = "游ゴシック"
    Cells(13, 5) = "OK"
    Cells(13, 5).Font.Name = "游ゴシック"
    Cells(13, 6).Font.Name = "游ゴシック"
    Cells(13, 7) = "リスト"
    Cells(13, 7).Font.Name = "游ゴシック"
end Function
Function revival2()
    Cells(14, 1) = "表示形式"
    Cells(14, 1).Font.Name = "游ゴシック"
    Cells(14, 2) = "D-{00000000}"
    Cells(14, 2).Font.Name = "游ゴシック"
    Cells(14, 3).Font.Name = "游ゴシック"
    Cells(14, 4) = "D-{00000000}"
    Cells(14, 4).Font.Name = "游ゴシック"
    Cells(14, 5).Font.Name = "游ゴシック"
    Cells(14, 6).Font.Name = "游ゴシック"
    Cells(14, 7).Font.Name = "游ゴシック"
    Cells(15, 1).Font.Name = "游ゴシック"
    Cells(15, 2).Font.Name = "游ゴシック"
    Cells(15, 3).Font.Name = "游ゴシック"
    Cells(15, 4).Font.Name = "游ゴシック"
    Cells(15, 5).Font.Name = "游ゴシック"
    Cells(15, 6).Font.Name = "游ゴシック"
    Cells(15, 7).Font.Name = "游ゴシック"
    Cells(16, 1) = "追加の機能"
    Cells(16, 1).Font.Name = "游ゴシック"
    Cells(16, 1).Interior.Color = 13431551
    Cells(16, 1).Font.Bold = True
    Cells(16, 2).Font.Name = "游ゴシック"
    Cells(16, 2).Interior.Color = 13431551
    Cells(16, 2).Font.Bold = True
    Cells(16, 3).Font.Name = "游ゴシック"
    Cells(16, 3).Interior.Color = 13431551
    Cells(16, 4).Font.Name = "游ゴシック"
    Cells(16, 4).Interior.Color = 13431551
    Cells(16, 5).Font.Name = "游ゴシック"
    Cells(16, 6).Font.Name = "游ゴシック"
    Cells(16, 7).Font.Name = "游ゴシック"
    Cells(17, 1) = "レポートを許可"
    Cells(17, 1).Font.Name = "游ゴシック"
    Cells(17, 2) = "しない"
    Cells(17, 2).Font.Name = "游ゴシック"
    Cells(17, 2).Validation.Delete
    Cells(17, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(17, 3).Font.Name = "游ゴシック"
    Cells(17, 4) = "False"
    Cells(17, 4).Font.Name = "游ゴシック"
    Cells(17, 5) = "OK"
    Cells(17, 5).Font.Name = "游ゴシック"
    Cells(17, 6).Font.Name = "游ゴシック"
    Cells(17, 7) = "boolean"
    Cells(17, 7).Font.Name = "游ゴシック"
    Cells(18, 1) = "活動を許可"
    Cells(18, 1).Font.Name = "游ゴシック"
    Cells(18, 2) = "する"
    Cells(18, 2).Font.Name = "游ゴシック"
    Cells(18, 2).Validation.Delete
    Cells(18, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
end Function
Function revival3()
    Cells(18, 3).Font.Name = "游ゴシック"
    Cells(18, 4) = "True"
    Cells(18, 4).Font.Name = "游ゴシック"
    Cells(18, 5) = "OK"
    Cells(18, 5).Font.Name = "游ゴシック"
    Cells(18, 6).Font.Name = "游ゴシック"
    Cells(18, 7) = "boolean"
    Cells(18, 7).Font.Name = "游ゴシック"
    Cells(19, 1) = "項目履歴管理"
    Cells(19, 1).Font.Name = "游ゴシック"
    Cells(19, 2) = "しない"
    Cells(19, 2).Font.Name = "游ゴシック"
    Cells(19, 2).Validation.Delete
    Cells(19, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(19, 3).Font.Name = "游ゴシック"
    Cells(19, 4) = "False"
    Cells(19, 4).Font.Name = "游ゴシック"
    Cells(19, 5) = "OK"
    Cells(19, 5).Font.Name = "游ゴシック"
    Cells(19, 6).Font.Name = "游ゴシック"
    Cells(19, 7) = "boolean"
    Cells(19, 7).Font.Name = "游ゴシック"
    Cells(20, 1) = "Chatterグループ内で許可"
    Cells(20, 1).Font.Name = "游ゴシック"
    Cells(20, 2) = "する"
    Cells(20, 2).Font.Name = "游ゴシック"
    Cells(20, 2).Validation.Delete
    Cells(20, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(20, 3).Font.Name = "游ゴシック"
    Cells(20, 4) = "True"
    Cells(20, 4).Font.Name = "游ゴシック"
    Cells(20, 5) = "OK"
    Cells(20, 5).Font.Name = "游ゴシック"
    Cells(20, 6).Font.Name = "游ゴシック"
    Cells(20, 7) = "boolean"
    Cells(20, 7).Font.Name = "游ゴシック"
    Cells(21, 1) = "ライセンスの有効化"
    Cells(21, 1).Font.Name = "游ゴシック"
    Cells(21, 1).Interior.Color = 14277081
    Cells(21, 2) = "する"
    Cells(21, 2).Font.Name = "游ゴシック"
    Cells(21, 2).Interior.Color = 14277081
    Cells(21, 2).Validation.Delete
    Cells(21, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(21, 3).Font.Name = "游ゴシック"
    Cells(21, 3).Interior.Color = 14277081
    Cells(21, 4) = "True"
    Cells(21, 4).Font.Name = "游ゴシック"
    Cells(21, 4).Interior.Color = 14277081
    Cells(21, 5) = "OK"
    Cells(21, 5).Font.Name = "游ゴシック"
    Cells(21, 6).Font.Name = "游ゴシック"
    Cells(21, 7) = "boolean"
    Cells(21, 7).Font.Name = "游ゴシック"
    Cells(22, 1).Font.Name = "游ゴシック"
    Cells(22, 2).Font.Name = "游ゴシック"
    Cells(22, 3).Font.Name = "游ゴシック"
    Cells(22, 4).Font.Name = "游ゴシック"
end Function
Function revival4()
    Cells(22, 5).Font.Name = "游ゴシック"
    Cells(22, 6).Font.Name = "游ゴシック"
    Cells(22, 7).Font.Name = "游ゴシック"
    Cells(23, 1) = "オブジェクトの分類"
    Cells(23, 1).Font.Name = "游ゴシック"
    Cells(23, 1).Interior.Color = 13431551
    Cells(23, 1).Font.Bold = True
    Cells(23, 2).Font.Name = "游ゴシック"
    Cells(23, 2).Interior.Color = 13431551
    Cells(23, 2).Font.Bold = True
    Cells(23, 3).Font.Name = "游ゴシック"
    Cells(23, 3).Interior.Color = 13431551
    Cells(23, 4).Font.Name = "游ゴシック"
    Cells(23, 4).Interior.Color = 13431551
    Cells(23, 5).Font.Name = "游ゴシック"
    Cells(23, 6).Font.Name = "游ゴシック"
    Cells(23, 7).Font.Name = "游ゴシック"
    Cells(24, 1) = "共有を許可"
    Cells(24, 1).Font.Name = "游ゴシック"
    Cells(24, 2) = "する"
    Cells(24, 2).Font.Name = "游ゴシック"
    Cells(24, 2).Validation.Delete
    Cells(24, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(24, 3).Font.Name = "游ゴシック"
    Cells(24, 4) = "True"
    Cells(24, 4).Font.Name = "游ゴシック"
    Cells(24, 5) = "OK"
    Cells(24, 5).Font.Name = "游ゴシック"
    Cells(24, 6).Font.Name = "游ゴシック"
    Cells(24, 7) = "boolean"
    Cells(24, 7).Font.Name = "游ゴシック"
    Cells(25, 1) = "Bulk API アクセスを許可"
    Cells(25, 1).Font.Name = "游ゴシック"
    Cells(25, 2) = "する"
    Cells(25, 2).Font.Name = "游ゴシック"
    Cells(25, 2).Validation.Delete
    Cells(25, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(25, 3).Font.Name = "游ゴシック"
    Cells(25, 4) = "True"
    Cells(25, 4).Font.Name = "游ゴシック"
    Cells(25, 5) = "OK"
    Cells(25, 5).Font.Name = "游ゴシック"
    Cells(25, 6).Font.Name = "游ゴシック"
    Cells(25, 7) = "boolean"
    Cells(25, 7).Font.Name = "游ゴシック"
    Cells(26, 1) = "ストリーミング API アクセスを許可"
    Cells(26, 1).Font.Name = "游ゴシック"
    Cells(26, 2) = "する"
    Cells(26, 2).Font.Name = "游ゴシック"
    Cells(26, 2).Validation.Delete
    Cells(26, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(26, 3).Font.Name = "游ゴシック"
    Cells(26, 4) = "True"
    Cells(26, 4).Font.Name = "游ゴシック"
    Cells(26, 5) = "OK"
    Cells(26, 5).Font.Name = "游ゴシック"
    Cells(26, 6).Font.Name = "游ゴシック"
end Function
Function revival5()
    Cells(26, 7) = "boolean"
    Cells(26, 7).Font.Name = "游ゴシック"
    Cells(27, 1).Font.Name = "游ゴシック"
    Cells(27, 2).Font.Name = "游ゴシック"
    Cells(27, 3).Font.Name = "游ゴシック"
    Cells(27, 4).Font.Name = "游ゴシック"
    Cells(27, 5).Font.Name = "游ゴシック"
    Cells(27, 6).Font.Name = "游ゴシック"
    Cells(27, 7).Font.Name = "游ゴシック"
    Cells(28, 1) = "リリース状況"
    Cells(28, 1).Font.Name = "游ゴシック"
    Cells(28, 1).Interior.Color = 13431551
    Cells(28, 1).Font.Bold = True
    Cells(28, 2).Font.Name = "游ゴシック"
    Cells(28, 2).Interior.Color = 13431551
    Cells(28, 2).Font.Bold = True
    Cells(28, 3).Font.Name = "游ゴシック"
    Cells(28, 3).Interior.Color = 13431551
    Cells(28, 4).Font.Name = "游ゴシック"
    Cells(28, 4).Interior.Color = 13431551
    Cells(28, 5).Font.Name = "游ゴシック"
    Cells(28, 6).Font.Name = "游ゴシック"
    Cells(28, 7).Font.Name = "游ゴシック"
    Cells(29, 1) = "リリース状況"
    Cells(29, 1).Font.Name = "游ゴシック"
    Cells(29, 2) = "リリース済み"
    Cells(29, 2).Font.Name = "游ゴシック"
    Cells(29, 2).Validation.Delete
    Cells(29, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$G$2:$G$3"
    Cells(29, 3).Font.Name = "游ゴシック"
    Cells(29, 4) = "Deployed"
    Cells(29, 4).Font.Name = "游ゴシック"
    Cells(29, 5).Font.Name = "游ゴシック"
    Cells(29, 6).Font.Name = "游ゴシック"
    Cells(29, 7) = "リスト"
    Cells(29, 7).Font.Name = "游ゴシック"
    Cells(30, 1).Font.Name = "游ゴシック"
    Cells(30, 2).Font.Name = "游ゴシック"
    Cells(30, 3).Font.Name = "游ゴシック"
    Cells(30, 4).Font.Name = "游ゴシック"
    Cells(30, 5).Font.Name = "游ゴシック"
    Cells(30, 6).Font.Name = "游ゴシック"
    Cells(30, 7).Font.Name = "游ゴシック"
    Cells(31, 1) = "検索状況"
    Cells(31, 1).Font.Name = "游ゴシック"
    Cells(31, 1).Interior.Color = 13431551
    Cells(31, 1).Font.Bold = True
end Function
Function revival6()
    Cells(31, 2).Font.Name = "游ゴシック"
    Cells(31, 2).Interior.Color = 13431551
    Cells(31, 2).Font.Bold = True
    Cells(31, 3).Font.Name = "游ゴシック"
    Cells(31, 3).Interior.Color = 13431551
    Cells(31, 4).Font.Name = "游ゴシック"
    Cells(31, 4).Interior.Color = 13431551
    Cells(31, 5).Font.Name = "游ゴシック"
    Cells(31, 6).Font.Name = "游ゴシック"
    Cells(31, 7).Font.Name = "游ゴシック"
    Cells(32, 1) = "検索を許可"
    Cells(32, 1).Font.Name = "游ゴシック"
    Cells(32, 2) = "する"
    Cells(32, 2).Font.Name = "游ゴシック"
    Cells(32, 2).Validation.Delete
    Cells(32, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(32, 3).Font.Name = "游ゴシック"
    Cells(32, 4) = "True"
    Cells(32, 4).Font.Name = "游ゴシック"
    Cells(32, 5) = "OK"
    Cells(32, 5).Font.Name = "游ゴシック"
    Cells(32, 6).Font.Name = "游ゴシック"
    Cells(32, 7) = "boolean"
    Cells(32, 7).Font.Name = "游ゴシック"
    Cells(33, 1).Font.Name = "游ゴシック"
    Cells(33, 2).Font.Name = "游ゴシック"
    Cells(33, 3).Font.Name = "游ゴシック"
    Cells(33, 4).Font.Name = "游ゴシック"
    Cells(33, 5).Font.Name = "游ゴシック"
    Cells(33, 6).Font.Name = "游ゴシック"
    Cells(33, 7).Font.Name = "游ゴシック"
    Cells(34, 1) = "フィード追跡"
    Cells(34, 1).Font.Name = "游ゴシック"
    Cells(34, 1).Interior.Color = 13431551
    Cells(34, 1).Font.Bold = True
    Cells(34, 2).Font.Name = "游ゴシック"
    Cells(34, 2).Interior.Color = 13431551
    Cells(34, 2).Font.Bold = True
    Cells(34, 3).Font.Name = "游ゴシック"
    Cells(34, 3).Interior.Color = 13431551
    Cells(34, 4).Font.Name = "游ゴシック"
    Cells(34, 4).Interior.Color = 13431551
    Cells(34, 5).Font.Name = "游ゴシック"
    Cells(34, 6).Font.Name = "游ゴシック"
    Cells(34, 7).Font.Name = "游ゴシック"
    Cells(35, 1) = "フィード追跡"
    Cells(35, 1).Font.Name = "游ゴシック"
    Cells(35, 2) = "する"
    Cells(35, 2).Font.Name = "游ゴシック"
    Cells(35, 2).Validation.Delete
    Cells(35, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(35, 3).Font.Name = "游ゴシック"
end Function
Function revival7()
    Cells(35, 4) = "True"
    Cells(35, 4).Font.Name = "游ゴシック"
    Cells(35, 5).Font.Name = "游ゴシック"
    Cells(35, 6).Font.Name = "游ゴシック"
    Cells(35, 7) = "boolean"
    Cells(35, 7).Font.Name = "游ゴシック"
    Cells(36, 1).Font.Name = "游ゴシック"
    Cells(36, 2).Font.Name = "游ゴシック"
    Cells(36, 3).Font.Name = "游ゴシック"
    Cells(36, 4).Font.Name = "游ゴシック"
    Cells(36, 5).Font.Name = "游ゴシック"
    Cells(36, 6).Font.Name = "游ゴシック"
    Cells(36, 7).Font.Name = "游ゴシック"
    Cells(37, 1) = "共有設定"
    Cells(37, 1).Font.Name = "游ゴシック"
    Cells(37, 1).Interior.Color = 13431551
    Cells(37, 1).Font.Bold = True
    Cells(37, 2).Font.Name = "游ゴシック"
    Cells(37, 2).Interior.Color = 13431551
    Cells(37, 2).Font.Bold = True
    Cells(37, 3).Font.Name = "游ゴシック"
    Cells(37, 3).Interior.Color = 13431551
    Cells(37, 4).Font.Name = "游ゴシック"
    Cells(37, 4).Interior.Color = 13431551
    Cells(37, 5).Font.Name = "游ゴシック"
    Cells(37, 6).Font.Name = "游ゴシック"
    Cells(37, 7).Font.Name = "游ゴシック"
    Cells(38, 1) = "共有設定"
    Cells(38, 1).Font.Name = "游ゴシック"
    Cells(38, 2) = "公開/参照・更新可能"
    Cells(38, 2).Font.Name = "游ゴシック"
    Cells(38, 2).Validation.Delete
    Cells(38, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$E$2:$E$4"
    Cells(38, 3).Font.Name = "游ゴシック"
    Cells(38, 4) = "ReadWrite"
    Cells(38, 4).Font.Name = "游ゴシック"
    Cells(38, 5).Font.Name = "游ゴシック"
    Cells(38, 6).Font.Name = "游ゴシック"
    Cells(38, 7) = "リスト"
    Cells(38, 7).Font.Name = "游ゴシック"
    Cells(39, 1).Font.Name = "游ゴシック"
    Cells(39, 2).Font.Name = "游ゴシック"
    Cells(39, 3).Font.Name = "游ゴシック"
    Cells(39, 4).Font.Name = "游ゴシック"
    Cells(39, 5).Font.Name = "游ゴシック"
end Function
Function revival8()
    Cells(39, 6).Font.Name = "游ゴシック"
    Cells(39, 7).Font.Name = "游ゴシック"
    Cells(40, 1) = "コンパクトレイアウト"
    Cells(40, 1).Font.Name = "游ゴシック"
    Cells(40, 1).Interior.Color = 13431551
    Cells(40, 1).Font.Bold = True
    Cells(40, 2).Font.Name = "游ゴシック"
    Cells(40, 2).Interior.Color = 13431551
    Cells(40, 2).Font.Bold = True
    Cells(40, 3).Font.Name = "游ゴシック"
    Cells(40, 3).Interior.Color = 13431551
    Cells(40, 4).Font.Name = "游ゴシック"
    Cells(40, 4).Interior.Color = 13431551
    Cells(40, 5).Font.Name = "游ゴシック"
    Cells(40, 6).Font.Name = "游ゴシック"
    Cells(40, 7).Font.Name = "游ゴシック"
    Cells(41, 1) = "コンパクトレイアウト"
    Cells(41, 1).Font.Name = "游ゴシック"
    Cells(41, 2) = "SYSTEM"
    Cells(41, 2).Font.Name = "游ゴシック"
    Cells(41, 3).Font.Name = "游ゴシック"
    Cells(41, 4) = "SYSTEM"
    Cells(41, 4).Font.Name = "游ゴシック"
    Cells(41, 5).Font.Name = "游ゴシック"
    Cells(41, 6).Font.Name = "游ゴシック"
    Cells(41, 7).Font.Name = "游ゴシック"
    Cells(42, 1).Font.Name = "游ゴシック"
    Cells(42, 2).Font.Name = "游ゴシック"
    Cells(42, 3).Font.Name = "游ゴシック"
    Cells(42, 4).Font.Name = "游ゴシック"
    Cells(42, 5).Font.Name = "游ゴシック"
    Cells(42, 6).Font.Name = "游ゴシック"
    Cells(42, 7).Font.Name = "游ゴシック"
    Cells(43, 1) = "検索レイアウト"
    Cells(43, 1).Font.Name = "游ゴシック"
    Cells(43, 1).Interior.Color = 13431551
    Cells(43, 1).Font.Bold = True
    Cells(43, 2).Font.Name = "游ゴシック"
    Cells(43, 2).Interior.Color = 13431551
    Cells(43, 2).Font.Bold = True
    Cells(43, 3).Font.Name = "游ゴシック"
    Cells(43, 3).Interior.Color = 13431551
    Cells(43, 4).Font.Name = "游ゴシック"
    Cells(43, 4).Interior.Color = 13431551
    Cells(43, 5).Font.Name = "游ゴシック"
    Cells(43, 6).Font.Name = "游ゴシック"
    Cells(43, 7).Font.Name = "游ゴシック"
end Function
Function revival9()
    Cells(44, 1) = "項目1"
    Cells(44, 1).Font.Name = "游ゴシック"
    Cells(44, 2) = "UPDATEDBY_USER"
    Cells(44, 2).Font.Name = "游ゴシック"
    Cells(44, 3).Font.Name = "游ゴシック"
    Cells(44, 4) = "UPDATEDBY_USER"
    Cells(44, 4).Font.Name = "游ゴシック"
    Cells(44, 5).Font.Name = "游ゴシック"
    Cells(44, 6).Font.Name = "游ゴシック"
    Cells(44, 7).Font.Name = "游ゴシック"
    Cells(45, 1).Font.Name = "游ゴシック"
    Cells(45, 2).Font.Name = "游ゴシック"
    Cells(45, 3).Font.Name = "游ゴシック"
    Cells(45, 4).Font.Name = "游ゴシック"
    Cells(45, 5).Font.Name = "游ゴシック"
    Cells(45, 6).Font.Name = "游ゴシック"
    Cells(45, 7).Font.Name = "游ゴシック"
    Cells(46, 1) = "タブ情報"
    Cells(46, 1).Font.Name = "游ゴシック"
    Cells(46, 1).Interior.Color = 13431551
    Cells(46, 1).Font.Bold = True
    Cells(46, 2).Font.Name = "游ゴシック"
    Cells(46, 2).Interior.Color = 13431551
    Cells(46, 2).Font.Bold = True
    Cells(46, 3).Font.Name = "游ゴシック"
    Cells(46, 3).Interior.Color = 13431551
    Cells(46, 4).Font.Name = "游ゴシック"
    Cells(46, 4).Interior.Color = 13431551
    Cells(46, 5).Font.Name = "游ゴシック"
    Cells(46, 6).Font.Name = "游ゴシック"
    Cells(46, 7).Font.Name = "游ゴシック"
    Cells(47, 1) = "タブを作成"
    Cells(47, 1).Font.Name = "游ゴシック"
    Cells(47, 2) = "する"
    Cells(47, 2).Font.Name = "游ゴシック"
    Cells(47, 2).Validation.Delete
    Cells(47, 2).Validation.Add Type:=xlValidateList, _
        Operator:=xlEqual, _
        Formula1:="=マスタ!$A$2:$A$3"
    Cells(47, 3).Font.Name = "游ゴシック"
    Cells(47, 4) = "True"
    Cells(47, 4).Font.Name = "游ゴシック"
    Cells(47, 5).Font.Name = "游ゴシック"
    Cells(47, 6).Font.Name = "游ゴシック"
    Cells(47, 7).Font.Name = "游ゴシック"
    Cells(48, 1).Font.Name = "游ゴシック"
    Cells(48, 2).Font.Name = "游ゴシック"
end Function
Function revival10()
    Cells(48, 3).Font.Name = "游ゴシック"
    Cells(48, 4).Font.Name = "游ゴシック"
    Cells(48, 5).Font.Name = "游ゴシック"
    Cells(48, 6).Font.Name = "游ゴシック"
    Cells(48, 7).Font.Name = "游ゴシック"
    Cells(49, 1) = "アクション設定"
    Cells(49, 1).Font.Name = "游ゴシック"
    Cells(49, 1).Interior.Color = 13431551
    Cells(49, 1).Font.Bold = True
    Cells(49, 2).Font.Name = "游ゴシック"
    Cells(49, 2).Interior.Color = 13431551
    Cells(49, 2).Font.Bold = True
    Cells(49, 3).Font.Name = "游ゴシック"
    Cells(49, 3).Interior.Color = 13431551
    Cells(49, 4).Font.Name = "游ゴシック"
    Cells(49, 4).Interior.Color = 13431551
    Cells(49, 5).Font.Name = "游ゴシック"
    Cells(49, 6).Font.Name = "游ゴシック"
    Cells(49, 7).Font.Name = "游ゴシック"
    Rows(1).RowHeight = 18.75
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
    Rows(16).RowHeight = 18.75
    Rows(17).RowHeight = 18.75
    Rows(18).RowHeight = 18.75
    Rows(19).RowHeight = 18.75
    Rows(20).RowHeight = 18.75
    Rows(21).RowHeight = 18.75
    Rows(22).RowHeight = 18.75
    Rows(23).RowHeight = 18.75
    Rows(24).RowHeight = 18.75
    Rows(25).RowHeight = 18.75
    Rows(26).RowHeight = 18.75
    Rows(27).RowHeight = 18.75
    Rows(28).RowHeight = 18.75
    Rows(29).RowHeight = 18.75
    Rows(30).RowHeight = 18.75
    Rows(31).RowHeight = 18.75
    Rows(32).RowHeight = 18.75
    Rows(33).RowHeight = 18.75
    Rows(34).RowHeight = 18.75
    Rows(35).RowHeight = 18.75
    Rows(36).RowHeight = 18.75
    Rows(37).RowHeight = 18.75
    Rows(38).RowHeight = 18.75
    Rows(39).RowHeight = 18.75
    Rows(40).RowHeight = 18.75
    Rows(41).RowHeight = 18.75
    Rows(42).RowHeight = 18.75
    Rows(43).RowHeight = 18.75
    Rows(44).RowHeight = 18.75
    Rows(45).RowHeight = 18.75
    Rows(46).RowHeight = 18.75
    Rows(47).RowHeight = 18.75
    Rows(48).RowHeight = 18.75
    Rows(49).RowHeight = 18.75
    Columns(1).ColumnWidth = 31.13
    Columns(2).ColumnWidth = 33.88
    Columns(3).ColumnWidth = 8.38
    Columns(4).ColumnWidth = 16.75
    Columns(5).ColumnWidth = 8.38
    Columns(6).ColumnWidth = 26.13
    Columns(7).ColumnWidth = 8.38
    Dim onShape As Object
    Set onShape = ActiveSheet.Buttons.Add(0,112.5,190.5,18.75)
    onShape.Name = "Drop Down 4"
    onShape.Visible = 0
    onShape.Characters.Text = ""
End Function
Sub revival()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    CALL revival0()
    CALL revival1()
    CALL revival2()
    CALL revival3()
    CALL revival4()
    CALL revival5()
    CALL revival6()
    CALL revival7()
    CALL revival8()
    CALL revival9()
    CALL revival10()
    Worksheets(1).select
end sub