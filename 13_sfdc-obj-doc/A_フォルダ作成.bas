Attribute VB_Name = "A_フォルダ作成"
'オブジェクト情報シートのシート名（Indexで定義した方がいいか？）
Public Const OBJECT_SHEET = "オブジェクト"
'オブジェクトmetaファイルの情報を管理するシートのシート名
Public Const OBJECT_META_SHEET = "CustomObject"
'項目情報シートのシート名（Indexで定義した方がいいか？）
Public Const ITEM_SHEET = "項目"
'項目のタグ情報を管理するシートのシート名
Public Const ITEM_META_SHEET = "CustomItem"
'ページレイアウト情報シートのシート名
Public Const LAYOUT_SHEET = "ページレイアウト"
'フォルダ作成
Sub A_フォルダ作成()
    
    'TODO:もっと厳密なエラーチェックが必要か
    If Not Sheets(OBJECT_SHEET).Cells(4, 4) Like "*__c" Then
        MsgBox "オブジェクト名が「__c」で終わっていません。"
        Exit Sub
    End If
    
    'オブジェクトの親フォルダ作成
    If Dir(ThisWorkbook.Path & "\objects\", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\objects\"
    End If
    
    'オブジェクト名のフォルダをルートとする
    Dim rootDirName As String: rootDirName = ThisWorkbook.Path & "\objects\" & Sheets(OBJECT_SHEET).Cells(4, 4) & "\"
    
    'オブジェクト名フォルダ
    If Dir(rootDirName, vbDirectory) = "" Then
        MkDir rootDirName
    End If
    
    'コンパクトレイアウト
    If Dir(rootDirName & "compactLayouts\", vbDirectory) = "" Then
        MkDir rootDirName & "compactLayouts\"
    End If
    
    '項目
    If Dir(rootDirName & "fields\", vbDirectory) = "" Then
        MkDir rootDirName & "fields\"
    End If
    
    'リストビュー
    If Dir(rootDirName & "listViews\", vbDirectory) = "" Then
        MkDir rootDirName & "listViews\"
    End If
    
    '入力規則
    If Dir(rootDirName & "validationRules\", vbDirectory) = "" Then
        MkDir rootDirName & "validationRules\"
    End If
    
    'レコードタイプ
    If Dir(rootDirName & "recordTypes\", vbDirectory) = "" Then
        MkDir rootDirName & "recordTypes\"
    End If
    
    'パス（これはオブジェクト配下ではない）
    If Dir(ThisWorkbook.Path & "\tabs\", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\tabs\"
    End If

'    '権限付与はデータローダからした方が安全なため非推奨
'    If Dir(ThisWorkbook.Path & "\profiles\", vbDirectory) = "" Then
'        MkDir ThisWorkbook.Path & "\profiles\"
'    End If

'    'レイアウトは画面から作成した方が自動設定が加味されるため非推奨
'    If Dir(ThisWorkbook.Path & "\layouts\", vbDirectory) = "" Then
'        MkDir ThisWorkbook.Path & "\layouts\"
'    End If
    
    MsgBox "フォルダを作成しました。"
End Sub
