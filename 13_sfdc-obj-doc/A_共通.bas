Attribute VB_Name = "A_共通"
'オブジェクト情報シートのシート名（Indexで定義した方がいいか？）
Public Const OBJECT_SHEET = "オブジェクト"

'オブジェクトmetaファイルの情報を管理するシートのシート名
Public Const OBJECT_META_SHEET = "CustomObject"

'項目情報シートのシート名（Indexで定義した方がいいか？）
Public Const ITEM_SHEET = "項目"

'項目のタグ情報を管理するシートのシート名
Public Const ITEM_META_SHEET = "CustomItem"

'権限情報シートのシート名
Public Const PERMISSION_SHEET = "権限"

'ページレイアウト情報シートのシート名
Public Const LAYOUT_SHEET = "ページレイアウト"

'Trueを示す文字
Public Const ON_TRUE = "〇"

'テキストファイル出力用
Public stream As Object

'正規表現
Public regexp As Object

'オブジェクト情報シート
Public objSheet As Worksheet

'オブジェクトメタ情報シート
Public objMetaSheet As Worksheet

'項目情報シート
Public itemSheet As Worksheet

'項目メタ情報シート
Public itemMetaSheet As Worksheet

'権限情報シート
Public permissionSheet As Worksheet

'オブジェクトAPI名
Public objApiName As String

'ファイル名
Public fileName As String

'項目フォルダパス
Public fieldsDirPath As String

'初期化
Public Function initiarize()
    Set regexp = CreateObject("VBScript.RegExp")
    Set objSheet = Sheets(OBJECT_SHEET)
    Set objMetaSheet = Sheets(OBJECT_META_SHEET)
    Set itemSheet = Sheets(ITEM_SHEET)
    Set itemMetaSheet = Sheets(ITEM_META_SHEET)
    Set permissionSheet = Sheets(PERMISSION_SHEET)
    objApiName = Sheets(OBJECT_SHEET).Cells(4, 4).Value
    fieldsDirPath = ThisWorkbook.path & "\objects\" & objApiName & "\fields\"
End Function
'テキスト出力用
Public Function openStream()
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "UTF-8"
    stream.Open
End Function
'正規表現
Public Function setupRegexp(argPattern As String)
    '置換文字抽出用パターン（VBAで肯定先読みは使えない）
    regexp.Pattern = argPattern
    '英大文字小文字を区別しない
    regexp.IgnoreCase = True
    '文字列全体に対してパターンマッチさせる
    regexp.Global = True
End Function
'UTF-8で保存するときの保存処理をストリームobjectとファイル名で行う
Public Function saveTextWithUTF8(stream As Object, fileFullName As String)
    'Streamオブジェクトの先頭からの位置を指定する。Typeに値を設定するときは0である必要がある
    stream.Position = 0
    '扱うデータ種類をバイナリデータに変更する
    stream.Type = 1
    '読み取り開始位置？を3バイト目に移動する（3バイトはBOM付き部分を削除するため）
    stream.Position = 3
    'バイト文字を一時保存
    bytetmp = stream.Read
    'ここでは保存は不要。一度閉じて書き込んだ内容をリセットする目的がある
    stream.Close
    '再度開いて
    stream.Open
    'バイト形式で書き込むんで
    stream.write bytetmp
    Call checkExistDir(getDirPath(fileFullName))
    '保存
    stream.SaveToFile fileFullName, 2
    'コピー先ファイルを閉じる
    stream.Close
End Function
'ファイルパスが存在するかチェックする。なければつくる
Public Function checkExistDir(path As String)
    'ファイル操作オブジェクト
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'フォルダ名配列
    Dim dirs() As String: dirs = Split(path, "\")
    '検査フォルダ
    Dim incrementDir As String
    
    For i = LBound(dirs) To UBound(dirs)
        If incrementDir <> "" And Not objFSO.FolderExists(incrementDir) Then
            objFSO.CreateFolder (incrementDir)
        End If
        incrementDir = incrementDir & dirs(i) & "\"
    Next
End Function
'ファイルのフルパスからフォルダパスを取得する
Public Function getDirPath(argFilePath As String)
    Dim dirs As Variant: dirs = Split(argFilePath, "\")
    getDirPath = Left(argFilePath, Len(argFilePath) - Len(dirs(UBound(dirs))) - 1) & "\"
End Function
