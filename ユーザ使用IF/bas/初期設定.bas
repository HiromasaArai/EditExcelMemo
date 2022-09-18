Attribute VB_Name = "初期設定"
Option Explicit

' Pythonモジュールを実行するためのiniを作成します。
Sub 初期設定()
    Dim errMsg As String: errMsg = ""
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fl As Folder
    
    ' 対象フォルダがなければ、新規作成
    If Not IsExistDirA(constIniDirPath) Then: Set fl = fso.CreateFolder(constIniDirPath)
        
    ' 書き込み処理
    ' ts.Write ("文字列")       ' 文字列の書き込み
    ' ts.WriteLine ("改行付き") ' 最後に改行を付けて書き込み
    ' ts.WriteBlankLines (3)    ' 空行を指定の数だけ書き込み
    Dim ib As String: ib = InputBox("Pythonモジュールが存在するプロジェクトパスを入力して下さい。")
    
    If Not IsExistDirA(ib) Or ib = "" Then
        ' フールプルーフ
        errMsg = "入力されたプロジェクトパス[" + ib + "]が存在しません。"
    Else
        ' ファイルがあったら削除
        If IsExistFileA(constIniFilePath) Then: Call fso.DeleteFile(constIniFilePath, True)
    
        ' 新規ファイルを作成
        Dim ts As TextStream: Set ts = fso.OpenTextFile(constIniFilePath, ForWriting, True)
        ts.WriteLine ("[PythonProject]")
        ts.Write ("path = " + ib)
        
        ' ファイルを閉じる
        ts.Close
        
        ' ガベージコレクション
        Set ts = Nothing
    End If

    ' ガベージコレクション
    Set fso = Nothing
    
    ' エラーメッセージがあれば出力する
    If errMsg <> "" Then MsgBox errMsg
End Sub
