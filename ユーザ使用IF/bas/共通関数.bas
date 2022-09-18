Attribute VB_Name = "共通関数"
Option Explicit

' iniファイル取得用
Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Function get_ini_value(ini_section As String, ini_key As String)
    Dim sPath                   '// INIファイルのパス
    Dim sValue      As String   '// 取得値
    Dim lSize       As Long     '// 取得値のサイズ
    Dim lRet        As Long     '// 戻り値
    
    lSize = 2000
    sPath = constIniFilePath

    ' 取得バッファを初期化
    sValue = Space(lSize)
    lRet = GetPrivateProfileString(ini_section, ini_key, "non", sValue, lSize, sPath)
    
    ' 値を返す
    get_ini_value = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function


' FileSystemObjectを使用する場合は、「CreateObject 関数」を使います。
' [ツール] から [参照設定] をクリック、[Microsoft Scripting Runtime] をチェックして [OK] をクリックします。

Public Function IsExistDirA(a_sFolder As String) As Boolean
    ' フォルダが存在すればTrueを返す。
    Dim result: result = Dir(a_sFolder, vbDirectory)
    
    If result = "" Then
        IsExistDirA = False
    Else
        IsExistDirA = True
    End If
    
End Function


Public Function IsExistFileA(a_sFile As String) As Boolean
    ' ファイルが存在すればTrueを返す。
    Dim result: result = Dir(a_sFile)
    
    If result = "" Then
        IsExistFileA = False
    Else
        IsExistFileA = True
    End If
    
End Function


Public Sub GoPython(exeName As String)
    ' コマンドプロンプト経由でPython実行
    
    Dim errMsg As String: errMsg = "予期しないエラーです。"

    On Error GoTo Catch
        If Not IsExistFileA(constIniFilePath) Then
            ' iniファイルが存在しなかったらエラー処理
            errMsg = "iniファイル[" + constIniFilePath + "]が存在しません。初期設定を行ってください。"
            GoTo Catch
        End If
    
        ' 実行したいexeファイルが存在するプロジェクトディレクトリを取得
        Dim pyProject As String: pyProject = get_ini_value("PythonProject", "path")
        
        ' 実行したいexeファイルのフルパスを作成
        Dim exeDir As String: exeDir = pyProject + "\exe"
        Dim exeFullName As String:  exeFullName = exeDir + "\" + exeName
        
        If Not IsExistFileA(exeFullName) Then
            ' exeファイルが存在しなかったらエラー処理
            errMsg = "exeファイル[" + exeFullName + "]が存在しません。"
            GoTo Catch
        End If
        
        ' exe実行
        Dim cmdStr As String: cmdStr = "cd " + exeDir + " & " + exeName
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        
        ' 「Windows Script Host Object Model」を参照設定
        Dim result As WshExec
        Set result = wsh.exec("%ComSpec% /c " & cmdStr)

        Do While result.Status = 0
            DoEvents
        Loop
        
        Dim result2 As String: result2 = result.StdOut.ReadAll
        If result2 <> "" Then MsgBox result2
        
        Exit Sub
    
Catch:
    ' 例外処理やフールプルーフをキャッチ
    MsgBox errMsg
    
End Sub

