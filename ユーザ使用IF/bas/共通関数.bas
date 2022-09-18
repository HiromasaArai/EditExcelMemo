Attribute VB_Name = "���ʊ֐�"
Option Explicit

' ini�t�@�C���擾�p
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
    Dim sPath                   '// INI�t�@�C���̃p�X
    Dim sValue      As String   '// �擾�l
    Dim lSize       As Long     '// �擾�l�̃T�C�Y
    Dim lRet        As Long     '// �߂�l
    
    lSize = 2000
    sPath = constIniFilePath

    ' �擾�o�b�t�@��������
    sValue = Space(lSize)
    lRet = GetPrivateProfileString(ini_section, ini_key, "non", sValue, lSize, sPath)
    
    ' �l��Ԃ�
    get_ini_value = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function


' FileSystemObject���g�p����ꍇ�́A�uCreateObject �֐��v���g���܂��B
' [�c�[��] ���� [�Q�Ɛݒ�] ���N���b�N�A[Microsoft Scripting Runtime] ���`�F�b�N���� [OK] ���N���b�N���܂��B

Public Function IsExistDirA(a_sFolder As String) As Boolean
    ' �t�H���_�����݂����True��Ԃ��B
    Dim result: result = Dir(a_sFolder, vbDirectory)
    
    If result = "" Then
        IsExistDirA = False
    Else
        IsExistDirA = True
    End If
    
End Function


Public Function IsExistFileA(a_sFile As String) As Boolean
    ' �t�@�C�������݂����True��Ԃ��B
    Dim result: result = Dir(a_sFile)
    
    If result = "" Then
        IsExistFileA = False
    Else
        IsExistFileA = True
    End If
    
End Function


Public Sub GoPython(exeName As String)
    ' �R�}���h�v�����v�g�o�R��Python���s
    
    Dim errMsg As String: errMsg = "�\�����Ȃ��G���[�ł��B"

    On Error GoTo Catch
        If Not IsExistFileA(constIniFilePath) Then
            ' ini�t�@�C�������݂��Ȃ�������G���[����
            errMsg = "ini�t�@�C��[" + constIniFilePath + "]�����݂��܂���B�����ݒ���s���Ă��������B"
            GoTo Catch
        End If
    
        ' ���s������exe�t�@�C�������݂���v���W�F�N�g�f�B���N�g�����擾
        Dim pyProject As String: pyProject = get_ini_value("PythonProject", "path")
        
        ' ���s������exe�t�@�C���̃t���p�X���쐬
        Dim exeDir As String: exeDir = pyProject + "\exe"
        Dim exeFullName As String:  exeFullName = exeDir + "\" + exeName
        
        If Not IsExistFileA(exeFullName) Then
            ' exe�t�@�C�������݂��Ȃ�������G���[����
            errMsg = "exe�t�@�C��[" + exeFullName + "]�����݂��܂���B"
            GoTo Catch
        End If
        
        ' exe���s
        Dim cmdStr As String: cmdStr = "cd " + exeDir + " & " + exeName
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        
        ' �uWindows Script Host Object Model�v���Q�Ɛݒ�
        Dim result As WshExec
        Set result = wsh.exec("%ComSpec% /c " & cmdStr)

        Do While result.Status = 0
            DoEvents
        Loop
        
        Dim result2 As String: result2 = result.StdOut.ReadAll
        If result2 <> "" Then MsgBox result2
        
        Exit Sub
    
Catch:
    ' ��O������t�[���v���[�t���L���b�`
    MsgBox errMsg
    
End Sub

