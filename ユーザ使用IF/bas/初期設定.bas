Attribute VB_Name = "�����ݒ�"
Option Explicit

' Python���W���[�������s���邽�߂�ini���쐬���܂��B
Sub �����ݒ�()
    Dim errMsg As String: errMsg = ""
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fl As Folder
    
    ' �Ώۃt�H���_���Ȃ���΁A�V�K�쐬
    If Not IsExistDirA(constIniDirPath) Then: Set fl = fso.CreateFolder(constIniDirPath)
        
    ' �������ݏ���
    ' ts.Write ("������")       ' ������̏�������
    ' ts.WriteLine ("���s�t��") ' �Ō�ɉ��s��t���ď�������
    ' ts.WriteBlankLines (3)    ' ��s���w��̐�������������
    Dim ib As String: ib = InputBox("Python���W���[�������݂���v���W�F�N�g�p�X����͂��ĉ������B")
    
    If Not IsExistDirA(ib) Or ib = "" Then
        ' �t�[���v���[�t
        errMsg = "���͂��ꂽ�v���W�F�N�g�p�X[" + ib + "]�����݂��܂���B"
    Else
        ' �t�@�C������������폜
        If IsExistFileA(constIniFilePath) Then: Call fso.DeleteFile(constIniFilePath, True)
    
        ' �V�K�t�@�C�����쐬
        Dim ts As TextStream: Set ts = fso.OpenTextFile(constIniFilePath, ForWriting, True)
        ts.WriteLine ("[PythonProject]")
        ts.Write ("path = " + ib)
        
        ' �t�@�C�������
        ts.Close
        
        ' �K�x�[�W�R���N�V����
        Set ts = Nothing
    End If

    ' �K�x�[�W�R���N�V����
    Set fso = Nothing
    
    ' �G���[���b�Z�[�W������Ώo�͂���
    If errMsg <> "" Then MsgBox errMsg
End Sub
