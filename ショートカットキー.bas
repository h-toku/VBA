Attribute VB_Name = "�V���[�g�J�b�g�L�["
Option Explicit

Sub ChangeHeadersOrRenameSheets()
    Dim renameSheet As worksheet
    Dim sheetNameChangeSheet As worksheet
    
    ' �A�N�e�B�u�ȃ��[�N�u�b�N���g�p
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' �V�[�g���`�F�b�N
    On Error Resume Next
    Set renameSheet = wb.Sheets("�w�b�_�[���ꊇ�ύX")
    Set sheetNameChangeSheet = wb.Sheets("�V�[�g���ꊇ�ύX")
    On Error GoTo 0
    
    ' �w�b�_�[���ꊇ�ύX�V�[�g�����݂���ꍇ��ChangeHeaders�����s
    If Not renameSheet Is Nothing Then
        ChangeHeaders wb ' �A�N�e�B�u�ȃ��[�N�u�b�N��n��
    End If
    
    ' �V�[�g���ꊇ�ύX�V�[�g�����݂���ꍇ��RenameSheets�����s
    If Not sheetNameChangeSheet Is Nothing Then
        RenameSheets wb ' �A�N�e�B�u�ȃ��[�N�u�b�N��n��
    End If
    
    ' �����̃V�[�g�����݂��Ȃ��ꍇ�̃G���[���b�Z�[�W
    If renameSheet Is Nothing And sheetNameChangeSheet Is Nothing Then
        MsgBox "�u�w�b�_�[���ꊇ�ύX�v�V�[�g�܂��́u�V�[�g���ꊇ�ύX�v�V�[�g�����݂��܂���B", vbExclamation
    End If
End Sub

Sub ChangeHeaders(wb As Workbook)
    Dim renameSheet As worksheet
    Dim ws As worksheet
    Dim oldHeader As String, newHeader As String
    Dim i As Integer

    ' �u�w�b�_�[���ꊇ�ύX�v�V�[�g���擾
    Set renameSheet = wb.Sheets("�w�b�_�[���ꊇ�ύX")

    ' A��̃V�[�g���ɑΉ�����B��̃w�b�_�[��C��̐V�����w�b�_�[���ɕύX
    i = 2
    While renameSheet.Cells(i, 1).value <> ""
        Dim sheetName As String
        sheetName = Trim(renameSheet.Cells(i, 1).value) ' �V�[�g�����g���~���O

        ' �V�[�g�����݂��邩�m�F
        On Error Resume Next
        Set ws = wb.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            MsgBox "�V�[�g '" & sheetName & "' �͑��݂��܂���B", vbExclamation
            i = i + 1
            GoTo ContinueLoop ' ���̍s�֐i��
        End If
        
        ' B��̌Â��w�b�_�[�����擾
        oldHeader = renameSheet.Cells(i, 2).value
        newHeader = renameSheet.Cells(i, 3).value

        ' C�񂪋�łȂ��ꍇ�̂ݏ������s��
        If newHeader <> "" Then
            ' �w�b�_�[�͈̔͂�ݒ�
            Dim headerRange As Range
            Set headerRange = ws.Rows(1).Find(oldHeader, LookIn:=xlValues, LookAt:=xlWhole)

            If Not headerRange Is Nothing Then
                ' �Y������w�b�_�[�����������ꍇ�A�V�����w�b�_�[�ɕύX
                headerRange.value = newHeader
            End If
        End If

        i = i + 1
ContinueLoop: ' ���x�����`
    Wend
    
    ' ����������A�u�w�b�_�[���ꊇ�ύX�v�V�[�g���폜
    Application.DisplayAlerts = False
    renameSheet.Delete
    Application.DisplayAlerts = True

    MsgBox "�w�b�_�[�����ύX����܂����B", vbInformation
End Sub

Sub RenameSheets(wb As Workbook)
    Dim renameSheet As worksheet
    Dim ws As worksheet
    Dim i As Integer
    Dim newName As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim errorMsg As String

    ' �u�V�[�g���ꊇ�ύX�v�V�[�g���擾
    Set renameSheet = wb.Sheets("�V�[�g���ꊇ�ύX")

    If renameSheet Is Nothing Then
        MsgBox "�V�[�g���ꊇ�ύX�V�[�g��������܂���B", vbCritical
        Exit Sub
    End If

    ' A��̃V�[�g���ɑ΂���B��̐V�����V�[�g���ɕύX
    i = 2
    While renameSheet.Cells(i, 1).value <> ""
        newName = Trim(renameSheet.Cells(i, 2).value) ' �V�[�g�����g���~���O

        ' B�񂪋�łȂ��ꍇ�̂ݏ������s��
        If newName <> "" Then
            Dim sheetName As String
            sheetName = Trim(renameSheet.Cells(i, 1).value)

            ' �V�[�g�����݂��邩�m�F
            On Error Resume Next
            Set ws = wb.Sheets(sheetName)
            On Error GoTo 0
            
            If ws Is Nothing Then
                MsgBox "�V�[�g '" & sheetName & "' �͑��݂��܂���B", vbExclamation
                i = i + 1
                GoTo ContinueLoopSheets ' ���̍s�֐i��
            End If
            
            ' �d������ꍇ�A�A�Ԃ�t����
            If dict.Exists(newName) Then
                dict(newName) = dict(newName) + 1
                newName = newName & "_" & dict(newName)
            Else
                dict.Add newName, 1
            End If

            On Error Resume Next
            ws.Name = newName
            If Err.Number <> 0 Then
                MsgBox "�V�[�g�� '" & newName & "' �ɕύX�ł��܂���B", vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        i = i + 1
ContinueLoopSheets: ' ���x�����`
    Wend

    ' ����������A�u�V�[�g���ꊇ�ύX�v�V�[�g���폜
    Application.DisplayAlerts = False
    renameSheet.Delete
    Application.DisplayAlerts = True

    MsgBox "�V�[�g�����ύX����܂����B", vbInformation
End Sub
