Attribute VB_Name = "�V�[�g���ꊇ�ύX"
Option Explicit

Sub CreateRenameSheet()
    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Integer
    Dim wb As Workbook

    ' �A�N�e�B�u�ȃ��[�N�u�b�N���擾
    Set wb = Application.ActiveWorkbook

    ' ���Ɂu�V�[�g���ꊇ�ύX�v�V�[�g�����݂���ꍇ�A�폜
    On Error Resume Next
    Set renameSheet = wb.Sheets("�V�[�g���ꊇ�ύX")
    If Not renameSheet Is Nothing Then
        Application.DisplayAlerts = False
        renameSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' �V�����V�[�g�u�V�[�g���ꊇ�ύX�v���쐬
    Set renameSheet = wb.Sheets.Add
    renameSheet.Name = "�V�[�g���ꊇ�ύX"

    ' �V�[�g�^�u�̐F�����F�ɐݒ�
    renameSheet.Tab.Color = RGB(255, 255, 0)

    ' A1�Ɂu�V�[�g���v�AB1�Ɂu�V�����V�[�g���v����͂��A���o���̐F�����F�ɐݒ�
    With renameSheet
        .Range("A1").value = "�V�[�g��"
        .Range("B1").value = "�V�����V�[�g��"
        .Range("A1:B1").Interior.Color = RGB(255, 255, 0)
        .Range("G3").value = "�uCtrl+Shift+R�v�Ŏ��s"
        .Range("G3").Font.Bold = True  ' �����ɐݒ�
        .Columns("A:G").AutoFit '�񕝂̎�������

        ' �S�ẴV�[�g����A��ɏ����o��
        i = 2
        For Each ws In wb.Sheets
            If ws.Name <> "�V�[�g���ꊇ�ύX" Then
                .Cells(i, 1).value = ws.Name
                i = i + 1
            End If
        Next ws
    End With
End Sub
