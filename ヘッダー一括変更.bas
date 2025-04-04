Attribute VB_Name = "�w�b�_�[�ꊇ�ύX"
Option Explicit

Sub CreateRenameHeaders()
    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Integer, j As Integer
    Dim wb As Workbook
    Dim lastCol As Long

    ' ����Ώۂ̃��[�N�u�b�N���A�N�e�B�u�ȃ��[�N�u�b�N�ɐݒ�
    Set wb = Application.ActiveWorkbook

    ' ���Ɂu�w�b�_�[���ꊇ�ύX�v�V�[�g�����݂���ꍇ�A�폜
    On Error Resume Next
    Set renameSheet = wb.Sheets("�w�b�_�[���ꊇ�ύX")
    If Not renameSheet Is Nothing Then
        Application.DisplayAlerts = False
        renameSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' �V�����V�[�g�u�w�b�_�[���ꊇ�ύX�v���쐬
    Set renameSheet = wb.Sheets.Add
    renameSheet.Name = "�w�b�_�[���ꊇ�ύX"

    ' �V�[�g�^�u�̐F�����F�ɐݒ�
    renameSheet.Tab.Color = RGB(255, 255, 0)

    ' A1�Ɂu�V�[�g���v�AB1�Ɂu�w�b�_�[���v�AC1�Ɂu�V�����w�b�_�[���v����͂��A���o���̐F�����F�ɐݒ�
    With renameSheet
        .Range("A1").value = "�V�[�g��"
        .Range("B1").value = "�w�b�_�[��"
        .Range("C1").value = "�V�����w�b�_�[��"
        .Range("A1:C1").Interior.Color = RGB(255, 255, 0)
        .Range("G3").value = "�uCtrl+Shift+R�v�Ŏ��s"
        .Range("G3").Font.Bold = True  ' �����ɐݒ�
        .Columns("A:G").AutoFit '�񕝂̎�������

        ' �S�ẴV�[�g���ƃw�b�_�[�����擾���AA���B��ɏ����o��
        i = 2
        For Each ws In wb.Sheets
            If ws.Name <> "�w�b�_�[���ꊇ�ύX" Then
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                For j = 1 To lastCol
                    .Cells(i, 1).value = ws.Name
                    .Cells(i, 2).value = ws.Cells(1, j).value
                    i = i + 1
                Next j
            End If
        Next ws
    End With
End Sub
