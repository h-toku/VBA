VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �w�b�_�[�ύX�t�H�[�� 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "�w�b�_�[�ύX�t�H�[��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�w�b�_�[�ύX�t�H�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As worksheet
    Dim wb As Workbook

    ' ����Ώۂ̃��[�N�u�b�N���A�N�e�B�u�ȃ��[�N�u�b�N�ɐݒ�
    Set wb = Application.ActiveWorkbook

    ' ���X�g�{�b�N�X�ɃV�[�g����ǉ�
    ListBox1.Clear ' �����̍��ڂ��N���A
    For Each ws In wb.Sheets
        ' �u�w�b�_�[���ꊇ�ύX�v�V�[�g�����O
        If ws.Name <> "�w�b�_�[���ꊇ�ύX" Then
            ListBox1.AddItem ws.Name
        End If
    Next ws
    
End Sub

Private Sub CommandButton1_Click()

    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Long, j As Long
    Dim wb As Workbook
    Dim lastCol As Long
    Dim selectedSheet As String
    Dim iSelected As Integer
    Dim dataRow As Long

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
    End With

    ' ���X�g�{�b�N�X�őI�����ꂽ�V�[�g�����擾���A���[�v�ŏ���
    If ListBox1.ListCount > 0 Then
        For iSelected = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(iSelected) Then
                selectedSheet = ListBox1.List(iSelected) ' �I�����ꂽ�V�[�g�����擾

                ' �I�����ꂽ�V�[�g�̃w�b�_�[�����擾
                Set ws = wb.Sheets(selectedSheet)
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                dataRow = renameSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1 ' �������ݍs��ݒ�

                ' �V�[�g�̃w�b�_�[�����擾���ď�������
                For j = 1 To lastCol
                    renameSheet.Cells(dataRow, 1).value = ws.Name
                    renameSheet.Cells(dataRow, 2).value = ws.Cells(1, j).value
                    dataRow = dataRow + 1
                Next j
            End If
        Next iSelected
    Else
        MsgBox "���X�g�{�b�N�X�ŃV�[�g��I�����Ă��������B", vbExclamation
    End If
    
    Unload Me

End Sub
