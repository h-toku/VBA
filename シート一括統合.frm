VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �V�[�g�ꊇ���� 
   Caption         =   "�V�[�g�ꊇ����"
   ClientHeight    =   4590
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6860
   OleObjectBlob   =   "�V�[�g�ꊇ����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�V�[�g�ꊇ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ���[�U�[�t�H�[���̃��[�h���ɁA�}�N�������s���Ă���u�b�N�ȊO�̂��ׂẴV�[�g�����X�g�{�b�N�X�ɒǉ�����
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim ws As worksheet
    Dim thisWb As Workbook
    
    ' �}�N�������s���Ă���A�N�e�B�u���[�N�u�b�N���擾
    Set thisWb = ActiveWorkbook
    
    ' TextBox1�ɃA�N�e�B�u���[�N�u�b�N�̖��O��\��
    TextBox1.value = thisWb.Name
    
    ' ListBox1�ɃA�N�e�B�u���[�N�u�b�N�ȊO�̃u�b�N�ƃV�[�g����ǉ�
    For Each wb In Workbooks
        If wb.Name <> thisWb.Name Then ' �}�N�������s���Ă���u�b�N�ȊO
            For Each ws In wb.Sheets
                ' ���X�g�{�b�N�X�Ƀu�b�N���ƃV�[�g����ǉ�
                ListBox1.AddItem wb.Name & " - " & ws.Name
            Next ws
        End If
    Next wb
End Sub


' OK�{�^�����N���b�N���ꂽ�Ƃ��ɁA�I�����ꂽ�V�[�g���}�N�������s���Ă���u�b�N�Ɉړ�����
Private Sub CommandButton1_Click()
    Dim thisWb As Workbook
    Dim ws As worksheet
    Dim i As Integer
    Dim sheetInfo() As String
    
    ' �}�N�������s���Ă���u�b�N���擾
    Set thisWb = ActiveWorkbook
    
    ' �I�����ꂽ�V�[�g�����̃u�b�N�̖����Ɉړ�
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            sheetInfo = Split(ListBox1.List(i), " - ")
            Workbooks(sheetInfo(0)).Sheets(sheetInfo(1)).Move after:=thisWb.Sheets(thisWb.Sheets.Count)
        End If
    Next i
    
    ' ���[�U�[�t�H�[�������
    Unload Me
End Sub


