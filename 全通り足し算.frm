VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �S�ʂ葫���Z 
   Caption         =   "�S�ʂ葫���Z"
   ClientHeight    =   1910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5050
   OleObjectBlob   =   "�S�ʂ葫���Z.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�S�ʂ葫���Z"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As worksheet
    Dim lastCol As Long
    Dim i As Long
    
    ' �A�N�e�B�u�V�[�g�̎Q��
    Set ws = ActiveSheet
    
    ' �Ō�̗���擾�i�w�b�_�[�̂���s��z��j
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' ComboBox1��ComboBox2�Ƀw�b�_�[���Z�b�g
    For i = 1 To lastCol
        ComboBox1.AddItem ws.Cells(1, i).value  ' ���ڗ�̌��
        ComboBox2.AddItem ws.Cells(1, i).value  ' �l��̌��
    Next i
End Sub

Private Sub CommandButton1_Click()
    Dim ws As worksheet
    Dim result As worksheet
    Dim itemCol As Long, valueCol As Long
    Dim lastrow As Long
    Dim items As Variant
    Dim nums As Variant
    Dim i As Long, j As Long
    Dim rowCounter As Long
    
    ' �A�N�e�B�u�V�[�g�̎Q��
    Set ws = ActiveSheet
    
    ' ComboBox1��ComboBox2�őI�����ꂽ����擾
    itemCol = Application.Match(ComboBox1.value, ws.Rows(1), 0)  ' ���ڗ�
    valueCol = Application.Match(ComboBox2.value, ws.Rows(1), 0) ' �l��
    
    ' �ŏI�s�̎擾
    lastrow = ws.Cells(ws.Rows.Count, itemCol).End(xlUp).Row
    
    ' �f�[�^��z��Ɋi�[�i2�s�ڂ���ŏI�s�܂Łj
    items = ws.Range(ws.Cells(2, itemCol), ws.Cells(lastrow, itemCol)).value
    nums = ws.Range(ws.Cells(2, valueCol), ws.Cells(lastrow, valueCol)).value
    
    ' ���ʂ��o�͂���V�����V�[�g���쐬
    Set result = Sheets.Add
    result.Name = "Pair Sum Combinations"
    
    ' �w�b�_�[��ݒ�
    result.Cells(1, 1).value = ComboBox1.value & "�i����1�j"
    result.Cells(1, 2).value = ComboBox1.value & "�i����2�j"
    result.Cells(1, 3).value = ComboBox2.value & "�i���v�j"
    
    rowCounter = 2
    
    ' 2�̍��ڂ̑g�ݍ��킹�Ƃ���ɑΉ�����l�̍��v��񋓁i�d�����ȗ����A�Z�����g�̑����Z���܂ށj
    For i = 1 To UBound(nums, 1)
        For j = i To UBound(nums, 1)
            ' �g�ݍ��킹�̍��ږ����V�[�g�ɏo��
            result.Cells(rowCounter, 1).value = items(i, 1)  ' ����1
            result.Cells(rowCounter, 2).value = items(j, 1)  ' ����2
            ' ����ɑΉ�����B��̒l�̍��v���V�[�g�ɏo��
            result.Cells(rowCounter, 3).value = nums(i, 1) + nums(j, 1)  ' �l�̍��v
            rowCounter = rowCounter + 1
        Next j
    Next i
     
     Unload Me
     
End Sub

