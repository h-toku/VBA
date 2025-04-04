VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �h��Ԃ��t�H�[�� 
   Caption         =   "�h��Ԃ�"
   ClientHeight    =   5360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9250.001
   OleObjectBlob   =   "�h��Ԃ��t�H�[��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�h��Ԃ��t�H�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim savedSettings As Object ' �ݒ��ۑ����邽�߂̎����I�u�W�F�N�g
Dim settingsFilePath As String
Dim previousComboBoxValues(1 To 7) As Variant ' ComboBox�̒l��ۑ�����z��
Dim previousTextBoxBackColors(1 To 7) As Long ' TextBox�̔w�i�F��ۑ�����z��

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler ' �G���[�n���h�����O�̊J�n

    settingsFilePath = ThisWorkbook.Path & "\settings.xlsx" ' �p�X�𖾎��I�Ɏw��
    Dim i As Long
    Dim headerValue As String
    Dim lastCol As Long
    Dim j As Integer
    
    'HEADERBOX��1�`10�̒l��ǉ�
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.value = "1"
    
    ' �ŏI����擾
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    ' Dictionary�̏�����
    Set savedSettings = CreateObject("Scripting.Dictionary")
    
    ' �ݒ���t�@�C������ǂݍ���
    LoadSettingsFromFile
    
    ' �R���{�{�b�N�XA�`G�̏����ݒ�
    Dim comboBoxNames As Variant
    comboBoxNames = Array("ComboBoxA", "ComboBoxB", "ComboBoxC", "ComboBoxD", "ComboBoxE", "ComboBoxF", "ComboBoxG")
    
    For i = LBound(comboBoxNames) To UBound(comboBoxNames)
        With Me.Controls(comboBoxNames(i))
            .Clear ' �����̃A�C�e�����N���A
            .AddItem "��v" ' ��v��ǉ�
            .AddItem "�ȏ�" ' �ȏ��ǉ�
            .AddItem "�ȉ�" ' �ȉ���ǉ�
            .AddItem "�܂�" ' �܂ނ�ǉ�
        End With
    Next i

    Exit Sub ' ����I��

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description ' �G���[���b�Z�[�W��\��
End Sub
Private Sub HEADERBOX_Change()

    Dim targetRow As Long
    Dim lastColumn As Long
    Dim wsActive As worksheet
    Dim i As Long
    Dim j As Variant

    ' �A�N�e�B�u�V�[�g���擾
    Set wsActive = ActiveWorkbook.ActiveSheet
    
    ' HEADERBOX�̒l���擾
    targetRow = HEADERBOX.value
    
    ' HEADERBOX�̒l��1����10�͈͓̔����m�F
    If targetRow < 1 Or targetRow > 10 Then
        MsgBox "HEADERBOX�̒l��1����10�͈̔͂ł���K�v������܂��B", vbExclamation
        Exit Sub
    End If

    ' �Ō�̗���擾
    lastColumn = wsActive.Cells(targetRow, wsActive.Columns.Count).End(xlToLeft).Column

    ' ComboBox1���N���A
    ListBox1.Clear

    ' �w�肵���s (targetRow) ��1��ځ`lastColumn�܂ł̒l��ListBox1�ɒǉ�
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).value
        
        ' �Z���̒l����łȂ��ꍇ�ɂ̂ݒǉ�
        If Not IsEmpty(j) Then
            ListBox1.AddItem CStr(j)  ' �l�𕶎���ɕϊ����Ēǉ�
        End If
    Next i

End Sub

Private Sub LoadSettingsFromFile()
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    Dim key As String
    Dim value As String
    
    fileNum = FreeFile
    Open settingsFilePath For Input As #fileNum
    
    ' �t�@�C���̊e�s��ǂݍ���
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "|")
        If UBound(parts) = 1 Then
            key = parts(0)
            value = parts(1)
            savedSettings(key) = value
            
            ' �I�v�V�����{�^���̏�Ԃ�ݒ�
            If key = "OptionButton1" Then
                OptionButton1.value = (value = "1")
            ElseIf key = "OptionButton2" Then
                OptionButton2.value = (value = "1")
            End If
        End If
    Loop
    
    Close #fileNum
End Sub

Private Sub UserForm_Activate()
    Dim wb As Workbook
    Dim ws As worksheet
    
    ' ListBox2���N���A
    ListBox2.Clear
    
    ' `setting.xlsx` ���J��
    On Error Resume Next
    Set wb = Workbooks.Open(settingsFilePath)
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub
    
    ' �e�V�[�g�̖��O��ListBox2�ɒǉ�
    For Each ws In wb.Sheets
        ListBox2.AddItem ws.Name
    Next ws
    
    ' ���[�N�u�b�N�����
    wb.Close False
End Sub

Private Sub ListBox1_Click()
    ' ListBox1�ŗ񂪑I�����ꂽ�Ƃ���ComboBox1?ComboBox7���X�V
    LoadUniqueValuesToComboBoxes
End Sub
Private Sub CommandButton8_Click()
    ' CommandButton8���N���b�N���ꂽ�Ƃ��ɓh��Ԃ������s
    Dim ws As worksheet
    Dim columnNumber As Long
    Dim selectedValue As String
    Dim lastrow As Long
    Dim cell As Range
    Dim i As Long
    Dim filterCondition As String

    Set ws = ActiveSheet

    ' ListBox1�����ԍ����擾
    If ListBox1.ListIndex <> -1 Then
        columnNumber = ListBox1.ListIndex + 1 ' �I�����ꂽ��ԍ����擾
    Else
        MsgBox "��ԍ���I�����Ă��������B"
        Exit Sub
    End If

    ' �I��������̍ŏI�s���擾
    lastrow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    ' �eComboBox����I�������l���擾���A�Ή�����TextBox�̔w�i�F�Ŕ͈͂�h��Ԃ�
    For i = 1 To 7
        selectedValue = Me.Controls("ComboBox" & i).value ' ���݂�ComboBox�̑I��l���擾

        ' ComboBox����łȂ��ꍇ�̂ݏ��������s
        If selectedValue <> "" Then
            ' �t�B���^�������擾�iComboBoxA, ComboBoxB, ...�j
            filterCondition = Me.Controls("ComboBox" & Chr(64 + i)).value
            
            ' �f�o�b�O�p���b�Z�[�W��\��
            Debug.Print "filterCondition: " & filterCondition ' ������filterCondition��\��

            ' �t�B���^�������L�����m�F
            If filterCondition <> "��v" And filterCondition <> "�ȏ�" And filterCondition <> "�ȉ�" And filterCondition <> "�܂�" Then
                MsgBox "�����ȃt�B���^�����ł��B������������I�����Ă��������B"
                Exit Sub
            End If

            For Each cell In ws.Range(ws.Cells(2, columnNumber), ws.Cells(lastrow, columnNumber))
                Dim shouldFill As Boolean
                shouldFill = False ' ������

                ' �t�B���^�����ɂ��`�F�b�N
                If filterCondition = "��v" Then
                    If cell.value = selectedValue Then
                        shouldFill = True
                    End If
                ElseIf filterCondition = "�ȏ�" Then
                    If IsNumeric(cell.value) And IsNumeric(selectedValue) Then
                        If cell.value >= CDbl(selectedValue) Then
                            shouldFill = True
                        End If
                    End If
                ElseIf filterCondition = "�ȉ�" Then
                    If IsNumeric(cell.value) And IsNumeric(selectedValue) Then
                        If cell.value <= CDbl(selectedValue) Then
                            shouldFill = True
                        End If
                    End If
                ElseIf filterCondition = "�܂�" Then
                    If InStr(1, cell.value, selectedValue) > 0 Then
                        shouldFill = True
                    End If
                End If

                ' �h��Ԃ�����
                If shouldFill Then
                    If OptionButton1.value Then
                        ' OptionButton1���I������Ă���ꍇ�A�Y���Z����h��Ԃ�
                        cell.Interior.Color = Me.Controls("TextBox" & i).BackColor ' �Z����h��Ԃ�
                    ElseIf OptionButton2.value Then
                        ' OptionButton2���I������Ă���ꍇ�AA�񂩂�E�[�܂ł̍s��h��Ԃ�
                        ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, ws.Columns.Count).End(xlToLeft)).Interior.Color = Me.Controls("TextBox" & i).BackColor ' �s�S�̂�h��Ԃ�
                    End If
                End If
            Next cell
        End If
    Next i

    Unload Me
End Sub


' CommandButton9�œh��Ԃ��𖳂��ɂ���
Private Sub CommandButton9_Click()
    Dim ws As worksheet
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' �V�[�g�S�̂̓h��Ԃ��𖳂��ɂ���
    ws.Cells.Interior.ColorIndex = xlNone
End Sub

Private Sub CommandButton1_Click()
    SetTextBoxColor 1
End Sub

Private Sub CommandButton2_Click()
    SetTextBoxColor 2
End Sub

Private Sub CommandButton3_Click()
    SetTextBoxColor 3
End Sub

Private Sub CommandButton4_Click()
    SetTextBoxColor 4
End Sub

Private Sub CommandButton5_Click()
    SetTextBoxColor 5
End Sub

Private Sub CommandButton6_Click()
    SetTextBoxColor 6
End Sub

Private Sub CommandButton7_Click()
    SetTextBoxColor 7
End Sub

Private Sub SetTextBoxColor(textBoxIndex As Integer)
    Dim intresult As Long

    ' �J���[�_�C�A���O��\�����ĐF��I��
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        ' �I�������F���w�肳�ꂽTextBox�̔w�i�F�ɐݒ�
        intresult = ActiveWorkbook.colors(1) ' �I�����ꂽ�F���擾
        Me.Controls("TextBox" & textBoxIndex).BackColor = intresult ' �w�i�F��ݒ�
        
        ' �w�i�F��z��ɕۑ�
        previousTextBoxBackColors(textBoxIndex) = intresult
    End If
End Sub

Private Sub CommandButton10_Click()
    Dim settingName As String
    Dim wb As Workbook
    Dim ws As worksheet
    Dim i As Long
    Dim result As VbMsgBoxResult
    
    ' �ݒ薼���擾
    settingName = InputBox("�ۑ�����ݒ�̖��O����͂��Ă��������B")
    
    If settingName = "" Then
        MsgBox "�ݒ薼����͂��Ă��������B"
        Exit Sub
    End If
    
    ' `settings.xlsx` ���J�����A�V�K�쐬
    On Error Resume Next
    Set wb = Workbooks.Open(settingsFilePath)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Add
        wb.SaveAs ThisWorkbook.Path & "\settings.xlsx"
    End If
    
    ' �V�����V�[�g���쐬
    On Error Resume Next
    Set ws = wb.Sheets(settingName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        ws.Name = settingName
    Else
        result = MsgBox("�������O�̐ݒ肪���ɑ��݂��܂��B" & vbCrLf & "�㏑���ۑ����܂����H", vbYesNo + vbExclamation)
        If result = vbYes Then
        Application.DisplayAlerts = False ' �x�����b�Z�[�W���\���ɂ���
        ws.Delete ' �����̃V�[�g���폜
        Application.DisplayAlerts = True ' �x�����b�Z�[�W���ĕ\��
        Else
        Exit Sub
        End If
    End If
    
    ' ComboBox1�`7�̒l��A��ɕۑ�
    For i = 1 To 7
        ws.Cells(i, 1).value = Me.Controls("ComboBox" & i).value ' ComboBox1�`7��A���
    Next i
    
    ' ComboBoxA�`G�̒l��B��ɕۑ�
    For i = 1 To 7
        ws.Cells(i, 2).value = Me.Controls("ComboBox" & Chr(64 + i)).value ' ComboBoxA�`G��B���
    Next i
    
    ' TextBox1�`7�̔w�i�F��C��ɕۑ�
    For i = 1 To 7
        ws.Cells(i, 3).value = Me.Controls("TextBox" & i).BackColor ' TextBox1�`7�̔w�i�F��C���
    Next i
    
    ' �I�v�V�����{�^���̏�Ԃ�D��ɕۑ�
    ws.Cells(1, 4).value = IIf(OptionButton1.value, "1", "0") ' �I�v�V�����{�^��1�̏��
    ws.Cells(2, 4).value = IIf(OptionButton2.value, "1", "0") ' �I�v�V�����{�^��2�̏��
    
    ' �ۑ����ĕ���
    wb.Save
    wb.Close
    
    ' ListBox2�ɐݒ薼��ǉ�
    ListBox2.AddItem settingName
End Sub



Private Function GetComboBoxValues() As String
    Dim i As Integer
    Dim values As String
    For i = 1 To 7
        values = values & Me.Controls("ComboBox" & i).value & ","
    Next i
    GetComboBoxValues = Left(values, Len(values) - 1) ' �Ō�̃J���}���폜
End Function

Private Function GetTextBoxColors() As String
    Dim i As Integer
    Dim colors As String
    For i = 1 To 7
        colors = colors & Me.Controls("TextBox" & i).BackColor & ","
    Next i
    GetTextBoxColors = Left(colors, Len(colors) - 1) ' �Ō�̃J���}���폜
End Function

Private Sub SaveSettingsToFile()
    Dim fileNum As Integer
    Dim key As Variant
    
    fileNum = FreeFile
    Open settingsFilePath For Output As #fileNum
    
    ' ��������ݒ���t�@�C���ɏ�������
    For Each key In savedSettings.Keys
        Print #fileNum, key & "|" & savedSettings(key)
    Next key
    
    Close #fileNum
End Sub

Private Sub ListBox2_Click()
    Dim settingName As String
    Dim wb As Workbook
    Dim ws As worksheet
    Dim i As Long
    
    If ListBox2.ListIndex = -1 Then Exit Sub
    
    settingName = ListBox2.value
    
    ' `settings.xlsx` ���J��
    Set wb = Workbooks.Open(settingsFilePath)
    
    ' �Ή�����V�[�g���擾
    Set ws = wb.Sheets(settingName)
    
    ' ComboBox��TextBox�ɒl��ݒ�
    For i = 1 To 7
        Me.Controls("ComboBox" & i).value = ws.Cells(i, 1).value ' ComboBox1�`7��ݒ�
        Me.Controls("ComboBox" & Chr(64 + i)).value = ws.Cells(i, 2).value ' ComboBoxA�`G��ݒ�
        Me.Controls("TextBox" & i).BackColor = ws.Cells(i, 3).value ' TextBox1�`7�̔w�i�F��ݒ�
    Next i
    
    ' �I�v�V�����{�^���̏�Ԃ�ݒ�
    OptionButton1.value = (ws.Cells(1, 4).value = "1") ' D�񂩂�I�v�V�����{�^��1�̏�Ԃ�ݒ�
    OptionButton2.value = (ws.Cells(2, 4).value = "1") ' D�񂩂�I�v�V�����{�^��2�̏�Ԃ�ݒ�
    
    ' ���[�N�u�b�N�����
    wb.Close False
End Sub

Private Sub CommandButton11_Click()
    Dim key As Variant
    
    ' ListBox2����I�����ꂽ�ݒ薼���擾
    If ListBox2.ListIndex <> -1 Then
        key = ListBox2.value
        
        ' �ݒ���폜
        If savedSettings.Exists(key) Then
            savedSettings.Remove key
            ListBox2.RemoveItem ListBox2.ListIndex ' ListBox����폜
            SaveSettingsToFile ' �ݒ���t�@�C��������폜
        End If
    End If
End Sub

Private Sub UpdateListBox2()
    ' ListBox2���X�V����֐�
    Dim key As Variant
    
    ' ListBox2���N���A
    ListBox2.Clear
    
    ' �ۑ����ꂽ�ݒ薼�����X�g�{�b�N�X�ɒǉ�
    For Each key In savedSettings.Keys
        ListBox2.AddItem key
    Next key
End Sub

Private Sub LoadUniqueValuesToComboBoxes()
    Dim ws As worksheet
    Dim columnNumber As Variant
    Dim lastrow As Long
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim i As Long
    Dim item As Variant ' item�ϐ���錾
    Dim headernum As Long

    Set ws = ActiveSheet
    
    headernum = Val(HEADERBOX.value)
    
        ' ListBox1�̓��e���f�o�b�O�o��
    For i = 0 To ListBox1.ListCount - 1
        Debug.Print "Item " & i & ": " & ListBox1.List(i, 0) ' 1��ڂ̒l��\��
    Next i
    

    ' ListBox1�����ԍ����擾
    If ListBox1.ListIndex <> -1 Then  ' -1�͉����I������Ă��Ȃ����
        columnNumber = ListBox1.ListIndex + 1 ' �I�����ꂽ��ԍ����擾
    Else
        MsgBox "��ԍ���I�����Ă��������B"
        Exit Sub
    End If

    Debug.Print "Column Number: " & columnNumber

    ' �ŏI�s���擾
    lastrow = ws.Cells(ws.Rows.Count, columnNumber + 1).End(xlUp).Row

    ' ���j�[�N�Ȓl���i�[����R���N�V������������
    Set uniqueValues = New Collection

    ' �w�肵����̃��j�[�N�Ȓl�����W
    On Error Resume Next ' �d���G���[�𖳎�
    For Each cell In ws.Range(ws.Cells(headernum + 1, columnNumber), ws.Cells(lastrow, columnNumber))
        If cell.value <> "" Then
            uniqueValues.Add cell.value, CStr(cell.value) ' ���j�[�N�Ȓl��ǉ�
        End If
    Next cell
    On Error GoTo 0 ' �G���[�n���h�����O������

    ' ComboBox1�`7�Ƀ��j�[�N�Ȓl��ǉ�
    For i = 1 To 7
        With Me.Controls("ComboBox" & i)
            .Clear ' �����̃A�C�e�����N���A
            For Each item In uniqueValues
                .AddItem item ' ���j�[�N�Ȓl��ǉ�
            Next item
        End With
    Next i
End Sub



