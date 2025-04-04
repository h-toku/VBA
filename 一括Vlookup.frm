VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ꊇVlookup 
   Caption         =   "�ꊇVlookup"
   ClientHeight    =   5830
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11500
   OleObjectBlob   =   "�ꊇVlookup.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ꊇVlookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    ComboBox1.Clear

    ' �w�肵���s (targetRow) ��1��ځ`lastColumn�܂ł̒l��ComboBox1�ɒǉ�
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).value
        
        ' �Z���̒l����łȂ��ꍇ�ɂ̂ݒǉ�
        If Not IsEmpty(j) Then
            ComboBox1.AddItem CStr(j)  ' �l�𕶎���ɕϊ����Ēǉ�
        End If
    Next i
    
End Sub

Private Sub DeleteButton_Click()
    On Error Resume Next ' �G���[�𖳎����đ��s

    If ListBox1.ListIndex <> -1 Then ' �I������Ă���A�C�e��������ꍇ
        ListBox1.RemoveItem ListBox1.ListIndex
    End If

    On Error GoTo 0 ' �G���[�n���h�����O�����ɖ߂��i�ʏ�̃G���[�n���h�����O�ɖ߂��j
End Sub

Private Sub UserForm_Initialize()
    Dim wsActive As worksheet
    Dim wsTree As worksheet
    Dim wsList As worksheet
    Dim lastColumn As Long
    Dim i As Long
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim ws As worksheet
    Dim targetRow As Long
    Dim j As Integer

    ' �A�N�e�B�u�V�[�g���擾
    Set wsActive = ActiveSheet
    
    TextBox1.value = ""

    'HEADERBOX��1�`10�̒l��ǉ�
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.value = "1"
    
    ' ��������.xlsx �̃p�X��ݒ�
    searchBookPath = ThisWorkbook.Path & "\��������.xlsx"
    
    ' ��������.xlsx �����݂���ꍇ�A�V�[�g����ListBox2�ɕ\��
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath, ReadOnly:=True)
        
        ' �V�[�g����ListBox2�ɒǉ�
        For Each wsList In searchBook.Worksheets
            Me.ListBox2.AddItem wsList.Name
        Next wsList
        
        ' �u�b�N�����
        searchBook.Close SaveChanges:=False
    End If
    
    ' ListBoxP�ɃA�N�e�B�u�V�[�g�ȊO�̃V�[�g����ǉ�
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> wsActive.Name Then
            ListBoxP.AddItem ws.Name
        End If
    Next ws
End Sub

Private Sub ComboBox1_Change()
    Dim selectedValue As String
    Dim columnNumber As Variant ' �ύX: Long����Variant��
    Dim ws As worksheet
    
    ' �A�N�e�B�u�V�[�g���擾
    Set ws = ActiveSheet
    
    ' ComboBox1�őI�����ꂽ�l���擾
    selectedValue = ComboBox1.value
    
    ' �I�����ꂽ�l�ɑΉ������ԍ����擾
    On Error Resume Next ' �G���[�n���h�����O��L���ɂ���
    columnNumber = Application.Match(selectedValue, ws.Rows(1), 0)
    On Error GoTo 0 ' �G���[�n���h�����O�𖳌��ɂ���
    
End Sub

Private Sub ListBoxP_Change()
    ' ���X�g�{�b�N�XC�ɑI�����ꂽ�V�[�g��1�s�ڂ̍��ڂ�\��
    Dim ws As worksheet
    Dim i As Long
    
    ListBoxC.Clear ' ���X�g�{�b�N�XC���N���A
    Set ws = ActiveWorkbook.Sheets(ListBoxP.value) ' �I�����ꂽ�V�[�g���擾
    
    ' 1�s�ڂ̍��ڂ����X�g�{�b�N�XC�ɒǉ�
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ListBoxC.AddItem ws.Cells(1, i).value
    Next i
End Sub

Private Sub ADDButton_Click()
    Dim selectedItem As String
    Dim selectedSheet As String
    Dim combinationExists As Boolean
    Dim i As Long
    Dim optionState As String ' �I�v�V�����{�^���̏�Ԃ�ێ�
    
    If ListBoxC.ListIndex = -1 Then
        MsgBox "�ΏۃV�[�g��I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
        ' ���X�g�{�b�N�X���I������Ă��邩�m�F
    If ListBoxC.ListIndex = -1 Then
        MsgBox "�Ώۗ񖼂�I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �I�����ꂽ�A�C�e�����擾
    selectedItem = ListBoxC.value
    selectedSheet = ListBoxP.value
    combinationExists = False
    
    ' �I�v�V�����{�^���̏�Ԃ��m�F
    If OptionButton1.value Then
        optionState = "V(�G���[)"
    ElseIf OptionButton2.value Then
        optionState = "V(0)"
    ElseIf OptionButton3.value Then
        optionState = "SUMIF"
    Else: MsgBox "�������@��I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' ���X�g�{�b�N�X1�̊e�A�C�e�����`�F�b�N
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.List(i, 0) = selectedSheet And _
            ListBox1.List(i, 1) = selectedItem And _
            ListBox1.List(i, 2) = optionState And _
            ListBox1.List(i, 3) = TextBox1.value Then
            combinationExists = True
            Exit For
        End If
    Next i
    
    ' �����g�ݍ��킹�����݂��Ȃ��ꍇ�̂݁A���X�g�{�b�N�X1�ɒǉ�
    If Not combinationExists Then
        ListBox1.AddItem selectedSheet ' 1��ڂɃ��X�g�{�b�N�XP�̒l��ǉ�
        ListBox1.List(ListBox1.ListCount - 1, 1) = selectedItem ' 2��ڂɃ��X�g�{�b�N�XC�̑I�����ڂ�ǉ�
        ListBox1.List(ListBox1.ListCount - 1, 2) = optionState ' 3��ڂɃI�v�V�����{�^���̏�Ԃ�ǉ�
        If Not TextBox1.value = "" Then
            ListBox1.List(ListBox1.ListCount - 1, 3) = TextBox1.value ' 4��ڂɍ��ږ���ǉ�
            Else: ListBox1.List(ListBox1.ListCount - 1, 3) = selectedSheet & selectedItem
        End If
    End If
    
    TextBox1.value = ""
    
End Sub

' UPButton�N���b�N�C�x���g: �I�����Ă��鍀�ڂ�1��Ɉړ�
Private Sub UPButton_Click()
    Dim selectedIndex As Long
    Dim temp1 As String, temp2 As String, temp3 As String

    ' �I������Ă��鍀�ڂ̃C���f�b�N�X���擾
    selectedIndex = Me.ListBox1.ListIndex
    
    ' �C���f�b�N�X��1�ȏ�̏ꍇ�i2�ڈȍ~�̍��ڂ��I������Ă���ꍇ�j�̂ݏ��������s
    If selectedIndex > 0 Then
        ' ���ݑI������Ă��鍀�ڂ̓��e���ꎞ�I�ɕێ�
        temp1 = Me.ListBox1.List(selectedIndex, 0)
        temp2 = Me.ListBox1.List(selectedIndex, 1)
        temp3 = Me.ListBox1.List(selectedIndex, 2)
        temp4 = Me.ListBox1.List(selectedIndex, 3)
        
        ' �I�����ꂽ���ڂƂ���1��̍��ڂ����ւ�
        Me.ListBox1.List(selectedIndex, 0) = Me.ListBox1.List(selectedIndex - 1, 0)
        Me.ListBox1.List(selectedIndex, 1) = Me.ListBox1.List(selectedIndex - 1, 1)
        Me.ListBox1.List(selectedIndex, 2) = Me.ListBox1.List(selectedIndex - 1, 2)
        Me.ListBox1.List(selectedIndex, 3) = Me.ListBox1.List(selectedIndex - 1, 3)
        
        Me.ListBox1.List(selectedIndex - 1, 0) = temp1
        Me.ListBox1.List(selectedIndex - 1, 1) = temp2
        Me.ListBox1.List(selectedIndex - 1, 2) = temp3
        Me.ListBox1.List(selectedIndex - 1, 3) = temp4
        
        ' ���ڂ�I����Ԃɖ߂�
        Me.ListBox1.ListIndex = selectedIndex - 1
    End If
End Sub

' DownButton�N���b�N�C�x���g: �I�����Ă��鍀�ڂ�1���Ɉړ�
Private Sub DownButton_Click()
    Dim selectedIndex As Long
    Dim temp1 As String, temp2 As String, temp3 As String

    ' �I������Ă��鍀�ڂ̃C���f�b�N�X���擾
    selectedIndex = Me.ListBox1.ListIndex
    
    ' �C���f�b�N�X���ŏI�s�����̏ꍇ�i���ɍ��ڂ�����ꍇ�j�̂ݏ��������s
    If selectedIndex <> -1 And selectedIndex < Me.ListBox1.ListCount - 1 Then
        ' ���ݑI������Ă��鍀�ڂ̓��e���ꎞ�I�ɕێ�
        temp1 = Me.ListBox1.List(selectedIndex, 0)
        temp2 = Me.ListBox1.List(selectedIndex, 1)
        temp3 = Me.ListBox1.List(selectedIndex, 2)
        temp4 = Me.ListBox1.List(selectedIndex, 3)
        
        ' �I�����ꂽ���ڂƂ���1���̍��ڂ����ւ�
        Me.ListBox1.List(selectedIndex, 0) = Me.ListBox1.List(selectedIndex + 1, 0)
        Me.ListBox1.List(selectedIndex, 1) = Me.ListBox1.List(selectedIndex + 1, 1)
        Me.ListBox1.List(selectedIndex, 2) = Me.ListBox1.List(selectedIndex + 1, 2)
        Me.ListBox1.List(selectedIndex, 3) = Me.ListBox1.List(selectedIndex + 1, 3)
        
        Me.ListBox1.List(selectedIndex + 1, 0) = temp1
        Me.ListBox1.List(selectedIndex + 1, 1) = temp2
        Me.ListBox1.List(selectedIndex + 1, 2) = temp3
        Me.ListBox1.List(selectedIndex + 1, 3) = temp4
        
        ' ���ڂ�I����Ԃɖ߂�
        Me.ListBox1.ListIndex = selectedIndex + 1
    End If
End Sub

Private Function itemExists(value As String) As Boolean
    Dim i As Long
    itemExists = False
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.List(i) = value Then
            itemExists = True
            Exit Function
        End If
    Next i
End Function

' �w�肵���V�[�g�������݂��邩�m�F����֐�
Function sheetExists(wb As Workbook, sheetName As String) As Boolean
    On Error Resume Next
    sheetExists = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub CommandButton3_Click()
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim ws As worksheet
    Dim settingName As String
    Dim lastrow As Long
    Dim i As Long
    Dim result As VbMsgBoxResult
    
    ' �ݒ薼����͂�����C���v�b�g�{�b�N�X��\��
    settingName = InputBox("�ݒ薼����͂��Ă�������:", "�ݒ薼�̓���")
    
    ' �ݒ薼�����͂���Ă��Ȃ��ꍇ�A�������I��
    If settingName = "" Then
        MsgBox "�ݒ薼�����͂���Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' ��������.xlsx �̃p�X��ݒ�
    searchBookPath = ThisWorkbook.Path & "\��������.xlsx"
    
    ' ��������.xlsx �����݂���ꍇ�A�u�b�N���J���B���݂��Ȃ��ꍇ�A�V�K�쐬
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath)
    Else
        Set searchBook = Workbooks.Add
        searchBook.SaveAs searchBookPath
    End If
    
    ' �w�肳�ꂽ�ݒ薼�̃V�[�g�����݂��邩�m�F
    On Error Resume Next
    Set ws = searchBook.Worksheets(settingName)
    On Error GoTo 0
    
    ' �������O�̃V�[�g�����ɑ��݂���ꍇ
    If Not ws Is Nothing Then
        result = MsgBox("�������O�̃V�[�g�����ɑ��݂��܂��B�㏑�����܂����H", vbYesNo + vbExclamation)
        If result = vbYes Then
        Application.DisplayAlerts = False ' �x�����b�Z�[�W���\���ɂ���
        ws.Delete ' �����̃V�[�g���폜
        Application.DisplayAlerts = True ' �x�����b�Z�[�W���ĕ\��
        Else
        Exit Sub
        End If
    End If
    
    ' �V�����V�[�g��ǉ����A�ݒ薼��ݒ�
    Set ws = searchBook.Worksheets.Add
    ws.Name = settingName
    
    ' ComboBox1�̒l��A1�Z���ɁAListBox1�̒l��B���C��ɏ�������
    ws.Cells(1, 1).value = HEADERBOX.value
    ws.Cells(2, 1).value = ComboBox1.value
    
    lastrow = ListBox1.ListCount
    For i = 0 To lastrow - 1
        ws.Cells(i + 2, 2).value = ListBox1.List(i, 0)
        ws.Cells(i + 2, 3).value = ListBox1.List(i, 1)
        ws.Cells(i + 2, 4).value = ListBox1.List(i, 2)
        ws.Cells(i + 2, 5).value = ListBox1.List(i, 3)
        
    Next i
    
    ' �u�b�N��ۑ����ĕ���
    searchBook.Close SaveChanges:=True
    
    MsgBox "�ݒ肪�ۑ�����܂����B", vbInformation
End Sub

Private Sub CommandButton4_Click()
    Dim searchBookPath As String
    Dim newWb As Workbook
    Dim wsName As String
    
    ' ListBox2�ō��ڂ��I������Ă��邩�m�F
    If Me.ListBox2.ListIndex = -1 Then
        MsgBox "�폜����V�[�g��I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �I�����ꂽ�V�[�g�����擾
    wsName = Me.ListBox2.value
    
    ' ��������.xlsx �̃p�X���쐬
    searchBookPath = ThisWorkbook.Path & "\��������.xlsx"
    
    ' ��������.xlsx ���J��
    Set newWb = Workbooks.Open(searchBookPath)
    
    ' ListBox2�Ɏc���Ă��鍀�ڐ����m�F
    If Me.ListBox2.ListCount = 1 Then
        ' ���ڂ�1�����̏ꍇ�A�u�b�N���폜
        newWb.Close SaveChanges:=False
        Kill searchBookPath ' �u�b�N���폜
        MsgBox "�u�b�N '" & searchBookPath & "' ���폜����܂����B", vbInformation
    Else
        ' ���ڂ������̏ꍇ�A�I�����ꂽ�V�[�g�̂ݍ폜
        Application.DisplayAlerts = False ' �폜�m�F�_�C�A���O���\���ɂ���
        On Error Resume Next
        newWb.Worksheets(wsName).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        
        ' �u�b�N��ۑ����ĕ���
        newWb.Save
        newWb.Close SaveChanges:=True
        
        ' ListBox2���X�V
        Me.ListBox2.RemoveItem Me.ListBox2.ListIndex
        
        MsgBox "�V�[�g '" & wsName & "' ���폜����܂����B", vbInformation
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim selectedSheetName As String
    Dim ws As worksheet
    Dim i As Long
    
    ' ��������.xlsx �̃p�X��ݒ�
    searchBookPath = ThisWorkbook.Path & "\��������.xlsx"
    
    ' �I�����ꂽ�V�[�g�����擾
    selectedSheetName = Me.ListBox2.value
    
    ' ��������.xlsx ���J���đI�����ꂽ�V�[�g���擾
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath, ReadOnly:=True)
        Set ws = searchBook.Worksheets(selectedSheetName)
        
        ' HEADERBOX��A1�Z���̒l�𔽉f
        Me.ComboBox1.value = ws.Cells(1, 1).value
        ' ComboBox1��A2�Z���̒l�𔽉f
        Me.ComboBox1.value = ws.Cells(2, 1).value
        
        ' ListBox1�ɃV�[�g��B���C��̒l�𔽉f
        Me.ListBox1.Clear
        For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            Me.ListBox1.AddItem
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = ws.Cells(i, 2).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 3).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 4).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = ws.Cells(i, 5).value
        Next i
        
        ' �u�b�N�����
        searchBook.Close SaveChanges:=False
    End If
End Sub

Private Sub UserForm_Terminate()
    Dim searchBookPath As String
    Dim newWb As Workbook
    
    ' ��������.xlsx �̃p�X���쐬
    searchBookPath = ThisWorkbook.Path & "\��������.xlsx"
    
    ' ��������.xlsx �����ɊJ����Ă��邩�m�F
    On Error Resume Next
    Set newWb = Workbooks("��������.xlsx")
    On Error GoTo 0
    
    ' �u�b�N���J����Ă���ꍇ�̂ݕ���
    If Not newWb Is Nothing Then
        On Error Resume Next ' ���łɃu�b�N���Ȃ���Ԃ͖���
        newWb.Close SaveChanges:=True ' �K�v�ɉ����� SaveChanges �� False �ɕύX
        On Error GoTo 0
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim i As Long
    Dim listItem As String
    Dim sheetName As String
    Dim optionValue As String
    Dim header As String
    Dim resultColumn As Long
    
    ' ListBox1�̑S���ڂɑ΂��ď������s��
    For i = 0 To ListBox1.ListCount - 1
        ' 1��ڂ��V�[�g���A2��ڂ����ځA3��ڂ��I�v�V�����{�^���̏�ԂƉ���
        sheetName = ListBox1.List(i, 0) ' 1��ځi�V�[�g���j
        listItem = ListBox1.List(i, 1) ' 2��ځi���X�g���ځj
        optionValue = ListBox1.List(i, 2) ' 3��ځi�I�v�V�����{�^���̏�ԁj
        header = ListBox1.List(i, 3) ' 4��ځi���ږ��j
        
        ' 3��ڂ̒l�ɉ����ď�����U�蕪����
        If optionValue = "V(�G���[)" Then
            Call ProcessOption1(sheetName, listItem, header)
        ElseIf optionValue = "V(0)" Then
            Call ProcessOption2(sheetName, listItem, header)
        Else
            Call ProcessOption3(sheetName, listItem, header)
        End If
        
    Next i
    
    If MsgBox("�������������܂����B" & vbCrLf & "������ۑ����܂����H", vbYesNo + vbQuestion) = vbYes Then
        Call CommandButton3_Click
        End If
        
        Unload Me
        
End Sub

Private Sub ProcessOption1(sheetName As String, listItem As String, header As String)
    ' VLOOKUP����
    Dim ws As worksheet
    Dim lastrow As Long
    Dim resultColumn As Long
    Dim lookupValue As String
    Dim lookupColumn As Long
    Dim targetWorkbook As Workbook
    Dim j As Long
    Dim selectedValue As String
    Dim searchRange As Range
    Dim lookupResult As Variant

    Set targetWorkbook = Application.ActiveWorkbook
    lookupValue = ComboBox1.value

    ' �A�N�e�B�u�V�[�g�̌�����ԍ����擾
    lookupColumn = Application.Match(lookupValue, ActiveSheet.Rows(1), 0)
    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, lookupColumn).End(xlUp).Row
    resultColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1
    
    ' �Ώۂ̃V�[�g���擾
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

        If Not ws Is Nothing Then
            selectedValue = listItem ' ListBox1�̎q�m�[�h�̃w�b�_�[

            ' �e�m�[�h�i�R���{�{�b�N�X�̒l�j�ɑΉ������ԍ����擾
            lookupColumn = Application.Match(lookupValue, ws.Rows(1), 0)
            Dim childColumn As Long
            childColumn = Application.Match(selectedValue, ws.Rows(1), 0)

            If Not IsError(lookupColumn) And Not IsError(childColumn) Then
                ' �����͈͂�ݒ�
            Set searchRange = ws.Range(ws.Cells(2, lookupColumn), ws.Cells(ws.Rows.Count, lookupColumn).End(xlUp)).Resize(, childColumn - lookupColumn + 1)

                ' �A�N�e�B�u�V�[�g�̊e�s�����[�v���AVLOOKUP�����s
                For j = 2 To lastrow ' 2�s�ڂ���ŏI�s�܂�
                    Dim skuValue As Variant
                    skuValue = ActiveSheet.Cells(j, Application.Match(lookupValue, ActiveSheet.Rows(1), 0)).value ' �����l���擾
                    
                        If Not IsEmpty(skuValue) And Len(Trim(skuValue)) > 0 Then
                        '�Z�����󔒂łȂ�or�X�y�[�X��^�u�̋󔒕����݂̂łȂ��̏ꍇ
                
                    ' VLOOKUP�̎��s
                    lookupResult = Application.Vlookup(skuValue, searchRange, childColumn - lookupColumn + 1, False)

                    ' ���ʂ��A�N�e�B�u�V�[�g�ɒǉ�
                    If Not IsError(lookupResult) Then
                        ActiveSheet.Cells(j, resultColumn).value = lookupResult
                    Else
                    ActiveSheet.Cells(j, resultColumn).value = CVErr(xlErrNA)
                        End If
                    End If
            Next j
            
            ' �w�b�_�[��ǉ�
            ActiveSheet.Cells(1, resultColumn).value = header
                resultColumn = resultColumn + 1 ' ���̗�Ɉړ�

            Else
                MsgBox "�񂪌�����܂���: " & lookupValue & " �܂��� " & selectedValue
            End If
        Else
            MsgBox "�V�[�g��������܂���"
        End If

End Sub

Private Sub ProcessOption2(sheetName As String, listItem As String, header As String)
    ' VLOOKUP�G���[��0�ɒu���������
    Dim ws As worksheet
    Dim lastrow As Long
    Dim resultColumn As Long
    Dim lookupValue As String
    Dim lookupColumn As Long
    Dim targetWorkbook As Workbook
    Dim i As Long, j As Long
    Dim selectedValue As String
    Dim searchRange As Range
    Dim lookupResult As Variant

    Set targetWorkbook = Application.ActiveWorkbook
    lookupValue = ComboBox1.value

    ' �A�N�e�B�u�V�[�g�̌�����ԍ����擾
    lookupColumn = Application.Match(lookupValue, ActiveSheet.Rows(1), 0)
    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, lookupColumn).End(xlUp).Row
    resultColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1

    ' �Ώۂ̃V�[�g���擾
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

        If Not ws Is Nothing Then
            selectedValue = listItem ' ListBox1�̎q�m�[�h�̃w�b�_�[

            ' �e�m�[�h�i�R���{�{�b�N�X�̒l�j�ɑΉ������ԍ����擾
            lookupColumn = Application.Match(lookupValue, ws.Rows(1), 0)
            Dim childColumn As Long
            childColumn = Application.Match(selectedValue, ws.Rows(1), 0)

            If Not IsError(lookupColumn) And Not IsError(childColumn) Then
                ' �����͈͂�ݒ�
                Set searchRange = ws.Range(ws.Cells(2, lookupColumn), ws.Cells(ws.Rows.Count, lookupColumn).End(xlUp)).Resize(, childColumn - lookupColumn + 1)

                ' �A�N�e�B�u�V�[�g�̊e�s�����[�v���AVLOOKUP�����s
                For j = 2 To lastrow ' 2�s�ڂ���ŏI�s�܂�
                    Dim skuValue As Variant
                    skuValue = ActiveSheet.Cells(j, Application.Match(lookupValue, ActiveSheet.Rows(1), 0)).value ' �����l���擾
                    
                        If Not IsEmpty(skuValue) And Len(Trim(skuValue)) > 0 Then
                        '�Z�����󔒂łȂ�or�X�y�[�X��^�u�̋󔒕����݂̂łȂ��̏ꍇ
                
                    ' VLOOKUP�̎��s
                    lookupResult = Application.Vlookup(skuValue, searchRange, childColumn - lookupColumn + 1, False)

                    ' ���ʂ��A�N�e�B�u�V�[�g�ɒǉ�
                    If Not IsError(lookupResult) Then
                        ActiveSheet.Cells(j, resultColumn).value = lookupResult
                    Else
                            ActiveSheet.Cells(j, resultColumn).value = 0
                        End If
                End If
            Next j

                ' �w�b�_�[��ǉ�
                ActiveSheet.Cells(1, resultColumn).value = header
                resultColumn = resultColumn + 1 ' ���̗�Ɉړ�
            Else
                MsgBox "�񂪌�����܂���: " & lookupValue & " �܂��� " & selectedValue
            End If
        Else
            MsgBox "�V�[�g��������܂���"
        End If
End Sub

Private Sub ProcessOption3(sheetName As String, listItem As String, header As String)
    ' SUMIF�������s����
    Dim wsActive As worksheet
    Dim wsTarget As worksheet
    Dim comboValue As String
    Dim foundCellList As Range
    Dim criteriaRange As Range
    Dim sumRange As Range
    Dim rng As Range
    Dim lastrow As Long
    Dim lastCol As Long
    Dim criteriaCol As Long
    Dim sumResult As Double
    Dim j As Long
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set wsActive = ActiveSheet
    
    ' ComboBox1�̒l���擾
    comboValue = Me.ComboBox1.value
    
    ' ComboBox1�̒l�ƈ�v����������
    criteriaCol = Application.Match(comboValue, wsActive.Rows(1), 0)
    
    If IsError(criteriaCol) Then
        MsgBox "���������̗񂪌�����܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �A�N�e�B�u�V�[�g�̍ŏI�s���擾
    lastrow = wsActive.Cells(wsActive.Rows.Count, criteriaCol).End(xlUp).Row
    
    ' ���������͈͂�ݒ�i�A�N�e�B�u�V�[�g��2�s�ڂ���ŏI�s�܂Łj
    Set criteriaRange = wsActive.Range(wsActive.Cells(2, criteriaCol), wsActive.Cells(lastrow, criteriaCol))
    
    ' �A�N�e�B�u�V�[�g�̍ŏI����擾
    lastCol = wsActive.Cells(2, wsActive.Columns.Count).End(xlToLeft).Column + 1
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set wsTarget = Worksheets(sheetName) ' sheetName ���������񋟂���Ă���Ɖ���
    
    ' ListBox�̒l�ƈ�v����Z��������
    Dim selectedCellValue As String
    selectedCellValue = listItem ' listItem �� ListBox ����̑I�����ꂽ�l���܂ނƉ���
    Set foundCellList = wsTarget.Cells.Find(What:=selectedCellValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCellList Is Nothing Then
        ' ComboBox1�̒l�ƈ�v������ΏۃV�[�g�Ō���
        Dim targetCriteriaCol As Long
        targetCriteriaCol = Application.Match(comboValue, wsTarget.Rows(1), 0)
        
        If IsError(targetCriteriaCol) Then
            MsgBox "�ΏۃV�[�g�Ō��������̗񂪌�����܂���B", vbExclamation
            Exit Sub
        End If
        
        ' �����͈́irng�j�ƍ��v�͈́isumRange�j��ݒ�
        Set rng = wsTarget.Range(wsTarget.Cells(2, targetCriteriaCol), wsTarget.Cells(wsTarget.Rows.Count, targetCriteriaCol).End(xlUp))
        Set sumRange = wsTarget.Range(wsTarget.Cells(2, foundCellList.Column), wsTarget.Cells(wsTarget.Rows.Count, foundCellList.Column).End(xlUp))
        
        ' ���ʂ��A�N�e�B�u�V�[�g�̍ŉE��ɔ��f
        wsActive.Cells(1, lastCol).value = header
        
        ' criteria�̃Z���ɑΉ�����S�Ă̌��ʂ��v�Z���ăA�N�e�B�u�V�[�g�ɑ}��
        For j = 1 To criteriaRange.Rows.Count
            ' �󔒃Z�����X�L�b�v
            If Len(Trim(criteriaRange.Cells(j, 1).value)) > 0 Then
                On Error Resume Next
                ' SUMIF�����s���A���ʂ��v�Z
                sumResult = Application.WorksheetFunction.SumIf(rng, criteriaRange.Cells(j, 1).value, sumRange)
                
                wsActive.Cells(j + 1, lastCol).value = sumResult
                On Error GoTo 0
            End If
        Next j
    Else
        MsgBox "ListBox1�̒l�ƈ�v����Z����������܂���B", vbExclamation
    End If
End Sub



