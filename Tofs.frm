VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tofs 
   Caption         =   "THOPS�f�[�^���o�t�H�[��"
   ClientHeight    =   4770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6340
   OleObjectBlob   =   "Tofs.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Tofs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As Object ' �ڑ��I�u�W�F�N�g�����W���[�����x���Ő錾

Private Sub UserForm_Initialize()
    Dim wsActive As Worksheet
    Dim j As Integer
    Dim rs As Object
    Dim sql As String
    
    ' �A�N�e�B�u�V�[�g���擾
    Set wsActive = ActiveSheet

    ' HEADERBOX��1�`10�̒l��ǉ�
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.Value = "1"
    
    On Error GoTo ErrorHandler
    
    ' MySQL�ڑ���� (DSN���X�ڑ�)
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Driver={MySQL ODBC 9.1 Unicode Driver};" & _
                            "Server=localhost;" & _
                            "port=33061;" & _
                            "Database=tofs;" & _
                            "User=root;" & _
                            "Password=password;" & _
                            "Option=3;"
    conn.Open
    
    ' SQL�N�G��: �e�[�u���̃t�B�[���h�����擾
    sql = "SHOW FIELDS FROM items"
    Set rs = conn.Execute(sql)
    
    ' ListBox1�Ƀt�B�[���h�� (Field��) ��ǉ�
    ListBox1.Clear
    ListBox1.ColumnCount = 1 ' 1��̂ݕ\��
    
    Do Until rs.EOF
        ListBox1.AddItem rs.Fields("Field").Value
        rs.MoveNext
    Loop
    
    ' �㏈��
    rs.Close
    Set rs = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

Private Sub ComboBox2_Change()
    Dim selectedValue As String
    Dim columnNumber As Variant
    Dim ws As Worksheet
    
    ' �A�N�e�B�u�V�[�g���擾
    Set ws = ActiveSheet
    
    ' ComboBox2�őI�����ꂽ�l���擾
    selectedValue = ComboBox2.Value
    
    ' �I�����ꂽ�l�ɑΉ������ԍ����擾
    On Error Resume Next ' �G���[�n���h�����O��L���ɂ���
    columnNumber = Application.Match(selectedValue, ws.Rows(1), 0)
    On Error GoTo 0 ' �G���[�n���h�����O�𖳌��ɂ���
End Sub

Private Sub HEADERBOX_Change()
    Dim targetRow As Long
    Dim lastColumn As Long
    Dim wsActive As Worksheet
    Dim i As Long
    Dim j As Variant

    ' �A�N�e�B�u�V�[�g���擾
    Set wsActive = ActiveWorkbook.ActiveSheet
    
    ' HEADERBOX�̒l���擾
    targetRow = HEADERBOX.Value
    
    ' HEADERBOX�̒l��1����10�͈͓̔����m�F
    If targetRow < 1 Or targetRow > 10 Then
        MsgBox "HEADERBOX�̒l��1����10�͈̔͂ł���K�v������܂��B", vbExclamation
        Exit Sub
    End If

    ' �Ō�̗���擾
    lastColumn = wsActive.Cells(targetRow, wsActive.Columns.Count).End(xlToLeft).Column

    ' ComboBox2���N���A
    ComboBox2.Clear

    ' �w�肵���s (targetRow) ��1��ځ`lastColumn�܂ł̒l��ComboBox2�ɒǉ�
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).Value
        
        ' �Z���̒l����łȂ��ꍇ�ɂ̂ݒǉ�
        If Not IsEmpty(j) Then
            ComboBox2.AddItem CStr(j)  ' �l�𕶎���ɕϊ����Ēǉ�
        End If
    Next i
    
End Sub

Private Sub btnFetchData_Click()
    Dim rs As Object
    Dim sql As String
    Dim selectedFields As String
    Dim skuValue As String
    Dim lastRow As Long
    Dim i As Long, j As Long
    
    On Error GoTo ErrorHandler
    
    ' �ڑ����m�F���A�K�v�Ȃ�Đڑ�
    If conn Is Nothing Or conn.State = 0 Then
        MsgBox "�ڑ����m������Ă��܂���B�Đڑ����܂��B", vbExclamation
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionString = "Driver={MySQL ODBC 9.1 Unicode Driver};" & _
                                "Server=localhost;" & _
                                "port=33061;" & _
                                "Database=tofs;" & _
                                "User=root;" & _
                                "Password=password;" & _
                                "Option=3;"
        conn.Open
    End If
    
    ' ListBox����I�����ꂽ�t�B�[���h���擾
    If ListBox1.ListIndex = -1 Then
        MsgBox "���Ȃ��Ƃ�1�̃t�B�[���h��I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �����I�����ꂽ�t�B�[���h���J���}�ŘA��
    selectedFields = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            If selectedFields <> "" Then
                selectedFields = selectedFields & ", "
            End If
            selectedFields = selectedFields & ListBox1.List(i, 0)
        End If
    Next i
    
    If selectedFields = "" Then
        MsgBox "�t�B�[���h���I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Dim ws As Worksheet
    Set ws = ActiveSheet ' �A�N�e�B�u�V�[�g���g�p
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 1�s�ڂɃt�B�[���h����}��
    Dim fieldArray() As String
    fieldArray = Split(selectedFields, ", ")
    
    For i = 0 To UBound(fieldArray)
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = fieldArray(i)
    Next i
    
    ' 2�s�ڂ���ŏI�s�܂Ńf�[�^���擾
    For i = 2 To lastRow
        skuValue = ws.Cells(i, ComboBox2.ListIndex + 1).Value ' ComboBox2�őI�����ꂽ��
        
        If skuValue <> "" Then
            sql = "SELECT " & selectedFields & " FROM items WHERE sku = '" & skuValue & "'"
            Set rs = conn.Execute(sql)
            
            If Not rs.EOF Then
                For j = 0 To UBound(fieldArray)
                    ws.Cells(i, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = rs.Fields(j).Value
                Next j
            Else
                ws.Cells(i, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = ""
            End If
        End If
    Next i
    
    MsgBox "�f�[�^�𐳏�Ɏ擾���܂����I", vbInformation

    ' �N���[���A�b�v
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub


