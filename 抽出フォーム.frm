VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���o�t�H�[�� 
   Caption         =   "�t�H�[��"
   ClientHeight    =   4310
   ClientLeft      =   40
   ClientTop       =   150
   ClientWidth     =   6450
   OleObjectBlob   =   "���o�t�H�[��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���o�t�H�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim headerValue As String
    Dim lastCol As Long
    
    ' �ŏI����擾
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    ' ���X�g�{�b�N�X�Ƀw�b�_�[�̒l��ǉ�
    For i = 1 To lastCol ' �ŏI��܂Ń��[�v
        headerValue = ActiveSheet.Cells(1, i).value ' �w�b�_�[�̒l���擾
        If headerValue <> "" Then ' �w�b�_�[����łȂ��ꍇ�̂ݒǉ�
            ListBox1.AddItem headerValue ' �w�b�_�[�̒l��ǉ�
            ListBox1.List(ListBox1.ListCount - 1, 1) = i ' ��ԍ����\���̗�ɐݒ�
        End If
    Next i
End Sub

Sub CommandButton1_Click()

    Dim ws As worksheet, ws2 As worksheet
    Set ws = ActiveSheet ' �A�N�e�B�u�V�[�g��ݒ�
    
    Dim lrow As Long
    lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    Dim sheetName As String
    Dim validSheetName As Boolean
    Dim columnNumber As Long
    Dim errorMsg As String ' �G���[���b�Z�[�W�p�̕ϐ�
    Dim hasError As Boolean ' �G���[�t���O

    hasError = False ' �G���[�t���O��������
    
    ' ���X�g�{�b�N�X�����ԍ����擾
    If ListBox1.ListIndex <> -1 Then
        columnNumber = ListBox1.ListIndex + 1 ' �I�����ꂽ��ԍ����擾�i�C���f�b�N�X����1�𑫂��j
    Else
        MsgBox "��ԍ���I�����Ă��������B"
        Exit Sub
    End If
    
    Dim sheetNames As Collection
    Set sheetNames = New Collection ' �V�[�g�����i�[����R���N�V����
    
    ' �V�[�g�쐬����
    On Error Resume Next ' �G���[�n���h�����O�J�n
    For i = 2 To lrow
        sheetName = ws.Cells(i, columnNumber).value
        
        ' �w�肵���񂪋󔒂̏ꍇ�̓X�L�b�v
        If Trim(sheetName) = "" Then GoTo NextIteration
        
        ' �V�[�g���̗L�������`�F�b�N
        validSheetName = IsValidSheetName(sheetName)
        If Not validSheetName Then
            errorMsg = errorMsg & "�V�[�g���Ɏg�p�ł��Ȃ��������܂܂�Ă��܂�: " & sheetName & vbCrLf
            hasError = True ' �G���[�����������̂Ńt���O��ݒ�
            GoTo NextIteration
        End If
        
        ' �V�[�g�̑��݊m�F
        Set ws2 = Nothing
        On Error Resume Next
        Set ws2 = Worksheets(sheetName)
        On Error GoTo 0
        
        ' �V�[�g�����݂��Ȃ��ꍇ�ɐV�K�쐬
        If ws2 Is Nothing Then
            Set ws2 = Worksheets.Add
            On Error GoTo SheetNameError ' ���O�������ȏꍇ�̃G���[�n���h�����O
            ws2.Name = sheetName
            On Error GoTo 0 ' �G���[�n���h�����O�I��
            
            ' �w�b�_�[�s���R�s�[
            ws.Rows(1).Copy Destination:=ws2.Rows(1)
        End If
        
        ' �f�[�^�s���R�s�[�i�󔒂łȂ��ꍇ�j
        Dim lrow2 As Long
        lrow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row + 1
        ws.Rows(i).Copy Destination:=ws2.Rows(lrow2) ' �ŏI�s�̎��ɃR�s�[
        
NextIteration:
    Next i
    
    On Error GoTo 0  ' �G���[�n���h�����O�I��

    ' �G���[���������ꍇ�A�������s��Ȃ�
    If hasError Then
        MsgBox errorMsg
        Exit Sub
    End If
    
    MsgBox "�f�[�^�̒��o���������܂����B"
    Exit Sub

SheetNameError:
    MsgBox "�V�[�g�� '" & sheetName & "' �̐ݒ�Ɏ��s���܂����B���O���m�F���Ă��������B"
    Resume Next
    
    Unload Me

End Sub

Function IsValidSheetName(sheetName As String) As Boolean
    ' �V�[�g�����L�����ǂ������`�F�b�N
    Dim invalidChars As String
    invalidChars = "[]\/:*?""<>|"
    Dim i As Integer
    
    IsValidSheetName = True
    For i = 1 To Len(invalidChars)
        If InStr(sheetName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidSheetName = False
            Exit Function
        End If
    Next i
    
    If Len(sheetName) = 0 Or Len(sheetName) > 31 Then
        IsValidSheetName = False
    End If
End Function
