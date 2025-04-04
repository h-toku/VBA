VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ʃu�b�N�쐬 
   Caption         =   "�ʃu�b�N�쐬"
   ClientHeight    =   3910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5880
   OleObjectBlob   =   "�ʃu�b�N�쐬.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ʃu�b�N�쐬"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As worksheet

    ' ���X�g�{�b�N�X���N���A
    ListBox1.Clear

    ' �e�V�[�g�̖��O�����X�g�{�b�N�X�ɒǉ�
    For Each ws In ActiveWorkbook.Worksheets
        ListBox1.AddItem ws.Name
    Next ws
End Sub


Private Sub CommandButton1_Click()
    Dim folderPath As String
    Dim folderDialog As FileDialog

    ' �t�H���_�s�b�J�[�_�C�A���O���쐬
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)

    ' �_�C�A���O��\�����A���[�U�[���I�������ꍇ
    If folderDialog.Show = -1 Then
        ' �I�����ꂽ�t�H���_�̃p�X���擾
        folderPath = folderDialog.SelectedItems(1)
        ' �e�L�X�g�{�b�N�X�Ƀp�X��ǉ�
        TextBox1.Text = folderPath
    Else
        MsgBox "�t�H���_�͑I������܂���ł����B"
    End If

    ' �_�C�A���O�I�u�W�F�N�g�̉��
    Set folderDialog = Nothing
End Sub

Sub CommandButton2_Click()
    Dim ws As worksheet
    Dim newWorkbook As Workbook
    Dim sh As worksheet
    Dim Path As String
    Dim sheetName As String
    Dim fileExtension As String
    Dim i As Long
    Dim sourceWorkbook As Workbook
    
    Set sourceWorkbook = ActiveWorkbook
    
    ' TextBox1����p�X���擾�B�󔒂̏ꍇ�͎��s�u�b�N�̃p�X���g�p
    Path = Me.TextBox1.value
    If Path = "" Then
        Path = ActiveWorkbook.Path
    End If

    ' Path�̖����Ƀo�b�N�X���b�V�����Ȃ��ꍇ�ɒǉ�
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

    ' �I�����ꂽ�t�@�C���̊g���q���擾
    If OptionButton1.value Then
        fileExtension = "xlsx"
    ElseIf OptionButton2.value Then
        fileExtension = "csv"
    ElseIf OptionButton3.value Then
        fileExtension = "txt"
    Else
        MsgBox "�t�@�C���̊g���q���I������Ă��܂���B"
        Exit Sub
    End If

    ' ListBox1����I�����ꂽ�V�[�g�ɑ΂��ď��������s
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            ' ListBox1�̒l���V�[�g���Ƃ��Ď擾
            sheetName = ListBox1.List(i)
            
            ' �A�N�e�B�u�ȃu�b�N���̃V�[�g��ݒ�
            On Error Resume Next
            Set ws = sourceWorkbook.Sheets(sheetName)
            On Error GoTo 0

            ' �V�[�g��������Ȃ������ꍇ�̃G���[�n���h�����O
            If ws Is Nothing Then
                MsgBox "�V�[�g " & sheetName & " ��������܂���B"
                Exit Sub
            End If
        
            ' �V�����u�b�N���쐬
            Set newWorkbook = Workbooks.Add
        
        ' �V�[�g��V�����u�b�N�ɃR�s�[
        ws.Copy Before:=newWorkbook.Sheets(1)

        ' �f�t�H���g�ō쐬�����Sheet1�`3���폜
        Application.DisplayAlerts = False
        For Each sh In newWorkbook.Sheets
            If sh.Name <> sheetName Then
                sh.Delete
            End If
        Next sh
        Application.DisplayAlerts = True
        
        ' �V�����u�b�N���V�[�g���ƑI�������g���q�ŕۑ�
        newWorkbook.SaveAs Path & sheetName & "." & fileExtension, FileFormat:=GetFileFormat(fileExtension)
        newWorkbook.Close
        
        End If
    
    Next i
    
    MsgBox "�I�����܂����B"
    
    Unload Me
    
End Sub

Private Function GetFileFormat(ext As String) As XlFileFormat
    Select Case LCase(ext)
        Case "xlsx"
            GetFileFormat = xlOpenXMLWorkbook
        Case "csv"
            GetFileFormat = xlCSV
        Case "txt"
            GetFileFormat = xlText
        Case Else
            GetFileFormat = xlOpenXMLWorkbook
    End Select
End Function
