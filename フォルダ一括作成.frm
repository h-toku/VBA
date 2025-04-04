VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �t�H���_�ꊇ�쐬 
   Caption         =   "�t�H���_�ꊇ�쐬"
   ClientHeight    =   2200
   ClientLeft      =   -60
   ClientTop       =   -300
   ClientWidth     =   5820
   OleObjectBlob   =   "�t�H���_�ꊇ�쐬.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�t�H���_�ꊇ�쐬"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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
    Dim Path As String ' �쐬�\��t�H���_�̏�ʃp�X
    
    ' TextBox1����p�X���擾�B�󔒂̏ꍇ�͎��s�u�b�N�̃p�X���g�p
    Path = Me.TextBox1.value
    If Path = "" Then
        Path = ActiveWorkbook.Path
    End If
    
    ' Path�̖����Ƀo�b�N�X���b�V�����Ȃ��ꍇ�ɒǉ�
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    Dim folderNames As Collection
    Set folderNames = New Collection
    
    Dim cell As Range ' �I���Z�������[�v���邽�߂̕ϐ�
    
    ' �I�������Z������t�H���_�������W�i�d�����폜�j
    On Error Resume Next ' �d���G���[�𖳎�
    For Each cell In Selection
        If cell.value <> "" Then
            folderNames.Add cell.value, CStr(cell.value) ' �t�H���_�����R���N�V�����ɒǉ�
        End If
    Next cell
    On Error GoTo 0 ' �G���[�n���h�����O�I��

    Dim folderName As Variant
    On Error Resume Next ' �G���[�n���h�����O�J�n
    For Each folderName In folderNames
        Dim NewDirPath As String ' �쐬�\��̃t�H���_�p�X
        NewDirPath = Path & folderName
        
        ' �쐬�\��t�H���_�Ɠ����̃t�H���_�̑��ݗL�����m�F
        If Dir(NewDirPath, vbDirectory) = "" Then
            MkDir NewDirPath
            If Err.Number <> 0 Then
                MsgBox "�G���[: �t�H���_ " & NewDirPath & " �̍쐬�Ɏ��s���܂����B", vbExclamation
                Err.Clear
            End If
        End If
    Next folderName
    On Error GoTo 0 ' �G���[�n���h�����O�I��
    
    MsgBox "�I�����܂����B"

    Unload Me
    
End Sub
