VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} フォルダ一括作成 
   Caption         =   "フォルダ一括作成"
   ClientHeight    =   2200
   ClientLeft      =   -60
   ClientTop       =   -300
   ClientWidth     =   5820
   OleObjectBlob   =   "フォルダ一括作成.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "フォルダ一括作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Dim folderPath As String
    Dim folderDialog As FileDialog

    ' フォルダピッカーダイアログを作成
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)

    ' ダイアログを表示し、ユーザーが選択した場合
    If folderDialog.Show = -1 Then
        ' 選択されたフォルダのパスを取得
        folderPath = folderDialog.SelectedItems(1)
        ' テキストボックスにパスを追加
        TextBox1.Text = folderPath
    Else
        MsgBox "フォルダは選択されませんでした。"
    End If

    ' ダイアログオブジェクトの解放
    Set folderDialog = Nothing
End Sub


Sub CommandButton2_Click()
    Dim Path As String ' 作成予定フォルダの上位パス
    
    ' TextBox1からパスを取得。空白の場合は実行ブックのパスを使用
    Path = Me.TextBox1.value
    If Path = "" Then
        Path = ActiveWorkbook.Path
    End If
    
    ' Pathの末尾にバックスラッシュがない場合に追加
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    Dim folderNames As Collection
    Set folderNames = New Collection
    
    Dim cell As Range ' 選択セルをループするための変数
    
    ' 選択したセルからフォルダ名を収集（重複を削除）
    On Error Resume Next ' 重複エラーを無視
    For Each cell In Selection
        If cell.value <> "" Then
            folderNames.Add cell.value, CStr(cell.value) ' フォルダ名をコレクションに追加
        End If
    Next cell
    On Error GoTo 0 ' エラーハンドリング終了

    Dim folderName As Variant
    On Error Resume Next ' エラーハンドリング開始
    For Each folderName In folderNames
        Dim NewDirPath As String ' 作成予定のフォルダパス
        NewDirPath = Path & folderName
        
        ' 作成予定フォルダと同名のフォルダの存在有無を確認
        If Dir(NewDirPath, vbDirectory) = "" Then
            MkDir NewDirPath
            If Err.Number <> 0 Then
                MsgBox "エラー: フォルダ " & NewDirPath & " の作成に失敗しました。", vbExclamation
                Err.Clear
            End If
        End If
    Next folderName
    On Error GoTo 0 ' エラーハンドリング終了
    
    MsgBox "終了しました。"

    Unload Me
    
End Sub
