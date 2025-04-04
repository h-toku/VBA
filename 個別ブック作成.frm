VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 個別ブック作成 
   Caption         =   "個別ブック作成"
   ClientHeight    =   3910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5880
   OleObjectBlob   =   "個別ブック作成.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "個別ブック作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As worksheet

    ' リストボックスをクリア
    ListBox1.Clear

    ' 各シートの名前をリストボックスに追加
    For Each ws In ActiveWorkbook.Worksheets
        ListBox1.AddItem ws.Name
    Next ws
End Sub


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
    Dim ws As worksheet
    Dim newWorkbook As Workbook
    Dim sh As worksheet
    Dim Path As String
    Dim sheetName As String
    Dim fileExtension As String
    Dim i As Long
    Dim sourceWorkbook As Workbook
    
    Set sourceWorkbook = ActiveWorkbook
    
    ' TextBox1からパスを取得。空白の場合は実行ブックのパスを使用
    Path = Me.TextBox1.value
    If Path = "" Then
        Path = ActiveWorkbook.Path
    End If

    ' Pathの末尾にバックスラッシュがない場合に追加
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

    ' 選択されたファイルの拡張子を取得
    If OptionButton1.value Then
        fileExtension = "xlsx"
    ElseIf OptionButton2.value Then
        fileExtension = "csv"
    ElseIf OptionButton3.value Then
        fileExtension = "txt"
    Else
        MsgBox "ファイルの拡張子が選択されていません。"
        Exit Sub
    End If

    ' ListBox1から選択されたシートに対して処理を実行
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            ' ListBox1の値をシート名として取得
            sheetName = ListBox1.List(i)
            
            ' アクティブなブック内のシートを設定
            On Error Resume Next
            Set ws = sourceWorkbook.Sheets(sheetName)
            On Error GoTo 0

            ' シートが見つからなかった場合のエラーハンドリング
            If ws Is Nothing Then
                MsgBox "シート " & sheetName & " が見つかりません。"
                Exit Sub
            End If
        
            ' 新しいブックを作成
            Set newWorkbook = Workbooks.Add
        
        ' シートを新しいブックにコピー
        ws.Copy Before:=newWorkbook.Sheets(1)

        ' デフォルトで作成されるSheet1～3を削除
        Application.DisplayAlerts = False
        For Each sh In newWorkbook.Sheets
            If sh.Name <> sheetName Then
                sh.Delete
            End If
        Next sh
        Application.DisplayAlerts = True
        
        ' 新しいブックをシート名と選択した拡張子で保存
        newWorkbook.SaveAs Path & sheetName & "." & fileExtension, FileFormat:=GetFileFormat(fileExtension)
        newWorkbook.Close
        
        End If
    
    Next i
    
    MsgBox "終了しました。"
    
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
