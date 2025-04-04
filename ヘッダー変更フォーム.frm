VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ヘッダー変更フォーム 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "ヘッダー変更フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ヘッダー変更フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As worksheet
    Dim wb As Workbook

    ' 操作対象のワークブックをアクティブなワークブックに設定
    Set wb = Application.ActiveWorkbook

    ' リストボックスにシート名を追加
    ListBox1.Clear ' 既存の項目をクリア
    For Each ws In wb.Sheets
        ' 「ヘッダー名一括変更」シートを除外
        If ws.Name <> "ヘッダー名一括変更" Then
            ListBox1.AddItem ws.Name
        End If
    Next ws
    
End Sub

Private Sub CommandButton1_Click()

    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Long, j As Long
    Dim wb As Workbook
    Dim lastCol As Long
    Dim selectedSheet As String
    Dim iSelected As Integer
    Dim dataRow As Long

    ' 操作対象のワークブックをアクティブなワークブックに設定
    Set wb = Application.ActiveWorkbook

    ' 既に「ヘッダー名一括変更」シートが存在する場合、削除
    On Error Resume Next
    Set renameSheet = wb.Sheets("ヘッダー名一括変更")
    If Not renameSheet Is Nothing Then
        Application.DisplayAlerts = False
        renameSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' 新しいシート「ヘッダー名一括変更」を作成
    Set renameSheet = wb.Sheets.Add
    renameSheet.Name = "ヘッダー名一括変更"

    ' シートタブの色を黄色に設定
    renameSheet.Tab.Color = RGB(255, 255, 0)

    ' A1に「シート名」、B1に「ヘッダー名」、C1に「新しいヘッダー名」を入力し、見出しの色を黄色に設定
    With renameSheet
        .Range("A1").value = "シート名"
        .Range("B1").value = "ヘッダー名"
        .Range("C1").value = "新しいヘッダー名"
        .Range("A1:C1").Interior.Color = RGB(255, 255, 0)
        .Range("G3").value = "「Ctrl+Shift+R」で実行"
        .Range("G3").Font.Bold = True  ' 太字に設定
        .Columns("A:G").AutoFit '列幅の自動調整
    End With

    ' リストボックスで選択されたシート名を取得し、ループで処理
    If ListBox1.ListCount > 0 Then
        For iSelected = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(iSelected) Then
                selectedSheet = ListBox1.List(iSelected) ' 選択されたシート名を取得

                ' 選択されたシートのヘッダー名を取得
                Set ws = wb.Sheets(selectedSheet)
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                dataRow = renameSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1 ' 書き込み行を設定

                ' シートのヘッダー名を取得して書き込み
                For j = 1 To lastCol
                    renameSheet.Cells(dataRow, 1).value = ws.Name
                    renameSheet.Cells(dataRow, 2).value = ws.Cells(1, j).value
                    dataRow = dataRow + 1
                Next j
            End If
        Next iSelected
    Else
        MsgBox "リストボックスでシートを選択してください。", vbExclamation
    End If
    
    Unload Me

End Sub
