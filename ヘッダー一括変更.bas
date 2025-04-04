Attribute VB_Name = "ヘッダー一括変更"
Option Explicit

Sub CreateRenameHeaders()
    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Integer, j As Integer
    Dim wb As Workbook
    Dim lastCol As Long

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

        ' 全てのシート名とヘッダー名を取得し、A列とB列に書き出し
        i = 2
        For Each ws In wb.Sheets
            If ws.Name <> "ヘッダー名一括変更" Then
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                For j = 1 To lastCol
                    .Cells(i, 1).value = ws.Name
                    .Cells(i, 2).value = ws.Cells(1, j).value
                    i = i + 1
                Next j
            End If
        Next ws
    End With
End Sub
