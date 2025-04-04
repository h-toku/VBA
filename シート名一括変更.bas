Attribute VB_Name = "シート名一括変更"
Option Explicit

Sub CreateRenameSheet()
    Dim ws As worksheet
    Dim renameSheet As worksheet
    Dim i As Integer
    Dim wb As Workbook

    ' アクティブなワークブックを取得
    Set wb = Application.ActiveWorkbook

    ' 既に「シート名一括変更」シートが存在する場合、削除
    On Error Resume Next
    Set renameSheet = wb.Sheets("シート名一括変更")
    If Not renameSheet Is Nothing Then
        Application.DisplayAlerts = False
        renameSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' 新しいシート「シート名一括変更」を作成
    Set renameSheet = wb.Sheets.Add
    renameSheet.Name = "シート名一括変更"

    ' シートタブの色を黄色に設定
    renameSheet.Tab.Color = RGB(255, 255, 0)

    ' A1に「シート名」、B1に「新しいシート名」を入力し、見出しの色を黄色に設定
    With renameSheet
        .Range("A1").value = "シート名"
        .Range("B1").value = "新しいシート名"
        .Range("A1:B1").Interior.Color = RGB(255, 255, 0)
        .Range("G3").value = "「Ctrl+Shift+R」で実行"
        .Range("G3").Font.Bold = True  ' 太字に設定
        .Columns("A:G").AutoFit '列幅の自動調整

        ' 全てのシート名をA列に書き出し
        i = 2
        For Each ws In wb.Sheets
            If ws.Name <> "シート名一括変更" Then
                .Cells(i, 1).value = ws.Name
                i = i + 1
            End If
        Next ws
    End With
End Sub
