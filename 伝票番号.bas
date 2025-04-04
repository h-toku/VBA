Attribute VB_Name = "伝票番号"
Option Explicit

Sub renban()

    Dim lastrow As Long
    Dim ws As worksheet
    Dim currentValue As String
    Dim currentAlpha As String
    Dim currentNum As Long
    Dim i As Long
    Dim newValue As String

    Set ws = ActiveSheet

    ' H列の最終行を取得
    lastrow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' A2の値（例: "A1"）を取得
    currentValue = ws.Range("A2").value
    currentAlpha = Left(currentValue, 1) ' アルファベット部分
    currentNum = CLng(Mid(currentValue, 2)) ' 数字部分
    
    ' A列の3行目から最終行まで、H列の値が変わるたびにA列に値をコピー
    For i = 3 To lastrow
        If ws.Cells(i, "H").value <> ws.Cells(i - 1, "H").value Then
            currentNum = currentNum + 1 ' H列の値が変わった場合、数字部分を+1
        End If
        
        ' 新しい値を作成（アルファベット部分はそのままで数字部分をインクリメント）
        newValue = currentAlpha & currentNum
        ws.Cells(i, "A").value = newValue
    Next i
    
End Sub
