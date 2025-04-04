Attribute VB_Name = "TA伝票"
Option Explicit

Sub TA()
    
    Dim ws As worksheet
    Dim cell As Range
    Dim i As Long
    Dim j As Long
    Dim lastrow As Long
    Dim value As String
    Dim parts As Variant
    Dim first As Long
    Dim last As Long
    Dim col As Long
    Dim rowData As Variant
    
    Set ws = ActiveSheet
    
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row '最終の値
    
    For i = lastrow To 1 Step -1  '逆順で処理
        Set cell = ws.Cells(i, 1)
                ' 結合セルの先頭の値を取得（エラー防止）
        If cell.MergeCells Then
            value = cell.MergeArea.Cells(1, 1).value
        Else
            value = cell.value
        End If
        
        If InStr(value, "-") > 0 Then  'ハイフンがある処理
            parts = Split(value, "-")
            first = Val(Trim(parts(0)))
            last = Val(Trim(parts(1)))
            
            cell.value = first '最初の行
            rowData = ws.Rows(i).value  '他の列
            
            For j = last To first + 1 Step -1 '行の挿入と値の展開
                cell.Offset(1, 0).EntireRow.Insert shift:=xlDown
                cell.Offset(1, 0).value = j
                
            For col = 2 To ws.UsedRange.Columns.Count
                cell.Offset(1, col - 1).value = rowData(1, col)
            Next col
            Next j
        End If
    Next i

End Sub
