VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 全通り足し算 
   Caption         =   "全通り足し算"
   ClientHeight    =   1910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5050
   OleObjectBlob   =   "全通り足し算.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "全通り足し算"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As worksheet
    Dim lastCol As Long
    Dim i As Long
    
    ' アクティブシートの参照
    Set ws = ActiveSheet
    
    ' 最後の列を取得（ヘッダーのある行を想定）
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' ComboBox1とComboBox2にヘッダーをセット
    For i = 1 To lastCol
        ComboBox1.AddItem ws.Cells(1, i).value  ' 項目列の候補
        ComboBox2.AddItem ws.Cells(1, i).value  ' 値列の候補
    Next i
End Sub

Private Sub CommandButton1_Click()
    Dim ws As worksheet
    Dim result As worksheet
    Dim itemCol As Long, valueCol As Long
    Dim lastrow As Long
    Dim items As Variant
    Dim nums As Variant
    Dim i As Long, j As Long
    Dim rowCounter As Long
    
    ' アクティブシートの参照
    Set ws = ActiveSheet
    
    ' ComboBox1とComboBox2で選択された列を取得
    itemCol = Application.Match(ComboBox1.value, ws.Rows(1), 0)  ' 項目列
    valueCol = Application.Match(ComboBox2.value, ws.Rows(1), 0) ' 値列
    
    ' 最終行の取得
    lastrow = ws.Cells(ws.Rows.Count, itemCol).End(xlUp).Row
    
    ' データを配列に格納（2行目から最終行まで）
    items = ws.Range(ws.Cells(2, itemCol), ws.Cells(lastrow, itemCol)).value
    nums = ws.Range(ws.Cells(2, valueCol), ws.Cells(lastrow, valueCol)).value
    
    ' 結果を出力する新しいシートを作成
    Set result = Sheets.Add
    result.Name = "Pair Sum Combinations"
    
    ' ヘッダーを設定
    result.Cells(1, 1).value = ComboBox1.value & "（項目1）"
    result.Cells(1, 2).value = ComboBox1.value & "（項目2）"
    result.Cells(1, 3).value = ComboBox2.value & "（合計）"
    
    rowCounter = 2
    
    ' 2つの項目の組み合わせとそれに対応する値の合計を列挙（重複を省略し、セル自身の足し算も含む）
    For i = 1 To UBound(nums, 1)
        For j = i To UBound(nums, 1)
            ' 組み合わせの項目名をシートに出力
            result.Cells(rowCounter, 1).value = items(i, 1)  ' 項目1
            result.Cells(rowCounter, 2).value = items(j, 1)  ' 項目2
            ' それに対応するB列の値の合計をシートに出力
            result.Cells(rowCounter, 3).value = nums(i, 1) + nums(j, 1)  ' 値の合計
            rowCounter = rowCounter + 1
        Next j
    Next i
     
     Unload Me
     
End Sub

