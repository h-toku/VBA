VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tofs 
   Caption         =   "THOPSデータ抽出フォーム"
   ClientHeight    =   4770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6340
   OleObjectBlob   =   "Tofs.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Tofs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As Object ' 接続オブジェクトをモジュールレベルで宣言

Private Sub UserForm_Initialize()
    Dim wsActive As Worksheet
    Dim j As Integer
    Dim rs As Object
    Dim sql As String
    
    ' アクティブシートを取得
    Set wsActive = ActiveSheet

    ' HEADERBOXに1～10の値を追加
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.Value = "1"
    
    On Error GoTo ErrorHandler
    
    ' MySQL接続情報 (DSNレス接続)
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Driver={MySQL ODBC 9.1 Unicode Driver};" & _
                            "Server=localhost;" & _
                            "port=33061;" & _
                            "Database=tofs;" & _
                            "User=root;" & _
                            "Password=password;" & _
                            "Option=3;"
    conn.Open
    
    ' SQLクエリ: テーブルのフィールド名を取得
    sql = "SHOW FIELDS FROM items"
    Set rs = conn.Execute(sql)
    
    ' ListBox1にフィールド名 (Field列) を追加
    ListBox1.Clear
    ListBox1.ColumnCount = 1 ' 1列のみ表示
    
    Do Until rs.EOF
        ListBox1.AddItem rs.Fields("Field").Value
        rs.MoveNext
    Loop
    
    ' 後処理
    rs.Close
    Set rs = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

Private Sub ComboBox2_Change()
    Dim selectedValue As String
    Dim columnNumber As Variant
    Dim ws As Worksheet
    
    ' アクティブシートを取得
    Set ws = ActiveSheet
    
    ' ComboBox2で選択された値を取得
    selectedValue = ComboBox2.Value
    
    ' 選択された値に対応する列番号を取得
    On Error Resume Next ' エラーハンドリングを有効にする
    columnNumber = Application.Match(selectedValue, ws.Rows(1), 0)
    On Error GoTo 0 ' エラーハンドリングを無効にする
End Sub

Private Sub HEADERBOX_Change()
    Dim targetRow As Long
    Dim lastColumn As Long
    Dim wsActive As Worksheet
    Dim i As Long
    Dim j As Variant

    ' アクティブシートを取得
    Set wsActive = ActiveWorkbook.ActiveSheet
    
    ' HEADERBOXの値を取得
    targetRow = HEADERBOX.Value
    
    ' HEADERBOXの値が1から10の範囲内か確認
    If targetRow < 1 Or targetRow > 10 Then
        MsgBox "HEADERBOXの値は1から10の範囲である必要があります。", vbExclamation
        Exit Sub
    End If

    ' 最後の列を取得
    lastColumn = wsActive.Cells(targetRow, wsActive.Columns.Count).End(xlToLeft).Column

    ' ComboBox2をクリア
    ComboBox2.Clear

    ' 指定した行 (targetRow) の1列目～lastColumnまでの値をComboBox2に追加
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).Value
        
        ' セルの値が空でない場合にのみ追加
        If Not IsEmpty(j) Then
            ComboBox2.AddItem CStr(j)  ' 値を文字列に変換して追加
        End If
    Next i
    
End Sub

Private Sub btnFetchData_Click()
    Dim rs As Object
    Dim sql As String
    Dim selectedFields As String
    Dim skuValue As String
    Dim lastRow As Long
    Dim i As Long, j As Long
    
    On Error GoTo ErrorHandler
    
    ' 接続を確認し、必要なら再接続
    If conn Is Nothing Or conn.State = 0 Then
        MsgBox "接続が確立されていません。再接続します。", vbExclamation
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionString = "Driver={MySQL ODBC 9.1 Unicode Driver};" & _
                                "Server=localhost;" & _
                                "port=33061;" & _
                                "Database=tofs;" & _
                                "User=root;" & _
                                "Password=password;" & _
                                "Option=3;"
        conn.Open
    End If
    
    ' ListBoxから選択されたフィールドを取得
    If ListBox1.ListIndex = -1 Then
        MsgBox "少なくとも1つのフィールドを選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 複数選択されたフィールドをカンマで連結
    selectedFields = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            If selectedFields <> "" Then
                selectedFields = selectedFields & ", "
            End If
            selectedFields = selectedFields & ListBox1.List(i, 0)
        End If
    Next i
    
    If selectedFields = "" Then
        MsgBox "フィールドが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' アクティブシートを設定
    Dim ws As Worksheet
    Set ws = ActiveSheet ' アクティブシートを使用
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 1行目にフィールド名を挿入
    Dim fieldArray() As String
    fieldArray = Split(selectedFields, ", ")
    
    For i = 0 To UBound(fieldArray)
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = fieldArray(i)
    Next i
    
    ' 2行目から最終行までデータを取得
    For i = 2 To lastRow
        skuValue = ws.Cells(i, ComboBox2.ListIndex + 1).Value ' ComboBox2で選択された列
        
        If skuValue <> "" Then
            sql = "SELECT " & selectedFields & " FROM items WHERE sku = '" & skuValue & "'"
            Set rs = conn.Execute(sql)
            
            If Not rs.EOF Then
                For j = 0 To UBound(fieldArray)
                    ws.Cells(i, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = rs.Fields(j).Value
                Next j
            Else
                ws.Cells(i, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = ""
            End If
        End If
    Next i
    
    MsgBox "データを正常に取得しました！", vbInformation

    ' クリーンアップ
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub


