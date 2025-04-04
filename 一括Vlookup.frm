VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 一括Vlookup 
   Caption         =   "一括Vlookup"
   ClientHeight    =   5830
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11500
   OleObjectBlob   =   "一括Vlookup.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "一括Vlookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HEADERBOX_Change()

    Dim targetRow As Long
    Dim lastColumn As Long
    Dim wsActive As worksheet
    Dim i As Long
    Dim j As Variant

    ' アクティブシートを取得
    Set wsActive = ActiveWorkbook.ActiveSheet
    
    ' HEADERBOXの値を取得
    targetRow = HEADERBOX.value
    
    ' HEADERBOXの値が1から10の範囲内か確認
    If targetRow < 1 Or targetRow > 10 Then
        MsgBox "HEADERBOXの値は1から10の範囲である必要があります。", vbExclamation
        Exit Sub
    End If

    ' 最後の列を取得
    lastColumn = wsActive.Cells(targetRow, wsActive.Columns.Count).End(xlToLeft).Column

    ' ComboBox1をクリア
    ComboBox1.Clear

    ' 指定した行 (targetRow) の1列目～lastColumnまでの値をComboBox1に追加
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).value
        
        ' セルの値が空でない場合にのみ追加
        If Not IsEmpty(j) Then
            ComboBox1.AddItem CStr(j)  ' 値を文字列に変換して追加
        End If
    Next i
    
End Sub

Private Sub DeleteButton_Click()
    On Error Resume Next ' エラーを無視して続行

    If ListBox1.ListIndex <> -1 Then ' 選択されているアイテムがある場合
        ListBox1.RemoveItem ListBox1.ListIndex
    End If

    On Error GoTo 0 ' エラーハンドリングを元に戻す（通常のエラーハンドリングに戻す）
End Sub

Private Sub UserForm_Initialize()
    Dim wsActive As worksheet
    Dim wsTree As worksheet
    Dim wsList As worksheet
    Dim lastColumn As Long
    Dim i As Long
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim ws As worksheet
    Dim targetRow As Long
    Dim j As Integer

    ' アクティブシートを取得
    Set wsActive = ActiveSheet
    
    TextBox1.value = ""

    'HEADERBOXに1～10の値を追加
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.value = "1"
    
    ' 検索条件.xlsx のパスを設定
    searchBookPath = ThisWorkbook.Path & "\検索条件.xlsx"
    
    ' 検索条件.xlsx が存在する場合、シート名をListBox2に表示
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath, ReadOnly:=True)
        
        ' シート名をListBox2に追加
        For Each wsList In searchBook.Worksheets
            Me.ListBox2.AddItem wsList.Name
        Next wsList
        
        ' ブックを閉じる
        searchBook.Close SaveChanges:=False
    End If
    
    ' ListBoxPにアクティブシート以外のシート名を追加
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> wsActive.Name Then
            ListBoxP.AddItem ws.Name
        End If
    Next ws
End Sub

Private Sub ComboBox1_Change()
    Dim selectedValue As String
    Dim columnNumber As Variant ' 変更: LongからVariantに
    Dim ws As worksheet
    
    ' アクティブシートを取得
    Set ws = ActiveSheet
    
    ' ComboBox1で選択された値を取得
    selectedValue = ComboBox1.value
    
    ' 選択された値に対応する列番号を取得
    On Error Resume Next ' エラーハンドリングを有効にする
    columnNumber = Application.Match(selectedValue, ws.Rows(1), 0)
    On Error GoTo 0 ' エラーハンドリングを無効にする
    
End Sub

Private Sub ListBoxP_Change()
    ' リストボックスCに選択されたシートの1行目の項目を表示
    Dim ws As worksheet
    Dim i As Long
    
    ListBoxC.Clear ' リストボックスCをクリア
    Set ws = ActiveWorkbook.Sheets(ListBoxP.value) ' 選択されたシートを取得
    
    ' 1行目の項目をリストボックスCに追加
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ListBoxC.AddItem ws.Cells(1, i).value
    Next i
End Sub

Private Sub ADDButton_Click()
    Dim selectedItem As String
    Dim selectedSheet As String
    Dim combinationExists As Boolean
    Dim i As Long
    Dim optionState As String ' オプションボタンの状態を保持
    
    If ListBoxC.ListIndex = -1 Then
        MsgBox "対象シートを選択してください。", vbExclamation
        Exit Sub
    End If
    
        ' リストボックスが選択されているか確認
    If ListBoxC.ListIndex = -1 Then
        MsgBox "対象列名を選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 選択されたアイテムを取得
    selectedItem = ListBoxC.value
    selectedSheet = ListBoxP.value
    combinationExists = False
    
    ' オプションボタンの状態を確認
    If OptionButton1.value Then
        optionState = "V(エラー)"
    ElseIf OptionButton2.value Then
        optionState = "V(0)"
    ElseIf OptionButton3.value Then
        optionState = "SUMIF"
    Else: MsgBox "検索方法を選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' リストボックス1の各アイテムをチェック
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.List(i, 0) = selectedSheet And _
            ListBox1.List(i, 1) = selectedItem And _
            ListBox1.List(i, 2) = optionState And _
            ListBox1.List(i, 3) = TextBox1.value Then
            combinationExists = True
            Exit For
        End If
    Next i
    
    ' 同じ組み合わせが存在しない場合のみ、リストボックス1に追加
    If Not combinationExists Then
        ListBox1.AddItem selectedSheet ' 1列目にリストボックスPの値を追加
        ListBox1.List(ListBox1.ListCount - 1, 1) = selectedItem ' 2列目にリストボックスCの選択項目を追加
        ListBox1.List(ListBox1.ListCount - 1, 2) = optionState ' 3列目にオプションボタンの状態を追加
        If Not TextBox1.value = "" Then
            ListBox1.List(ListBox1.ListCount - 1, 3) = TextBox1.value ' 4列目に項目名を追加
            Else: ListBox1.List(ListBox1.ListCount - 1, 3) = selectedSheet & selectedItem
        End If
    End If
    
    TextBox1.value = ""
    
End Sub

' UPButtonクリックイベント: 選択している項目を1つ上に移動
Private Sub UPButton_Click()
    Dim selectedIndex As Long
    Dim temp1 As String, temp2 As String, temp3 As String

    ' 選択されている項目のインデックスを取得
    selectedIndex = Me.ListBox1.ListIndex
    
    ' インデックスが1以上の場合（2つ目以降の項目が選択されている場合）のみ処理を実行
    If selectedIndex > 0 Then
        ' 現在選択されている項目の内容を一時的に保持
        temp1 = Me.ListBox1.List(selectedIndex, 0)
        temp2 = Me.ListBox1.List(selectedIndex, 1)
        temp3 = Me.ListBox1.List(selectedIndex, 2)
        temp4 = Me.ListBox1.List(selectedIndex, 3)
        
        ' 選択された項目とその1つ上の項目を入れ替え
        Me.ListBox1.List(selectedIndex, 0) = Me.ListBox1.List(selectedIndex - 1, 0)
        Me.ListBox1.List(selectedIndex, 1) = Me.ListBox1.List(selectedIndex - 1, 1)
        Me.ListBox1.List(selectedIndex, 2) = Me.ListBox1.List(selectedIndex - 1, 2)
        Me.ListBox1.List(selectedIndex, 3) = Me.ListBox1.List(selectedIndex - 1, 3)
        
        Me.ListBox1.List(selectedIndex - 1, 0) = temp1
        Me.ListBox1.List(selectedIndex - 1, 1) = temp2
        Me.ListBox1.List(selectedIndex - 1, 2) = temp3
        Me.ListBox1.List(selectedIndex - 1, 3) = temp4
        
        ' 項目を選択状態に戻す
        Me.ListBox1.ListIndex = selectedIndex - 1
    End If
End Sub

' DownButtonクリックイベント: 選択している項目を1つ下に移動
Private Sub DownButton_Click()
    Dim selectedIndex As Long
    Dim temp1 As String, temp2 As String, temp3 As String

    ' 選択されている項目のインデックスを取得
    selectedIndex = Me.ListBox1.ListIndex
    
    ' インデックスが最終行未満の場合（下に項目がある場合）のみ処理を実行
    If selectedIndex <> -1 And selectedIndex < Me.ListBox1.ListCount - 1 Then
        ' 現在選択されている項目の内容を一時的に保持
        temp1 = Me.ListBox1.List(selectedIndex, 0)
        temp2 = Me.ListBox1.List(selectedIndex, 1)
        temp3 = Me.ListBox1.List(selectedIndex, 2)
        temp4 = Me.ListBox1.List(selectedIndex, 3)
        
        ' 選択された項目とその1つ下の項目を入れ替え
        Me.ListBox1.List(selectedIndex, 0) = Me.ListBox1.List(selectedIndex + 1, 0)
        Me.ListBox1.List(selectedIndex, 1) = Me.ListBox1.List(selectedIndex + 1, 1)
        Me.ListBox1.List(selectedIndex, 2) = Me.ListBox1.List(selectedIndex + 1, 2)
        Me.ListBox1.List(selectedIndex, 3) = Me.ListBox1.List(selectedIndex + 1, 3)
        
        Me.ListBox1.List(selectedIndex + 1, 0) = temp1
        Me.ListBox1.List(selectedIndex + 1, 1) = temp2
        Me.ListBox1.List(selectedIndex + 1, 2) = temp3
        Me.ListBox1.List(selectedIndex + 1, 3) = temp4
        
        ' 項目を選択状態に戻す
        Me.ListBox1.ListIndex = selectedIndex + 1
    End If
End Sub

Private Function itemExists(value As String) As Boolean
    Dim i As Long
    itemExists = False
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.List(i) = value Then
            itemExists = True
            Exit Function
        End If
    Next i
End Function

' 指定したシート名が存在するか確認する関数
Function sheetExists(wb As Workbook, sheetName As String) As Boolean
    On Error Resume Next
    sheetExists = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub CommandButton3_Click()
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim ws As worksheet
    Dim settingName As String
    Dim lastrow As Long
    Dim i As Long
    Dim result As VbMsgBoxResult
    
    ' 設定名を入力させるインプットボックスを表示
    settingName = InputBox("設定名を入力してください:", "設定名の入力")
    
    ' 設定名が入力されていない場合、処理を終了
    If settingName = "" Then
        MsgBox "設定名が入力されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 検索条件.xlsx のパスを設定
    searchBookPath = ThisWorkbook.Path & "\検索条件.xlsx"
    
    ' 検索条件.xlsx が存在する場合、ブックを開く。存在しない場合、新規作成
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath)
    Else
        Set searchBook = Workbooks.Add
        searchBook.SaveAs searchBookPath
    End If
    
    ' 指定された設定名のシートが存在するか確認
    On Error Resume Next
    Set ws = searchBook.Worksheets(settingName)
    On Error GoTo 0
    
    ' 同じ名前のシートが既に存在する場合
    If Not ws Is Nothing Then
        result = MsgBox("同じ名前のシートが既に存在します。上書きしますか？", vbYesNo + vbExclamation)
        If result = vbYes Then
        Application.DisplayAlerts = False ' 警告メッセージを非表示にする
        ws.Delete ' 既存のシートを削除
        Application.DisplayAlerts = True ' 警告メッセージを再表示
        Else
        Exit Sub
        End If
    End If
    
    ' 新しいシートを追加し、設定名を設定
    Set ws = searchBook.Worksheets.Add
    ws.Name = settingName
    
    ' ComboBox1の値をA1セルに、ListBox1の値をB列とC列に書き込む
    ws.Cells(1, 1).value = HEADERBOX.value
    ws.Cells(2, 1).value = ComboBox1.value
    
    lastrow = ListBox1.ListCount
    For i = 0 To lastrow - 1
        ws.Cells(i + 2, 2).value = ListBox1.List(i, 0)
        ws.Cells(i + 2, 3).value = ListBox1.List(i, 1)
        ws.Cells(i + 2, 4).value = ListBox1.List(i, 2)
        ws.Cells(i + 2, 5).value = ListBox1.List(i, 3)
        
    Next i
    
    ' ブックを保存して閉じる
    searchBook.Close SaveChanges:=True
    
    MsgBox "設定が保存されました。", vbInformation
End Sub

Private Sub CommandButton4_Click()
    Dim searchBookPath As String
    Dim newWb As Workbook
    Dim wsName As String
    
    ' ListBox2で項目が選択されているか確認
    If Me.ListBox2.ListIndex = -1 Then
        MsgBox "削除するシートを選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 選択されたシート名を取得
    wsName = Me.ListBox2.value
    
    ' 検索条件.xlsx のパスを作成
    searchBookPath = ThisWorkbook.Path & "\検索条件.xlsx"
    
    ' 検索条件.xlsx を開く
    Set newWb = Workbooks.Open(searchBookPath)
    
    ' ListBox2に残っている項目数を確認
    If Me.ListBox2.ListCount = 1 Then
        ' 項目が1つだけの場合、ブックを削除
        newWb.Close SaveChanges:=False
        Kill searchBookPath ' ブックを削除
        MsgBox "ブック '" & searchBookPath & "' が削除されました。", vbInformation
    Else
        ' 項目が複数の場合、選択されたシートのみ削除
        Application.DisplayAlerts = False ' 削除確認ダイアログを非表示にする
        On Error Resume Next
        newWb.Worksheets(wsName).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        
        ' ブックを保存して閉じる
        newWb.Save
        newWb.Close SaveChanges:=True
        
        ' ListBox2を更新
        Me.ListBox2.RemoveItem Me.ListBox2.ListIndex
        
        MsgBox "シート '" & wsName & "' が削除されました。", vbInformation
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim searchBookPath As String
    Dim searchBook As Workbook
    Dim selectedSheetName As String
    Dim ws As worksheet
    Dim i As Long
    
    ' 検索条件.xlsx のパスを設定
    searchBookPath = ThisWorkbook.Path & "\検索条件.xlsx"
    
    ' 選択されたシート名を取得
    selectedSheetName = Me.ListBox2.value
    
    ' 検索条件.xlsx を開いて選択されたシートを取得
    If Dir(searchBookPath) <> "" Then
        Set searchBook = Workbooks.Open(searchBookPath, ReadOnly:=True)
        Set ws = searchBook.Worksheets(selectedSheetName)
        
        ' HEADERBOXにA1セルの値を反映
        Me.ComboBox1.value = ws.Cells(1, 1).value
        ' ComboBox1にA2セルの値を反映
        Me.ComboBox1.value = ws.Cells(2, 1).value
        
        ' ListBox1にシートのB列とC列の値を反映
        Me.ListBox1.Clear
        For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            Me.ListBox1.AddItem
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = ws.Cells(i, 2).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 3).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 4).value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = ws.Cells(i, 5).value
        Next i
        
        ' ブックを閉じる
        searchBook.Close SaveChanges:=False
    End If
End Sub

Private Sub UserForm_Terminate()
    Dim searchBookPath As String
    Dim newWb As Workbook
    
    ' 検索条件.xlsx のパスを作成
    searchBookPath = ThisWorkbook.Path & "\検索条件.xlsx"
    
    ' 検索条件.xlsx が既に開かれているか確認
    On Error Resume Next
    Set newWb = Workbooks("検索条件.xlsx")
    On Error GoTo 0
    
    ' ブックが開かれている場合のみ閉じる
    If Not newWb Is Nothing Then
        On Error Resume Next ' すでにブックがない状態は無視
        newWb.Close SaveChanges:=True ' 必要に応じて SaveChanges を False に変更
        On Error GoTo 0
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim i As Long
    Dim listItem As String
    Dim sheetName As String
    Dim optionValue As String
    Dim header As String
    Dim resultColumn As Long
    
    ' ListBox1の全項目に対して処理を行う
    For i = 0 To ListBox1.ListCount - 1
        ' 1列目がシート名、2列目が項目、3列目がオプションボタンの状態と仮定
        sheetName = ListBox1.List(i, 0) ' 1列目（シート名）
        listItem = ListBox1.List(i, 1) ' 2列目（リスト項目）
        optionValue = ListBox1.List(i, 2) ' 3列目（オプションボタンの状態）
        header = ListBox1.List(i, 3) ' 4列目（項目名）
        
        ' 3列目の値に応じて処理を振り分ける
        If optionValue = "V(エラー)" Then
            Call ProcessOption1(sheetName, listItem, header)
        ElseIf optionValue = "V(0)" Then
            Call ProcessOption2(sheetName, listItem, header)
        Else
            Call ProcessOption3(sheetName, listItem, header)
        End If
        
    Next i
    
    If MsgBox("処理が完了しました。" & vbCrLf & "条件を保存しますか？", vbYesNo + vbQuestion) = vbYes Then
        Call CommandButton3_Click
        End If
        
        Unload Me
        
End Sub

Private Sub ProcessOption1(sheetName As String, listItem As String, header As String)
    ' VLOOKUP処理
    Dim ws As worksheet
    Dim lastrow As Long
    Dim resultColumn As Long
    Dim lookupValue As String
    Dim lookupColumn As Long
    Dim targetWorkbook As Workbook
    Dim j As Long
    Dim selectedValue As String
    Dim searchRange As Range
    Dim lookupResult As Variant

    Set targetWorkbook = Application.ActiveWorkbook
    lookupValue = ComboBox1.value

    ' アクティブシートの検索列番号を取得
    lookupColumn = Application.Match(lookupValue, ActiveSheet.Rows(1), 0)
    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, lookupColumn).End(xlUp).Row
    resultColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1
    
    ' 対象のシートを取得
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

        If Not ws Is Nothing Then
            selectedValue = listItem ' ListBox1の子ノードのヘッダー

            ' 親ノード（コンボボックスの値）に対応する列番号を取得
            lookupColumn = Application.Match(lookupValue, ws.Rows(1), 0)
            Dim childColumn As Long
            childColumn = Application.Match(selectedValue, ws.Rows(1), 0)

            If Not IsError(lookupColumn) And Not IsError(childColumn) Then
                ' 検索範囲を設定
            Set searchRange = ws.Range(ws.Cells(2, lookupColumn), ws.Cells(ws.Rows.Count, lookupColumn).End(xlUp)).Resize(, childColumn - lookupColumn + 1)

                ' アクティブシートの各行をループし、VLOOKUPを実行
                For j = 2 To lastrow ' 2行目から最終行まで
                    Dim skuValue As Variant
                    skuValue = ActiveSheet.Cells(j, Application.Match(lookupValue, ActiveSheet.Rows(1), 0)).value ' 検索値を取得
                    
                        If Not IsEmpty(skuValue) And Len(Trim(skuValue)) > 0 Then
                        'セルが空白でないorスペースやタブの空白文字のみでないの場合
                
                    ' VLOOKUPの実行
                    lookupResult = Application.Vlookup(skuValue, searchRange, childColumn - lookupColumn + 1, False)

                    ' 結果をアクティブシートに追加
                    If Not IsError(lookupResult) Then
                        ActiveSheet.Cells(j, resultColumn).value = lookupResult
                    Else
                    ActiveSheet.Cells(j, resultColumn).value = CVErr(xlErrNA)
                        End If
                    End If
            Next j
            
            ' ヘッダーを追加
            ActiveSheet.Cells(1, resultColumn).value = header
                resultColumn = resultColumn + 1 ' 次の列に移動

            Else
                MsgBox "列が見つかりません: " & lookupValue & " または " & selectedValue
            End If
        Else
            MsgBox "シートが見つかりません"
        End If

End Sub

Private Sub ProcessOption2(sheetName As String, listItem As String, header As String)
    ' VLOOKUPエラーを0に置き換える例
    Dim ws As worksheet
    Dim lastrow As Long
    Dim resultColumn As Long
    Dim lookupValue As String
    Dim lookupColumn As Long
    Dim targetWorkbook As Workbook
    Dim i As Long, j As Long
    Dim selectedValue As String
    Dim searchRange As Range
    Dim lookupResult As Variant

    Set targetWorkbook = Application.ActiveWorkbook
    lookupValue = ComboBox1.value

    ' アクティブシートの検索列番号を取得
    lookupColumn = Application.Match(lookupValue, ActiveSheet.Rows(1), 0)
    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, lookupColumn).End(xlUp).Row
    resultColumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1

    ' 対象のシートを取得
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

        If Not ws Is Nothing Then
            selectedValue = listItem ' ListBox1の子ノードのヘッダー

            ' 親ノード（コンボボックスの値）に対応する列番号を取得
            lookupColumn = Application.Match(lookupValue, ws.Rows(1), 0)
            Dim childColumn As Long
            childColumn = Application.Match(selectedValue, ws.Rows(1), 0)

            If Not IsError(lookupColumn) And Not IsError(childColumn) Then
                ' 検索範囲を設定
                Set searchRange = ws.Range(ws.Cells(2, lookupColumn), ws.Cells(ws.Rows.Count, lookupColumn).End(xlUp)).Resize(, childColumn - lookupColumn + 1)

                ' アクティブシートの各行をループし、VLOOKUPを実行
                For j = 2 To lastrow ' 2行目から最終行まで
                    Dim skuValue As Variant
                    skuValue = ActiveSheet.Cells(j, Application.Match(lookupValue, ActiveSheet.Rows(1), 0)).value ' 検索値を取得
                    
                        If Not IsEmpty(skuValue) And Len(Trim(skuValue)) > 0 Then
                        'セルが空白でないorスペースやタブの空白文字のみでないの場合
                
                    ' VLOOKUPの実行
                    lookupResult = Application.Vlookup(skuValue, searchRange, childColumn - lookupColumn + 1, False)

                    ' 結果をアクティブシートに追加
                    If Not IsError(lookupResult) Then
                        ActiveSheet.Cells(j, resultColumn).value = lookupResult
                    Else
                            ActiveSheet.Cells(j, resultColumn).value = 0
                        End If
                End If
            Next j

                ' ヘッダーを追加
                ActiveSheet.Cells(1, resultColumn).value = header
                resultColumn = resultColumn + 1 ' 次の列に移動
            Else
                MsgBox "列が見つかりません: " & lookupValue & " または " & selectedValue
            End If
        Else
            MsgBox "シートが見つかりません"
        End If
End Sub

Private Sub ProcessOption3(sheetName As String, listItem As String, header As String)
    ' SUMIF処理を行う例
    Dim wsActive As worksheet
    Dim wsTarget As worksheet
    Dim comboValue As String
    Dim foundCellList As Range
    Dim criteriaRange As Range
    Dim sumRange As Range
    Dim rng As Range
    Dim lastrow As Long
    Dim lastCol As Long
    Dim criteriaCol As Long
    Dim sumResult As Double
    Dim j As Long
    
    ' アクティブシートを設定
    Set wsActive = ActiveSheet
    
    ' ComboBox1の値を取得
    comboValue = Me.ComboBox1.value
    
    ' ComboBox1の値と一致する列を検索
    criteriaCol = Application.Match(comboValue, wsActive.Rows(1), 0)
    
    If IsError(criteriaCol) Then
        MsgBox "検索条件の列が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' アクティブシートの最終行を取得
    lastrow = wsActive.Cells(wsActive.Rows.Count, criteriaCol).End(xlUp).Row
    
    ' 検索条件範囲を設定（アクティブシートの2行目から最終行まで）
    Set criteriaRange = wsActive.Range(wsActive.Cells(2, criteriaCol), wsActive.Cells(lastrow, criteriaCol))
    
    ' アクティブシートの最終列を取得
    lastCol = wsActive.Cells(2, wsActive.Columns.Count).End(xlToLeft).Column + 1
    
    ' 対象のシートを設定
    Set wsTarget = Worksheets(sheetName) ' sheetName が正しく提供されていると仮定
    
    ' ListBoxの値と一致するセルを検索
    Dim selectedCellValue As String
    selectedCellValue = listItem ' listItem が ListBox からの選択された値を含むと仮定
    Set foundCellList = wsTarget.Cells.Find(What:=selectedCellValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCellList Is Nothing Then
        ' ComboBox1の値と一致する列を対象シートで検索
        Dim targetCriteriaCol As Long
        targetCriteriaCol = Application.Match(comboValue, wsTarget.Rows(1), 0)
        
        If IsError(targetCriteriaCol) Then
            MsgBox "対象シートで検索条件の列が見つかりません。", vbExclamation
            Exit Sub
        End If
        
        ' 検索範囲（rng）と合計範囲（sumRange）を設定
        Set rng = wsTarget.Range(wsTarget.Cells(2, targetCriteriaCol), wsTarget.Cells(wsTarget.Rows.Count, targetCriteriaCol).End(xlUp))
        Set sumRange = wsTarget.Range(wsTarget.Cells(2, foundCellList.Column), wsTarget.Cells(wsTarget.Rows.Count, foundCellList.Column).End(xlUp))
        
        ' 結果をアクティブシートの最右列に反映
        wsActive.Cells(1, lastCol).value = header
        
        ' criteriaのセルに対応する全ての結果を計算してアクティブシートに挿入
        For j = 1 To criteriaRange.Rows.Count
            ' 空白セルをスキップ
            If Len(Trim(criteriaRange.Cells(j, 1).value)) > 0 Then
                On Error Resume Next
                ' SUMIFを実行し、結果を計算
                sumResult = Application.WorksheetFunction.SumIf(rng, criteriaRange.Cells(j, 1).value, sumRange)
                
                wsActive.Cells(j + 1, lastCol).value = sumResult
                On Error GoTo 0
            End If
        Next j
    Else
        MsgBox "ListBox1の値と一致するセルが見つかりません。", vbExclamation
    End If
End Sub



