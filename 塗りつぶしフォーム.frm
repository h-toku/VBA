VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 塗りつぶしフォーム 
   Caption         =   "塗りつぶし"
   ClientHeight    =   5360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9250.001
   OleObjectBlob   =   "塗りつぶしフォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "塗りつぶしフォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim savedSettings As Object ' 設定を保存するための辞書オブジェクト
Dim settingsFilePath As String
Dim previousComboBoxValues(1 To 7) As Variant ' ComboBoxの値を保存する配列
Dim previousTextBoxBackColors(1 To 7) As Long ' TextBoxの背景色を保存する配列

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler ' エラーハンドリングの開始

    settingsFilePath = ThisWorkbook.Path & "\settings.xlsx" ' パスを明示的に指定
    Dim i As Long
    Dim headerValue As String
    Dim lastCol As Long
    Dim j As Integer
    
    'HEADERBOXに1～10の値を追加
    For j = 1 To 10
        HEADERBOX.AddItem j
    Next j
    
    HEADERBOX.value = "1"
    
    ' 最終列を取得
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    ' Dictionaryの初期化
    Set savedSettings = CreateObject("Scripting.Dictionary")
    
    ' 設定をファイルから読み込み
    LoadSettingsFromFile
    
    ' コンボボックスA～Gの初期設定
    Dim comboBoxNames As Variant
    comboBoxNames = Array("ComboBoxA", "ComboBoxB", "ComboBoxC", "ComboBoxD", "ComboBoxE", "ComboBoxF", "ComboBoxG")
    
    For i = LBound(comboBoxNames) To UBound(comboBoxNames)
        With Me.Controls(comboBoxNames(i))
            .Clear ' 既存のアイテムをクリア
            .AddItem "一致" ' 一致を追加
            .AddItem "以上" ' 以上を追加
            .AddItem "以下" ' 以下を追加
            .AddItem "含む" ' 含むを追加
        End With
    Next i

    Exit Sub ' 正常終了

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description ' エラーメッセージを表示
End Sub
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
    ListBox1.Clear

    ' 指定した行 (targetRow) の1列目～lastColumnまでの値をListBox1に追加
    For i = 1 To lastColumn
        j = wsActive.Cells(targetRow, i).value
        
        ' セルの値が空でない場合にのみ追加
        If Not IsEmpty(j) Then
            ListBox1.AddItem CStr(j)  ' 値を文字列に変換して追加
        End If
    Next i

End Sub

Private Sub LoadSettingsFromFile()
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    Dim key As String
    Dim value As String
    
    fileNum = FreeFile
    Open settingsFilePath For Input As #fileNum
    
    ' ファイルの各行を読み込む
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "|")
        If UBound(parts) = 1 Then
            key = parts(0)
            value = parts(1)
            savedSettings(key) = value
            
            ' オプションボタンの状態を設定
            If key = "OptionButton1" Then
                OptionButton1.value = (value = "1")
            ElseIf key = "OptionButton2" Then
                OptionButton2.value = (value = "1")
            End If
        End If
    Loop
    
    Close #fileNum
End Sub

Private Sub UserForm_Activate()
    Dim wb As Workbook
    Dim ws As worksheet
    
    ' ListBox2をクリア
    ListBox2.Clear
    
    ' `setting.xlsx` を開く
    On Error Resume Next
    Set wb = Workbooks.Open(settingsFilePath)
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub
    
    ' 各シートの名前をListBox2に追加
    For Each ws In wb.Sheets
        ListBox2.AddItem ws.Name
    Next ws
    
    ' ワークブックを閉じる
    wb.Close False
End Sub

Private Sub ListBox1_Click()
    ' ListBox1で列が選択されたときにComboBox1?ComboBox7を更新
    LoadUniqueValuesToComboBoxes
End Sub
Private Sub CommandButton8_Click()
    ' CommandButton8がクリックされたときに塗りつぶしを実行
    Dim ws As worksheet
    Dim columnNumber As Long
    Dim selectedValue As String
    Dim lastrow As Long
    Dim cell As Range
    Dim i As Long
    Dim filterCondition As String

    Set ws = ActiveSheet

    ' ListBox1から列番号を取得
    If ListBox1.ListIndex <> -1 Then
        columnNumber = ListBox1.ListIndex + 1 ' 選択された列番号を取得
    Else
        MsgBox "列番号を選択してください。"
        Exit Sub
    End If

    ' 選択した列の最終行を取得
    lastrow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    ' 各ComboBoxから選択した値を取得し、対応するTextBoxの背景色で範囲を塗りつぶす
    For i = 1 To 7
        selectedValue = Me.Controls("ComboBox" & i).value ' 現在のComboBoxの選択値を取得

        ' ComboBoxが空でない場合のみ処理を実行
        If selectedValue <> "" Then
            ' フィルタ条件を取得（ComboBoxA, ComboBoxB, ...）
            filterCondition = Me.Controls("ComboBox" & Chr(64 + i)).value
            
            ' デバッグ用メッセージを表示
            Debug.Print "filterCondition: " & filterCondition ' ここでfilterConditionを表示

            ' フィルタ条件が有効か確認
            If filterCondition <> "一致" And filterCondition <> "以上" And filterCondition <> "以下" And filterCondition <> "含む" Then
                MsgBox "無効なフィルタ条件です。正しい条件を選択してください。"
                Exit Sub
            End If

            For Each cell In ws.Range(ws.Cells(2, columnNumber), ws.Cells(lastrow, columnNumber))
                Dim shouldFill As Boolean
                shouldFill = False ' 初期化

                ' フィルタ条件によるチェック
                If filterCondition = "一致" Then
                    If cell.value = selectedValue Then
                        shouldFill = True
                    End If
                ElseIf filterCondition = "以上" Then
                    If IsNumeric(cell.value) And IsNumeric(selectedValue) Then
                        If cell.value >= CDbl(selectedValue) Then
                            shouldFill = True
                        End If
                    End If
                ElseIf filterCondition = "以下" Then
                    If IsNumeric(cell.value) And IsNumeric(selectedValue) Then
                        If cell.value <= CDbl(selectedValue) Then
                            shouldFill = True
                        End If
                    End If
                ElseIf filterCondition = "含む" Then
                    If InStr(1, cell.value, selectedValue) > 0 Then
                        shouldFill = True
                    End If
                End If

                ' 塗りつぶし処理
                If shouldFill Then
                    If OptionButton1.value Then
                        ' OptionButton1が選択されている場合、該当セルを塗りつぶし
                        cell.Interior.Color = Me.Controls("TextBox" & i).BackColor ' セルを塗りつぶし
                    ElseIf OptionButton2.value Then
                        ' OptionButton2が選択されている場合、A列から右端までの行を塗りつぶし
                        ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, ws.Columns.Count).End(xlToLeft)).Interior.Color = Me.Controls("TextBox" & i).BackColor ' 行全体を塗りつぶし
                    End If
                End If
            Next cell
        End If
    Next i

    Unload Me
End Sub


' CommandButton9で塗りつぶしを無しにする
Private Sub CommandButton9_Click()
    Dim ws As worksheet
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' シート全体の塗りつぶしを無しにする
    ws.Cells.Interior.ColorIndex = xlNone
End Sub

Private Sub CommandButton1_Click()
    SetTextBoxColor 1
End Sub

Private Sub CommandButton2_Click()
    SetTextBoxColor 2
End Sub

Private Sub CommandButton3_Click()
    SetTextBoxColor 3
End Sub

Private Sub CommandButton4_Click()
    SetTextBoxColor 4
End Sub

Private Sub CommandButton5_Click()
    SetTextBoxColor 5
End Sub

Private Sub CommandButton6_Click()
    SetTextBoxColor 6
End Sub

Private Sub CommandButton7_Click()
    SetTextBoxColor 7
End Sub

Private Sub SetTextBoxColor(textBoxIndex As Integer)
    Dim intresult As Long

    ' カラーダイアログを表示して色を選択
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        ' 選択した色を指定されたTextBoxの背景色に設定
        intresult = ActiveWorkbook.colors(1) ' 選択された色を取得
        Me.Controls("TextBox" & textBoxIndex).BackColor = intresult ' 背景色を設定
        
        ' 背景色を配列に保存
        previousTextBoxBackColors(textBoxIndex) = intresult
    End If
End Sub

Private Sub CommandButton10_Click()
    Dim settingName As String
    Dim wb As Workbook
    Dim ws As worksheet
    Dim i As Long
    Dim result As VbMsgBoxResult
    
    ' 設定名を取得
    settingName = InputBox("保存する設定の名前を入力してください。")
    
    If settingName = "" Then
        MsgBox "設定名を入力してください。"
        Exit Sub
    End If
    
    ' `settings.xlsx` を開くか、新規作成
    On Error Resume Next
    Set wb = Workbooks.Open(settingsFilePath)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Add
        wb.SaveAs ThisWorkbook.Path & "\settings.xlsx"
    End If
    
    ' 新しいシートを作成
    On Error Resume Next
    Set ws = wb.Sheets(settingName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        ws.Name = settingName
    Else
        result = MsgBox("同じ名前の設定が既に存在します。" & vbCrLf & "上書き保存しますか？", vbYesNo + vbExclamation)
        If result = vbYes Then
        Application.DisplayAlerts = False ' 警告メッセージを非表示にする
        ws.Delete ' 既存のシートを削除
        Application.DisplayAlerts = True ' 警告メッセージを再表示
        Else
        Exit Sub
        End If
    End If
    
    ' ComboBox1～7の値をA列に保存
    For i = 1 To 7
        ws.Cells(i, 1).value = Me.Controls("ComboBox" & i).value ' ComboBox1～7をA列に
    Next i
    
    ' ComboBoxA～Gの値をB列に保存
    For i = 1 To 7
        ws.Cells(i, 2).value = Me.Controls("ComboBox" & Chr(64 + i)).value ' ComboBoxA～GをB列に
    Next i
    
    ' TextBox1～7の背景色をC列に保存
    For i = 1 To 7
        ws.Cells(i, 3).value = Me.Controls("TextBox" & i).BackColor ' TextBox1～7の背景色をC列に
    Next i
    
    ' オプションボタンの状態をD列に保存
    ws.Cells(1, 4).value = IIf(OptionButton1.value, "1", "0") ' オプションボタン1の状態
    ws.Cells(2, 4).value = IIf(OptionButton2.value, "1", "0") ' オプションボタン2の状態
    
    ' 保存して閉じる
    wb.Save
    wb.Close
    
    ' ListBox2に設定名を追加
    ListBox2.AddItem settingName
End Sub



Private Function GetComboBoxValues() As String
    Dim i As Integer
    Dim values As String
    For i = 1 To 7
        values = values & Me.Controls("ComboBox" & i).value & ","
    Next i
    GetComboBoxValues = Left(values, Len(values) - 1) ' 最後のカンマを削除
End Function

Private Function GetTextBoxColors() As String
    Dim i As Integer
    Dim colors As String
    For i = 1 To 7
        colors = colors & Me.Controls("TextBox" & i).BackColor & ","
    Next i
    GetTextBoxColors = Left(colors, Len(colors) - 1) ' 最後のカンマを削除
End Function

Private Sub SaveSettingsToFile()
    Dim fileNum As Integer
    Dim key As Variant
    
    fileNum = FreeFile
    Open settingsFilePath For Output As #fileNum
    
    ' 辞書から設定をファイルに書き込む
    For Each key In savedSettings.Keys
        Print #fileNum, key & "|" & savedSettings(key)
    Next key
    
    Close #fileNum
End Sub

Private Sub ListBox2_Click()
    Dim settingName As String
    Dim wb As Workbook
    Dim ws As worksheet
    Dim i As Long
    
    If ListBox2.ListIndex = -1 Then Exit Sub
    
    settingName = ListBox2.value
    
    ' `settings.xlsx` を開く
    Set wb = Workbooks.Open(settingsFilePath)
    
    ' 対応するシートを取得
    Set ws = wb.Sheets(settingName)
    
    ' ComboBoxとTextBoxに値を設定
    For i = 1 To 7
        Me.Controls("ComboBox" & i).value = ws.Cells(i, 1).value ' ComboBox1～7を設定
        Me.Controls("ComboBox" & Chr(64 + i)).value = ws.Cells(i, 2).value ' ComboBoxA～Gを設定
        Me.Controls("TextBox" & i).BackColor = ws.Cells(i, 3).value ' TextBox1～7の背景色を設定
    Next i
    
    ' オプションボタンの状態を設定
    OptionButton1.value = (ws.Cells(1, 4).value = "1") ' D列からオプションボタン1の状態を設定
    OptionButton2.value = (ws.Cells(2, 4).value = "1") ' D列からオプションボタン2の状態を設定
    
    ' ワークブックを閉じる
    wb.Close False
End Sub

Private Sub CommandButton11_Click()
    Dim key As Variant
    
    ' ListBox2から選択された設定名を取得
    If ListBox2.ListIndex <> -1 Then
        key = ListBox2.value
        
        ' 設定を削除
        If savedSettings.Exists(key) Then
            savedSettings.Remove key
            ListBox2.RemoveItem ListBox2.ListIndex ' ListBoxから削除
            SaveSettingsToFile ' 設定をファイルからも削除
        End If
    End If
End Sub

Private Sub UpdateListBox2()
    ' ListBox2を更新する関数
    Dim key As Variant
    
    ' ListBox2をクリア
    ListBox2.Clear
    
    ' 保存された設定名をリストボックスに追加
    For Each key In savedSettings.Keys
        ListBox2.AddItem key
    Next key
End Sub

Private Sub LoadUniqueValuesToComboBoxes()
    Dim ws As worksheet
    Dim columnNumber As Variant
    Dim lastrow As Long
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim i As Long
    Dim item As Variant ' item変数を宣言
    Dim headernum As Long

    Set ws = ActiveSheet
    
    headernum = Val(HEADERBOX.value)
    
        ' ListBox1の内容をデバッグ出力
    For i = 0 To ListBox1.ListCount - 1
        Debug.Print "Item " & i & ": " & ListBox1.List(i, 0) ' 1列目の値を表示
    Next i
    

    ' ListBox1から列番号を取得
    If ListBox1.ListIndex <> -1 Then  ' -1は何も選択されていない状態
        columnNumber = ListBox1.ListIndex + 1 ' 選択された列番号を取得
    Else
        MsgBox "列番号を選択してください。"
        Exit Sub
    End If

    Debug.Print "Column Number: " & columnNumber

    ' 最終行を取得
    lastrow = ws.Cells(ws.Rows.Count, columnNumber + 1).End(xlUp).Row

    ' ユニークな値を格納するコレクションを初期化
    Set uniqueValues = New Collection

    ' 指定した列のユニークな値を収集
    On Error Resume Next ' 重複エラーを無視
    For Each cell In ws.Range(ws.Cells(headernum + 1, columnNumber), ws.Cells(lastrow, columnNumber))
        If cell.value <> "" Then
            uniqueValues.Add cell.value, CStr(cell.value) ' ユニークな値を追加
        End If
    Next cell
    On Error GoTo 0 ' エラーハンドリングを解除

    ' ComboBox1～7にユニークな値を追加
    For i = 1 To 7
        With Me.Controls("ComboBox" & i)
            .Clear ' 既存のアイテムをクリア
            For Each item In uniqueValues
                .AddItem item ' ユニークな値を追加
            Next item
        End With
    Next i
End Sub



