VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 抽出フォーム 
   Caption         =   "フォーム"
   ClientHeight    =   4310
   ClientLeft      =   40
   ClientTop       =   150
   ClientWidth     =   6450
   OleObjectBlob   =   "抽出フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "抽出フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim headerValue As String
    Dim lastCol As Long
    
    ' 最終列を取得
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    ' リストボックスにヘッダーの値を追加
    For i = 1 To lastCol ' 最終列までループ
        headerValue = ActiveSheet.Cells(1, i).value ' ヘッダーの値を取得
        If headerValue <> "" Then ' ヘッダーが空でない場合のみ追加
            ListBox1.AddItem headerValue ' ヘッダーの値を追加
            ListBox1.List(ListBox1.ListCount - 1, 1) = i ' 列番号を非表示の列に設定
        End If
    Next i
End Sub

Sub CommandButton1_Click()

    Dim ws As worksheet, ws2 As worksheet
    Set ws = ActiveSheet ' アクティブシートを設定
    
    Dim lrow As Long
    lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    Dim sheetName As String
    Dim validSheetName As Boolean
    Dim columnNumber As Long
    Dim errorMsg As String ' エラーメッセージ用の変数
    Dim hasError As Boolean ' エラーフラグ

    hasError = False ' エラーフラグを初期化
    
    ' リストボックスから列番号を取得
    If ListBox1.ListIndex <> -1 Then
        columnNumber = ListBox1.ListIndex + 1 ' 選択された列番号を取得（インデックスから1を足す）
    Else
        MsgBox "列番号を選択してください。"
        Exit Sub
    End If
    
    Dim sheetNames As Collection
    Set sheetNames = New Collection ' シート名を格納するコレクション
    
    ' シート作成処理
    On Error Resume Next ' エラーハンドリング開始
    For i = 2 To lrow
        sheetName = ws.Cells(i, columnNumber).value
        
        ' 指定した列が空白の場合はスキップ
        If Trim(sheetName) = "" Then GoTo NextIteration
        
        ' シート名の有効性をチェック
        validSheetName = IsValidSheetName(sheetName)
        If Not validSheetName Then
            errorMsg = errorMsg & "シート名に使用できない文字が含まれています: " & sheetName & vbCrLf
            hasError = True ' エラーが発生したのでフラグを設定
            GoTo NextIteration
        End If
        
        ' シートの存在確認
        Set ws2 = Nothing
        On Error Resume Next
        Set ws2 = Worksheets(sheetName)
        On Error GoTo 0
        
        ' シートが存在しない場合に新規作成
        If ws2 Is Nothing Then
            Set ws2 = Worksheets.Add
            On Error GoTo SheetNameError ' 名前が無効な場合のエラーハンドリング
            ws2.Name = sheetName
            On Error GoTo 0 ' エラーハンドリング終了
            
            ' ヘッダー行をコピー
            ws.Rows(1).Copy Destination:=ws2.Rows(1)
        End If
        
        ' データ行をコピー（空白でない場合）
        Dim lrow2 As Long
        lrow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row + 1
        ws.Rows(i).Copy Destination:=ws2.Rows(lrow2) ' 最終行の次にコピー
        
NextIteration:
    Next i
    
    On Error GoTo 0  ' エラーハンドリング終了

    ' エラーがあった場合、処理を行わない
    If hasError Then
        MsgBox errorMsg
        Exit Sub
    End If
    
    MsgBox "データの抽出が完了しました。"
    Exit Sub

SheetNameError:
    MsgBox "シート名 '" & sheetName & "' の設定に失敗しました。名前を確認してください。"
    Resume Next
    
    Unload Me

End Sub

Function IsValidSheetName(sheetName As String) As Boolean
    ' シート名が有効かどうかをチェック
    Dim invalidChars As String
    invalidChars = "[]\/:*?""<>|"
    Dim i As Integer
    
    IsValidSheetName = True
    For i = 1 To Len(invalidChars)
        If InStr(sheetName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidSheetName = False
            Exit Function
        End If
    Next i
    
    If Len(sheetName) = 0 Or Len(sheetName) > 31 Then
        IsValidSheetName = False
    End If
End Function
