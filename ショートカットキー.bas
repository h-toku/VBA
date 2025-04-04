Attribute VB_Name = "ショートカットキー"
Option Explicit

Sub ChangeHeadersOrRenameSheets()
    Dim renameSheet As worksheet
    Dim sheetNameChangeSheet As worksheet
    
    ' アクティブなワークブックを使用
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' シートをチェック
    On Error Resume Next
    Set renameSheet = wb.Sheets("ヘッダー名一括変更")
    Set sheetNameChangeSheet = wb.Sheets("シート名一括変更")
    On Error GoTo 0
    
    ' ヘッダー名一括変更シートが存在する場合はChangeHeadersを実行
    If Not renameSheet Is Nothing Then
        ChangeHeaders wb ' アクティブなワークブックを渡す
    End If
    
    ' シート名一括変更シートが存在する場合はRenameSheetsを実行
    If Not sheetNameChangeSheet Is Nothing Then
        RenameSheets wb ' アクティブなワークブックを渡す
    End If
    
    ' 両方のシートが存在しない場合のエラーメッセージ
    If renameSheet Is Nothing And sheetNameChangeSheet Is Nothing Then
        MsgBox "「ヘッダー名一括変更」シートまたは「シート名一括変更」シートが存在しません。", vbExclamation
    End If
End Sub

Sub ChangeHeaders(wb As Workbook)
    Dim renameSheet As worksheet
    Dim ws As worksheet
    Dim oldHeader As String, newHeader As String
    Dim i As Integer

    ' 「ヘッダー名一括変更」シートを取得
    Set renameSheet = wb.Sheets("ヘッダー名一括変更")

    ' A列のシート名に対応するB列のヘッダーをC列の新しいヘッダー名に変更
    i = 2
    While renameSheet.Cells(i, 1).value <> ""
        Dim sheetName As String
        sheetName = Trim(renameSheet.Cells(i, 1).value) ' シート名をトリミング

        ' シートが存在するか確認
        On Error Resume Next
        Set ws = wb.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            MsgBox "シート '" & sheetName & "' は存在しません。", vbExclamation
            i = i + 1
            GoTo ContinueLoop ' 次の行へ進む
        End If
        
        ' B列の古いヘッダー名を取得
        oldHeader = renameSheet.Cells(i, 2).value
        newHeader = renameSheet.Cells(i, 3).value

        ' C列が空でない場合のみ処理を行う
        If newHeader <> "" Then
            ' ヘッダーの範囲を設定
            Dim headerRange As Range
            Set headerRange = ws.Rows(1).Find(oldHeader, LookIn:=xlValues, LookAt:=xlWhole)

            If Not headerRange Is Nothing Then
                ' 該当するヘッダーが見つかった場合、新しいヘッダーに変更
                headerRange.value = newHeader
            End If
        End If

        i = i + 1
ContinueLoop: ' ラベルを定義
    Wend
    
    ' 処理完了後、「ヘッダー名一括変更」シートを削除
    Application.DisplayAlerts = False
    renameSheet.Delete
    Application.DisplayAlerts = True

    MsgBox "ヘッダー名が変更されました。", vbInformation
End Sub

Sub RenameSheets(wb As Workbook)
    Dim renameSheet As worksheet
    Dim ws As worksheet
    Dim i As Integer
    Dim newName As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim errorMsg As String

    ' 「シート名一括変更」シートを取得
    Set renameSheet = wb.Sheets("シート名一括変更")

    If renameSheet Is Nothing Then
        MsgBox "シート名一括変更シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' A列のシート名に対してB列の新しいシート名に変更
    i = 2
    While renameSheet.Cells(i, 1).value <> ""
        newName = Trim(renameSheet.Cells(i, 2).value) ' シート名をトリミング

        ' B列が空でない場合のみ処理を行う
        If newName <> "" Then
            Dim sheetName As String
            sheetName = Trim(renameSheet.Cells(i, 1).value)

            ' シートが存在するか確認
            On Error Resume Next
            Set ws = wb.Sheets(sheetName)
            On Error GoTo 0
            
            If ws Is Nothing Then
                MsgBox "シート '" & sheetName & "' は存在しません。", vbExclamation
                i = i + 1
                GoTo ContinueLoopSheets ' 次の行へ進む
            End If
            
            ' 重複する場合、連番を付ける
            If dict.Exists(newName) Then
                dict(newName) = dict(newName) + 1
                newName = newName & "_" & dict(newName)
            Else
                dict.Add newName, 1
            End If

            On Error Resume Next
            ws.Name = newName
            If Err.Number <> 0 Then
                MsgBox "シート名 '" & newName & "' に変更できません。", vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        i = i + 1
ContinueLoopSheets: ' ラベルを定義
    Wend

    ' 処理完了後、「シート名一括変更」シートを削除
    Application.DisplayAlerts = False
    renameSheet.Delete
    Application.DisplayAlerts = True

    MsgBox "シート名が変更されました。", vbInformation
End Sub
