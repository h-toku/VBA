VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} シート一括統合 
   Caption         =   "シート一括統合"
   ClientHeight    =   4590
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6860
   OleObjectBlob   =   "シート一括統合.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "シート一括統合"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ユーザーフォームのロード時に、マクロを実行しているブック以外のすべてのシートをリストボックスに追加する
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim ws As worksheet
    Dim thisWb As Workbook
    
    ' マクロを実行しているアクティブワークブックを取得
    Set thisWb = ActiveWorkbook
    
    ' TextBox1にアクティブワークブックの名前を表示
    TextBox1.value = thisWb.Name
    
    ' ListBox1にアクティブワークブック以外のブックとシート名を追加
    For Each wb In Workbooks
        If wb.Name <> thisWb.Name Then ' マクロを実行しているブック以外
            For Each ws In wb.Sheets
                ' リストボックスにブック名とシート名を追加
                ListBox1.AddItem wb.Name & " - " & ws.Name
            Next ws
        End If
    Next wb
End Sub


' OKボタンがクリックされたときに、選択されたシートをマクロを実行しているブックに移動する
Private Sub CommandButton1_Click()
    Dim thisWb As Workbook
    Dim ws As worksheet
    Dim i As Integer
    Dim sheetInfo() As String
    
    ' マクロを実行しているブックを取得
    Set thisWb = ActiveWorkbook
    
    ' 選択されたシートをこのブックの末尾に移動
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            sheetInfo = Split(ListBox1.List(i), " - ")
            Workbooks(sheetInfo(0)).Sheets(sheetInfo(1)).Move after:=thisWb.Sheets(thisWb.Sheets.Count)
        End If
    Next i
    
    ' ユーザーフォームを閉じる
    Unload Me
End Sub


