Attribute VB_Name = "自動結合"
Option Explicit

Sub 自動結合()

    Dim colNo As Long
    colNo = ActiveCell.Column
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, colNo).End(xlUp).Row
    
    Dim sRow As Long
    sRow = 1
    Dim i As Long
    Application.DisplayAlerts = False
    For i = 1 To lastrow
        If Cells(i, colNo).value <> Cells(i + 1, colNo).value Then
            Range(Cells(sRow, colNo), Cells(i, colNo)).Merge
            sRow = i + 1
        End If
    Next i
    Application.DisplayAlerts = True

End Sub
