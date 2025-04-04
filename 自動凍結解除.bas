Attribute VB_Name = "������������"
Option Explicit

Sub ������������()

    On Error GoTo ErrorHandler
    
    If TypeOf Selection Is Range Then
        
        Dim r As Range
        Dim unmergedCount As Long
        unmergedCount = 0
        
        For Each r In Selection
            If r.MergeCells Then
                Dim mr As Range
                Set mr = r.MergeArea
                Dim firstCellValue As Variant
                firstCellValue = mr.Cells(1, 1).value
                r.UnMerge
                mr.value = firstCellValue
                unmergedCount = unmergedCount + mr.Cells.Count
            End If
        Next r

    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description

End Sub


