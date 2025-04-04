Attribute VB_Name = "©“®“€Œ‹‰ğœ"
Option Explicit

Sub ©“®“€Œ‹‰ğœ()

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
    MsgBox "ƒGƒ‰[‚ª”­¶‚µ‚Ü‚µ‚½: " & Err.Description

End Sub


