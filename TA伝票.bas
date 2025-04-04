Attribute VB_Name = "TA�`�["
Option Explicit

Sub TA()
    
    Dim ws As worksheet
    Dim cell As Range
    Dim i As Long
    Dim j As Long
    Dim lastrow As Long
    Dim value As String
    Dim parts As Variant
    Dim first As Long
    Dim last As Long
    Dim col As Long
    Dim rowData As Variant
    
    Set ws = ActiveSheet
    
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row '�ŏI�̒l
    
    For i = lastrow To 1 Step -1  '�t���ŏ���
        Set cell = ws.Cells(i, 1)
                ' �����Z���̐擪�̒l���擾�i�G���[�h�~�j
        If cell.MergeCells Then
            value = cell.MergeArea.Cells(1, 1).value
        Else
            value = cell.value
        End If
        
        If InStr(value, "-") > 0 Then  '�n�C�t�������鏈��
            parts = Split(value, "-")
            first = Val(Trim(parts(0)))
            last = Val(Trim(parts(1)))
            
            cell.value = first '�ŏ��̍s
            rowData = ws.Rows(i).value  '���̗�
            
            For j = last To first + 1 Step -1 '�s�̑}���ƒl�̓W�J
                cell.Offset(1, 0).EntireRow.Insert shift:=xlDown
                cell.Offset(1, 0).value = j
                
            For col = 2 To ws.UsedRange.Columns.Count
                cell.Offset(1, col - 1).value = rowData(1, col)
            Next col
            Next j
        End If
    Next i

End Sub
